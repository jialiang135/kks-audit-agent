# -*- coding: utf-8 -*-
"""Optional second-stage AI review for uncertain KKS audit findings.

The module deliberately uses only the Python standard library so the audit
service remains runnable without an SDK.  It speaks the common
OpenAI-compatible ``/chat/completions`` protocol and never writes the API key
to a report or to the source workbook.
"""
from __future__ import annotations

import json
import logging
import os
import re
import time
import threading
import urllib.error
import urllib.parse
import urllib.request
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from functools import lru_cache
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from app_runtime import environment_value, load_explicit_app_config, load_skill_context


DEFAULT_BASE_URL = "https://www.ai.atyou.cn"
# 单次请求超时。**2026-10-10 由 90 提到 180**（`AI_TIMEOUT_SECONDS` 可覆盖）：
#   实测真实批量（40 行 / 39 KB 请求体）**端点空闲时串行 28~35s**，90s 只剩 3 倍余量，
#   而并发跑起来后端点半忙时抖动就会撞墙（一轮 20 批曾撞 5 次）。180s 给足余量且不影响正常路径
#   —— 正常请求 30s 就返回了，超时值只决定"卡住时等多久"。量测工具：`tools/probe_ai_endpoint.py`。
DEFAULT_TIMEOUT_SECONDS = 180
LOGGER = logging.getLogger("kks-audit.ai")
ProgressCallback = Callable[[dict[str, Any]], None]
INTERNAL_GROUP_SIZE = 12
REQUEST_ATTEMPTS = 2
RETRY_BACKOFF_SECONDS = 2
REVIEW_BATCH_RETRIES = 2


def _is_retriable_batch_error(error: str) -> bool:
    """Shared transient-error check for audit review and generation name review."""
    text = str(error).lower()
    return any(token in text for token in (
        "timeout", "timed out", "connection", "reset", "disconnected", "remote end",
        "json", "429", "408", "500", "502", "503", "504", "超时", "连接", "解析",
    ))
MAX_CONTEXT_ROWS = 12
MAX_ADJACENT_ROWS = 5
MAX_GLOBAL_TEXT = 1800
MAX_CONCURRENT_REQUESTS = 4
MODEL_REQUEST_SLOTS = threading.BoundedSemaphore(MAX_CONCURRENT_REQUESTS)
# 并发这个值的实测依据（真实 deepseek-flash 端点，2026-10-08）：
#   8 路并发时语义复核大量整组失败（"Remote end closed connection"/超时），15 个候选只复核到 1~4 个；
#   降到 4 路后同规模用例全部完成。批内重试救不回"端点被打崩"，所以宁可用较低并发换完成率。
# 端点够强时可以调高：AI_REVIEW_CONCURRENCY（上限 8）。
RULE_CONTEXT_HINTS = {
    "KKS-24": {"名称", "命名", "歧义"},
    "KKS-25": {"同一系统", "归一名称", "多个 KKS"},
    "KKS-25b": {"同物异名", "名称"},
    "KKS-25c": {"原码", "身份键", "多个新码"},
    "KKS-26": {"应编未编", "父设备"},
    "KKS-27": {"名称", "文本", "符号", "语义"},
    "KKS-28": {"机组", "名称", "父级", "三方"},
    "KKS-29": {"OCR", "易混", "字符"},
    "KKS-30": {"文本卫生", "空白", "控制符"},
    "KKS-31": {"父子", "名称", "语义"},
    "KKS-13/16": {"旧码", "迁移", "历史", "旧前缀"},
    "KKS-14": {"系统字母", "重分配"},
    "KKS-15": {"设备字母", "位置", "变化"},
    "KKS-16": {"旧前缀", "历史提示"},
    "KKS-18": {"历史", "原码", "前缀"},
    "KKS-17/22": {"扩展码", "12 位", "部件", "信号"},
    "KKS-TREE-REL": {"父级", "子级", "跨子系统", "变长"},
}


class AIReviewError(RuntimeError):
    """A safe-to-report model request or response error."""


@dataclass(frozen=True)
class AIConfig:
    base_url: str
    api_key: str
    model: str
    timeout_seconds: int
    enabled: bool
    max_candidates: int = 0
    summary_enabled: bool = True   # 单文件总体总结（AI_SUMMARY_ENABLED）
    trend_enabled: bool = True     # 台账跨文件趋势总结（AI_TREND_SUMMARY_ENABLED）
    review_concurrency: int = 4         # 语义复核组并发数（AI_REVIEW_CONCURRENCY）
    review_batch_size: int = 12
    context_rows: int = 12
    thinking_mode: str = ""
    brief_output: bool = False


def _env_int(name: str, default: int, minimum: int, maximum: int) -> int:
    try:
        value = int(os.environ.get(name, str(default)))
    except ValueError:
        value = default
    return max(minimum, min(value, maximum))


def _env_bool(name: str, default: bool) -> bool:
    raw = os.environ.get(name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "yes", "y", "on"}


def _parse_bool(raw: Any, default: bool) -> bool:
    if raw in (None, ""):
        return default
    return str(raw).strip().lower() in {"1", "true", "yes", "y", "on"}


def _parse_int(raw: Any, default: int, minimum: int, maximum: int) -> int:
    try:
        value = int(raw)
    except (TypeError, ValueError):
        value = default
    return max(minimum, min(value, maximum))


def _load_local_env() -> None:
    """Load an optional local-only config without overriding process env vars.

    🔴 这个函数**必须被调用**（此前定义了却从未调用，是交接文档里记的已知缺口）：
    不调用时 `.env.local` 里的 `KKS_ADMIN_TOKEN` 不会进入进程环境，后果是
    **管理接口一律 403（fail-closed）**，且 AI 配置退回默认值 → 界面上"勾了 AI"也跑不起来。
    调用点见文件末的 `_load_local_env()`（模块级，导入即生效；不覆盖真实环境变量）。
    """
    path = Path(__file__).with_name(".env.local")
    if not path.is_file():
        return
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        name, value = line.split("=", 1)
        name, value = name.strip(), value.strip()
        if name and name.isidentifier() and name not in os.environ:
            os.environ[name] = value.strip('"\'')


def normalize_base_url(value: str) -> str:
    """Normalize a provider root to a versioned OpenAI-compatible API root."""
    base = value.strip().rstrip("/")
    if not base:
        base = DEFAULT_BASE_URL
    parsed = urllib.parse.urlparse(base)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise AIReviewError("AI_BASE_URL 必须是 http:// 或 https:// 地址")
    if not re.search(r"/v\d+(?:\.\d+)?$", parsed.path.rstrip("/")):
        base += "/v1"
    return base


def load_config() -> AIConfig:
    """解析 AI 配置：进程环境 > config/app_config.json（文件里真写了的值）> .env.local > 默认值。

    ⚠️ 这里刻意用 load_explicit_app_config() 而不是 load_app_config()：
    后者会把 DEFAULT_APP_CONFIG 的默认值（如 ai.enabled=False）一起返回，默认值一旦被
    当成"管理员配过"，就会排在 .env.local 前面，使 .env.local 里的 AI_ENABLED=true 失效。
    """
    app_config = load_explicit_app_config()
    ai = app_config.get("ai", {}) if isinstance(app_config.get("ai"), dict) else {}
    api_key = environment_value("AI_API_KEY", ai.get("api_key", "")).strip()
    enabled = _parse_bool(environment_value("AI_ENABLED", ai.get("enabled", "")), bool(api_key))
    raw_base_url = environment_value("AI_BASE_URL", ai.get("base_url", "")).strip()
    return AIConfig(
        base_url=normalize_base_url(raw_base_url) if raw_base_url else "",
        api_key=api_key,
        model=environment_value("AI_MODEL", ai.get("model", "")).strip(),
        # 超时/并发/上限这类"调优参数"没有配置文件里的显式值时用代码默认值兜底
        timeout_seconds=_parse_int(environment_value("AI_TIMEOUT_SECONDS", ai.get("timeout_seconds", DEFAULT_TIMEOUT_SECONDS)), DEFAULT_TIMEOUT_SECONDS, 10, 300),
        enabled=enabled,
        max_candidates=_parse_int(environment_value("AI_REVIEW_MAX_CANDIDATES", ai.get("max_candidates", 0)), 0, 0, 2000),
        summary_enabled=_parse_bool(environment_value("AI_SUMMARY_ENABLED", ai.get("summary_enabled", "")), True),
        trend_enabled=_parse_bool(environment_value("AI_TREND_SUMMARY_ENABLED", ai.get("trend_enabled", "")), True),
        review_concurrency=_parse_int(environment_value("AI_REVIEW_CONCURRENCY", ai.get("review_concurrency", MAX_CONCURRENT_REQUESTS)), MAX_CONCURRENT_REQUESTS, 1, 8),
        review_batch_size=_parse_int(environment_value("AI_REVIEW_BATCH_SIZE", ai.get("review_batch_size", 12)), 12, 1, 32),
        context_rows=_parse_int(environment_value("AI_REVIEW_CONTEXT_ROWS", ai.get("context_rows", 12)), 12, 4, 24),
        thinking_mode=environment_value("AI_THINKING_MODE", ai.get("thinking_mode", "")).strip().lower(),
        brief_output=_parse_bool(environment_value("AI_REVIEW_BRIEF_OUTPUT", ai.get("brief_output", "")), False),
    )


class OpenAICompatibleClient:
    def __init__(self, config: AIConfig, *, audit_requests: bool = False) -> None:
        if not config.api_key:
            raise AIReviewError("未配置 AI_API_KEY")
        if not config.base_url:
            raise AIReviewError("未配置 AI_BASE_URL")
        self.config = config
        self.audit_requests = audit_requests

    def _request_json(self, path: str, payload: dict[str, Any] | None = None) -> dict[str, Any]:
        if (self.audit_requests and payload is not None and path == "chat/completions"
                and str(payload.get("model", "")).lower().startswith("deepseek")
                and self.config.thinking_mode in {"enabled", "disabled"}):
            payload = dict(payload)
            payload.setdefault("thinking", {"type": self.config.thinking_mode})
        url = f"{self.config.base_url}/{path.lstrip('/')}"
        body = None if payload is None else json.dumps(payload, ensure_ascii=False).encode("utf-8")
        request = urllib.request.Request(
            url,
            data=body,
            method="POST" if body is not None else "GET",
            headers={
                "Accept": "application/json",
                "Authorization": f"Bearer {self.config.api_key}",
                "Content-Type": "application/json",
                "User-Agent": "kks-audit-agent/0.2",
            },
        )
        last_error: AIReviewError | None = None
        for attempt in range(1, REQUEST_ATTEMPTS + 1):
            try:
                with MODEL_REQUEST_SLOTS:
                    with urllib.request.urlopen(request, timeout=self.config.timeout_seconds) as response:
                        raw = response.read().decode("utf-8", errors="replace")
                LOGGER.info("model_request path=%s status=200 bytes=%s attempt=%s", path, len(raw), attempt)
                break
            except urllib.error.HTTPError as exc:
                detail = exc.read().decode("utf-8", errors="replace")[:500]
                last_error = AIReviewError(f"模型接口 HTTP {exc.code}：{detail}")
                if exc.code not in {408, 425, 429, 500, 502, 503, 504} or attempt >= REQUEST_ATTEMPTS:
                    raise last_error from exc
            except urllib.error.URLError as exc:
                last_error = AIReviewError(f"模型接口连接失败：{exc.reason}")
                if attempt >= REQUEST_ATTEMPTS:
                    raise last_error from exc
            except TimeoutError as exc:
                last_error = AIReviewError("模型接口请求超时")
                if attempt >= REQUEST_ATTEMPTS:
                    raise last_error from exc
            LOGGER.warning("model_request_retry path=%s attempt=%s error=%s", path, attempt, last_error)
            time.sleep(RETRY_BACKOFF_SECONDS * attempt)
        else:
            raise last_error or AIReviewError("模型接口请求失败")
        try:
            decoded = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise AIReviewError("模型接口返回的不是 JSON") from exc
        if not isinstance(decoded, dict):
            raise AIReviewError("模型接口返回 JSON 类型异常")
        return decoded

    def choose_model(self) -> tuple[str, str]:
        if self.config.model:
            return self.config.model, "environment"
        payload = self._request_json("models")
        models = payload.get("data")
        if not isinstance(models, list):
            raise AIReviewError("模型接口未返回可用模型列表；请设置 AI_MODEL")
        ids = [str(item.get("id", "")).strip() for item in models if isinstance(item, dict)]
        ids = [item for item in ids if item]
        if not ids:
            raise AIReviewError("模型列表为空；请设置 AI_MODEL")
        return ids[0], "auto"

    def review(self, model: str, system_prompt: str, user_payload: dict[str, Any]) -> dict[str, Any]:
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": json.dumps(user_payload, ensure_ascii=False)},
        ]
        payload: dict[str, Any] = {
            "model": model,
            "messages": messages,
            "temperature": 0.1,
            "response_format": {"type": "json_object"},
        }
        try:
            response = self._request_json("chat/completions", payload)
        except AIReviewError as exc:
            # Some compatible gateways do not implement response_format. Retry
            # only for that compatibility case, never for arbitrary failures.
            if "response_format" not in str(exc).lower() and "json_object" not in str(exc).lower():
                raise
            payload.pop("response_format", None)
            response = self._request_json("chat/completions", payload)
        choices = response.get("choices")
        if not isinstance(choices, list) or not choices:
            raise AIReviewError("模型接口未返回 choices")
        message = choices[0].get("message", {}) if isinstance(choices[0], dict) else {}
        content = message.get("content") if isinstance(message, dict) else None
        if isinstance(content, list):
            content = "".join(str(part.get("text", "")) for part in content if isinstance(part, dict))
        if isinstance(content, dict):
            return content
        if not isinstance(content, str) or not content.strip():
            raise AIReviewError("模型接口未返回可解析的 message.content")
        return parse_json_content(content)


SYSTEM_PROMPT = """只审查现有编码，不生成或输出任何替代编码。审查顺序：格式、设备类型、系统归属、名称语义、编号、位置。明确强规则错误不得被语义判断撤销。
你是电厂 KKS 编码审核智能体的二次复核员。
规则引擎已经给出了候选问题。你只能依据候选证据包和本组相关 Skill 片段判断，不能凭空补充表外事实。
证据包中的当前行、父级、子级、同级编码和相邻行必须结合起来看；父子关系以 Excel 的显式父级列为准，不能用固定前 N 位臆测变长层级。
不要修改原 Excel，不要自动改码，不要自动关闭问题；证据不足时必须返回 needs_human。
classification_reference给出的分类释义及出处是分类核对依据；同级编码一致、父级存在、编号连续，仅说明内部结构自洽，不能证明分类符合国标。设备类型粗类目不能替代实际设备主体。
application_examples是国标应用实例，须与抽象分类定义一起核对；不得把明确列在实例中的设备排除。设备主体符合实例只支持本项分类，不证明编号、位置或其它工程事实。
必须区分设备本体、上下游修饰语、执行机构及测量对象；设备序号、加热器级号、抽汽段号不能当机组号。工程编号、备用、位置、未收录分类及工程索引缺少证据时只能needs_human。
名称以一次门、二次门或阀门等实体主体结尾时，前面的模拟量、液位、压力等词描述用途，不能仅因含测量词就判为指令或测点；明确的控制指令和末尾反馈/状态需另行识别。
suggestion只写核对资料和处理事项，不得包含任何长度的建议编码值，也不得给出带目标编码的迁移、重编或父级改码指令。
硬性语法问题通常保留为 confirmed_issue；变长码、历史映射、机组语义、名称异常要先检查证据包中的上下文，再判断是否可能误报。
每条结果必须返回结论、证据、原因和建议。证据只能引用证据包中真实存在的行号、编码、名称或规则字段，不能写“根据经验”作为唯一证据。
证据包的行引用从row_pool按引用键读取，信息与原行等价。只审本项规则；每项evidence、reason分别不超过80字，suggestion不超过40字；直接给出必要依据，避免复述整组上下文。
请只返回 JSON 对象，格式为：
{"reviews":[{"issue_index":0,"decision":"confirmed_issue|likely_false_positive|needs_human","priority":"P0|P1|P2","confidence":0.0,"evidence":"引用证据包中的具体行号/编码/关系","reason":"中文理由","suggestion":"中文建议"}]}
confidence 必须是 0 到 1 的数字。每个 issue_index 最多返回一次。"""

GLOBAL_REVIEW_PROMPT = """只审查现有编码，不生成或输出任何替代编码。审查顺序：格式、设备类型、系统归属、名称语义、编号、位置。明确强规则错误不得被语义判断撤销。
你是电厂 KKS 编码审核智能体的全局分析员。
你将收到一份 Excel 的结构画像、规则命中统计和候选问题摘要。先从整个文件的角度识别专业范围、层级组织、命名习惯、重复模式和可能的误报来源，帮助后续逐项复核。
不要凭空补充 Excel 外的事实，不要直接修改问题，也不要把某一条候选问题当成最终结论。
请只返回 JSON 对象，格式为：
{"global_review":{"summary":"整体审核观察","patterns":["全局模式或异常"],"focus":[{"rule_id":"规则编号","reason":"为什么需要重点核对"}],"consistency_checks":["建议后续重点验证的跨行关系"]}}
所有内容必须来自输入的结构画像、规则统计或候选摘要。"""

FINAL_ADJUDICATION_PROMPT = """只审查现有编码，不生成或输出任何替代编码。审查顺序：格式、设备类型、系统归属、名称语义、编号、位置。明确强规则错误不得被语义判断撤销。
你是电厂 KKS 编码审核智能体的最终归并员。
你将收到规则命中、AI 初审、疑似误报二次核验和全局分析。请为每个候选给出最终分类：
- confirmed_issue：证据足够，应该列为需要处理的问题；
- needs_human：存在冲突或证据不足，必须人工确认；
- likely_false_positive：只有在二次核验明确支持时才能使用。
程序规则问题不能被静默删除；如果 AI 没有完成复核，也必须归入 needs_human。
不要改变原 Excel，不要补充输入中不存在的事实。
请只返回 JSON 对象，格式为：
{"final_reviews":[{"issue_index":0,"final_decision":"confirmed_issue|needs_human|likely_false_positive","summary":"面向审核人员的简短结论","evidence":"引用输入中的真实行号、编码或关系","reason":"归并原因","suggestion":"下一步建议"}]}
每个 issue_index 最多返回一次。"""

VERIFICATION_PROMPT = """
现在进行第二次核验。以下结果曾被初审判断为 likely_false_positive，不能直接采信。
请重新检查原始证据包、规则片段和初审理由，确认它是 confirmed_issue、likely_false_positive 还是 needs_human。
如果初审没有足够的真实证据，必须改为 needs_human，不要为了减少问题而确认误报。
仍然只返回同样格式的 JSON，并且每条结果必须包含 evidence、reason、suggestion。
"""

def effective_system_prompt(rule_ids: set[str] | None = None, categories: set[str] | None = None, *, verification: bool = False) -> str:
    selected_rules = rule_ids or set()
    selected_categories = categories or set()
    terms = set(selected_rules) | set(selected_categories)
    for rule_id in selected_rules:
        terms.update(RULE_CONTEXT_HINTS.get(rule_id, set()))
    formal_context = load_skill_context(
        rule_ids=selected_rules,
        categories=selected_categories,
        keywords=terms,
        max_chars=12000,
    ).strip()
    prompt = SYSTEM_PROMPT + (VERIFICATION_PROMPT if verification else "")
    if formal_context:
        prompt += "\n\n以下是本组规则对应的 Skill 片段。只把它们作为规则解释，不得把参考资料中的厂站案例、固定码表或经验直接当成当前 Excel 的事实：\n" + formal_context
    return prompt


def _skill_context_suffix(rule_ids: set[str], categories: set[str], *, max_chars: int = 16000) -> str:
    terms = set(rule_ids) | set(categories)
    for rule_id in rule_ids:
        terms.update(RULE_CONTEXT_HINTS.get(rule_id, set()))
    formal_context = load_skill_context(
        rule_ids=rule_ids,
        categories=categories,
        keywords=terms,
        max_chars=max_chars,
    ).strip()
    if not formal_context:
        return ""
    return "\n\n以下是本次输入中相关规则对应的 Skill 片段。只把它们作为规则解释，不得把参考资料中的厂站案例、固定码表或经验直接当成当前 Excel 的事实：\n" + formal_context


def effective_global_system_prompt(rule_ids: set[str], categories: set[str]) -> str:
    return GLOBAL_REVIEW_PROMPT + _skill_context_suffix(rule_ids, categories)


def parse_json_content(content: str) -> dict[str, Any]:
    cleaned = content.strip()
    cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s*```$", "", cleaned)
    try:
        payload = json.loads(cleaned)
    except json.JSONDecodeError as exc:
        match = re.search(r"\{.*\}", cleaned, flags=re.DOTALL)
        if not match:
            raise AIReviewError("模型返回内容不是预期 JSON") from exc
        try:
            payload = json.loads(match.group(0))
        except json.JSONDecodeError as nested:
            raise AIReviewError("模型返回内容不是预期 JSON") from nested
    if isinstance(payload, list):
        return {"reviews": payload}
    if not isinstance(payload, dict):
        raise AIReviewError("模型返回 JSON 顶层类型异常")
    return payload


def _is_candidate(issue: dict[str, Any]) -> bool:
    if issue.get("compliance_locked") or issue.get("evidence_missing"):
        return False
    if issue.get("status") == "resolved":
        return False
    if issue.get("status") == "needs_review":
        return True
    # These are semantic checks where a model can add context, but the model
    # is never allowed to erase the deterministic issue itself.
    return issue.get("rule_id") in {
        "KKS-24", "KKS-25", "KKS-25b", "KKS-25c", "KKS-26", "KKS-28", "KKS-31",
        "KKS-A", "KKS-C", "KKS-E", "KKS-F", "KKS-G", "KKS-13/16", "KKS-14", "KKS-15", "KKS-16", "KKS-18",
    }


def _row_view(record: dict[str, Any]) -> dict[str, Any]:
    return {
        key: record.get(key, "")
        for key in (
            "sheet", "excel_row", "kks_code", "parent_code", "old_code", "name", "unit", "device_type", "system",
            "raw_kks", "raw_name",
        )
        if record.get(key, "") not in (None, "")
    }


def _context_records(result: dict[str, Any]) -> list[dict[str, Any]]:
    context = result.get("_ai_context", {})
    records = context.get("records", []) if isinstance(context, dict) else []
    return [record for record in records if isinstance(record, dict)]


@lru_cache(maxsize=1)
def _gb_device_index() -> dict[str, Any]:
    path = Path(__file__).resolve().parent / "config" / "gb_device_codes.json"
    return json.loads(path.read_text(encoding="utf-8")) if path.is_file() else {}


@lru_cache(maxsize=8192)
def _classification_reference(code: str) -> dict[str, Any]:
    """Give the model the local index meaning and its actual provenance."""
    if not re.fullmatch(r"[0-9]{2}[A-HJ-NP-Z]{3}[0-9]{2}[A-HJ-NP-Z]{2}[0-9]{3}", code):
        return {"status": "未按十二位设备骨架解析，禁止猜测分类段"}
    from kks_gen import identifiers as ids
    system, device = code[2:5], code[7:9]
    return {
        "standard": "GB/T 50549-2020",
        "system": {"code": system, "definition": ids.SYS_DICT.get(system, {}).get("cn", "未收录"),
                   "source": "国标附录A" if system in ids.GB_SYS_DICT else "项目补充或未收录，国标依据待核"},
        "device": {"code": device, "definition": ids.DEV_DICT.get(device, {}).get("cn", "未收录"),
                   "application_examples": _gb_device_index().get(device, {}).get("note", ""),
                   "source": "国标附录B" if device in _gb_device_index() else "项目补充或未收录，国标依据待核"},
    }


def _build_evidence_packet(issue_index: int, issue: dict[str, Any], result: dict[str, Any]) -> dict[str, Any]:
    index = result.get("_ai_evidence_index")
    if index is None:
        records = _context_records(result)
        by_location, by_code, by_parent = {}, defaultdict(list), defaultdict(list)
        for record in records:
            by_location[(record.get("sheet", ""), str(record.get("excel_row", "")))] = record
            by_code[str(record.get("kks_code", ""))].append(record)
            by_parent[str(record.get("parent_code", ""))].append(record)
        index = {"location": by_location, "code": by_code, "parent": by_parent, "records": records}
        result["_ai_evidence_index"] = index
    records = index["records"]
    context_limit = int(result.get("_ai_context_limit", MAX_CONTEXT_ROWS))
    issue_row, issue_code = str(issue.get("excel_row", "")), str(issue.get("kks_code", ""))
    issue_sheet = issue.get("sheet", "")
    current = index["location"].get((issue_sheet, issue_row))
    if current is None and not issue_sheet:
        current = next((r for r in records if str(r.get("excel_row", "")) == issue_row and issue_row), None)
    if current is None:
        matches = index["code"].get(issue_code, []) if issue_code else []
        current = matches[0] if matches else {key: issue.get(key, "") for key in ("sheet", "excel_row", "kks_code", "parent_code", "old_code", "name")}
    current_code = str(current.get("kks_code") or issue_code)
    parent_code = str(current.get("parent_code") or issue.get("parent_code") or "")
    parent_rows = index["code"].get(parent_code, [])[:1] if parent_code else []
    child_rows = index["parent"].get(current_code, [])[:context_limit] if current_code else []
    sibling_rows = [r for r in index["parent"].get(parent_code, [])
                    if (r.get("sheet", ""), str(r.get("excel_row", ""))) != (current.get("sheet", ""), str(current.get("excel_row", "")))][:context_limit] if parent_code else []
    adjacent_rows = []
    try:
        current_row_number = int(current.get("excel_row"))
        for number in range(current_row_number - 2, current_row_number + 3):
            if number != current_row_number:
                nearby = index["location"].get((current.get("sheet", ""), str(number)))
                if nearby:
                    adjacent_rows.append(nearby)
    except (TypeError, ValueError):
        pass

    def capped(values: list[dict[str, Any]], limit: int = MAX_CONTEXT_ROWS) -> list[dict[str, Any]]:
        return [_row_view(value) for value in values[:limit]]

    return {
        "current_row": _row_view(current),
        "classification_reference": _classification_reference(current_code),
        "parent_rows": capped(parent_rows, 1),
        "child_rows": capped(child_rows, context_limit),
        "same_level_rows": capped(sibling_rows, context_limit),
        "adjacent_rows": capped(adjacent_rows, MAX_ADJACENT_ROWS),
        "rule_finding": {
            "rule_id": issue.get("rule_id", ""),
            "category": issue.get("category", ""),
            "priority": issue.get("priority", ""),
            "message": issue.get("message", ""),
            "suggestion": issue.get("suggestion", ""),
            "existing_evidence": issue.get("evidence", ""),
        },
    }


def _candidate(issue_index: int, issue: dict[str, Any], result: dict[str, Any]) -> dict[str, Any]:
    return {
        "issue_index": issue_index,
        "rule_id": issue.get("rule_id", ""),
        "priority": issue.get("priority", ""),
        "category": issue.get("category", ""),
        "excel_row": issue.get("excel_row", ""),
        "kks_code": issue.get("kks_code", ""),
        "parent_code": issue.get("parent_code", ""),
        "old_code": issue.get("old_code", ""),
        "name": issue.get("name", ""),
        "rule_message": issue.get("message", ""),
        "rule_suggestion": issue.get("suggestion", ""),
        "evidence": issue.get("evidence", ""),
        "evidence_packet": _build_evidence_packet(issue_index, issue, result),
    }


def _compact_candidate(candidate: dict[str, Any], issue: dict[str, Any] | None = None) -> dict[str, Any]:
    """Return a compact candidate view for global analysis and final merge."""
    value = issue or candidate
    return {
        "issue_index": candidate.get("issue_index", ""),
        "rule_id": candidate.get("rule_id", ""),
        "category": candidate.get("category", ""),
        "priority": candidate.get("priority", ""),
        "excel_row": candidate.get("excel_row", ""),
        "kks_code": candidate.get("kks_code", ""),
        "parent_code": candidate.get("parent_code", ""),
        "name": candidate.get("name", ""),
        "rule_message": candidate.get("rule_message", value.get("message", "")),
        "evidence": candidate.get("evidence", value.get("evidence", "")),
    }


def _build_global_audit_context(result: dict[str, Any], candidates: list[dict[str, Any]]) -> dict[str, Any]:
    issues = [item for item in result.get("issues", []) if isinstance(item, dict)]
    rule_counts = defaultdict(int)
    category_counts = defaultdict(int)
    priority_counts = defaultdict(int)
    for issue in issues:
        rule_counts[str(issue.get("rule_id", "未标识"))] += 1
        category_counts[str(issue.get("category", "未分类"))] += 1
        priority_counts[str(issue.get("priority", "未标级"))] += 1
    return {
        "source_file": result.get("source_file", ""),
        "sheet": result.get("sheet", ""),
        "header_row": result.get("header_row", ""),
        "data_rows": result.get("data_rows", 0),
        "columns": result.get("columns", {}),
        "metrics": result.get("metrics", {}),
        "code_length_distribution": result.get("code_length_distribution", {}),
        "priority_counts": dict(sorted(priority_counts.items())),
        "rule_counts": dict(sorted(rule_counts.items(), key=lambda item: (-item[1], item[0]))),
        "category_counts": dict(sorted(category_counts.items(), key=lambda item: (-item[1], item[0]))),
        "structural_notes": [str(item) for item in result.get("structural_notes", [])],
        "candidate_count": len(candidates),
        "candidate_summary": [_compact_candidate(candidate) for candidate in candidates],
    }


def _compact_audit_profile(audit_profile: dict[str, Any]) -> dict[str, Any]:
    """Keep shared context small for repeated group and final calls.

    The full candidate summary is useful for the one global call, but sending
    it again with every evidence group multiplies prompt size and latency.
    Group payloads already contain their own complete candidates.
    """
    return {
        key: value
        for key, value in audit_profile.items()
        if key != "candidate_summary"
    }


def _safe_review(item: Any, allowed: set[int]) -> dict[str, Any] | None:
    if not isinstance(item, dict):
        return None
    try:
        issue_index = int(item.get("issue_index"))
        confidence = float(item.get("confidence", 0))
    except (TypeError, ValueError):
        return None
    if issue_index not in allowed:
        return None
    decision = str(item.get("decision", "needs_human"))
    if decision not in {"confirmed_issue", "likely_false_positive", "needs_human"}:
        decision = "needs_human"
    priority = str(item.get("priority", ""))
    if priority not in {"P0", "P1", "P2"}:
        priority = ""
    return {
        "issue_index": issue_index,
        "decision": decision,
        "priority": priority,
        "confidence": max(0.0, min(confidence, 1.0)),
        "evidence": str(item.get("evidence", item.get("supporting_evidence", "模型未提供明确证据")))[:1500],
        "reason": str(item.get("reason", "模型未提供理由"))[:1000],
        "suggestion": str(item.get("suggestion", "保留人工确认"))[:1000],
    }


def _safe_global_review(payload: Any) -> dict[str, Any]:
    value = payload.get("global_review") if isinstance(payload, dict) else None
    value = value if isinstance(value, dict) else {}

    def text_value(key: str, default: str = "") -> str:
        raw = value.get(key, default)
        return str(raw).strip()[:MAX_GLOBAL_TEXT] if raw not in (None, "") else default

    def list_value(key: str) -> list[str]:
        raw = value.get(key, [])
        if not isinstance(raw, list):
            return []
        return [str(item).strip()[:MAX_GLOBAL_TEXT] for item in raw if str(item).strip()][:20]

    focus: list[dict[str, str]] = []
    raw_focus = value.get("focus", [])
    if isinstance(raw_focus, list):
        for item in raw_focus[:20]:
            if isinstance(item, dict):
                rule_id = str(item.get("rule_id", "")).strip()
                reason = str(item.get("reason", "")).strip()
                if rule_id or reason:
                    focus.append({"rule_id": rule_id[:80], "reason": reason[:MAX_GLOBAL_TEXT]})
    return {
        "summary": text_value("summary", "模型未提供全局分析"),
        "patterns": list_value("patterns"),
        "focus": focus,
        "consistency_checks": list_value("consistency_checks"),
    }


def _safe_final_review(item: Any, allowed: set[int]) -> dict[str, Any] | None:
    if not isinstance(item, dict):
        return None
    try:
        issue_index = int(item.get("issue_index"))
    except (TypeError, ValueError):
        return None
    if issue_index not in allowed:
        return None
    decision = str(item.get("final_decision", "needs_human"))
    if decision not in {"confirmed_issue", "needs_human", "likely_false_positive"}:
        decision = "needs_human"
    return {
        "issue_index": issue_index,
        "final_decision": decision,
        "summary": str(item.get("summary", "需要人工确认"))[:MAX_GLOBAL_TEXT],
        "evidence": str(item.get("evidence", "模型未提供明确证据"))[:1500],
        "reason": str(item.get("reason", "模型未提供归并理由"))[:1000],
        "suggestion": str(item.get("suggestion", "保留人工确认"))[:1000],
    }


def _emit_progress(callback: ProgressCallback | None, event: dict[str, Any]) -> None:
    if callback is None:
        return
    try:
        callback(event)
    except Exception as exc:  # status reporting must not fail the audit
        LOGGER.warning("ai_progress_callback_failed error=%s", exc)


def _candidate_groups(candidates: list[dict[str, Any]], batch_size: int | None = None) -> list[list[dict[str, Any]]]:
    grouped: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for candidate in candidates:
        key = (str(candidate.get("rule_id", "")), str(candidate.get("category", "")))
        grouped[key].append(candidate)
    size = max(1, int(batch_size or INTERNAL_GROUP_SIZE))
    groups: list[list[dict[str, Any]]] = []
    for key in sorted(grouped):
        values = sorted(grouped[key], key=lambda item: int(item.get("issue_index", 0)))
        for start in range(0, len(values), size):
            groups.append(values[start:start + size])
    return groups


def _group_evidence_request(group: list[dict[str, Any]]) -> dict[str, Any]:
    """Share identical context rows within a batch without dropping evidence."""
    row_pool, row_keys, candidates = {}, {}, []
    for candidate in group:
        item = dict(candidate)
        packet = dict(candidate.get("evidence_packet", {}))
        for field in ("current_row", "parent_rows", "child_rows", "same_level_rows", "adjacent_rows"):
            rows = packet.get(field)
            if rows is None:
                continue
            refs = []
            for row in ([rows] if isinstance(rows, dict) else rows):
                fingerprint = json.dumps(row, ensure_ascii=False, sort_keys=True)
                if fingerprint not in row_keys:
                    key = f"r{len(row_pool)}"
                    row_keys[fingerprint] = key
                    row_pool[key] = row
                refs.append(row_keys[fingerprint])
            packet[field] = refs[0] if isinstance(rows, dict) else refs
        item["evidence_packet"] = packet
        candidates.append(item)
    return {"candidates": candidates, "row_pool": row_pool}


def _review_group(
    client: OpenAICompatibleClient,
    model: str,
    group: list[dict[str, Any]],
    *,
    verification: bool = False,
    audit_profile: dict[str, Any] | None = None,
    global_review: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """复核一个候选组；瞬时故障（连接被断/超时/非 JSON）自动重试。

    真实端点并发 8 时实测会出现整组丢失（15 个候选只复核到 1 个）。复核是幂等的只读请求，
    重试安全，且比"整组丢掉、报告少一批结论"划算得多。
    """
    rule_ids = {str(item.get("rule_id", "")) for item in group}
    categories = {str(item.get("category", "")) for item in group}
    request = {
        "review_type": "false_positive_verification" if verification else "candidate_review",
        "audit_profile": _compact_audit_profile(audit_profile or {}),
        "global_review": global_review or {},
        **_group_evidence_request(group),
    }
    prompt = effective_system_prompt(rule_ids, categories, verification=verification)
    if getattr(getattr(client, "config", None), "brief_output", False):
        prompt += "\n本次要求更简洁：evidence、reason各不超过40字，suggestion不超过20字。保留真实行号和必要分类依据；不能因简洁省略导致结论无依据。"
    payload: dict[str, Any] | None = None
    for attempt in range(1, REVIEW_BATCH_RETRIES + 1):
        try:
            payload = client.review(model, prompt, request)
            break
        except Exception as exc:
            # Transport retries already happen inside _request_json. Only
            # retry parsing here, avoiding four long transport attempts.
            if attempt < REVIEW_BATCH_RETRIES and any(term in str(exc).lower() for term in ("json", "解析")):
                LOGGER.warning(
                    "ai_review_group_retry attempt=%s candidates=%s error=%s",
                    attempt, len(group), str(exc)[:120],
                )
                time.sleep(RETRY_BACKOFF_SECONDS * attempt)
                continue
            raise
    if payload is None:
        raise AIReviewError("模型未返回可解析的复核结果")
    raw_reviews = payload.get("reviews", [])
    if not isinstance(raw_reviews, list):
        raise AIReviewError("模型返回的 reviews 不是数组")
    allowed = {int(item["issue_index"]) for item in group}
    reviews: list[dict[str, Any]] = []
    seen: set[int] = set()
    for item in raw_reviews:
        review = _safe_review(item, allowed)
        if review is None or review["issue_index"] in seen:
            continue
        seen.add(review["issue_index"])
        def original_row_reference(match):
            row = request.get("row_pool", {}).get(match.group())
            if row:
                return f"{row.get('sheet', '')}!{row.get('excel_row', '')}"
            return match.group()
        for field in ("evidence", "reason", "suggestion"):
            review[field] = re.sub(r"(?<![A-Za-z0-9])r[0-9]+(?![A-Za-z0-9])", original_row_reference, review[field])
        reviews.append(review)
    return reviews


def _default_final_review(issue: dict[str, Any], *, candidate: bool) -> dict[str, Any]:
    ai_decision = str(issue.get("ai_decision", ""))
    verification = str(issue.get("ai_verification_decision", ""))
    if issue.get("evidence_missing"):
        decision = "needs_human"
        source = "missing_evidence"
    elif not candidate:
        decision = "confirmed_issue"
        source = "rule"
    elif ai_decision == "likely_false_positive" and verification == "likely_false_positive":
        decision = "likely_false_positive"
        source = "ai_verification"
    elif ai_decision == "confirmed_issue":
        decision = "confirmed_issue"
        source = "ai"
    else:
        decision = "needs_human"
        source = "ai" if ai_decision else "rule_unreviewed"
    action = {
        "confirmed_issue": "需要处理",
        "needs_human": "需要人工确认",
        "likely_false_positive": "暂不列入整改，保留复核记录",
    }[decision]
    return {
        "final_decision": decision,
        "final_source": source,
        "final_action": action,
        "final_summary": str(issue.get("ai_reason") or issue.get("message", "")),
        "final_evidence": str(issue.get("ai_evidence") or issue.get("evidence", "")),
        "final_reason": str(issue.get("ai_reason") or ("程序规则命中；尚未获得可替代规则结论的 AI 归并结果。" if not candidate else "AI 复核未形成可直接采信的最终结论。")),
        "final_suggestion": str(issue.get("ai_suggestion") or issue.get("suggestion", "人工确认")),
    }


def _apply_final_review(issue: dict[str, Any], review: dict[str, Any], *, source: str) -> None:
    decision = review["final_decision"]
    issue["final_decision"] = decision
    issue["final_source"] = source
    issue["final_action"] = {
        "confirmed_issue": "需要处理",
        "needs_human": "需要人工确认",
        "likely_false_positive": "暂不列入整改，保留复核记录",
    }[decision]
    issue["final_summary"] = review["summary"]
    issue["final_evidence"] = review["evidence"]
    issue["final_reason"] = review["reason"]
    issue["final_suggestion"] = review["suggestion"]


def _finalize_reviews(
    client: OpenAICompatibleClient,
    model: str,
    candidates: list[dict[str, Any]],
    issues: list[dict[str, Any]],
    audit_profile: dict[str, Any],
    global_review: dict[str, Any],
) -> list[dict[str, Any]]:
    candidate_payload: list[dict[str, Any]] = []
    for candidate in candidates:
        issue = issues[int(candidate["issue_index"])]
        candidate_payload.append({
            **_compact_candidate(candidate, issue),
            "ai_initial_decision": issue.get("ai_initial_decision", ""),
            "ai_initial_evidence": issue.get("ai_initial_evidence", ""),
            "ai_initial_reason": issue.get("ai_initial_reason", ""),
            "ai_verification_decision": issue.get("ai_verification_decision", ""),
            "ai_verification_evidence": issue.get("ai_verification_evidence", ""),
            "ai_verification_reason": issue.get("ai_verification_reason", ""),
        })
    payload = client.review(
        model,
        FINAL_ADJUDICATION_PROMPT + _skill_context_suffix(
            {str(item.get("rule_id", "")) for item in candidates},
            {str(item.get("category", "")) for item in candidates},
        ),
        {
            "review_type": "final_adjudication",
            "audit_profile": _compact_audit_profile(audit_profile),
            "global_review": global_review,
            "candidates": candidate_payload,
        },
    )
    raw_reviews = payload.get("final_reviews", [])
    if not isinstance(raw_reviews, list):
        raise AIReviewError("模型返回的 final_reviews 不是数组")
    allowed = {int(item["issue_index"]) for item in candidates}
    reviews: list[dict[str, Any]] = []
    seen: set[int] = set()
    for item in raw_reviews:
        review = _safe_final_review(item, allowed)
        if review is None or review["issue_index"] in seen:
            continue
        seen.add(review["issue_index"])
        reviews.append(review)
    return reviews


def _attach_review(issue: dict[str, Any], review: dict[str, Any], model: str, *, verification: bool = False) -> None:
    prefix = "ai_verification" if verification else "ai_initial"
    issue[f"{prefix}_decision"] = review["decision"]
    issue[f"{prefix}_confidence"] = review["confidence"]
    issue[f"{prefix}_evidence"] = review["evidence"]
    issue[f"{prefix}_reason"] = review["reason"]
    issue[f"{prefix}_suggestion"] = review["suggestion"]
    issue["ai_decision"] = review["decision"]
    issue["ai_confidence"] = review["confidence"]
    issue["ai_evidence"] = review["evidence"]
    issue["ai_reason"] = review["reason"]
    issue["ai_suggestion"] = review["suggestion"]
    issue["ai_model"] = model


def _review_groups_parallel(
    client: OpenAICompatibleClient,
    model: str,
    groups: list[list[dict[str, Any]]],
    *,
    verification: bool,
    audit_profile: dict[str, Any],
    global_review: dict[str, Any],
    progress_callback: ProgressCallback | None,
    candidate_count: int,
    reviewed_count: int,
    percent_start: int,
    percent_end: int,
    review_concurrency: int = MAX_CONCURRENT_REQUESTS,
) -> tuple[dict[int, dict[str, Any]], list[str]]:
    """Review independent groups concurrently with a bounded worker pool.

    并发数可配（AI_REVIEW_CONCURRENCY）：实测真实端点在 8 路并发下会大量整组失败，
    降到 4 路后可全部完成——并发过高反而更慢（要重试、还丢结果）。
    """
    reviews_by_index: dict[int, dict[str, Any]] = {}
    errors: list[str] = []
    if not groups:
        return reviews_by_index, errors

    completed_groups = 0
    max_workers = min(review_concurrency, len(groups))
    with ThreadPoolExecutor(max_workers=max_workers, thread_name_prefix="kks-ai") as executor:
        futures = {
            executor.submit(
                _review_group,
                client,
                model,
                group,
                verification=verification,
                audit_profile=audit_profile,
                global_review=global_review,
            ): (group_index, group)
            for group_index, group in enumerate(groups, start=1)
        }
        for future in as_completed(futures):
            group_index, group = futures[future]
            try:
                # 组级重试放在 _review_group 内部（见 group_retries 参数）：
                # 真实端点并发 8 时出现过"只复核 1/15"——大量组因连接被断而整组丢失。
                # 复核组是幂等的（只读证据包、不改共享状态），重试安全。
                returned_reviews = future.result()
                returned_indexes = {review["issue_index"] for review in returned_reviews}
                missing_indexes = {int(item["issue_index"]) for item in group} - returned_indexes
                if missing_indexes:
                    errors.append(f"{'误报核验' if verification else '候选复核'}组{group_index}漏回{len(missing_indexes)}项，未返回项保留待核")
                for review in returned_reviews:
                    reviews_by_index[review["issue_index"]] = review
            except AIReviewError as exc:
                if verification:
                    error = f"疑似误报验证组 {group_index} 失败：{exc}"
                    LOGGER.error("ai_verification_group_failed group=%s candidates=%s error=%s", group_index, len(group), exc)
                else:
                    error = f"候选组 {group_index} 复核失败：{exc}"
                    LOGGER.error("ai_review_group_failed group=%s candidates=%s error=%s", group_index, len(group), exc)
                errors.append(error)
            except Exception as exc:
                error = f"AI {'误报验证' if verification else '候选'}组 {group_index} 发生未预期错误：{exc}"
                LOGGER.exception("ai_parallel_group_failed group=%s verification=%s", group_index, verification)
                errors.append(error)

            completed_groups += 1
            percent = percent_start + int((percent_end - percent_start) * completed_groups / len(groups))
            stage = "verification_running" if verification else "review_running"
            message = "AI 正在并行验证疑似误报" if verification else "AI 正在并行复核候选组"
            hint = f"已完成 {completed_groups}/{len(groups)} 组；当前并行数 {max_workers}"
            _emit_progress(progress_callback, {
                "phase": "ai",
                "stage": stage,
                "status": "running",
                "percent": percent,
                "message": message,
                "hint": hint,
                "model": model,
                "candidate_count": candidate_count,
                "reviewed_count": reviewed_count if verification else reviewed_count + len(reviews_by_index),
                "verified_count": len(reviews_by_index) if verification else 0,
                "completed_groups": completed_groups,
                "group_count": len(groups),
                "parallelism": max_workers,
            })
    return reviews_by_index, errors


def _review_parallelism(config: AIConfig) -> int:
    """语义复核的实际并发数（可配，见 AIConfig.review_concurrency）。"""
    try:
        value = int(getattr(config, "review_concurrency", MAX_CONCURRENT_REQUESTS) or MAX_CONCURRENT_REQUESTS)
    except (TypeError, ValueError):
        value = MAX_CONCURRENT_REQUESTS
    return max(1, min(value, 8))


def review_issue_candidates(result: dict[str, Any], progress_callback: ProgressCallback | None = None) -> dict[str, Any]:
    """Review candidates with evidence packets, internal grouping and verification."""
    config = load_config()
    LOGGER.info("ai_review_start enabled=%s model=%s", config.enabled, config.model or "auto")
    base = {
        "enabled": bool(config.enabled and config.api_key),
        "status": "disabled",
        "provider": "openai-compatible",
        "base_url": config.base_url,
        "model": config.model or None,
        "review_mode": "parallel_groups",
        "parallelism": _review_parallelism(config),
        "candidate_count": 0,
        "reviewed_count": 0,
        "group_count": 0,
        "verification_count": 0,
        "verified_count": 0,
        "global_status": "skipped",
        "global_review": {},
        "final_status": "programmatic",
        "finalized_count": 0,
        "final_decision_counts": {},
        "results": [],
        "errors": [],
    }
    issues = result.get("issues", [])
    if not isinstance(issues, list):
        issues = []
    for issue in issues:
        if isinstance(issue, dict):
            issue.update(_default_final_review(issue, candidate=_is_candidate(issue)))
    fallback_counts = Counter(str(issue.get("final_decision", "needs_human")) for issue in issues if isinstance(issue, dict) and issue.get("status") != "resolved")
    base["final_decision_counts"] = dict(sorted(fallback_counts.items()))
    base["actionable_count"] = fallback_counts.get("confirmed_issue", 0) + fallback_counts.get("needs_human", 0)
    if not config.enabled:
        base["reason"] = "AI_ENABLED=false"
        _emit_progress(progress_callback, {"phase": "ai", "stage": "disabled", "status": "disabled", "percent": 90, "message": "AI 复核未启用", "hint": base["reason"], "candidate_count": 0, "reviewed_count": 0})
        return base
    if not config.api_key:
        base["reason"] = "未配置 AI_API_KEY；仅运行本地规则审核"
        _emit_progress(progress_callback, {"phase": "ai", "stage": "disabled", "status": "disabled", "percent": 90, "message": "AI 复核未启用", "hint": base["reason"], "candidate_count": 0, "reviewed_count": 0})
        return base

    indexed = [(idx, issue) for idx, issue in enumerate(issues) if isinstance(issue, dict) and _is_candidate(issue)]
    # 送审上限：AI_REVIEW_MAX_CANDIDATES>0 时按优先级(P0>P1>P2)截断，0=不限制
    max_candidates = int(getattr(config, "max_candidates", 0) or 0)
    if max_candidates and len(indexed) > max_candidates:
        pr_rank = {"P0": 0, "P1": 1, "P2": 2}
        indexed.sort(key=lambda pair: pr_rank.get(str(pair[1].get("priority", "P2")), 2))
        base["candidate_total"] = len(indexed)
        LOGGER.info("ai_candidates_truncated total=%s kept=%s", len(indexed), max_candidates)
        indexed = indexed[:max_candidates]
    base["candidate_count"] = len(indexed)
    if not indexed:
        base["status"] = "completed"
        base["reason"] = "没有需要模型复核的候选项"
        _emit_progress(progress_callback, {"phase": "ai", "stage": "completed", "status": "completed", "percent": 90, "message": "没有需要 AI 复核的候选项", "hint": base["reason"], "candidate_count": 0, "reviewed_count": 0})
        return base

    _emit_progress(progress_callback, {"phase": "ai", "stage": "candidate_scan", "status": "running", "percent": 65, "message": "AI 正在整理候选证据", "hint": f"共发现 {len(indexed)} 个候选，正在生成行级上下文", "candidate_count": len(indexed), "reviewed_count": 0})
    client = OpenAICompatibleClient(config)
    client.audit_requests = True
    try:
        model, model_source = client.choose_model()
    except AIReviewError as exc:
        base["status"] = "error"
        base["errors"].append(str(exc))
        _emit_progress(progress_callback, {"phase": "ai", "stage": "error", "status": "error", "percent": 68, "message": "AI 模型准备失败", "hint": str(exc), "candidate_count": len(indexed), "reviewed_count": 0})
        return base

    base["model"] = model
    base["model_source"] = model_source
    result["_ai_context_limit"] = config.context_rows
    candidates = [_candidate(idx, issue, result) for idx, issue in indexed]
    # The program already has the workbook-wide metrics and structure.  Avoid
    # spending an extra model request on a global summary; send only compact
    # shared context with each evidence group.
    audit_profile = _compact_audit_profile(_build_global_audit_context(result, candidates))
    global_review: dict[str, Any] = {
        "summary": "未单独调用全局模型分析；本次以程序统计、规则证据和分组复核为准。",
        "patterns": [],
        "focus": [],
        "consistency_checks": [],
    }
    base["global_status"] = "skipped_programmatic"
    base["global_review"] = global_review
    groups = _candidate_groups(candidates, config.review_batch_size)
    base["group_count"] = len(groups)
    _emit_progress(progress_callback, {"phase": "ai", "stage": "model_ready", "status": "running", "percent": 66, "message": "AI 开始并行复核候选组", "hint": f"模型：{model}；共 {len(groups)} 组，最多并行 {_review_parallelism(config)} 组", "model": model, "candidate_count": len(indexed), "reviewed_count": 0, "parallelism": _review_parallelism(config)})

    reviews_by_index, group_errors = _review_groups_parallel(
        client,
        model,
        groups,
        verification=False,
        audit_profile=audit_profile,
        global_review=global_review,
        progress_callback=progress_callback,
        candidate_count=len(indexed),
        reviewed_count=0,
        percent_start=67,
        percent_end=82,
        review_concurrency=_review_parallelism(config),
    )
    base["errors"].extend(group_errors)
    for review in reviews_by_index.values():
        _attach_review(issues[review["issue_index"]], review, model)

    base["results"] = [reviews_by_index[index] for index in sorted(reviews_by_index)]
    base["reviewed_count"] = len(base["results"])
    false_positive_indexes = [index for index, review in reviews_by_index.items() if review["decision"] == "likely_false_positive"]
    base["verification_count"] = len(false_positive_indexes)

    if false_positive_indexes:
        verification_candidates = []
        candidates_by_index = {item["issue_index"]: item for item in candidates}
        for index in false_positive_indexes:
            candidate = candidates_by_index[index]
            verification_candidates.append({
                **candidate,
                "initial_review": reviews_by_index[index],
            })
        verification_groups = _candidate_groups(verification_candidates, config.review_batch_size)
        verification_by_index, verification_errors = _review_groups_parallel(
            client,
            model,
            verification_groups,
            verification=True,
            audit_profile=audit_profile,
            global_review=global_review,
            progress_callback=progress_callback,
            candidate_count=len(indexed),
            reviewed_count=len(reviews_by_index),
            percent_start=84,
            percent_end=90,
            review_concurrency=_review_parallelism(config),
        )
        base["errors"].extend(verification_errors)
        for review in verification_by_index.values():
            _attach_review(issues[review["issue_index"]], review, model, verification=True)
        base["verified_count"] = len(verification_by_index)
        for index in false_positive_indexes:
            if index in verification_by_index:
                continue
            issue = issues[index]
            issue["ai_decision"] = "needs_human"
            issue["ai_verification_decision"] = "needs_human"
            issue["ai_verification_confidence"] = 0.0
            issue["ai_verification_evidence"] = "二次核验未返回可解析证据"
            issue["ai_evidence"] = "二次核验未返回可解析证据"
            issue["ai_reason"] = "初审疑似误报，但二次核验未完成，保留人工确认。"
            issue["ai_suggestion"] = "请结合原 Excel 行、父子关系和同级编码人工确认。"

    candidate_indexes = {item["issue_index"] for item in candidates}
    for issue_index, issue in enumerate(issues):
        candidate = issue_index in candidate_indexes
        default_review = _default_final_review(issue, candidate=candidate)
        issue.update(default_review)

    # Final decisions are already deterministic after the initial review and
    # the dedicated false-positive verification.  Do not send all candidates
    # back to the model for a redundant final-merge request.
    base["final_status"] = "programmatic"
    base["finalized_count"] = len(indexed)
    _emit_progress(progress_callback, {"phase": "ai", "stage": "final_adjudication", "status": "completed", "percent": 91, "message": "程序已归并最终审核结论", "hint": "依据规则命中、AI 分组复核和误报验证结果完成归并", "model": model, "candidate_count": len(indexed), "reviewed_count": base["reviewed_count"], "finalized_count": base["finalized_count"]})

    decision_counts = Counter(str(issue.get("final_decision", "needs_human")) for issue in issues if issue.get("status") != "resolved")
    base["final_decision_counts"] = dict(sorted(decision_counts.items()))
    base["actionable_count"] = decision_counts.get("confirmed_issue", 0) + decision_counts.get("needs_human", 0)
    base["unreviewed_count"] = len(indexed) - base["reviewed_count"]
    base["status"] = "completed" if base["reviewed_count"] else "error"
    if base["unreviewed_count"]:
        base["errors"].append(f"{base['unreviewed_count']}项候选未获得可解析结果，保留待核")
    if base["errors"] and base["reviewed_count"]:
        base["status"] = "partial"
    _emit_progress(progress_callback, {"phase": "ai", "stage": "completed", "status": base["status"], "percent": 92, "message": "AI 语义复核完成", "hint": f"已复核 {base['reviewed_count']} 条；疑似误报二次核验 {base['verified_count']} 条", "model": model, "candidate_count": len(indexed), "reviewed_count": base["reviewed_count"]})
    LOGGER.info("ai_review_done status=%s candidates=%s reviewed=%s groups=%s verified=%s", base["status"], base["candidate_count"], base["reviewed_count"], base["group_count"], base["verified_count"])
    return base


def test_ai_connection() -> dict[str, Any]:
    """Make a harmless model ping for the configuration page."""
    config = load_config()
    if not config.enabled:
        return {"ok": False, "status": "disabled", "message": "AI_ENABLED=false"}
    client = OpenAICompatibleClient(config)
    client.audit_requests = True
    model, source = client.choose_model()
    payload = client.review(
        model,
        "你是接口测试助手。只返回 JSON 对象：{\"ok\":true,\"message\":\"连接正常\"}。",
        {"test": "kks-audit-connection"},
    )
    return {"ok": True, "status": "connected", "model": model, "model_source": source, "response": payload}


def summarize_audit(result: dict[str, Any], progress_callback: ProgressCallback | None = None) -> dict[str, Any]:
    """让模型对整份审核结果做总体总结（区别于逐条候选复核）。

    AI 未启用或调用失败时返回 {}，展示层据此省略"AI 总体总结"节。
    """
    config = load_config()
    if not (config.enabled and config.api_key):
        return {}
    if not config.summary_enabled:
        LOGGER.info("ai_summary_skipped reason=AI_SUMMARY_ENABLED=false")
        return {"status": "disabled", "reason": "AI_SUMMARY_ENABLED=false"}
    try:
        client = OpenAICompatibleClient(config)
        client.audit_requests = True
        model, _ = client.choose_model()
        issues = [i for i in (result.get("issues") or []) if isinstance(i, dict)]
        priority_counts: dict[str, int] = {}
        rule_counts: dict[str, int] = {}
        for issue in issues:
            pr = str(issue.get("priority", ""))
            if pr:
                priority_counts[pr] = priority_counts.get(pr, 0) + 1
            rid = str(issue.get("rule_id", ""))
            if rid:
                rule_counts[rid] = rule_counts.get(rid, 0) + 1
        scope = result.get("scope_validation", {}) or {}
        metrics = result.get("metrics", {}) or {}
        ai_review = result.get("ai_review", {}) or {}
        payload = {
            "compliance": result.get("compliance", {}),
            "audit_policy": result.get("audit_policy", {}),
            "evaluation_scope": result.get("evaluation_scope", {}),
            "file": Path(str(result.get("source_file", ""))).name,
            "conclusion": result.get("conclusion", ""),
            "data_rows": result.get("data_rows"),
            "reviewed_rows": result.get("reviewed_rows"),
            "issue_count": result.get("issue_count"),
            "priority_counts": priority_counts,
            "rule_top_counts": dict(sorted(rule_counts.items(), key=lambda kv: -kv[1])[:15]),
            "metrics": {k: metrics.get(k) for k in (
                "orphan_rows", "prefix_mismatch_rows", "converged_to_master_rows",
                "self_ref_rows", "parent_longer_rows", "duplicate_groups",
            ) if metrics.get(k) is not None},
            "incremental_batch": bool(scope.get("incremental_batch")),
            "external_parent_rows": scope.get("external_parent_rows", 0),
            "ai_review": {
                "candidate_count": ai_review.get("candidate_count", 0),
                "reviewed_count": ai_review.get("reviewed_count", 0),
            },
        }
        _emit_progress(progress_callback, {"phase": "ai", "stage": "summary_running", "status": "running", "percent": 93, "message": "AI 正在生成总体总结", "hint": f"模型：{model}"})
        messages = [
            {"role": "system", "content": (
                "你是资深电厂 KKS 编码审核专家。只输出审查结论和资料核对事项，禁止输出建议编码、替换编码或新编码值。强规则冲突不能判为通过，缺少工程资料必须待核，不能宣称全面符合国标。基于程序化审核结果，撰写面向编码管理负责人的总体总结："
                "统计数字由程序单独展示，overall和points只写定性分析，不复述具体数量。evaluation_scope非空时，这是抽样效果检查，规则频次、优先级和compliance只针对抽样，不能作为全工作簿统计；不得由内部自洽推定实际安装位置或工程归属。"
                "overall 为 3-6 句中文总结（整体质量、主要风险、数据口径要点）；"
                "points 为 3-5 条中文要点建议，每条一句话。只返回 JSON："
                '{"overall":"...","points":["...","..."]}'
            )},
            {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
        ]
        req = {"model": model, "messages": messages, "temperature": 0.2, "response_format": {"type": "json_object"}}
        try:
            response = client._request_json("chat/completions", req)
        except AIReviewError as exc:
            if "response_format" not in str(exc).lower() and "json_object" not in str(exc).lower():
                raise
            req.pop("response_format", None)
            response = client._request_json("chat/completions", req)
        content = str(((response.get("choices") or [{}])[0].get("message") or {}).get("content", ""))
        decoded = json.loads(content)
        if not isinstance(decoded, dict) or not str(decoded.get("overall", "")).strip():
            raise AIReviewError("AI 总体总结返回格式异常")
        points = [str(pt).strip() for pt in (decoded.get("points") or []) if str(pt).strip()]
        LOGGER.info("ai_summary_done model=%s points=%s", model, len(points))
        return {"overall": str(decoded["overall"]).strip(), "points": points, "model": model, "status": "completed"}
    except Exception as exc:
        LOGGER.error("ai_summary_failed error=%s", exc)
        return {"status": "failed", "error": str(exc)[:200]}


def summarize_ledger(ledger: list[dict[str, Any]], progress_callback: ProgressCallback | None = None) -> dict[str, Any]:
    """跨文件趋势总结：基于审计台账（历次审核记录）归纳质量趋势与改进建议。

    与 summarize_audit（单文件总体总结）互补——单文件总结回答"这一份怎么样"，
    趋势总结回答"多份之间在往哪个方向走"。记录不足 2 条、AI 未启用或
    AI_TREND_SUMMARY_ENABLED=false 时返回 {}，展示层据此省略该节。
    """
    entries = [e for e in (ledger or []) if isinstance(e, dict)]
    if len(entries) < 2:
        return {}
    config = load_config()
    if not (config.enabled and config.api_key):
        return {}
    if not config.trend_enabled:
        LOGGER.info("ai_trend_skipped reason=AI_TREND_SUMMARY_ENABLED=false")
        return {"status": "disabled", "reason": "AI_TREND_SUMMARY_ENABLED=false"}
    try:
        client = OpenAICompatibleClient(config)
        client.audit_requests = True
        model, _ = client.choose_model()
        total_rows = sum(int(e.get("rows", 0) or 0) for e in entries)
        total_issues = sum(int(e.get("issues", 0) or 0) for e in entries)
        total_p0 = sum(int((e.get("priority") or {}).get("P0", 0) or 0) for e in entries)
        total_p1 = sum(int((e.get("priority") or {}).get("P1", 0) or 0) for e in entries)
        series = []
        for e in entries[-30:]:  # 仅送最近 30 次，控制 token
            rows = int(e.get("rows", 0) or 0)
            issues = int(e.get("issues", 0) or 0)
            series.append({
                "ts": str(e.get("ts", "")),
                "file": str(e.get("file", "")),
                "rows": rows,
                "issues": issues,
                "rate": round(issues / rows * 100, 1) if rows else None,
                "p0": int((e.get("priority") or {}).get("P0", 0) or 0),
                "p1": int((e.get("priority") or {}).get("P1", 0) or 0),
                "p2": int((e.get("priority") or {}).get("P2", 0) or 0),
                "incremental": bool(e.get("incremental")),
                "conclusion": str(e.get("conclusion", ""))[:120],
            })
        payload = {
            "file_count": len(entries),
            "total_rows": total_rows,
            "total_issues": total_issues,
            "avg_rate": round(total_issues / total_rows * 100, 2) if total_rows else None,
            "total_p0": total_p0,
            "total_p1": total_p1,
            "series": series,
        }
        _emit_progress(progress_callback, {"phase": "ai", "stage": "trend_running", "status": "running", "percent": 96, "message": "AI 正在生成跨文件趋势总结", "hint": f"模型：{model}；累计 {len(entries)} 次审核"})
        messages = [
            {"role": "system", "content": (
                "你是资深电厂 KKS 编码质量负责人。以下是多个批次/文件的 KKS 编码审核台账，"
                "请做跨文件趋势总结，供编码管理负责人决策："
                "overall 为 3-6 句中文（整体质量趋势、批次间变化、反复出现的系统性问题）；"
                "points 为 3-5 条中文改进建议，每条一句话，指向管理动作而非具体某一行数据。"
                "只返回 JSON：{\"overall\":\"...\",\"points\":[\"...\",\"...\"]}"
            )},
            {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
        ]
        req = {"model": model, "messages": messages, "temperature": 0.2, "response_format": {"type": "json_object"}}
        try:
            response = client._request_json("chat/completions", req)
        except AIReviewError as exc:
            if "response_format" not in str(exc).lower() and "json_object" not in str(exc).lower():
                raise
            req.pop("response_format", None)
            response = client._request_json("chat/completions", req)
        content = str(((response.get("choices") or [{}])[0].get("message") or {}).get("content", ""))
        decoded = json.loads(content)
        if not isinstance(decoded, dict) or not str(decoded.get("overall", "")).strip():
            raise AIReviewError("AI 趋势总结返回格式异常")
        points = [str(pt).strip() for pt in (decoded.get("points") or []) if str(pt).strip()]
        LOGGER.info("ai_trend_done model=%s files=%s points=%s", model, len(entries), len(points))
        return {
            "overall": str(decoded["overall"]).strip(),
            "points": points,
            "model": model,
            "status": "completed",
            "ts": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "file_count": len(entries),
        }
    except Exception as exc:
        LOGGER.error("ai_trend_failed error=%s", exc)
        return {"status": "failed", "error": str(exc)[:200]}


# 🔴 模块级调用点（**不要删**）：导入本模块即把 `.env.local` 读进进程环境。
#    此前本函数定义了却从未被调用，导致：
#      · `KKS_ADMIN_TOKEN` 不生效 → 管理接口一律 403（fail-closed）
#      · AI 配置退回默认值 → 界面上"勾了 AI"也跑不起来
#    采用 `setdefault` 语义（函数内 `if name not in os.environ`），**不覆盖**真实环境变量，
#    因此 CI/测试里显式设置的变量仍然优先。
_load_local_env()
