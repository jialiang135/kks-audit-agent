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
import urllib.error
import urllib.parse
import urllib.request
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

from app_runtime import environment_value, load_app_config, load_skill_context


DEFAULT_BASE_URL = "https://www.ai.atyou.cn"
DEFAULT_TIMEOUT_SECONDS = 90
LOGGER = logging.getLogger("kks-audit.ai")
ProgressCallback = Callable[[dict[str, Any]], None]
INTERNAL_GROUP_SIZE = 12
REQUEST_ATTEMPTS = 2
RETRY_BACKOFF_SECONDS = 2
MAX_CONTEXT_ROWS = 12
MAX_ADJACENT_ROWS = 5
MAX_GLOBAL_TEXT = 1800
MAX_CONCURRENT_REQUESTS = 3
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
    "KKS-13/16": {"旧码", "迁移", "历史"},
    "KKS-14": {"系统字母", "重分配"},
    "KKS-15": {"设备字母", "位置", "变化"},
    "KKS-16": {"旧格式", "迁移"},
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
    """Load an optional local-only config without overriding process env vars."""
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
    app_config = load_app_config()
    ai = app_config.get("ai", {}) if isinstance(app_config.get("ai"), dict) else {}
    api_key = environment_value("AI_API_KEY", ai.get("api_key", "")).strip()
    enabled = _parse_bool(environment_value("AI_ENABLED", ai.get("enabled", "")), bool(api_key))
    raw_base_url = environment_value("AI_BASE_URL", ai.get("base_url", "")).strip()
    return AIConfig(
        base_url=normalize_base_url(raw_base_url) if raw_base_url else "",
        api_key=api_key,
        model=environment_value("AI_MODEL", ai.get("model", "")).strip(),
        timeout_seconds=_parse_int(environment_value("AI_TIMEOUT_SECONDS", ai.get("timeout_seconds", DEFAULT_TIMEOUT_SECONDS)), DEFAULT_TIMEOUT_SECONDS, 10, 300),
        enabled=enabled,
    )


class OpenAICompatibleClient:
    def __init__(self, config: AIConfig) -> None:
        if not config.api_key:
            raise AIReviewError("未配置 AI_API_KEY")
        if not config.base_url:
            raise AIReviewError("未配置 AI_BASE_URL")
        self.config = config

    def _request_json(self, path: str, payload: dict[str, Any] | None = None) -> dict[str, Any]:
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


SYSTEM_PROMPT = """你是电厂 KKS 编码审核智能体的二次复核员。
规则引擎已经给出了候选问题。你只能依据候选证据包和本组相关 Skill 片段判断，不能凭空补充表外事实。
证据包中的当前行、父级、子级、同级编码和相邻行必须结合起来看；父子关系以 Excel 的显式父级列为准，不能用固定前 N 位臆测变长层级。
不要修改原 Excel，不要自动改码，不要自动关闭问题；证据不足时必须返回 needs_human。
硬性语法问题通常保留为 confirmed_issue；变长码、历史映射、机组语义、名称异常要先检查证据包中的上下文，再判断是否可能误报。
每条结果必须返回结论、证据、原因和建议。证据只能引用证据包中真实存在的行号、编码、名称或规则字段，不能写“根据经验”作为唯一证据。
请只返回 JSON 对象，格式为：
{"reviews":[{"issue_index":0,"decision":"confirmed_issue|likely_false_positive|needs_human","priority":"P0|P1|P2","confidence":0.0,"evidence":"引用证据包中的具体行号/编码/关系","reason":"中文理由","suggestion":"中文建议"}]}
confidence 必须是 0 到 1 的数字。每个 issue_index 最多返回一次。"""

GLOBAL_REVIEW_PROMPT = """你是电厂 KKS 编码审核智能体的全局分析员。
你将收到一份 Excel 的结构画像、规则命中统计和候选问题摘要。先从整个文件的角度识别专业范围、层级组织、命名习惯、重复模式和可能的误报来源，帮助后续逐项复核。
不要凭空补充 Excel 外的事实，不要直接修改问题，也不要把某一条候选问题当成最终结论。
请只返回 JSON 对象，格式为：
{"global_review":{"summary":"整体审核观察","patterns":["全局模式或异常"],"focus":[{"rule_id":"规则编号","reason":"为什么需要重点核对"}],"consistency_checks":["建议后续重点验证的跨行关系"]}}
所有内容必须来自输入的结构画像、规则统计或候选摘要。"""

FINAL_ADJUDICATION_PROMPT = """你是电厂 KKS 编码审核智能体的最终归并员。
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
            "excel_row", "kks_code", "parent_code", "old_code", "name", "unit", "device_type",
            "raw_kks", "raw_name",
        )
        if record.get(key, "") not in (None, "")
    }


def _context_records(result: dict[str, Any]) -> list[dict[str, Any]]:
    context = result.get("_ai_context", {})
    records = context.get("records", []) if isinstance(context, dict) else []
    return [record for record in records if isinstance(record, dict)]


def _build_evidence_packet(issue_index: int, issue: dict[str, Any], result: dict[str, Any]) -> dict[str, Any]:
    records = _context_records(result)
    issue_row = str(issue.get("excel_row", ""))
    issue_code = str(issue.get("kks_code", ""))
    current = next((record for record in records if str(record.get("excel_row", "")) == issue_row and issue_row), None)
    if current is None and issue_code:
        current = next((record for record in records if str(record.get("kks_code", "")) == issue_code), None)
    if current is None:
        current = {key: issue.get(key, "") for key in ("excel_row", "kks_code", "parent_code", "old_code", "name")}

    current_code = str(current.get("kks_code") or issue_code)
    parent_code = str(current.get("parent_code") or issue.get("parent_code") or "")
    records_by_code: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for record in records:
        code = str(record.get("kks_code", ""))
        if code:
            records_by_code[code].append(record)
    parent_rows = records_by_code.get(parent_code, [])[:1] if parent_code else []
    child_rows = [record for record in records if current_code and str(record.get("parent_code", "")) == current_code]
    sibling_rows = [
        record for record in records
        if str(record.get("parent_code", "")) == parent_code
        and str(record.get("excel_row", "")) != str(current.get("excel_row", ""))
    ] if parent_code else []
    try:
        current_row_number = int(current.get("excel_row"))
    except (TypeError, ValueError):
        current_row_number = None
    adjacent_rows = []
    if current_row_number is not None:
        for record in records:
            try:
                row_number = int(record.get("excel_row"))
            except (TypeError, ValueError):
                continue
            if row_number != current_row_number and abs(row_number - current_row_number) <= 2:
                adjacent_rows.append(record)

    def capped(values: list[dict[str, Any]], limit: int = MAX_CONTEXT_ROWS) -> list[dict[str, Any]]:
        return [_row_view(value) for value in values[:limit]]

    return {
        "current_row": _row_view(current),
        "parent_rows": capped(parent_rows, 1),
        "child_rows": capped(child_rows),
        "same_level_rows": capped(sibling_rows),
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


def _candidate_groups(candidates: list[dict[str, Any]]) -> list[list[dict[str, Any]]]:
    grouped: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for candidate in candidates:
        key = (str(candidate.get("rule_id", "")), str(candidate.get("category", "")))
        grouped[key].append(candidate)
    groups: list[list[dict[str, Any]]] = []
    for key in sorted(grouped):
        values = sorted(grouped[key], key=lambda item: int(item.get("issue_index", 0)))
        for start in range(0, len(values), INTERNAL_GROUP_SIZE):
            groups.append(values[start:start + INTERNAL_GROUP_SIZE])
    return groups


def _review_group(
    client: OpenAICompatibleClient,
    model: str,
    group: list[dict[str, Any]],
    *,
    verification: bool = False,
    audit_profile: dict[str, Any] | None = None,
    global_review: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    rule_ids = {str(item.get("rule_id", "")) for item in group}
    categories = {str(item.get("category", "")) for item in group}
    payload = client.review(
        model,
        effective_system_prompt(rule_ids, categories, verification=verification),
        {
            "review_type": "false_positive_verification" if verification else "candidate_review",
            "audit_profile": _compact_audit_profile(audit_profile or {}),
            "global_review": global_review or {},
            "candidates": group,
        },
    )
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
        reviews.append(review)
    return reviews


def _default_final_review(issue: dict[str, Any], *, candidate: bool) -> dict[str, Any]:
    ai_decision = str(issue.get("ai_decision", ""))
    verification = str(issue.get("ai_verification_decision", ""))
    if not candidate:
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
        "final_summary": str(issue.get("message", "")),
        "final_evidence": str(issue.get("evidence", "")),
        "final_reason": "程序规则命中；尚未获得可替代规则结论的 AI 归并结果。" if not candidate else "AI 复核未形成可直接采信的最终结论。",
        "final_suggestion": str(issue.get("suggestion", "人工确认")),
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
) -> tuple[dict[int, dict[str, Any]], list[str]]:
    """Review independent groups concurrently with a small fixed worker pool."""
    reviews_by_index: dict[int, dict[str, Any]] = {}
    errors: list[str] = []
    if not groups:
        return reviews_by_index, errors

    completed_groups = 0
    max_workers = min(MAX_CONCURRENT_REQUESTS, len(groups))
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
                for review in future.result():
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
                "reviewed_count": reviewed_count + len(reviews_by_index),
                "completed_groups": completed_groups,
                "group_count": len(groups),
                "parallelism": max_workers,
            })
    return reviews_by_index, errors


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
        "parallelism": MAX_CONCURRENT_REQUESTS,
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
    base["candidate_count"] = len(indexed)
    if not indexed:
        base["status"] = "completed"
        base["reason"] = "没有需要模型复核的候选项"
        _emit_progress(progress_callback, {"phase": "ai", "stage": "completed", "status": "completed", "percent": 90, "message": "没有需要 AI 复核的候选项", "hint": base["reason"], "candidate_count": 0, "reviewed_count": 0})
        return base

    _emit_progress(progress_callback, {"phase": "ai", "stage": "candidate_scan", "status": "running", "percent": 65, "message": "AI 正在整理候选证据", "hint": f"共发现 {len(indexed)} 个候选，正在生成行级上下文", "candidate_count": len(indexed), "reviewed_count": 0})
    client = OpenAICompatibleClient(config)
    try:
        model, model_source = client.choose_model()
    except AIReviewError as exc:
        base["status"] = "error"
        base["errors"].append(str(exc))
        _emit_progress(progress_callback, {"phase": "ai", "stage": "error", "status": "error", "percent": 68, "message": "AI 模型准备失败", "hint": str(exc), "candidate_count": len(indexed), "reviewed_count": 0})
        return base

    base["model"] = model
    base["model_source"] = model_source
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
    groups = _candidate_groups(candidates)
    base["group_count"] = len(groups)
    _emit_progress(progress_callback, {"phase": "ai", "stage": "model_ready", "status": "running", "percent": 66, "message": "AI 开始并行复核候选组", "hint": f"模型：{model}；共 {len(groups)} 组，最多并行 {MAX_CONCURRENT_REQUESTS} 组", "model": model, "candidate_count": len(indexed), "reviewed_count": 0, "parallelism": MAX_CONCURRENT_REQUESTS})

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
        for index in false_positive_indexes:
            candidate = next(item for item in candidates if item["issue_index"] == index)
            verification_candidates.append({
                **candidate,
                "initial_review": reviews_by_index[index],
            })
        verification_groups = _candidate_groups(verification_candidates)
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

    for issue_index, issue in enumerate(issues):
        candidate = any(item["issue_index"] == issue_index for item in candidates)
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
    base["status"] = "completed" if base["reviewed_count"] else "error"
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
    model, source = client.choose_model()
    payload = client.review(
        model,
        "你是接口测试助手。只返回 JSON 对象：{\"ok\":true,\"message\":\"连接正常\"}。",
        {"test": "kks-audit-connection"},
    )
    return {"ok": True, "status": "connected", "model": model, "model_source": source, "response": payload}
