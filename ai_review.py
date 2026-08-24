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
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

from app_runtime import environment_value, load_app_config, load_skill_context


DEFAULT_BASE_URL = "https://www.ai.atyou.cn"
DEFAULT_TIMEOUT_SECONDS = 90
LOGGER = logging.getLogger("kks-audit.ai")
ProgressCallback = Callable[[dict[str, Any]], None]


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
        try:
            with urllib.request.urlopen(request, timeout=self.config.timeout_seconds) as response:
                raw = response.read().decode("utf-8", errors="replace")
            LOGGER.info("model_request path=%s status=200 bytes=%s", path, len(raw))
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")[:500]
            raise AIReviewError(f"模型接口 HTTP {exc.code}：{detail}") from exc
        except urllib.error.URLError as exc:
            raise AIReviewError(f"模型接口连接失败：{exc.reason}") from exc
        except TimeoutError as exc:
            raise AIReviewError("模型接口请求超时") from exc
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
规则引擎已经给出了候选问题。你只判断提供的候选证据，不能凭空补充表外事实。
不要修改原 Excel，不要自动改码，不要自动关闭问题；证据不足时必须返回 needs_human。
硬性语法问题通常保留为 confirmed_issue；变长码、历史映射、机组语义、名称异常可以结合给定字段判断是否可能误报。
请只返回 JSON 对象，格式为：
{"reviews":[{"issue_index":0,"decision":"confirmed_issue|likely_false_positive|needs_human","priority":"P0|P1|P2","confidence":0.0,"reason":"中文理由","suggestion":"中文建议"}]}
confidence 必须是 0 到 1 的数字。每个 issue_index 最多返回一次。"""


def effective_system_prompt() -> str:
    formal_context = load_skill_context().strip()
    prompt = SYSTEM_PROMPT
    if formal_context:
        prompt += "\n\n以下是正式 Skill 和审核参考资料。正式 Skill 用于说明审核范围；参考资料只能在候选证据相符时辅助理解，不得把其中的厂站案例、固定码表或经验直接当成当前 Excel 的硬规则：\n" + formal_context
    return prompt


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


def _candidate(issue_index: int, issue: dict[str, Any]) -> dict[str, Any]:
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
        "reason": str(item.get("reason", "模型未提供理由"))[:1000],
        "suggestion": str(item.get("suggestion", "保留人工确认"))[:1000],
    }


def _emit_progress(callback: ProgressCallback | None, event: dict[str, Any]) -> None:
    if callback is None:
        return
    try:
        callback(event)
    except Exception as exc:  # status reporting must not fail the audit
        LOGGER.warning("ai_progress_callback_failed error=%s", exc)


def review_issue_candidates(result: dict[str, Any], progress_callback: ProgressCallback | None = None) -> dict[str, Any]:
    """Review every semantic candidate in one model request and attach results."""
    config = load_config()
    LOGGER.info("ai_review_start enabled=%s model=%s", config.enabled, config.model or "auto")
    base = {
        "enabled": bool(config.enabled and config.api_key),
        "status": "disabled",
        "provider": "openai-compatible",
        "base_url": config.base_url,
        "model": config.model or None,
        "candidate_count": 0,
        "reviewed_count": 0,
        "results": [],
        "errors": [],
    }
    if not config.enabled:
        base["reason"] = "AI_ENABLED=false"
        _emit_progress(progress_callback, {"phase": "ai", "stage": "disabled", "status": "disabled", "percent": 90, "message": "AI 复核未启用", "hint": base["reason"], "candidate_count": 0, "reviewed_count": 0})
        return base
    if not config.api_key:
        base["reason"] = "未配置 AI_API_KEY；仅运行本地规则审核"
        _emit_progress(progress_callback, {"phase": "ai", "stage": "disabled", "status": "disabled", "percent": 90, "message": "AI 复核未启用", "hint": base["reason"], "candidate_count": 0, "reviewed_count": 0})
        return base

    issues = result.get("issues", [])
    indexed = [(idx, issue) for idx, issue in enumerate(issues) if isinstance(issue, dict) and _is_candidate(issue)]
    base["candidate_count"] = len(indexed)
    if not indexed:
        base["status"] = "completed"
        base["reason"] = "没有需要模型复核的候选项"
        _emit_progress(progress_callback, {"phase": "ai", "stage": "completed", "status": "completed", "percent": 90, "message": "没有需要 AI 复核的候选项", "hint": base["reason"], "candidate_count": 0, "reviewed_count": 0})
        return base

    _emit_progress(progress_callback, {"phase": "ai", "stage": "candidate_scan", "status": "running", "percent": 65, "message": "AI 正在准备全部候选", "hint": f"共发现 {len(indexed)} 个候选，将全部提交模型", "candidate_count": len(indexed), "reviewed_count": 0})
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
    candidates = [_candidate(idx, issue) for idx, issue in indexed]
    allowed = {item["issue_index"] for item in candidates}
    _emit_progress(progress_callback, {"phase": "ai", "stage": "model_ready", "status": "running", "percent": 68, "message": "AI 模型已连接，开始复核全部候选", "hint": f"模型：{model}；候选 {len(candidates)} 条", "model": model, "candidate_count": len(indexed), "reviewed_count": 0})
    _emit_progress(progress_callback, {"phase": "ai", "stage": "review_running", "status": "running", "percent": 70, "message": "AI 正在复核全部候选", "hint": f"正在提交 {len(candidates)} 条候选证据", "model": model, "candidate_count": len(indexed), "reviewed_count": 0})
    try:
        payload = client.review(model, effective_system_prompt(), {"candidates": candidates})
        raw_reviews = payload.get("reviews", [])
        if not isinstance(raw_reviews, list):
            raise AIReviewError("模型返回的 reviews 不是数组")
        for item in raw_reviews:
            review = _safe_review(item, allowed)
            if review is None:
                continue
            issue = issues[review["issue_index"]]
            issue["ai_decision"] = review["decision"]
            issue["ai_confidence"] = review["confidence"]
            issue["ai_reason"] = review["reason"]
            issue["ai_suggestion"] = review["suggestion"]
            issue["ai_model"] = model
            base["results"].append(review)
        _emit_progress(progress_callback, {"phase": "ai", "stage": "review_completed", "status": "running", "percent": 90, "message": "AI 全部候选复核完成", "hint": f"已提交 {len(candidates)} 条候选；得到 {len(base['results'])} 条复核结果", "model": model, "candidate_count": len(indexed), "reviewed_count": len(base["results"])})
    except AIReviewError as exc:
        base["errors"].append(f"全部候选复核失败：{exc}")
        LOGGER.error("ai_review_failed candidates=%s error=%s", len(candidates), exc)
        _emit_progress(progress_callback, {"phase": "ai", "stage": "review_error", "status": "running", "percent": 90, "message": "AI 全部候选复核失败", "hint": str(exc), "model": model, "candidate_count": len(indexed), "reviewed_count": len(base["results"])})

    base["reviewed_count"] = len(base["results"])
    base["status"] = "completed" if base["reviewed_count"] else "error"
    if base["errors"] and base["reviewed_count"]:
        base["status"] = "partial"
    _emit_progress(progress_callback, {"phase": "ai", "stage": "completed", "status": base["status"], "percent": 92, "message": "AI 语义复核完成", "hint": f"已得到 {base['reviewed_count']} 条复核结果", "model": model, "candidate_count": len(indexed), "reviewed_count": base["reviewed_count"]})
    LOGGER.info("ai_review_done status=%s candidates=%s reviewed=%s", base["status"], base["candidate_count"], base["reviewed_count"])
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
