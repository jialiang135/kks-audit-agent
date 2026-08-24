# -*- coding: utf-8 -*-
"""Runtime paths, editable configuration, Skill text and logging."""
from __future__ import annotations

import copy
import json
import logging
import logging.handlers
import os
import shutil
import sys
import time
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Any


def app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_ROOT = app_root()
CONFIG_DIR = APP_ROOT / "config"
APP_CONFIG_PATH = CONFIG_DIR / "app_config.json"
FORMAL_SKILL_PATH = APP_ROOT / "kks-audit" / "SKILL.md"
SKILL_REFERENCES_DIR = APP_ROOT / "kks-audit" / "references"
SKILL_RULE_SCRIPT_PATH = APP_ROOT / "kks-audit" / "scripts" / "audit_template.py"
SKILL_BACKUP_DIR = APP_ROOT / "skill_backups"
LOG_DIR = APP_ROOT / "logs"
LOG_PATH = LOG_DIR / "kks-audit.log"

DEFAULT_APP_CONFIG: dict[str, Any] = {
    "ai": {
        "enabled": False,
        "base_url": "",
        "api_key": "",
        "model": "",
        "timeout_seconds": 90,
    },
    "audit": {
        "max_upload_mb": 50,
        "runs_dir": "runs",
    },
}


def _merge(base: dict[str, Any], extra: dict[str, Any]) -> dict[str, Any]:
    result = copy.deepcopy(base)
    for key, value in extra.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = _merge(result[key], value)
        else:
            result[key] = value
    return result


def load_app_config() -> dict[str, Any]:
    config = copy.deepcopy(DEFAULT_APP_CONFIG)
    if not APP_CONFIG_PATH.is_file():
        return config
    try:
        payload = json.loads(APP_CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return config
    return _merge(config, payload) if isinstance(payload, dict) else config


def save_app_config(patch: dict[str, Any]) -> dict[str, Any]:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    config = _merge(load_app_config(), patch)
    temp_path = APP_CONFIG_PATH.with_suffix(".tmp")
    temp_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
    temp_path.replace(APP_CONFIG_PATH)
    return config


def _skill_context_documents() -> list[tuple[str, str]]:
    documents: list[tuple[str, str]] = []
    if FORMAL_SKILL_PATH.is_file():
        documents.append((f"正式 Skill：{FORMAL_SKILL_PATH.relative_to(APP_ROOT)}", FORMAL_SKILL_PATH.read_text(encoding="utf-8")))
    if SKILL_REFERENCES_DIR.is_dir():
        for path in sorted(SKILL_REFERENCES_DIR.glob("*.md")):
            documents.append((f"审核参考资料：{path.relative_to(APP_ROOT)}", path.read_text(encoding="utf-8")))
    return documents


def load_skill_context(
    *,
    rule_ids: set[str] | None = None,
    categories: set[str] | None = None,
    keywords: set[str] | None = None,
    max_chars: int | None = None,
) -> str:
    """Load Skill text, optionally selecting only relevant sections.

    The executable rule module is intentionally not copied into the prompt. It
    is run by ``run_audit.py``; the model receives the rule finding and the
    matching human-readable guidance instead.
    """
    documents = _skill_context_documents()
    terms = {
        str(value).strip().casefold()
        for value in (*sorted(rule_ids or set()), *sorted(categories or set()), *sorted(keywords or set()))
        if str(value).strip()
    }
    if not terms:
        sections = [f"【{title}】\n{content}" for title, content in documents]
        return "\n\n".join(sections)

    sections: list[str] = []
    for title, content in documents:
        lines = content.splitlines()
        def match_score(line: str) -> int:
            folded = line.casefold()
            score = 0
            for term in rule_ids or set():
                if str(term).casefold() in folded:
                    score = max(score, 100)
            for term in categories or set():
                if str(term).casefold() in folded:
                    score = max(score, 80)
            for term in keywords or set():
                if str(term).casefold() in folded:
                    score = max(score, min(60, len(str(term))))
            return score

        matches = [(match_score(line), index) for index, line in enumerate(lines) if any(term in line.casefold() for term in terms)]
        matches = [index for _, index in sorted(matches, key=lambda item: (-item[0], item[1]))[:12]]
        matches.sort()
        if not matches:
            continue
        ranges: list[tuple[int, int]] = []
        for index in matches:
            start = max(0, index - 2)
            end = min(len(lines), index + 7)
            if start and lines[start - 1].lstrip().startswith("#"):
                start -= 1
            if ranges and start <= ranges[-1][1] + 1:
                ranges[-1] = (ranges[-1][0], max(ranges[-1][1], end))
            else:
                ranges.append((start, end))
        excerpt = "\n".join("\n".join(lines[start:end]) for start, end in ranges)
        sections.append(f"【{title}｜相关片段】\n{excerpt}")

    if not sections and documents:
        title, content = documents[0]
        sections.append(f"【{title}｜摘要】\n{content[:2000]}")
    context = "\n\n".join(sections)
    return context if max_chars is None else context[:max_chars]


def skill_documents() -> list[dict[str, Any]]:
    """Return the Skill files that can be viewed or edited from the admin page."""
    documents: list[dict[str, Any]] = []
    candidates = [("kks-audit/SKILL.md", FORMAL_SKILL_PATH, "formal")]
    if SKILL_REFERENCES_DIR.is_dir():
        candidates.extend(
            (f"kks-audit/references/{path.name}", path, "reference")
            for path in sorted(SKILL_REFERENCES_DIR.glob("*.md"))
        )
    candidates.append(("kks-audit/scripts/audit_template.py", SKILL_RULE_SCRIPT_PATH, "executable_rule"))
    for relative, path, kind in candidates:
        if not path.is_file():
            continue
        documents.append(
            {
                "path": relative,
                "kind": kind,
                "editable": True,
                "content": path.read_text(encoding="utf-8"),
            }
        )
    return documents


def skill_target_path(relative_path: str) -> Path:
    """Resolve one allow-listed Skill path; reject traversal and arbitrary files."""
    normalized = str(relative_path or "").replace("\\", "/").strip().lstrip("/")
    if normalized == "kks-audit/SKILL.md":
        return FORMAL_SKILL_PATH
    if normalized == "kks-audit/scripts/audit_template.py":
        return SKILL_RULE_SCRIPT_PATH
    if normalized.startswith("kks-audit/references/") and normalized.count("/") == 2 and normalized.endswith(".md"):
        name = normalized.rsplit("/", 1)[-1]
        if name and name not in {".", ".."}:
            return SKILL_REFERENCES_DIR / name
    raise ValueError("不允许修改该文件，只能修改正式 Skill、references/*.md 或 audit_template.py")


def _backup_skill_file(path: Path) -> None:
    if not path.is_file():
        return
    stamp = time.strftime("%Y%m%dT%H%M%S") + f"_{time.time_ns() % 1_000_000:06d}"
    backup = SKILL_BACKUP_DIR / stamp / path.relative_to(APP_ROOT)
    backup.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(path, backup)


def save_skill_document(relative_path: str, content: str) -> dict[str, Any]:
    path = skill_target_path(relative_path)
    if not isinstance(content, str):
        raise ValueError("Skill 内容必须是文本")
    if len(content.encode("utf-8")) > 5 * 1024 * 1024:
        raise ValueError("单个 Skill 文件不能超过 5 MB")
    _backup_skill_file(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_suffix(path.suffix + ".tmp")
    temp_path.write_text(content, encoding="utf-8")
    temp_path.replace(path)
    logging.getLogger("kks-audit").info("skill_updated path=%s", relative_path)
    return {"path": relative_path, "requires_restart": relative_path.endswith("audit_template.py")}


def install_skill_zip(data: bytes, filename: str = "skill.zip") -> dict[str, Any]:
    """Install only allow-listed Skill files from a zip and keep backups."""
    if len(data) > 10 * 1024 * 1024:
        raise ValueError("Skill 压缩包不能超过 10 MB")
    staged: list[tuple[str, bytes]] = []
    with zipfile.ZipFile(BytesIO(data)) as archive:
        for info in archive.infolist():
            if info.is_dir():
                continue
            raw_name = info.filename.replace("\\", "/").lstrip("/")
            if raw_name.startswith("kks-audit/"):
                relative = raw_name
            else:
                relative = f"kks-audit/{raw_name}"
            target = skill_target_path(relative)
            if info.file_size > 5 * 1024 * 1024:
                raise ValueError(f"Skill 文件过大：{raw_name}")
            staged.append((relative, archive.read(info)))
            if target == FORMAL_SKILL_PATH:
                pass
    if not staged:
        raise ValueError("压缩包中没有可安装的 Skill 文件")
    total = sum(len(content) for _, content in staged)
    if total > 10 * 1024 * 1024:
        raise ValueError("Skill 压缩包中文件总大小不能超过 10 MB")
    updated: list[dict[str, Any]] = []
    for relative, content in staged:
        path = skill_target_path(relative)
        _backup_skill_file(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        temp_path = path.with_suffix(path.suffix + ".tmp")
        temp_path.write_bytes(content)
        temp_path.replace(path)
        updated.append({"path": relative, "requires_restart": relative.endswith("audit_template.py")})
    logging.getLogger("kks-audit").info("skill_zip_installed file=%s paths=%s", filename, ",".join(item["path"] for item in updated))
    return {"updated": updated, "requires_restart": any(item["requires_restart"] for item in updated)}


def _read_local_env() -> dict[str, str]:
    path = APP_ROOT / ".env.local"
    if not path.is_file():
        return {}
    values: dict[str, str] = {}
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        name, value = line.split("=", 1)
        name, value = name.strip(), value.strip().strip('"\'')
        if name.isidentifier():
            values[name] = value
    return values


def environment_value(name: str, config_value: Any = "") -> str:
    """Process environment > app_config.json > .env.local."""
    if name in os.environ:
        return os.environ[name]
    config_names = {
        "AI_ENABLED": ("enabled",),
        "AI_BASE_URL": ("base_url",),
        "AI_API_KEY": ("api_key",),
        "AI_MODEL": ("model",),
        "AI_TIMEOUT_SECONDS": ("timeout_seconds",),
    }
    local = _read_local_env()
    if config_value not in (None, ""):
        return str(config_value)
    return local.get(name, "")


def configure_logging() -> logging.Logger:
    logger = logging.getLogger("kks-audit")
    logger.setLevel(logging.INFO)
    logger.propagate = False
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    if not any(isinstance(handler, logging.handlers.RotatingFileHandler) for handler in logger.handlers):
        file_handler = logging.handlers.RotatingFileHandler(
            LOG_PATH, maxBytes=5 * 1024 * 1024, backupCount=5, encoding="utf-8"
        )
        file_handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s %(message)s"))
        logger.addHandler(file_handler)
    if not any(isinstance(handler, logging.StreamHandler) and not isinstance(handler, logging.FileHandler) for handler in logger.handlers):
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        logger.addHandler(console_handler)
    return logger


def tail_log(limit: int = 300) -> list[str]:
    if not LOG_PATH.is_file():
        return []
    try:
        lines = LOG_PATH.read_text(encoding="utf-8", errors="replace").splitlines()
    except OSError:
        return []
    return lines[-max(1, min(limit, 2000)) :]


def mask_secret(value: str) -> str:
    if not value:
        return ""
    if len(value) <= 8:
        return "已配置"
    return f"{value[:3]}***{value[-4:]}"
