# -*- coding: utf-8 -*-
"""KKS Excel audit runner.

The ``kks-audit`` Skill is the executable rule source. Its reusable rule
template is loaded at runtime and every rule result is converted into an
auditable issue row instead of being left as documentation-only code.
"""
from __future__ import annotations

import argparse
import html
import importlib.util
import json
import logging
import re
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Callable

import openpyxl
from ai_review import review_issue_candidates
from app_runtime import APP_ROOT, configure_logging
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter


VERSION = "0.2.0"
LOGGER = logging.getLogger("kks-audit")
ProgressCallback = Callable[[dict[str, Any]], None]
ROOT_PARENTS = {"", "-1"}
OLD_PREFIX = {"50": "05", "60": "06", "L0": "61", "J0": "61"}
COMMON_UNITS = {"00", "61", "L0"}
ALLOWED_CODE = re.compile(r"^[A-Z0-9\-/~]+$")
NAME_UNIT = re.compile(r"([一二三四五六七八九十\d])\s*号\s*(机|炉|机组)")
CN_NUM = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}


def output_artifact_names(input_path: Path) -> tuple[str, str]:
    """Return user-facing report names derived from the uploaded workbook."""
    stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", input_path.stem).strip(" .")
    stem = (stem or "审核文件")[:120]
    return f"{stem}_审核报告.html", f"{stem}_问题清单.xlsx"


def _emit_progress(callback: ProgressCallback | None, event: dict[str, Any]) -> None:
    if callback is None:
        return
    try:
        callback(event)
    except Exception as exc:  # progress reporting must not fail the audit
        LOGGER.warning("progress_callback_failed error=%s", exc)


def _load_executable_skill_rules():
    """Load the formal kks-audit rule module, including frozen EXE paths."""
    template_path = APP_ROOT / "kks-audit" / "scripts" / "audit_template.py"
    if not template_path.is_file():
        raise RuntimeError(f"正式 Skill 规则文件不存在：{template_path}")
    spec = importlib.util.spec_from_file_location("kks_audit_skill_rules", template_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法加载正式 Skill 规则文件：{template_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


SKILL_RULES = _load_executable_skill_rules()
SKILL_RULE_SOURCE = "kks-audit/SKILL.md"


def text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def nonempty(row: tuple[Any, ...]) -> bool:
    return any(text(v) for v in row)


def header_score(row: tuple[Any, ...]) -> int:
    s = "|".join(text(v) for v in row)
    return sum(token in s for token in ("KKS", "名称", "父级", "上级", "机组"))


def detect_header(rows: list[tuple[Any, ...]]) -> int:
    candidates = [(header_score(row), i) for i, row in enumerate(rows[:30])]
    score, idx = max(candidates, default=(0, 0))
    if score < 2:
        raise ValueError("未找到包含 KKS/名称/父级等关键字段的表头")
    return idx


def find_column(header: tuple[Any, ...], kind: str) -> int | None:
    labels = [text(v) for v in header]

    def first(predicate):
        for i, label in enumerate(labels):
            if predicate(label):
                return i
        return None

    if kind == "kks":
        return first(lambda s: "KKS" in s and not any(x in s for x in ("原", "旧", "上级", "父级", "设计")))
    if kind == "old":
        return first(lambda s: "KKS" in s and any(x in s for x in ("原", "旧")))
    if kind == "parent":
        return first(lambda s: "父级" in s or "上级" in s)
    if kind == "name":
        preferred = first(lambda s: ("修改后" in s or "新名称" in s or "新设备名称" in s) and "原" not in s)
        if preferred is not None:
            return preferred
        return first(lambda s: "名称" in s and not any(x in s for x in ("原", "机组", "系统", "上级")))
    if kind == "unit":
        return first(lambda s: ("机组" in s or "机组号" in s) and "名称" not in s)
    if kind == "dtype":
        return first(lambda s: "设备类型" in s or s == "类型")
    if kind == "profession":
        return first(lambda s: "专业" in s or "所属专业" in s)
    return None


def inspect_sheets(wb) -> list[dict[str, Any]]:
    result = []
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        nonempty_rows = [i + 1 for i, row in enumerate(rows) if nonempty(row)]
        result.append({
            "sheet": ws.title,
            "max_row": ws.max_row,
            "max_col": ws.max_column,
            "nonempty_rows": len(nonempty_rows),
            "first_nonempty_row": nonempty_rows[0] if nonempty_rows else None,
            "last_nonempty_row": nonempty_rows[-1] if nonempty_rows else None,
            "header_preview": [text(v) for v in rows[0]] if rows else [],
        })
    return result


def choose_sheet(wb, requested: str | None) -> tuple[str, list[tuple[Any, ...]], int]:
    if requested:
        if requested not in wb.sheetnames:
            raise ValueError(f"Sheet 不存在：{requested}")
        names = [requested]
    else:
        names = list(wb.sheetnames)
    best = None
    for name in names:
        rows = list(wb[name].iter_rows(values_only=True))
        count = sum(1 for row in rows if nonempty(row))
        candidate = (count, name, rows)
        if best is None or candidate[0] > best[0]:
            best = candidate
    if best is None or best[0] == 0:
        raise ValueError("没有可审核的非空 Sheet")
    _, name, rows = best
    return name, rows, detect_header(rows)


def make_record(excel_row: int, row: tuple[Any, ...], cols: dict[str, int | None]) -> dict[str, Any]:
    def get(kind: str) -> str:
        idx = cols.get(kind)
        return text(row[idx]) if idx is not None and idx < len(row) else ""

    return {
        "excel_row": excel_row,
        "kks_code": get("kks"),
        "old_code": get("old"),
        "parent_code": get("parent"),
        "name": get("name"),
        "unit": get("unit"),
        "device_type": get("dtype"),
        "raw_name": str(row[cols["name"]]) if cols.get("name") is not None and row[cols["name"]] is not None else "",
        "raw_kks": str(row[cols["kks"]]) if cols.get("kks") is not None and row[cols["kks"]] is not None else "",
    }


def add_issue(issues: list[dict[str, Any]], *, priority: str, rule_id: str, category: str,
              rec: dict[str, Any] | None = None, message: str, suggestion: str,
              status: str = "open", evidence: str = "") -> None:
    rec = rec or {}
    issues.append({
        "priority": priority,
        "status": status,
        "rule_id": rule_id,
        "category": category,
        "excel_row": rec.get("excel_row", ""),
        "kks_code": rec.get("kks_code", ""),
        "parent_code": rec.get("parent_code", ""),
        "old_code": rec.get("old_code", ""),
        "name": rec.get("name", ""),
        "message": message,
        "suggestion": suggestion,
        "evidence": evidence,
    })


def code_unit(prefix: str, common_units: set[str] | None = None) -> int | None:
    if prefix in (common_units or COMMON_UNITS):
        return None
    if prefix.isdigit():
        value = int(prefix)
        return value if 1 <= value <= 12 else None
    return None


def classify_extension(code: str) -> str:
    if re.search(r"-KF\d{2}$", code):
        return "DCS 点号"
    if re.search(r"-[A-Z]{2}\d{2}$", code):
        return "部件级附加码"
    if len(code) == 13 and code[-1] in "AB":
        return "A/B 位置扩展"
    if len(code) == 13 and code[-1] in "C":
        return "末位 A/B/C 扩展"
    if "~" in code:
        return "区间设备组"
    return "其他超长码"


STANDARD_PROCESS_PROFESSIONS = {
    "锅炉", "汽机", "电气", "热控", "燃料", "灰硫", "输煤", "化学", "脱硫", "脱硝",
}


def _normalise_profession(value: str) -> str:
    value = text(value).replace("专业", "").replace("及", "/")
    return value.strip()


def summarize_scope(input_path: Path) -> dict[str, Any]:
    """Summarize one workbook for the Skill's cross-scope comparison rules."""
    wb = openpyxl.load_workbook(input_path, read_only=True, data_only=True)
    sheet, rows, header_idx = choose_sheet(wb, None)
    header = rows[header_idx]
    cols = {kind: find_column(header, kind) for kind in ("kks", "old", "parent", "name", "unit", "dtype", "profession")}
    if cols["kks"] is None:
        raise ValueError(f"对比文件主表缺少 KKS 码列：{input_path.name}")
    records: list[dict[str, Any]] = []
    professions: Counter[str] = Counter()
    depth: Counter[str] = Counter()
    codes: set[str] = set()
    for excel_row, row in enumerate(rows[header_idx + 2:], start=header_idx + 2):
        rec = make_record(excel_row, row, cols)
        code = rec["kks_code"]
        if not code:
            continue
        records.append(rec)
        codes.add(code)
        if cols.get("profession") is not None and cols["profession"] < len(row):
            profession = _normalise_profession(text(row[cols["profession"]]))
            if profession:
                professions[profession] += 1
        if len(code) <= 12:
            depth["设备/系统级"] += 1
        else:
            depth[classify_extension(code)] += 1
    non_process = sum(
        count for profession, count in professions.items()
        if profession not in STANDARD_PROCESS_PROFESSIONS
        and not any(token in profession for token in STANDARD_PROCESS_PROFESSIONS)
    )
    return {
        "file": str(input_path),
        "sheet": sheet,
        "header_row": header_idx + 1,
        "total_codes": len(records),
        "unique_codes": len(codes),
        "professional_counts": dict(sorted(professions.items())),
        "non_process_profession_codes": non_process,
        "depth_counts": dict(sorted(depth.items())),
        "profession_column_detected": cols.get("profession") is not None,
    }


def compare_scope(primary_path: Path, comparison_path: Path) -> dict[str, Any]:
    """Compare professional coverage and code depth with a closed quantity delta."""
    primary = summarize_scope(primary_path)
    comparison = summarize_scope(comparison_path)
    p_depth = primary["depth_counts"]
    c_depth = comparison["depth_counts"]
    depth_keys = sorted(set(p_depth) | set(c_depth))
    depth_delta = {key: p_depth.get(key, 0) - c_depth.get(key, 0) for key in depth_keys}
    quantity_delta = primary["total_codes"] - comparison["total_codes"]
    closed_delta = sum(depth_delta.values())
    p_prof = set(primary["professional_counts"])
    c_prof = set(comparison["professional_counts"])
    return {
        "primary": primary,
        "comparison": comparison,
        "professional_coverage": {
            "primary_only": sorted(p_prof - c_prof),
            "comparison_only": sorted(c_prof - p_prof),
            "shared": sorted(p_prof & c_prof),
            "non_process_profession_delta": primary["non_process_profession_codes"] - comparison["non_process_profession_codes"],
        },
        "encoding_depth": {
            "delta_by_category": depth_delta,
            "signal_point_delta": depth_delta.get("DCS 点号", 0),
            "component_extension_delta": depth_delta.get("部件级附加码", 0),
            "base_device_delta": depth_delta.get("设备/系统级", 0),
            "other_extension_delta": sum(value for key, value in depth_delta.items() if key not in {"DCS 点号", "部件级附加码", "设备/系统级"}),
        },
        "quantity_waterfall": {
            "primary_minus_comparison": quantity_delta,
            "signal_point_increment": depth_delta.get("DCS 点号", 0),
            "component_increment": depth_delta.get("部件级附加码", 0),
            "base_device_increment": depth_delta.get("设备/系统级", 0),
            "other_extension_increment": sum(value for key, value in depth_delta.items() if key not in {"DCS 点号", "部件级附加码", "设备/系统级"}),
            "closed_sum": closed_delta,
            "closure_delta": quantity_delta - closed_delta,
            "closed": quantity_delta == closed_delta,
        },
    }


def cycle_nodes(records_by_code: dict[str, dict[str, Any]], root_parents: set[str] | None = None) -> list[str]:
    root_parents = root_parents or ROOT_PARENTS
    cycles: set[str] = set()
    for start in records_by_code:
        seen: list[str] = []
        current = start
        while current and current not in root_parents and current in records_by_code:
            if current in seen:
                cycles.update(seen[seen.index(current):])
                break
            seen.append(current)
            current = records_by_code[current].get("parent_code", "")
    return sorted(cycles)


def semantic_unit_review(rec: dict[str, Any], expected: int) -> tuple[str, str] | None:
    name = rec["name"]
    if expected == 1 and ("至2号机" in name or "1/2号机" in name or "1、2号机" in name):
        return ("resolved", "名称中的另一机组号是联络/去向对象，不是本行所属机组；编码、机组列和父级均支持 1 号机。")
    return None


def _template_columns(header: tuple[Any, ...], cols: dict[str, int | None]) -> dict[str, int | None]:
    """Expose both semantic column names and original headers to the Skill template."""
    result: dict[str, int | None] = dict(cols)
    for idx, value in enumerate(header):
        label = text(value)
        if label:
            result[label] = idx
    # The reliable identity key is the 1:1 old/original KKS, never a drawing number.
    result["key"] = cols.get("old")
    return result


def _append_skill_issue(
    issues: list[dict[str, Any]],
    seen: set[tuple[str, Any, str]],
    records_by_index: dict[int, dict[str, Any]],
    *,
    rule_id: str,
    priority: str,
    category: str,
    index: int | None = None,
    message: str,
    suggestion: str,
    status: str = "open",
    evidence: str = "",
) -> bool:
    rec = records_by_index.get(index, {}) if index is not None else {}
    key = (rule_id, rec.get("excel_row", ""), category)
    if key in seen:
        return False
    seen.add(key)
    add_issue(
        issues,
        priority=priority,
        rule_id=rule_id,
        category=category,
        rec=rec,
        message=message,
        suggestion=suggestion,
        status=status,
        evidence=evidence,
    )
    return True


def _load_rule_config() -> dict[str, Any]:
    path = APP_ROOT / "config" / "kks_rules.json"
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return payload if isinstance(payload, dict) else {}


def _apply_template_skill_rules(
    template_rows: list[tuple[Any, ...]],
    template_cols: dict[str, int | None],
    records_by_index: dict[int, dict[str, Any]],
    records_by_code: dict[str, dict[str, Any]],
    issues: list[dict[str, Any]],
) -> dict[str, Any]:
    """Execute every reusable rule in kks-audit/scripts/audit_template.py."""
    seen = {(str(i.get("rule_id")), i.get("excel_row", ""), str(i.get("category"))) for i in issues}
    diagnostics: dict[str, Any] = {"executed_rules": [], "duplicate_diagnosis": {}}
    rules = SKILL_RULES

    suffix_loss, collisions = rules.dup_split(template_cols, template_rows)
    diagnostics["duplicate_diagnosis"] = {
        "same_source_suffix_loss_groups": len(suffix_loss),
        "cross_device_collision_groups": len(collisions),
        "same_source_suffix_loss_samples": suffix_loss[:20],
        "cross_device_collision_samples": collisions[:20],
    }
    diagnostics["executed_rules"].append("KKS-07")

    for index, name in rules.naming_ambiguity(template_cols, template_rows):
        _append_skill_issue(
            issues, seen, records_by_index, rule_id="KKS-24", priority="P1", category="命名歧义", index=index,
            message=f"名称含相邻机组/设备序号，可能省略限定词：{name}",
            suggestion="补齐机组、设备或位置限定词，避免名称存在多种解释。", status="needs_review",
        )
    diagnostics["executed_rules"].append("KKS-24")

    for system_key, normalized_name, codes in rules.near_dup_by_name(template_cols, template_rows):
        for code in codes:
            index = next((i for i, rec in records_by_index.items() if rec.get("kks_code") == code), None)
            _append_skill_issue(
                issues, seen, records_by_index, rule_id="KKS-25", priority="P1", category="归一名称对应多个 KKS 码", index=index,
                message=f"同一系统 {system_key} 内，归一名称 {normalized_name!r} 对应多个编码：{', '.join(codes)}。",
                suggestion="仅作为人工抽检线索；结合原KKS/设计编码、位置和图纸证据确认，不自动合并。",
                status="needs_review", evidence=f"系统前9位={system_key}；编码={codes}",
            )
    diagnostics["executed_rules"].append("KKS-25")

    semantic_candidates = rules.semantic_dup_by_name(template_cols, template_rows)
    for index, other_index, code_a, code_b, core_a, core_b, similarity, shared in semantic_candidates:
        other = records_by_index.get(other_index, {})
        _append_skill_issue(
            issues, seen, records_by_index, rule_id="KKS-25b", priority="P1", category="同物异名候选", index=index,
            message=f"名称语义高度相近但编码不同：{code_a} / {code_b}，相似度 {similarity}。",
            suggestion="凭原KKS/设计编码、位置和图纸人工裁定；模型和规则都不得自动合并设备。",
            status="needs_review", evidence=f"另一行={other.get('excel_row', other_index)}；核心名={core_a} / {core_b}；共有片段={shared}",
        )
    diagnostics["executed_rules"].append("KKS-25b")

    for identity_key, items in rules.identity_dup_by_key(template_cols, template_rows):
        codes = sorted({item[1] for item in items})
        for _, code, name in items:
            index = next((i for i, rec in records_by_index.items() if rec.get("kks_code") == code), None)
            _append_skill_issue(
                issues, seen, records_by_index, rule_id="KKS-25c", priority="P1", category="结构身份键一对多", index=index,
                message=f"1:1 身份键 {identity_key!r} 对应多个新码：{', '.join(codes)}。",
                suggestion="以原KKS/旧版设计编码作为身份键人工裁定，不自动合并。",
                evidence=f"身份键={identity_key}; 名称={name}",
            )
    diagnostics["executed_rules"].append("KKS-25c")

    for parent_code, parent_name, parent_type, expected in rules.expected_missing(template_cols, template_rows):
        parent = records_by_code.get(parent_code, {})
        index = next((i for i, rec in records_by_index.items() if rec.get("kks_code") == parent_code), None)
        _append_skill_issue(
            issues, seen, records_by_index, rule_id="KKS-26", priority="P1", category="应编未编", index=index,
            message=f"父设备 {parent_code}（{parent_name}）被识别为 {parent_type}，下挂子项未发现期望类型。",
            suggestion=f"人工确认是否缺少：{'、'.join(expected)}；规则库只作提示，不自动补码。",
            status="needs_review", evidence=f"父级编码={parent_code}; 子部件期望={expected}; 父级行={parent.get('excel_row', '')}",
        )
    diagnostics["executed_rules"].append("KKS-26")

    symbol_issues, semantic_voids = rules.name_text_hygiene(template_cols, template_rows)
    for index, sample, chars in symbol_issues:
        _append_skill_issue(
            issues, seen, records_by_index, rule_id="KKS-27", priority="P1", category="名称非常规符号", index=index,
            message=f"名称含控制符、替换符或不在白名单的字符：{chars}。",
            suggestion="清理名称文本并人工确认语义，不直接覆盖源 Excel。", evidence=sample,
        )
    for index, sample in semantic_voids:
        _append_skill_issue(
            issues, seen, records_by_index, rule_id="KKS-27", priority="P2", category="名称语义需人核", index=index,
            message="名称过短、纯符号/数字或无法提取中文/字母语义。",
            suggestion="人工确认名称；该项不自动判定设备编码错误。", status="needs_review", evidence=sample,
        )
    diagnostics["executed_rules"].append("KKS-27")

    for index, code, reason in rules.unit_consistency(template_cols, template_rows):
        _append_skill_issue(
            issues, seen, records_by_index, rule_id="KKS-28", priority="P1", category="机组三方不一致", index=index,
            message=reason, suggestion="核对编码前缀、机组列和名称三方，不自动改码。",
        )
    diagnostics["executed_rules"].append("KKS-28")

    for index, code, reason, suggestion in rules.segment_type_check(template_cols, template_rows):
        _append_skill_issue(
            issues, seen, records_by_index, rule_id="KKS-04b", priority="P0", category="分段字符类型错误", index=index,
            message=reason, suggestion=f"候选修正：{suggestion or code}；必须结合兄弟码/图纸人工确认。",
        )
    diagnostics["executed_rules"].append("KKS-04b")

    for index, code, reason in rules.ocr_confusable(template_cols, template_rows):
        # I/O in a numeric segment is already a P0; lower-case l in an otherwise
        # valid segment remains the Skill's P1 human-review hint.
        if any(i.get("excel_row") == records_by_index.get(index, {}).get("excel_row") and i.get("rule_id") in {"KKS-04b", "KKS-06"} for i in issues):
            continue
        priority = "P1" if "小写 l" in reason else "P0"
        _append_skill_issue(
            issues, seen, records_by_index, rule_id="KKS-29", priority=priority, category="OCR 易混字符待核", index=index,
            message=reason, suggestion="只标记待核，不自动把字符替换为数字。",
        )
    diagnostics["executed_rules"].append("KKS-29")

    for index, label, reason, sample in rules.text_hygiene(template_cols, template_rows):
        category = "编码文本卫生" if label == "码" else "名称文本卫生"
        _append_skill_issue(
            issues, seen, records_by_index, rule_id="KKS-30", priority="P1", category=category, index=index,
            message=f"{label}存在{reason}。", suggestion="按 trim、去控制符、全角空格转半角规则清洗后重新审核，不直接覆盖源 Excel。", evidence=sample,
        )
    diagnostics["executed_rules"].append("KKS-30")

    for code, parent_name, child_name, reason in rules.parent_child_name_conflict(template_cols, template_rows):
        index = next((i for i, rec in records_by_index.items() if rec.get("kks_code") == code), None)
        _append_skill_issue(
            issues, seen, records_by_index, rule_id="KKS-31", priority="P1", category="父子名称语义矛盾", index=index,
            message=f"直接父子名称标识不一致：{reason}。", suggestion="核对父子设备是否挂错或名称是否复制错误；不自动改码。", evidence=f"父名={parent_name}; 子名={child_name}",
        )
    diagnostics["executed_rules"].append("KKS-31")
    return diagnostics


DEVICE_LETTER_SEMANTICS = {
    "泵": {"BG", "AP", "CP"},
    "阀": {"QM", "AA"},
    "门": {"QM", "AA"},
    "执行机构": {"MA", "CA"},
    "测温": {"BT"},
    "温度": {"BT"},
    "流量": {"BW", "FT"},
    "液位": {"LS"},
    "给煤机": {"GL"},
    "压缩机": {"CM"},
}
KNOWN_DEVICE_LETTERS = set().union(*DEVICE_LETTER_SEMANTICS.values())
NAME_CONTEXT_TERMS = ("管道", "系统", "水室", "用汽", "冷却水", "油处理")
NAME_TYPE_TERMS = {
    "阀": ("阀", "截止阀", "闸阀", "止回阀", "蝶阀"),
    "泵": ("泵",),
    "电机": ("电机", "电动机", "马达"),
    "风机": ("风机", "送风机", "引风机"),
}


def _apply_additional_skill_rules(
    records: list[dict[str, Any]],
    records_by_index: dict[int, dict[str, Any]],
    issues: list[dict[str, Any]],
    config: dict[str, Any],
) -> dict[str, Any]:
    """Implement the checklist's A-G, migration, sequence and import guards."""
    seen = {(str(i.get("rule_id")), i.get("excel_row", ""), str(i.get("category"))) for i in issues}
    diagnostics: dict[str, Any] = {"executed_rules": [], "sequence_gap_groups": 0, "style_variants": {}}
    old_prefix_migrations = config.get("old_prefix_migrations", OLD_PREFIX)
    if not isinstance(old_prefix_migrations, dict):
        old_prefix_migrations = OLD_PREFIX

    for index, rec in records_by_index.items():
        code = rec.get("kks_code", "")
        name = rec.get("name", "")
        if len(code) >= 9 and code[7:9] in KNOWN_DEVICE_LETTERS:
            letters = code[7:9]
            matches = [(name.rfind(term), term, allowed) for term, allowed in DEVICE_LETTER_SEMANTICS.items() if term in name]
            if matches:
                position, term, allowed = max(matches)
                # “供给水泵管道/泵系统”中的泵是工艺上下文，不是本行设备主体；
                # 只有最后一个明确设备词才进入名称-设备字母一致性检查。
                if not (term == "泵" and any(context in name[position:] for context in NAME_CONTEXT_TERMS)) and letters not in allowed:
                    _append_skill_issue(
                        issues, seen, records_by_index, rule_id="KKS-A", priority="P1", category="名称语义与设备字母疑点", index=index,
                        message=f"名称主体词为“{term}”，但编码设备字母为 {letters}。", suggestion="结合 KKS 设备字母表和名称语义人工确认，不自动改码。", status="needs_review",
                    )
    diagnostics["executed_rules"].append("KKS-A")

    style_counts = Counter()
    for rec in records:
        name = rec.get("name", "")
        if re.search(r"\d+#", name):
            style_counts["1#"] += 1
        if re.search(r"#\d+", name):
            style_counts["#1"] += 1
        if re.search(r"\d+号", name):
            style_counts["1号"] += 1
        if re.search(r"一号|二号|三号|四号|五号|六号", name):
            style_counts["汉字号"] += 1
    diagnostics["style_variants"] = dict(style_counts)
    if len([key for key, value in style_counts.items() if value]) > 1:
        _append_skill_issue(
            issues, seen, records_by_index, rule_id="KKS-C", priority="P2", category="命名风格不一",
            message="同一文件同时出现多种机组/序号命名风格。",
            suggestion="统一使用项目约定的机组和序号表达方式；该项是风格提示，不直接判错。",
            evidence=json.dumps(dict(style_counts), ensure_ascii=False),
        )
    diagnostics["executed_rules"].append("KKS-C")

    for index, rec in records_by_index.items():
        dtype = rec.get("device_type", "")
        name = rec.get("name", "")
        if not dtype or not name:
            continue
        dtype_group = next((group for group, terms in NAME_TYPE_TERMS.items() if any(term in dtype for term in terms)), None)
        name_groups = [group for group, terms in NAME_TYPE_TERMS.items() if any(term in name for term in terms)]
        if dtype_group and name_groups and dtype_group not in name_groups:
            _append_skill_issue(
                issues, seen, records_by_index, rule_id="KKS-E", priority="P2", category="设备类型与名称矛盾", index=index,
                message=f"设备类型列为“{dtype}”，名称疑似属于“{'、'.join(name_groups)}”。", suggestion="人工核对设备类型列和名称，不自动覆盖任一字段。", status="needs_review",
            )
    diagnostics["executed_rules"].append("KKS-E")

    sequence_groups: defaultdict[tuple[str, str], list[tuple[int, dict[str, Any], int]]] = defaultdict(list)
    for index, rec in records_by_index.items():
        code = rec.get("kks_code", "")
        if len(code) < 12 or not code[9:12].isdigit():
            continue
        sequence_groups[(rec.get("parent_code", ""), code[:9])].append((index, rec, int(code[9:12])))
    for (parent, prefix), items in sequence_groups.items():
        numbers = sorted({number for _, _, number in items})
        if len(numbers) < 2 or numbers[-1] - numbers[0] > 500:
            continue
        gaps = [number for number in range(numbers[0], numbers[-1] + 1) if number not in numbers]
        if not gaps or len(gaps) > 30:
            continue
        index, _, _ = items[0]
        _append_skill_issue(
            issues, seen, records_by_index, rule_id="KKS-F", priority="P2", category="同父级编号断号", index=index,
            message=f"同父级/同设备类型序列存在缺号：{', '.join(f'{gap:03d}' for gap in gaps[:30])}。",
            suggestion="确认是否为漏编、已废弃编号或有意留空；不按断号自动补码。", status="needs_review", evidence=f"父级={parent}; 前缀={prefix}; 已有编号={numbers[:30]}",
        )
        diagnostics["sequence_gap_groups"] += 1
    diagnostics["executed_rules"].append("KKS-F")

    pair_rules = config.get("valid_letter_pairs", {}) if isinstance(config.get("valid_letter_pairs"), dict) else {}
    for index, rec in records_by_index.items():
        code = rec.get("kks_code", "")
        if len(code) < 9 or not pair_rules:
            continue
        system, device = code[2:5], code[7:9]
        allowed_devices = pair_rules.get(system, [])
        if allowed_devices and device not in allowed_devices:
            _append_skill_issue(
                issues, seen, records_by_index, rule_id="KKS-G", priority="P2", category="系统与设备字母组合疑点", index=index,
                message=f"系统字母 {system} 与设备字母 {device} 不在配置的合法组合中。", suggestion="更新项目字母组合字典或人工确认该编码。", status="needs_review",
            )
    diagnostics["executed_rules"].append("KKS-G")

    migration = config.get("letter_reallocations", {}) if isinstance(config.get("letter_reallocations"), dict) else {}
    system_map = migration.get("system", {}) if isinstance(migration.get("system"), dict) else {}
    device_map = migration.get("device", {}) if isinstance(migration.get("device"), dict) else {}
    for index, rec in records_by_index.items():
        old, new = rec.get("old_code", ""), rec.get("kks_code", "")
        if not old or not new:
            continue
        if old[:2] in old_prefix_migrations and new[:2] != old_prefix_migrations[old[:2]]:
            _append_skill_issue(
                issues, seen, records_by_index, rule_id="KKS-16", priority="P1", category="旧格式码未正确迁移", index=index,
                message=f"原码/新码前缀迁移疑点：{old[:2]} → {new[:2]}。", suggestion="核对 10 版→20 版迁移表，不直接覆盖历史码。",
            )
        if old[2:5] in system_map and system_map[old[2:5]] != new[2:5]:
            _append_skill_issue(
                issues, seen, records_by_index, rule_id="KKS-14", priority="P1", category="系统字母重分配疑点", index=index,
                message=f"系统字母应按配置从 {old[2:5]} 映射为 {system_map[old[2:5]]}，当前为 {new[2:5]}。", suggestion="核对版本迁移字母表和设备语义。",
            )
        if old[7:9] in device_map and device_map[old[7:9]] != new[7:9]:
            _append_skill_issue(
                issues, seen, records_by_index, rule_id="KKS-14", priority="P1", category="设备字母重分配疑点", index=index,
                message=f"设备字母应按配置从 {old[7:9]} 映射为 {device_map[old[7:9]]}，当前为 {new[7:9]}。", suggestion="核对版本迁移字母表和设备语义。",
            )
        if len(old) != len(new):
            _append_skill_issue(
                issues, seen, records_by_index, rule_id="KKS-15", priority="P1", category="层级位置变化需人工核", index=index,
                message=f"原码长度 {len(old)} 与新码长度 {len(new)} 不同，存在层级重排或扩展建模变化。", suggestion="结合设备语义确认层级是否重排，不自动判为错误。", status="needs_review",
            )
    diagnostics["executed_rules"].extend(["KKS-14", "KKS-15", "KKS-16"])
    return diagnostics


def detect_tree_collection_columns(header: tuple[Any, ...]) -> dict[str, int] | None:
    """Detect the Skill's system-device-tree collection format."""
    labels = [text(value) for value in header]

    def first(predicate):
        return next((index for index, label in enumerate(labels) if predicate(label)), None)

    parent_code = first(lambda label: "KKS" in label and "子" not in label and ("设备" in label or "编码" in label))
    child_code = first(lambda label: "KKS" in label and "子" in label)
    parent_name = first(lambda label: ("名称" in label or "设备名" in label) and "子" not in label and "设备" in label)
    child_name = first(lambda label: ("名称" in label or "设备名" in label) and "子" in label)
    if None in (parent_code, child_code, parent_name, child_name):
        return None
    return {"parent_code": parent_code, "child_code": child_code, "parent_name": parent_name, "child_name": child_name}


def _tree_value(row: tuple[Any, ...], index: int) -> str:
    return text(row[index]) if index < len(row) else ""


def _tree_record(excel_row: int, code: str, parent: str, name: str) -> dict[str, Any]:
    return {
        "excel_row": excel_row,
        "kks_code": code,
        "old_code": "",
        "parent_code": parent,
        "name": name,
        "unit": "",
        "device_type": "",
        "raw_name": name,
        "raw_kks": code,
    }


def audit_tree_collection(
    input_path: Path,
    output_dir: Path,
    sheet: str,
    rows: list[tuple[Any, ...]],
    header_idx: int,
    tree_cols: dict[str, int],
    comparison_path: Path | None = None,
    progress_callback: ProgressCallback | None = None,
) -> dict[str, Any]:
    """Audit the two-code-per-row system-device-tree format from the Skill."""
    _emit_progress(progress_callback, {"phase": "audit", "stage": "reading", "percent": 12, "message": "正在读取系统设备树", "hint": "识别父、子设备列和工作表结构"})
    sheet_info = []
    try:
        workbook = openpyxl.load_workbook(input_path, read_only=True, data_only=True)
        sheet_info = inspect_sheets(workbook)
    except Exception:
        sheet_info = [{"sheet": sheet, "max_row": len(rows), "max_col": max((len(row) for row in rows), default=0), "nonempty_rows": len(rows), "first_nonempty_row": 1, "last_nonempty_row": len(rows), "header_preview": list(rows[header_idx])}]
    issues: list[dict[str, Any]] = []
    resolved: list[dict[str, Any]] = []
    child_records: list[dict[str, Any]] = []
    child_groups: defaultdict[str, list[dict[str, Any]]] = defaultdict(list)
    relation_counts: Counter[str] = Counter()
    missing_child = 0
    accepted_lengths = {2, 3, 4, 5, 7, 12, 13}

    def normalized(code: str) -> str:
        return code.replace("O", "0").replace("o", "0").replace("I", "1").replace("i", "1")

    for excel_row, row in enumerate(rows[header_idx + 2:], start=header_idx + 2):
        parent_code = _tree_value(row, tree_cols["parent_code"])
        child_code = _tree_value(row, tree_cols["child_code"])
        parent_name = _tree_value(row, tree_cols["parent_name"])
        child_name = _tree_value(row, tree_cols["child_name"])
        if not parent_code and not child_code:
            if child_name:
                missing_child += 1
                add_issue(issues, priority="P0", rule_id="KKS-00", category="子设备缺 KKS 码", rec=_tree_record(excel_row, "", "", child_name), message="系统设备树行存在子设备名称但没有子设备 KKS 码。", suggestion="补齐子设备 KKS，或明确该行仅为父设备节点。")
            continue
        if not child_code:
            missing_child += 1
            add_issue(issues, priority="P0", rule_id="KKS-00", category="子设备缺 KKS 码", rec=_tree_record(excel_row, parent_code, "", parent_name), message="系统设备树行存在子设备名称但没有子设备 KKS 码。", suggestion="补齐子设备 KKS，或明确该行仅为父设备节点。")
            continue
        rec = _tree_record(excel_row, child_code, parent_code, child_name)
        child_records.append(rec)
        child_groups[child_code].append(rec)
        if not ALLOWED_CODE.fullmatch(child_code):
            add_issue(issues, priority="P0", rule_id="KKS-04", category="编码含非法字符", rec=rec, message=f"子设备编码含非允许字符：{child_code}", suggestion="只保留大写字母、数字和约定分隔符。")
        if len(child_code) not in accepted_lengths:
            add_issue(issues, priority="P0", rule_id="KKS-05", category="编码长度异常", rec=rec, message=f"子设备编码长度为 {len(child_code)}，不在已识别范围内。", suggestion="核对变长层级或扩展建模。")
        if len(child_code) >= 12:
            for start, end, label in ((0, 2, "机组位"), (2, 5, "系统字母"), (5, 7, "系统编号"), (7, 9, "设备字母"), (9, 12, "设备顺序号")):
                part = child_code[start:end]
                should_digit = label in {"机组位", "系统编号", "设备顺序号"}
                if (part.isdigit() if should_digit else part.isalpha()) is False:
                    add_issue(issues, priority="P0", rule_id="KKS-04b", category="分段字符类型错误", rec=rec, message=f"{label}片段 {part!r} 含错误字符。", suggestion="按同层兄弟码和图纸人工确认修正值。")
                    break
            if any(char in child_code for char in "IiOo"):
                add_issue(issues, priority="P0", rule_id="KKS-06", category="KKS 禁用易混字母", rec=rec, message=f"编码含 KKS 字母表排除的 I/O：{child_code}", suggestion="核对是否为 1/0 的 OCR 或录入错误。")
        if len(child_code) > 12:
            add_issue(issues, priority="P1", rule_id="KKS-17/22", category="12 位以上扩展码", rec=rec, message=f"{child_code} 属于 {classify_extension(child_code)}，不能直接按 12 位 LOCATIONS 主码处理。", suggestion="拆分位置、相位、部件或信号属性，不要直接截断。", status="needs_review")
        if not parent_code:
            relation_counts["MISSING"] += 1
            add_issue(issues, priority="P0", rule_id="KKS-08", category="树表缺父级 KKS", rec=rec, message="子设备存在但同行父设备 KKS 为空。", suggestion="补齐父级 KKS；不要通过改子设备码掩盖树关系缺失。")
            continue
        p_code = normalized(parent_code)
        c_code = normalized(child_code)
        if c_code.startswith(p_code):
            relation = "DESCENDS"
        elif len(c_code) >= 7 and len(p_code) >= 7 and c_code[:7] == p_code[:7] and c_code[7:9] != p_code[7:9]:
            relation = "SAME_SUBSYSTEM"
        elif c_code[:2] == p_code[:2]:
            relation = "CROSS_SUBSYSTEM"
        else:
            relation = "CROSS_UNIT"
        relation_counts[relation] += 1
        if relation == "CROSS_UNIT":
            add_issue(issues, priority="P0", rule_id="KKS-TREE-REL", category="跨机组树关系", rec=rec, message=f"父码 {parent_code} 与子码 {child_code} 的机组位不一致。", suggestion="先确认 O/I OCR 归一和真实机组归属；不要直接改子设备 KKS。", evidence=f"关系分类={relation}; 父名={parent_name}; 子名={child_name}")
        elif relation == "CROSS_SUBSYSTEM":
            add_issue(issues, priority="P1", rule_id="KKS-TREE-REL", category="跨子系统树关系", rec=rec, message=f"父码 {parent_code} 与子码 {child_code} 不在同一 7 位子系统前缀下。", suggestion="多数是运行/工艺关联与 KKS 归口不同；建议增加显式父级列，不改子码归口。", status="needs_review", evidence=f"关系分类={relation}; 父名={parent_name}; 子名={child_name}")

    for code, group in child_groups.items():
        if len(group) > 1:
            for rec in group:
                add_issue(issues, priority="P0", rule_id="KKS-07", category="重复子设备 KKS 码", rec=rec, message=f"子设备 KKS 码 {code} 在树表出现 {len(group)} 次。", suggestion="核对是否为重复挂接或不同设备碰撞。")
    priority_counts = Counter(item["priority"] for item in issues if item["status"] != "resolved")
    comparison = None
    if comparison_path:
        comparison = compare_scope(input_path, comparison_path)
        waterfall = comparison["quantity_waterfall"]
        add_issue(issues, priority="P2", rule_id="KKS-19/20/21", category="专业覆盖/编码深度/数量差异对比", message=f"与对比文件有效编码数量差 {waterfall['primary_minus_comparison']}；瀑布闭合={waterfall['closed']}。", suggestion="结合专业覆盖和编码深度解释数量，不以总量直接判断质量。", status="needs_review", evidence=json.dumps(comparison, ensure_ascii=False))
        priority_counts = Counter(item["priority"] for item in issues if item["status"] != "resolved")
    issue_count = sum(1 for item in issues if item["status"] != "resolved")
    conclusion = (
        f"本次共发现 {issue_count} 个问题，请按问题清单逐项核对；源 Excel 未修改。"
        if issue_count
        else "本次未发现需要列入问题清单的问题；源 Excel 未修改。"
    )
    result = {
        "version": VERSION,
        "source_file": str(input_path),
        "sheet": sheet,
        "header_row": header_idx + 1,
        "data_rows": len(child_records),
        "nonempty_rows_without_code": missing_child,
        "columns": tree_cols,
        "sheet_info": sheet_info,
        "code_length_distribution": dict(sorted(Counter(len(rec["kks_code"]) for rec in child_records).items())),
        "metrics": {"unique_codes": len(child_groups), "duplicate_code_groups": sum(1 for group in child_groups.values() if len(group) > 1), "duplicate_suffix_loss_groups": 0, "duplicate_cross_device_collision_groups": 0, "orphan_rows": relation_counts.get("MISSING", 0), "prefix_mismatch_rows": relation_counts.get("CROSS_SUBSYSTEM", 0) + relation_counts.get("CROSS_UNIT", 0), "blank_separator_rows": 0, "long_code_rows": sum(1 for rec in child_records if len(rec["kks_code"]) > 12), "unit_rule_hits_resolved": 0, "sequence_gap_groups": 0, "tree_relations": dict(relation_counts)},
        "priority_counts": dict(priority_counts),
        "issue_count": issue_count,
        "resolved_count": len(resolved),
        "conclusion": conclusion,
        "structural_notes": ["检测到系统设备树收集表格式：父/子设备通过同行两列表达，未按扁平表强行要求显式父级列。", "跨子系统多为运行关联与 KKS 归口冲突，建议增加显式父级 KKS 列，不自动改子设备码。", f"树关系统计：{dict(relation_counts)}"],
        "issues": issues,
        "resolved_reviews": resolved,
        "skill_coverage": {"source": SKILL_RULE_SOURCE, "template_loaded": True, "variant_mode": "system-device-tree", "template_rules": {"executed_rules": ["KKS-01", "KKS-02", "KKS-03", "KKS-04", "KKS-04b", "KKS-05", "KKS-06", "KKS-07", "KKS-08", "KKS-17/22", "KKS-TREE-REL"]}, "additional_rules": {"executed_rules": ["KKS-19/20/21"] if comparison else []}, "config_file": "config/kks_rules.json"},
        "import_governance": {"target": "DM8 LOCATIONS", "direct_import_ready": False, "long_code_requires_extension_model": bool(sum(1 for rec in child_records if len(rec["kks_code"]) > 12)), "historical_kks_rewrite_forbidden": True, "required_strategy": "保留历史 KKS；通过映射层/双码共存/冻结历史兼容两票与缺陷记录。", "runtime_dm8_validation": False},
        "scope_comparison": comparison,
    }
    _emit_progress(progress_callback, {"phase": "audit", "stage": "rules_completed", "percent": 62, "message": "本地规则审核完成", "hint": f"已发现 {len(issues)} 个规则问题，准备进行 AI 候选复核"})
    result["ai_review"] = review_issue_candidates(result, progress_callback=progress_callback)
    _emit_progress(progress_callback, {"phase": "audit", "stage": "report", "percent": 95, "message": "正在生成审核报告", "hint": "整理 HTML 报告和问题 Excel"})
    output_dir.mkdir(parents=True, exist_ok=True)
    html_name, xlsx_name = output_artifact_names(input_path)
    write_html(output_dir / html_name, result)
    write_issue_workbook_xlsx(output_dir / xlsx_name, result)
    return result


def audit_file(
    input_path: Path,
    output_dir: Path,
    requested_sheet: str | None = None,
    comparison_path: Path | None = None,
    progress_callback: ProgressCallback | None = None,
) -> dict[str, Any]:
    configure_logging()
    _emit_progress(progress_callback, {"phase": "audit", "stage": "reading", "percent": 8, "message": "正在读取 Excel", "hint": f"打开工作簿：{input_path.name}"})
    LOGGER.info("audit_start file=%s output_dir=%s", input_path.name, output_dir)
    wb = openpyxl.load_workbook(input_path, read_only=True, data_only=True)
    rule_config = _load_rule_config()
    root_parents = {text(item) for item in rule_config.get("root_parent_values", ROOT_PARENTS)} or ROOT_PARENTS
    old_prefix_migrations = rule_config.get("old_prefix_migrations", OLD_PREFIX)
    if not isinstance(old_prefix_migrations, dict):
        old_prefix_migrations = OLD_PREFIX
    common_units = {text(item) for item in rule_config.get("common_unit_prefixes", COMMON_UNITS)} or COMMON_UNITS
    accepted_lengths = {int(item) for item in rule_config.get("accepted_code_lengths", [2, 3, 4, 5, 7, 12, 13])}
    sheet_info = inspect_sheets(wb)
    sheet, rows, header_idx = choose_sheet(wb, requested_sheet)
    _emit_progress(progress_callback, {"phase": "audit", "stage": "structure", "percent": 25, "message": "正在识别工作表结构", "hint": f"当前工作表：{sheet}；数据行：{max(0, len(rows) - header_idx - 1)}"})
    header = rows[header_idx]
    cols = {kind: find_column(header, kind) for kind in ("kks", "old", "parent", "name", "unit", "dtype", "profession")}
    missing = [kind for kind in ("kks", "parent", "name") if cols[kind] is None]
    if missing:
        tree_cols = detect_tree_collection_columns(header)
        if tree_cols:
            LOGGER.info("audit_tree_collection_detected file=%s sheet=%s", input_path.name, sheet)
            return audit_tree_collection(input_path, output_dir, sheet, rows, header_idx, tree_cols, comparison_path, progress_callback)
        raise ValueError(f"主表缺少关键列：{', '.join(missing)}")

    records: list[dict[str, Any]] = []
    blank_rows: list[int] = []
    nonempty_without_code: list[dict[str, Any]] = []
    nonempty_indices = [i for i, row in enumerate(rows) if i > header_idx and nonempty(row)]
    last_data_index = max(nonempty_indices) if nonempty_indices else header_idx
    bounded_rows = rows[header_idx + 1:last_data_index + 1]
    template_cols = _template_columns(header, cols)
    template_records = {
        index: make_record(header_idx + 2 + index, row, cols)
        for index, row in enumerate(bounded_rows)
    }
    for i, row in enumerate(bounded_rows, start=header_idx + 2):
        if not nonempty(row):
            blank_rows.append(i)
            continue
        rec = make_record(i, row, cols)
        if rec["kks_code"]:
            records.append(rec)
        else:
            nonempty_without_code.append(rec)

    issues: list[dict[str, Any]] = []
    resolved: list[dict[str, Any]] = []
    by_code: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for rec in records:
        by_code[rec["kks_code"]].append(rec)
    code_set = set(by_code)

    for code, group in sorted(by_code.items()):
        if len(group) > 1:
            for rec in group:
                add_issue(issues, priority="P0", rule_id="KKS-07", category="重复 KKS 码", rec=rec,
                          message=f"KKS 码 {code} 在主表出现 {len(group)} 次。", suggestion="核对是否为后缀丢失或异设备真碰撞；修复前不得导入。")

    prefix_bad: list[dict[str, Any]] = []
    orphan_recs: list[dict[str, Any]] = []
    for rec in records:
        code, parent = rec["kks_code"], rec["parent_code"]
        if not code:
            continue
        if not ALLOWED_CODE.fullmatch(code):
            add_issue(issues, priority="P0", rule_id="KKS-04", category="编码含非法字符", rec=rec,
                      message=f"编码含非允许字符：{code}", suggestion="只保留大写字母、数字和约定分隔符；请人工确认。")
        if len(code) not in accepted_lengths:
            add_issue(issues, priority="P0", rule_id="KKS-05", category="编码长度异常", rec=rec,
                      message=f"编码长度为 {len(code)}，不在本表已识别的变长层级/扩展范围内。", suggestion="核对是否误拼接或缺失层级。")
        if parent and parent not in root_parents and parent not in code_set:
            orphan_recs.append(rec)
            add_issue(issues, priority="P0", rule_id="KKS-08", category="父级孤儿", rec=rec,
                      message=f"父级 {parent} 不存在于主表编码集合。", suggestion="补齐父级节点或确认是否引用了未迁移的旧码。")
        if len(code) == 2 and parent not in root_parents and parent in code_set:
            add_issue(issues, priority="P1", rule_id="KKS-12", category="根节点父级非哨兵", rec=rec,
                      message=f"根节点 {code} 的父级为 {parent}，不是空值或 -1 哨兵。", suggestion="核对根节点层级；根节点父级应使用空值或 -1。")
        if parent and parent not in root_parents and not code.startswith(parent):
            prefix_bad.append(rec)
            add_issue(issues, priority="P0", rule_id="KKS-09", category="父子前缀不一致", rec=rec,
                      message=f"子码 {code} 不以父码 {parent} 开头。", suggestion="核对显式父级指针；舟山变长层级不要按固定长度推导。")
        if parent == code:
            add_issue(issues, priority="P0", rule_id="KKS-10", category="自引用", rec=rec,
                      message="父级编码与本行编码相同。", suggestion="修正父级指针。")
        if parent and parent not in root_parents and len(parent) >= len(code):
            add_issue(issues, priority="P0", rule_id="KKS-11", category="父级长度异常", rec=rec,
                      message=f"父级长度 {len(parent)} 不小于子码长度 {len(code)}。", suggestion="核对层级关系。")

        if len(code) >= 12:
            segments = ((0, 2, "机组位"), (2, 5, "系统字母"), (5, 7, "系统编号"), (7, 9, "设备字母"), (9, 12, "设备顺序号"))
            seg_has_issue = False
            for start, end, label in segments:
                part = code[start:end]
                should_digit = label in {"机组位", "系统编号", "设备顺序号"}
                ok = part.isdigit() if should_digit else part.isalpha()
                if not ok:
                    seg_has_issue = True
                    suggestion = code
                    if label == "设备顺序号":
                        suggestion = code[:start] + part.replace("O", "0").replace("o", "0").replace("I", "1").replace("i", "1").replace("l", "1") + code[end:]
                    add_issue(issues, priority="P0", rule_id="KKS-04b", category="分段字符类型错误", rec=rec,
                              message=f"{label}片段 {part!r} 含错误字符。", suggestion=f"建议候选：{suggestion}；必须结合兄弟码/图纸人工确认。")
                    break
            if any(ch in code for ch in "IiOo") and not seg_has_issue:
                add_issue(issues, priority="P0", rule_id="KKS-06", category="KKS 禁用易混字母", rec=rec,
                          message=f"编码含 KKS 字母表排除的 I/O：{code}", suggestion="核对是否为 1/0 的 OCR 或录入错误，不自动替换。")

        if len(code) > 12:
            kind = classify_extension(code)
            add_issue(issues, priority="P1", rule_id="KKS-17/22", category="12 位以上扩展码", rec=rec,
                      message=f"{code} 被识别为“{kind}”；静态规则不能直接判断其是否应进入 LOCATIONS 主码。", suggestion="确认是否拆出位置/相位/部件/信号属性列，或进入扩展表；不要直接截断。", status="needs_review")

        if rec["name"] == "":
            add_issue(issues, priority="P0", rule_id="KKS-D", category="有码无名称", rec=rec,
                      message="有 KKS 码但设备名称为空。", suggestion="补齐设备名称后再审核。")
        raw = rec["raw_kks"]
        if raw != raw.strip() or any(ch in raw for ch in "\n\r\t\u3000\xa0"):
            add_issue(issues, priority="P1", rule_id="KKS-30", category="编码文本卫生", rec=rec,
                      message="编码含首尾空白、换行、制表符或全角空格。", suggestion="建议清洗文本后重新审核，不直接覆盖源表。")
        raw_name = rec["raw_name"]
        if any(ch in raw_name for ch in "\n\r\t\u3000\xa0"):
            add_issue(issues, priority="P1", rule_id="KKS-30", category="名称文本卫生", rec=rec,
                      message="名称含首尾空白、换行、制表符或全角空格。", suggestion="建议清洗文本后重新审核。")
        unusual_name_chars = re.findall(r"[^\u4e00-\u9fffA-Za-z0-9()（）\[\]【】.\-—_/\\,，、;；:：%#+*&~°℃㎡\"“”‘’!！?？\s]", raw_name)
        if unusual_name_chars:
            chars = "".join(dict.fromkeys(unusual_name_chars))
            if any("\u2e80" <= ch <= "\u2fff" for ch in unusual_name_chars):
                message = f"名称含 CJK 部首/兼容字符 {chars!r}，可能是普通汉字的 Unicode 变体。"
                suggestion = "建议将该字符与普通汉字（如‘水’）核对并标准化后再提交。"
            else:
                message = "名称含非常规符号或控制字符。"
                suggestion = "清理符号并人工确认名称语义。"
            add_issue(issues, priority="P1", rule_id="KKS-27", category="名称非常规符号", rec=rec,
                      message=message, suggestion=suggestion)
        if len(rec["name"]) < 2 or (not re.search(r"[\u4e00-\u9fffA-Za-z]", rec["name"])):
            add_issue(issues, priority="P2", rule_id="KKS-27", category="名称语义需人核", rec=rec,
                      message="名称过短或无法提取中文/字母语义。", suggestion="人工核对名称，不按此项自动判错。", status="needs_review")

        expected = code_unit(code[:2], common_units)
        if expected is not None:
            for field, label in (("name", "名称"), ("unit", "机组列")):
                m = NAME_UNIT.search(rec[field]) if rec[field] else None
                if not m:
                    continue
                key = m.group(1)
                found = CN_NUM[key] if key in CN_NUM else int(key)
                if found != expected:
                    review = semantic_unit_review(rec, expected)
                    if review and review[0] == "resolved":
                        resolved.append({"priority": "排除", "status": "resolved", "rule_id": "KKS-28", "category": "机组三方一致（语义复核）", "excel_row": rec["excel_row"], "kks_code": rec["kks_code"], "parent_code": rec["parent_code"], "old_code": rec["old_code"], "name": rec["name"], "message": f"{label}规则命中：{review[1]}", "suggestion": "不修改编码；保留为语义复核记录。", "evidence": "编码前缀、机组列、父级均指向当前机组。"})
                    else:
                        add_issue(issues, priority="P1", rule_id="KKS-28", category="机组三方不一致", rec=rec,
                                  message=f"{label}指向 {found} 号，但编码前缀指向 {expected} 号。", suggestion="核对编码、机组列与名称三方，不自动改码。")

        old, new = rec["old_code"], rec["kks_code"]
        if old and new and old[:2] != new[:2]:
            expected_prefix = old_prefix_migrations.get(old[:2])
            if expected_prefix and new[:2] != expected_prefix:
                add_issue(issues, priority="P1", rule_id="KKS-13/16", category="旧码迁移疑点", rec=rec,
                          message=f"原码前缀 {old[:2]} 按规则应迁移为 {expected_prefix}，当前新码为 {new[:2]}。", suggestion="核对 10 版→20 版迁移表和设备语义。")
            else:
                add_issue(issues, priority="P2", rule_id="KKS-18", category="历史原码与新码前缀变化", rec=rec,
                          message=f"原 KKS 前缀 {old[:2]} 与新 KKS 前缀 {new[:2]} 不同。", suggestion="确认历史映射；若新码、父级、机组三方一致，可作为历史提示保留。", status="needs_review")

        if code[:2] in old_prefix_migrations:
            add_issue(issues, priority="P1", rule_id="KKS-16", category="新码仍保留旧格式前缀", rec=rec,
                      message=f"当前 KKS 编码仍使用旧格式前缀 {code[:2]}。", suggestion="核对 10 版→20 版迁移结果；历史原码可以保留，但新码不应直接沿用。")

    unique_records_by_code = {code: group[0] for code, group in by_code.items() if len(group) == 1}
    skill_template_diagnostics = _apply_template_skill_rules(
        bounded_rows,
        template_cols,
        template_records,
        unique_records_by_code,
        issues,
    )
    additional_skill_diagnostics = _apply_additional_skill_rules(
        records,
        template_records,
        issues,
        rule_config,
    )

    scope_comparison = None
    if comparison_path:
        scope_comparison = compare_scope(input_path, comparison_path)
        waterfall = scope_comparison["quantity_waterfall"]
        add_issue(
            issues,
            priority="P2",
            rule_id="KKS-19/20/21",
            category="专业覆盖/编码深度/数量差异对比",
            message=(
                f"与对比文件的有效编码数量差 {waterfall['primary_minus_comparison']}；"
                f"瀑布闭合={waterfall['closed']}，闭合差={waterfall['closure_delta']}。"
            ),
            suggestion="结合专业覆盖差异和编码深度差异解释数量，不以总量直接判断编码质量。",
            status="needs_review",
            evidence=json.dumps(scope_comparison, ensure_ascii=False),
        )

    for rec in nonempty_without_code:
        add_issue(issues, priority="P0", rule_id="KKS-00", category="非空行缺 KKS 码", rec=rec,
                  message="非空数据行没有 KKS 编码。", suggestion="确认是待编设备、草稿行还是应删除的空记录；不直接删除。")

    records_by_code = {code: group[0] for code, group in by_code.items() if len(group) == 1}
    for code in cycle_nodes(records_by_code, root_parents):
        add_issue(issues, priority="P0", rule_id="KKS-10", category="父级环", rec=records_by_code[code],
                  message="父级链存在环。", suggestion="修正父级指针。")

    old_to_new: dict[str, set[str]] = defaultdict(set)
    for rec in records:
        if rec["old_code"]:
            old_to_new[rec["old_code"]].add(rec["kks_code"])
    for old_code, new_codes in old_to_new.items():
        if len(new_codes) > 1:
            for rec in records:
                if rec["old_code"] == old_code:
                    add_issue(issues, priority="P1", rule_id="KKS-25c", category="同一原码对应多个新码", rec=rec,
                              message=f"原 KKS {old_code} 对应多个新码：{', '.join(sorted(new_codes))}。", suggestion="用原码作为 1:1 身份键人工裁定，不自动合并。")

    # Preserve a structural note rather than treating intentionally blank separator rows as records.
    notes = []
    if len(sheet_info) > 1:
        empty_sheets = [x["sheet"] for x in sheet_info if x["nonempty_rows"] == 0]
        if empty_sheets:
            notes.append(f"空 Sheet 未参与审核：{', '.join(empty_sheets)}")
    if blank_rows:
        notes.append(f"主表跳过空白行 {len(blank_rows)} 行（示例：第 {blank_rows[0]} 行），未提前终止扫描。")
    notes.append("主表首个数据行被识别为有效根节点并纳入统计；这比附件模板的固定 data_start=2 更稳健。")

    priority_counts = Counter(i["priority"] for i in issues if i["status"] != "resolved")
    issue_count = sum(1 for item in issues if item["status"] != "resolved")
    conclusion = (
        f"本次共发现 {issue_count} 个问题，请按问题清单逐项核对；源 Excel 未修改。"
        if issue_count
        else "本次未发现需要列入问题清单的问题；源 Excel 未修改。"
    )
    result = {
        "version": VERSION,
        "source_file": str(input_path),
        "sheet": sheet,
        "header_row": header_idx + 1,
        "data_rows": len(records),
        "nonempty_rows_without_code": len(nonempty_without_code),
        "columns": cols,
        "sheet_info": sheet_info,
        "code_length_distribution": dict(sorted(Counter(len(r["kks_code"]) for r in records).items())),
        "metrics": {
            "unique_codes": len(code_set),
            "duplicate_code_groups": sum(1 for group in by_code.values() if len(group) > 1),
            "duplicate_suffix_loss_groups": skill_template_diagnostics["duplicate_diagnosis"].get("same_source_suffix_loss_groups", 0),
            "duplicate_cross_device_collision_groups": skill_template_diagnostics["duplicate_diagnosis"].get("cross_device_collision_groups", 0),
            "orphan_rows": len(orphan_recs),
            "prefix_mismatch_rows": len(prefix_bad),
            "blank_separator_rows": len(blank_rows),
            "long_code_rows": sum(1 for r in records if len(r["kks_code"]) > 12),
            "unit_rule_hits_resolved": len(resolved),
            "sequence_gap_groups": additional_skill_diagnostics.get("sequence_gap_groups", 0),
        },
        "priority_counts": dict(priority_counts),
        "issue_count": issue_count,
        "resolved_count": len(resolved),
        "conclusion": conclusion,
        "structural_notes": notes,
        "issues": issues,
        "resolved_reviews": resolved,
        "skill_coverage": {
            "source": SKILL_RULE_SOURCE,
            "template_loaded": True,
            "template_rules": skill_template_diagnostics,
            "additional_rules": additional_skill_diagnostics,
            "config_file": "config/kks_rules.json",
        },
        "import_governance": {
            "target": "DM8 LOCATIONS",
            "direct_import_ready": False,
            "long_code_requires_extension_model": bool(sum(1 for r in records if len(r["kks_code"]) > 12)),
            "historical_kks_rewrite_forbidden": True,
            "required_strategy": "保留历史 KKS；通过映射层/双码共存/冻结历史兼容两票与缺陷记录。",
            "runtime_dm8_validation": False,
        },
        "scope_comparison": scope_comparison,
    }
    # AI is an optional second stage.  With no AI_API_KEY the local audit is
    # unchanged; with a key, all semantic candidates are submitted together.
    _emit_progress(progress_callback, {"phase": "audit", "stage": "rules_completed", "percent": 62, "message": "本地规则审核完成", "hint": f"已发现 {len(issues)} 个规则问题，准备进行 AI 候选复核"})
    result["ai_review"] = review_issue_candidates(result, progress_callback=progress_callback)
    _emit_progress(progress_callback, {"phase": "audit", "stage": "report", "percent": 95, "message": "正在生成审核报告", "hint": "整理 HTML 报告和问题 Excel"})
    output_dir.mkdir(parents=True, exist_ok=True)
    html_name, xlsx_name = output_artifact_names(input_path)
    write_html(output_dir / html_name, result)
    write_issue_workbook_xlsx(output_dir / xlsx_name, result)
    LOGGER.info(
        "audit_done file=%s rows=%s issues=%s ai_status=%s ai_reviewed=%s",
        input_path.name,
        result["data_rows"],
        result.get("issue_count", 0),
        result.get("ai_review", {}).get("status", "disabled"),
        result.get("ai_review", {}).get("reviewed_count", 0),
    )
    return result


def write_issue_workbook_xlsx(path: Path, result: dict[str, Any]) -> None:
    """生成独立的问题清单工作簿，不修改上传的源 Excel。"""
    wb = Workbook()
    summary = wb.active
    summary.title = "概览"
    issues_ws = wb.create_sheet("问题清单")
    resolved_ws = wb.create_sheet("已澄清项")
    structure_ws = wb.create_sheet("文件结构")
    coverage_ws = wb.create_sheet("规则覆盖")
    teal, light_teal = "0F4C5C", "EAF4F5"
    gray = "F4F6F7"

    def header_style(ws, cell_range: str) -> None:
        for row in ws[cell_range]:
            for cell in row:
                cell.fill = PatternFill("solid", fgColor=teal)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.alignment = Alignment(vertical="center", wrap_text=True)

    def body_style(ws, cell_range: str) -> None:
        for row in ws[cell_range]:
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)

    def autosize(ws, maximum: int = 52) -> None:
        for col_idx in range(1, ws.max_column + 1):
            letter = get_column_letter(col_idx)
            max_len = max((len(str(ws.cell(row=row_idx, column=col_idx).value or "")) for row_idx in range(1, ws.max_row + 1)), default=0)
            ws.column_dimensions[letter].width = min(max(max_len + 2, 12), maximum)

    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False

    summary.merge_cells("A1:F1")
    summary["A1"] = "KKS 编码审核报告"
    summary["A1"].fill = PatternFill("solid", fgColor=teal)
    summary["A1"].font = Font(bold=True, color="FFFFFF", size=16)
    summary["A1"].alignment = Alignment(horizontal="center")
    summary.merge_cells("A2:F2")
    summary["A2"] = f"源文件：{result['source_file']}；主表：{result['sheet']}；表头行：{result['header_row']}"
    summary["A2"].fill = PatternFill("solid", fgColor=light_teal)
    summary["A2"].alignment = Alignment(wrap_text=True)
    summary.append([])
    summary.append(["指标", "值"])
    header_style(summary, "A4:B4")
    for label, value in (("有效编码行", result["data_rows"]), ("唯一 KKS 码", result["metrics"]["unique_codes"]), ("问题总数", result.get("issue_count", 0)), ("已澄清项", result.get("resolved_count", len(result.get("resolved_reviews", [])))), ("父级孤儿", result["metrics"]["orphan_rows"]), ("父子前缀不一致", result["metrics"]["prefix_mismatch_rows"]), ("12 位以上扩展码", result["metrics"]["long_code_rows"]), ("AI 复核候选", result.get("ai_review", {}).get("candidate_count", 0))):
        summary.append([label, value])
    body_style(summary, "A5:B12")
    summary.merge_cells("A14:F14")
    summary["A14"] = "总体结论"
    summary["A14"].fill = PatternFill("solid", fgColor=teal)
    summary["A14"].font = Font(bold=True, color="FFFFFF")
    summary.merge_cells("A15:F16")
    summary["A15"] = result["conclusion"]
    summary["A15"].fill = PatternFill("solid", fgColor=light_teal)
    summary["A15"].alignment = Alignment(vertical="center", wrap_text=True)
    summary.merge_cells("A18:F18")
    summary["A18"] = "结构与使用说明"
    summary["A18"].fill = PatternFill("solid", fgColor=teal)
    summary["A18"].font = Font(bold=True, color="FFFFFF")
    notes = result["structural_notes"] + ["源 Excel 未修改；13 位扩展码和历史映射需人工确认。", "本报告未做 DM8/LOCATIONS 真实导入验证。", "正式 Skill 模板已加载并执行；AI 对全部语义候选项进行二次复核。"]
    for row_no, note in enumerate(notes, start=19):
        summary.merge_cells(start_row=row_no, start_column=1, end_row=row_no, end_column=6)
        summary.cell(row_no, 1).value = f"• {note}"
        summary.cell(row_no, 1).fill = PatternFill("solid", fgColor=gray)
        summary.cell(row_no, 1).alignment = Alignment(wrap_text=True, vertical="top")
    summary.freeze_panes = "A5"
    autosize(summary)

    issue_headers = ["状态", "规则", "类别", "Excel行号", "KKS码", "父级码", "原KKS码", "设备名称", "审核意见", "建议", "证据", "AI决策", "AI置信度", "AI理由", "AI建议"]
    issues_ws.append(issue_headers)
    for item in result["issues"]:
        issues_ws.append([item.get(key, "") for key in ("status", "rule_id", "category", "excel_row", "kks_code", "parent_code", "old_code", "name", "message", "suggestion", "evidence", "ai_decision", "ai_confidence", "ai_reason", "ai_suggestion")])
    end_row = max(1, issues_ws.max_row)
    header_style(issues_ws, "A1:O1")
    body_style(issues_ws, f"A1:O{end_row}")
    issues_ws.freeze_panes = "A2"
    issues_ws.auto_filter.ref = f"A1:O{end_row}"
    autosize(issues_ws)
    issue_table = Table(displayName="KKSIssues", ref=f"A1:O{end_row}")
    issue_table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    issues_ws.add_table(issue_table)

    resolved_headers = ["状态", "规则", "类别", "Excel行号", "KKS码", "父级码", "原KKS码", "设备名称", "复核结论", "建议", "证据"]
    resolved_ws.append(resolved_headers)
    for item in result["resolved_reviews"]:
        resolved_ws.append([item.get(key, "") for key in ("status", "rule_id", "category", "excel_row", "kks_code", "parent_code", "old_code", "name", "message", "suggestion", "evidence")])
    resolved_end = max(1, resolved_ws.max_row)
    header_style(resolved_ws, "A1:K1")
    body_style(resolved_ws, f"A1:K{resolved_end}")
    resolved_ws.freeze_panes = "A2"
    resolved_ws.auto_filter.ref = f"A1:K{resolved_end}"
    autosize(resolved_ws)
    resolved_table = Table(displayName="KKSResolved", ref=f"A1:K{resolved_end}")
    resolved_table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    resolved_ws.add_table(resolved_table)

    structure_headers = ["Sheet", "最大行", "最大列", "非空行", "首个非空行", "最后非空行", "表头预览"]
    structure_ws.append(structure_headers)
    for item in result["sheet_info"]:
        preview = " | ".join(str(x) for x in item.get("header_preview", []))
        structure_ws.append([item.get(key, "") for key in ("sheet", "max_row", "max_col", "nonempty_rows", "first_nonempty_row", "last_nonempty_row")] + [preview])
    structure_end = max(1, structure_ws.max_row)
    header_style(structure_ws, "A1:G1")
    body_style(structure_ws, f"A1:G{structure_end}")
    structure_ws.freeze_panes = "A2"
    structure_ws.auto_filter.ref = f"A1:G{structure_end}"
    autosize(structure_ws)
    structure_table = Table(displayName="KKSSheets", ref=f"A1:G{structure_end}")
    structure_table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    structure_ws.add_table(structure_table)

    coverage_headers = ["维度", "规则", "执行状态", "命中数", "执行/证据来源"]
    coverage_ws.append(coverage_headers)
    issue_counts = Counter(item.get("rule_id", "") for item in result.get("issues", []))
    coverage_items = [
        ("文件与结构", "01-03", "已执行", "动态 Sheet、表头和列定位", "kks-audit/SKILL.md"),
        ("语法与格式", "KKS-04", "已执行", "非法字符", "kks-audit/SKILL.md"),
        ("语法与格式", "KKS-04b", "已执行", "12 位骨架分段类型", "kks-audit/scripts/audit_template.py"),
        ("语法与格式", "KKS-05", "已执行", "码长和变长层级", "kks-audit/SKILL.md"),
        ("语法与格式", "KKS-06/KKS-29", "已执行", "I/O 与 OCR 易混字符", "kks-audit/scripts/audit_template.py"),
        ("唯一性", "KKS-07", "已执行", "重复码、后缀丢失/跨设备诊断", "kks-audit/scripts/audit_template.py"),
        ("层级树", "KKS-08~KKS-12", "已执行", "孤儿、前缀、自引用/环、父级长度、根哨兵", "kks-audit/SKILL.md"),
        ("国标迁移", "KKS-13/14/15/16/18", "已执行", "前缀、字母、层级、旧码与历史提示", "kks-audit/SKILL.md + config/kks_rules.json"),
        ("语义深度", "KKS-17/22", "已执行", "信号点、部件级、A/B/C、区间扩展分类", "kks-audit/SKILL.md"),
        ("语义一致", "KKS-24~KKS-31", "已执行", "命名、同物异名、应编未编、卫生和父子语义", "kks-audit/scripts/audit_template.py"),
        ("轻量增强", "KKS-A/C/E/F/G", "已执行", "名称字母、风格、设备类型、断号、字母组合", "kks-audit/SKILL.md"),
        ("跨范围对比", "KKS-19/20/21", "已执行" if result.get("scope_comparison") else "未提供对比文件", "专业覆盖、编码深度、数量瀑布闭合", "--compare-file + kks-audit/SKILL.md"),
        ("DM8 可导入性", "KKS-22/23", "已执行", "LOCATIONS 约束和历史双码治理提示", "kks-audit/SKILL.md"),
    ]
    for dimension, rule_id, status, hit_note, source in coverage_items:
        hit_value = sum(issue_counts.get(rule, 0) for rule in re.split(r"[/~]", rule_id) if issue_counts.get(rule))
        if rule_id == "KKS-07":
            hit_value = issue_counts.get("KKS-07", 0)
        coverage_ws.append([dimension, rule_id, status, hit_value, f"{hit_note}；{source}"])
    coverage_end = max(1, coverage_ws.max_row)
    header_style(coverage_ws, "A1:E1")
    body_style(coverage_ws, f"A1:E{coverage_end}")
    coverage_ws.freeze_panes = "A2"
    coverage_ws.auto_filter.ref = f"A1:E{coverage_end}"
    autosize(coverage_ws)
    coverage_table = Table(displayName="KKSCoverage", ref=f"A1:E{coverage_end}")
    coverage_table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    coverage_ws.add_table(coverage_table)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def write_html(path: Path, result: dict[str, Any]) -> None:
    rows = [i for i in result["issues"] if i["status"] != "resolved"]
    ai = result.get("ai_review", {})
    rows_html = "".join(
        "<tr>" + "".join(f"<td>{html.escape(str(item.get(k, '')))}</td>" for k in ("status", "rule_id", "category", "excel_row", "kks_code", "name", "message", "suggestion")) +
        f"<td>{html.escape(str(item.get('ai_decision', '')))}" +
        (f"（{float(item['ai_confidence']):.0%}）：{html.escape(str(item.get('ai_reason', '')))}" if item.get("ai_decision") else "") +
        "</td></tr>"
        for item in rows
    ) or '<tr><td colspan="9">未发现问题</td></tr>'
    resolved_html = "".join(
        f"<tr><td>{item['excel_row']}</td><td>{html.escape(item['kks_code'])}</td><td>{html.escape(item['message'])}</td></tr>"
        for item in result["resolved_reviews"]
    ) or '<tr><td colspan="3">无</td></tr>'
    notes = "".join(f"<li>{html.escape(str(x))}</li>" for x in result["structural_notes"])
    coverage = result.get("skill_coverage", {})
    template_rules = coverage.get("template_rules", {}).get("executed_rules", [])
    additional_rules = coverage.get("additional_rules", {}).get("executed_rules", [])
    comparison = result.get("scope_comparison")
    comparison_html = ""
    if comparison:
        waterfall = comparison["quantity_waterfall"]
        comparison_html = (
            "<h2>跨范围对比</h2>"
            f"<p>对比文件：{html.escape(str(comparison['comparison']['file']))}；"
            f"有效编码数量差：{waterfall['primary_minus_comparison']}；"
            f"数量瀑布闭合：{html.escape(str(waterfall['closed']))}，闭合差：{waterfall['closure_delta']}。</p>"
            f"<pre>{html.escape(json.dumps(comparison, ensure_ascii=False, indent=2))}</pre>"
        )
    body = f"""<!doctype html><html lang='zh-CN'><head><meta charset='utf-8'><title>KKS 编码审核报告</title>
<style>body{{font-family:Segoe UI,Microsoft YaHei,sans-serif;color:#243447;margin:32px;line-height:1.5}}h1{{color:#0f4c5c}}h2{{border-bottom:2px solid #d7e7ea;padding-bottom:6px}}.kpi{{display:inline-block;min-width:130px;margin:6px;padding:14px 18px;border-radius:8px;background:#eef6f7}}table{{border-collapse:collapse;width:100%;font-size:13px}}th,td{{border-bottom:1px solid #dfe7ea;padding:7px;text-align:left;vertical-align:top}}th{{background:#0f4c5c;color:white;position:sticky;top:0}}.small{{color:#5b6770;font-size:13px}}</style></head><body>
<h1>KKS 编码审核报告</h1><p><b>源文件：</b>{html.escape(result['source_file'])}<br><b>主表：</b>{html.escape(result['sheet'])}；<b>表头行：</b>{result['header_row']}；<b>有效编码行：</b>{result['data_rows']}</p>
<p><b>总体结论：</b>{html.escape(result['conclusion'])}</p>
<div class='kpi'><b>问题总数</b><br><span>{result.get('issue_count', len(rows))}</span></div><div class='kpi'><b>已澄清项</b><br><span>{result.get('resolved_count', len(result['resolved_reviews']))}</span></div><div class='kpi'><b>AI 复核候选</b><br><span>{ai.get('candidate_count', 0)}</span></div><div class='kpi'><b>AI 已复核</b><br><span>{ai.get('reviewed_count', 0)}</span></div>
<h2>AI 二次复核</h2><p>{html.escape(str(ai.get('status', 'disabled')))}；候选项 {ai.get('candidate_count', 0)}，已复核 {ai.get('reviewed_count', 0)}，模型：{html.escape(str(ai.get('model') or '未配置'))}。AI 结果只作为复核建议，不自动关闭问题。</p>
<h2>关键指标</h2><ul><li>唯一 KKS 码：{result['metrics']['unique_codes']}；重复码组：{result['metrics']['duplicate_code_groups']}</li><li>父级孤儿：{result['metrics']['orphan_rows']}；父子前缀不一致：{result['metrics']['prefix_mismatch_rows']}</li><li>12 位以上扩展码：{result['metrics']['long_code_rows']}；长度分布：{html.escape(json.dumps(result['code_length_distribution'], ensure_ascii=False))}</li></ul>
<h2>Skill 执行覆盖</h2><p>规则来源：{html.escape(str(coverage.get('source', 'kks-audit/SKILL.md')))}；正式模板已加载：{html.escape(str(coverage.get('template_loaded', False)))}。</p><p>模板规则：{html.escape(', '.join(template_rules))}<br>增强规则：{html.escape(', '.join(additional_rules))}<br>DM8 导入策略：保留历史 KKS，使用映射层/双码共存；本次未连接真实 DM8。</p>
<h2>结构说明</h2><ul>{notes}</ul>
{comparison_html}
<h2>问题清单（全部问题）</h2><table><thead><tr><th>状态</th><th>规则</th><th>类别</th><th>行号</th><th>KKS</th><th>名称</th><th>审核意见</th><th>建议</th><th>AI复核</th></tr></thead><tbody>{rows_html}</tbody></table>
<h2>规则命中但已语义排除</h2><table><thead><tr><th>行号</th><th>KKS</th><th>复核结论</th></tr></thead><tbody>{resolved_html}</tbody></table>
<p class='small'>本报告由规则扫描与可选 AI 二次复核生成；源 Excel 未修改。未连接 DM8/LOCATIONS 做真实导入验证，AI 结果仅为候选判断，需人工确认。</p></body></html>"""
    path.write_text(body, encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="审核 KKS Excel 并生成 HTML 报告和 Excel 问题清单")
    parser.add_argument("input", type=Path)
    parser.add_argument("--output-dir", type=Path, default=Path("outputs/kks-audit"))
    parser.add_argument("--sheet", default=None)
    parser.add_argument("--compare-file", type=Path, default=None, help="可选：用于专业覆盖、编码深度和数量闭合对比的另一份 Excel")
    args = parser.parse_args()
    result = audit_file(args.input, args.output_dir, args.sheet, args.compare_file)
    print(json.dumps({"conclusion": result["conclusion"], "issue_count": result.get("issue_count", 0), "resolved_count": result.get("resolved_count", 0), "output_dir": str(args.output_dir)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
