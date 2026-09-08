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


VERSION = "0.4.0"
LOGGER = logging.getLogger("kks-audit")
ProgressCallback = Callable[[dict[str, Any]], None]
ROOT_PARENTS = {"", "-1"}
# 10 版旧前缀历史参考（G=5/6/L/J 本身合法，迁移为项目约定，不再作为 P1 阻断）
OLD_PREFIX = {"50": "05", "60": "06", "L0": "61", "J0": "61"}
# 全厂码 G：公用取值（期别公用 J-R / 多期公用 S-V / 全厂公用 Y）
COMMON_UNITS = {"J", "K", "L", "M", "N", "P", "Q", "R", "S", "T", "U", "V", "Y"}
# 全厂码 G：自由使用字母
FREE_UNIT_LETTERS = {"H", "W", "X", "Z"}
# 全厂码 G：机组映射（1-9 → 1~9 号；A-G → 10~16 号）
UNIT_LETTER_MAP = {"A": 10, "B": 11, "C": 12, "D": 13, "E": 14, "F": 15, "G": 16}
ALLOWED_CODE = re.compile(r"^[A-Z0-9-]+$")
NAME_UNIT = re.compile(r"(\d{1,2}|[一二三四五六七八九十])\s*号\s*(机|炉|机组)")
CN_NUM = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}


def output_artifact_names(input_path: Path) -> tuple[str, str]:
    """Return user-facing report names derived from the uploaded workbook."""
    stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", input_path.stem).strip(" .")
    stem = (stem or "审核文件")[:120]
    return f"{stem}_审核报告.html", f"{stem}_问题清单.xlsx"


FINAL_DECISION_LABELS = {
    "confirmed_issue": "确认问题",
    "needs_human": "待人工确认",
    "likely_false_positive": "疑似误报",
}


def final_decision_label(value: Any) -> str:
    return FINAL_DECISION_LABELS.get(str(value), "待人工确认")


def final_decision_counts(result: dict[str, Any]) -> dict[str, int]:
    counts: Counter[str] = Counter()
    for item in result.get("issues", []):
        if isinstance(item, dict) and item.get("status") != "resolved":
            counts[str(item.get("final_decision", "confirmed_issue"))] += 1
    return dict(sorted(counts.items()))


REPORT_METHODS = (
    "KKS 编码规则校验",
    "层级结构校验",
    "唯一性校验",
    "父子关系校验",
    "历史编码映射校验",
    "名称规范校验",
    "扩展编码分析",
)


def _report_issues(result: dict[str, Any]) -> list[dict[str, Any]]:
    return [
        item for item in result.get("issues", [])
        if isinstance(item, dict) and item.get("status") != "resolved"
    ]


# ---------------------------------------------------------------- 问题分组树
# 明细问题文本复用函数（HTML/Excel/Web 展示统一）
_ACTION_LABELS = {
    "confirmed_issue": "需整改",
    "needs_human": "需确认",
    "likely_false_positive": "建议抽查",
}


def action_label(item: dict[str, Any]) -> str:
    """按 AI 复核结论给出问题处置动作标签（需整改/需确认/建议抽查/待确认）。"""
    return _ACTION_LABELS.get(str(item.get("final_decision", "")), "待确认")


def detail_text(item: dict[str, Any]) -> str:
    """问题的具体描述（AI 复核后优先采用最终结论文本）。"""
    return str(item.get("final_summary") or item.get("message") or item.get("category") or "待核问题")


def suggestion_text(item: dict[str, Any]) -> str:
    """问题的整改建议。"""
    return str(item.get("final_suggestion") or item.get("suggestion") or "请结合原始 Excel 和业务资料确认。")


def _excel_row_num(item: dict[str, Any]) -> int:
    v = str(item.get("excel_row", "") or "").strip()
    return int(v) if v.lstrip("-").isdigit() else 0


def issue_rule_tree(result: dict[str, Any]) -> list[dict[str, Any]]:
    """把问题按「规则 → KKS → 行号明细」组织成展示树（HTML/Web 通用）。

    返回轻量可序列化节点列表：
      kind: "rule" | "code"；level 0/1
      rule 节点: title=规则号  sub=问题定义  action=处置标签(需整改/需确认/建议抽查)
                source=规则依据(文件名+章节)  children=该规则涉及的 KKS（去重，每个只出现一次）
      code 节点: title=KKS 编码  sub=设备名  issues=行号明细(行号/问题描述/整改建议)
    p0/p1/p2/count 为子树聚合计数。规则级属性上提到规则节点，明细行不再重复。
    """
    standard = str(result.get("standard", STANDARD_DEFAULT))
    priority_order = {"P0": 0, "P1": 1, "P2": 2}
    issues = sorted(
        _report_issues(result),
        key=lambda it: (
            priority_order.get(str(it.get("priority")), 9),
            str(it.get("rule_id", "")),
            str(it.get("kks_code", "")),
            _excel_row_num(it),
        ),
    )

    rules: list[dict[str, Any]] = []
    rule_index: dict[str, dict[str, Any]] = {}

    def bump(node: dict[str, Any], pr: str) -> None:
        if pr == "P0":
            node["p0"] += 1
        elif pr == "P1":
            node["p1"] += 1
        else:
            node["p2"] += 1

    for item in issues:
        rid = str(item.get("rule_id") or "未编号规则")
        pr = str(item.get("priority", "P2"))
        rule = rule_index.get(rid)
        if rule is None:
            rule = {
                "kind": "rule", "level": 0,
                "title": rid,
                "sub": str(item.get("definition") or item.get("category") or ""),
                "note": "",
                "action": action_label(item),
                "source": rule_source(rid, standard),
                "priority": pr,
                "p0": 0, "p1": 0, "p2": 0, "count": 0,
                "issues": [], "children": [],
            }
            rules.append(rule)
            rule_index[rid] = rule

        code = str(item.get("kks_code") or "").strip()
        node: dict[str, Any] | None = None
        for child in rule["children"]:  # 同一规则下编码数量有限，线性查找即可
            if child["title"] == code:
                node = child
                break
        if node is None:
            node = {
                "kind": "code", "level": 1,
                "title": code or "未挂接编码的问题",
                "sub": str(item.get("name") or ""),
                "note": "",
                "action": "", "source": "",
                "priority": pr,
                "p0": 0, "p1": 0, "p2": 0, "count": 0,
                "issues": [], "children": [],
            }
            rule["children"].append(node)

        node["issues"].append({
            "priority": pr,
            "rule_id": rid,
            "kks_code": code,
            "name": str(item.get("name") or ""),
            "excel_row": item.get("excel_row", ""),
            "message": detail_text(item),
            "suggestion": suggestion_text(item),
            "status": item.get("status", ""),
            "ai_decision": item.get("ai_decision", ""),
        })
        bump(rule, pr)
        bump(node, pr)
        rule["count"] += 1
        node["count"] += 1

    for rule in rules:
        rule["note"] = f"{len(rule['children'])} 个编码 · {rule['count']} 条"
        for node in rule["children"]:
            node["note"] = f"{node['count']} 条"
    return rules


def rule_tree_html(nodes: list[dict[str, Any]]) -> str:
    """把 issue_rule_tree 的分组树渲染为 <details> 折叠 HTML。

    规则节点默认展开（呈现该规则涉及的 KKS 清单）；KKS 节点默认收起，
    点击展开后才显示行号明细——默认页面短，且规则字段/KKS 不再重复出现。
    """
    parts = []
    for node in nodes:
        kind = node.get("kind")
        title = html.escape(str(node.get("title") or ""))
        sub = html.escape(str(node.get("sub") or ""))
        note = html.escape(str(node.get("note") or ""))
        badge = "".join(
            f"<span class='tb tb-p{c}'>{node.get('p' + c, 0)}</span>"
            for c in ("0", "1", "2") if node.get("p" + c, 0)
        )
        if kind == "code":
            rows = node.get("issues") or []
            summary = (
                f"<span class='it-name'>{title or '未挂接编码的问题'}</span>"
                + (f"<span class='it-title'>{sub}</span>" if sub else "")
                + f"<span class='it-count'>{note}</span>{badge}"
            )
            inner = ""
            if rows:
                inner = "<div class='it-issues'>" + "".join(
                    f"<div class='it-issue'><span class='it-row'>行号 {html.escape(str(row.get('excel_row', '') or '—'))}</span>"
                    f"<span class='it-msg'>{html.escape(str(row.get('message') or '待核问题'))}</span>"
                    f"<div class='it-suggest'>建议：{html.escape(str(row.get('suggestion') or ''))}</div></div>"
                    for row in rows
                ) + "</div>"
            parts.append(f"<details class='it-node it-code'><summary>{summary}</summary>{inner}</details>")
            continue

        # 规则节点：P 徽标 + 规则号 + 问题定义 + 处置标签 + 计数；展开出 KKS 清单
        p = node.get("priority", "P2")
        act = html.escape(str(node.get("action") or "待确认"))
        act_cls = {"需整改": "act-fix", "需确认": "act-hm"}.get(str(node.get("action")), "act-ck")
        src = html.escape(str(node.get("source") or ""))
        summary = (
            f"<span class='priority p{str(p)[-1]}'>{html.escape(str(p))}</span>"
            f"<span class='it-name tg-rid'>{title}</span>"
            f"<span class='tg-def'>{sub}</span>"
            f"<span class='tg-act {act_cls}'>{act}</span>"
            f"<span class='it-count'>{note}</span>"
        )
        kids = rule_tree_html(node.get("children") or [])
        src_html = f"<div class='tg-src'>规则依据：{src}</div>" if src else ""
        parts.append(
            f"<details class='it-node tg-rule' open><summary>{summary}</summary>{src_html}"
            f"<div class='it-kids'>{kids}</div></details>"
        )
    return "".join(parts)





def quality_metrics(result: dict[str, Any]) -> dict[str, int]:
    """Return business-facing quality metrics used by the report and UI.

    P0/P1/P2 均为真实规则级计数（与详细问题清单自洽）。
    另保留两个业务治理重点的聚焦计数（扩展码治理 KKS-17/22、历史迁移 KKS-18）
    供报告明细作为治理提示，不作为首页大数字口径。
    """
    records = result.get("_ai_context", {}).get("records", [])
    codes = [text(item.get("kks_code")) for item in records if isinstance(item, dict)]
    metrics = result.get("metrics", {})
    priority_counts = Counter(item.get("priority", "P2") for item in _report_issues(result))
    return {
        "effective_kks": int(result.get("data_rows", len(codes)) or 0),
        "device_level_codes": sum(1 for code in codes if len(code) == 12),
        "duplicate_codes": int(metrics.get("duplicate_code_groups", 0) or 0),
        "parent_errors": int(metrics.get("orphan_rows", 0) or 0) + int(metrics.get("prefix_mismatch_rows", 0) or 0),
        "p0": int(priority_counts.get("P0", 0)),
        "p1": int(priority_counts.get("P1", 0)),
        "p2": int(priority_counts.get("P2", 0)),
        "rule_p1": int(priority_counts.get("P1", 0)),
        "rule_p2": int(priority_counts.get("P2", 0)),
        "focused_p1": _rule_issue_count(result, {"KKS-17", "KKS-22"}),
        "focused_p2": _rule_issue_count(result, {"KKS-18"}),
    }


def _rule_issue_count(result: dict[str, Any], rule_ids: set[str]) -> int:
    total = 0
    for item in _report_issues(result):
        rules = {part.strip() for part in re.split(r"[/~]", str(item.get("rule_id", "")))}
        if rules & rule_ids:
            total += 1
    return total


def report_quality_dimensions(result: dict[str, Any]) -> list[dict[str, str]]:
    """Build the eight quality dimensions shown to non-technical reviewers."""
    metrics = result.get("metrics", {})
    dimensions: list[dict[str, str]] = []

    def add(dimension: str, check: str, count: int, detail: str, *, zero: str = "通过") -> None:
        dimensions.append({
            "dimension": dimension,
            "check": check,
            "result": zero if count == 0 else f"发现 {count} 项",
            "detail": detail,
        })

    structure_ok = bool(result.get("sheet_info")) and bool(result.get("columns"))
    dimensions.append({
        "dimension": "结构层",
        "check": "表结构、字段识别",
        "result": "通过" if structure_ok else "待处理",
        "detail": f"主表 {result.get('sheet', '—')}；已识别表头和 KKS、父级、名称字段。" if structure_ok else "未能完整识别主表结构。",
    })
    add("语法层", "字符合法性", _rule_issue_count(result, {"KKS-04", "KKS-04b", "KKS-05", "KKS-06", "KKS-29", "KKS-30"}), "检查字符、分段类型、长度和文本卫生。")
    duplicate_count = int(metrics.get("duplicate_code_groups", 0) or 0)
    add("唯一性", "重复 KKS", duplicate_count, "按完整 KKS 码分组，未将不同层级的相似码误判为重复。")
    parent_count = int(metrics.get("orphan_rows", 0) or 0) + int(metrics.get("prefix_mismatch_rows", 0) or 0)
    add("层级", "父级关系", parent_count, "检查孤儿节点、父子归属和层级关系。")
    migration_count = _rule_issue_count(result, {"KKS-13", "KKS-14", "KKS-15", "KKS-16", "KKS-18", "KKS-25c"})
    add("迁移", "旧码映射", migration_count, "检查历史原码、新码前缀变化和一对多映射。")
    semantic_count = max(int(metrics.get("long_code_rows", 0) or 0), _rule_issue_count(result, {"KKS-17", "KKS-22", "KKS-A", "KKS-C", "KKS-E", "KKS-F", "KKS-G"}))
    add("语义", "扩展码分析", semantic_count, "检查扩展编码、设备类型、编号连续性和语义疑点。")
    naming_count = _rule_issue_count(result, {"KKS-24", "KKS-25", "KKS-25b", "KKS-26", "KKS-27", "KKS-28", "KKS-31"})
    add("命名", "名称规范", naming_count, "检查名称歧义、同物异名、机组三方一致和父子语义。")
    import_governance = result.get("import_governance", {})
    import_ready = bool(import_governance.get("direct_import_ready"))
    dimensions.append({
        "dimension": "导入",
        "check": "数据库兼容性",
        "result": "通过" if import_ready else "待处理",
        "detail": "已通过导入前检查。" if import_ready else "本次未连接真实 DM8；扩展码、历史映射和 P0/P1 问题需处理后再导入。",
    })
    return dimensions


def report_import_conclusion(result: dict[str, Any]) -> str:
    metrics = quality_metrics(result)
    if metrics["p0"] or metrics["p1"] or not result.get("import_governance", {}).get("direct_import_ready", False):
        return "当前不建议直接导入。完成 P0/P1 整改，确认历史映射和扩展编码，并通过目标库导入验证后再进入系统。"
    return "本批 KKS 编码整体结构完整，可在完成导入前抽查和目标库验证后进入系统导入。"


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

# 规则问题定义表：每条审核规则的一句话定义（判定语义）。
# 由 references/KKS编码审核标准检查清单.md 的"校验原理"提炼，
# 作为 HTML 报告 / Excel 问题清单 / Web 预览"问题定义"列的统一数据源。
RULE_DEFINITIONS: dict[str, str] = {
    "KKS-00": "数据行缺少 KKS 编码：非空记录应含设备 KKS 码，缺失则无法入库定位。",
    "KKS-01": "多 Sheet 定位主表：应区分主表与导出/草稿表，防止重复或漏审。",
    "KKS-02": "表头与数据起始行定位：表头不总在第 1 行，错读将导致整列错位。",
    "KKS-03": "关键列动态定位：按表头名匹配 KKS/父级/名称/类型列，不依赖固定列序。",
    "KKS-04": "字符合法性：仅允许大写字母 A-Z、数字 0-9 与分隔符 -；小写、中文、全角及 ~、/ 等未定义字符判非法。",
    "KKS-04b": "分段字符类型校验：12 位骨架逐段定类型——G/F1F2F3/A1A2 为字母段，F0/FN/AN 为数字段，数字段出现字母或字母段出现数字即错。",
    "KKS-05": "编码长度合规：设备级 12 位、系统级 5/6/7 位、根节点 1 位；超过 12 位不立即判错，转编码深度定性。",
    "KKS-06": "结构位取值合规：全厂码 G 须在 1-9/A-G/J-R/S-V/Y 或自由字母 H/W/X/Z 取值内；编码不得含 I/O 等易混字母。",
    "KKS-07": "重复 KKS 码：主键须唯一；区分位置后缀丢失（可机械修复）与异设备真碰撞（须人工重编）。",
    "KKS-08": "父级参照完整性：非根节点父级码须作为某行真实存在，否则层级断链、破坏外键。",
    "KKS-09": "子码前缀一致性：子码须以父码为前缀（结合父级列值判定，不按固定位数截取推导）。",
    "KKS-10": "自引用/环检测：父级不得指向自身，编码树不得形成闭合环。",
    "KKS-11": "父级长度约束：父码长度应小于子码（父码应是子码的短前缀）。",
    "KKS-12": "根节点哨兵：根节点父级应为空或 -1，不得指向不存在的编码。",
    "KKS-13/16": "旧前缀迁移提示：50/60/L0/J0 等旧前缀编码仅作历史提示，需人工核对是否迁移遗漏。",
    "KKS-14": "字母重分配：系统/设备类型字母变更应在重分配映射表内，避免随意改字母导致语义漂移。",
    "KKS-15": "层级重排：版本间同一设备的层级位置调整，须人工核验语义未变。",
    "KKS-16": "旧码遗留：父级或个别码仍命中旧前缀表，提示核对是否为历史遗留。",
    "KKS-17/22": "12 位以上扩展码定性：区分纯 DCS 点号（-KF）、部件级附加码（-XXnn）与电气 A/B/C 分相，评估导入方式。",
    "KKS-18": "旧→新码语义一致：迁移后设备语义应与原码对应，避免码变设备不对。",
    "KKS-19/20/21": "跨厂/范围对比：专业覆盖差异、编码深度差异与数量瀑布分解闭合。",
    "KKS-22": "码长与导入结构：超过 12 位的编码无法原样套用 LOCATIONS 结构，需扩展表/分层承载。",
    "KKS-23": "两票/缺陷历史兼容：纸质票面已打印的历史 KKS 不可改写，采用映射层+双码共存+冻结历史。",
    "KKS-24": "命名模糊/歧义：名称省略关键限定词致多解（如 1号2号高加3号液位计），需补全机组/位置限定。",
    "KKS-25": "同设备异码：同一物理设备被编了多个 KKS 码，按归一名称聚类提示人工确认。",
    "KKS-25b": "同物异名候选：名称风格不一（如 电机/电动机/马达）聚簇出的高可疑对，仅供人工抽检。",
    "KKS-25c": "结构身份键一对多：同一旧码对应多个不同新码，疑似真重复，须人工裁定。",
    "KKS-26": "应编未编：父设备按惯例应含可独立标识的子部件（隔离点/回路/执行机构）却缺失。",
    "KKS-27": "名称文本卫生与语义可解性：名称含非常规符号/控制符判脏；提取不到任何已知设备名词判语义需人核。",
    "KKS-28": "机组三方一致：机组列、名称中的机组指代、全厂码 G 三方须一致；公用/自由字母跳过不误报。",
    "KKS-29": "OCR 易混字符待核：编码含 I/i/L/l/O/o 等易混字符，人工核对是否为数字 1/0 的识别错误。",
    "KKS-30": "文本卫生：码与名称的首尾空白、换行、制表符、全角空格等脏数据，标记并附清洗。",
    "KKS-31": "父子名称语义一致性：直接父子一级、同类设备的名称标识 token 应一致，矛盾即报。",
    "KKS-A": "名称语义与设备字母一致：名称含泵/阀等词应对应泵类/阀类字母，不符即名实疑点。",
    "KKS-C": "命名风格归一：1号机/1#机/#1机/一号机 等风格混用，提示统一风格。",
    "KKS-D": "空/缺名称：名称列为空、-、无、None 等，设为必填字段时按 P0 阻断。",
    "KKS-E": "设备类型列与名称一致：类型列值（如 截止阀）与名称推导类型（如 闸阀）矛盾。",
    "KKS-F": "同父级编号断号：同父设备编号序列缺号（001/002/004），可能漏编。",
    "KKS-G": "字母组合合法性：系统字母与设备字母的标准配对校验。",
}


# ---- 规则依据出处（word 标准文档 · 文件 + 章节）----
# 三份依据文档（2026-09-02 从 doc/docx 原始文件提取核对）：
#   火电导则：电力火电厂标识系统（KKS）编码导则-0813（.doc）
#   风电规范：中国电力投资集团公司风电场标识系统（KKS）编码规范 CPI（.docx）
#   光伏规范：中国电力投资集团公司光伏发电站标识系统（KKS）编码规范 CPI（.docx）
# RULE_SOURCES 值为"文件短名 + 条款号（表号）"，直接用于 HTML/Excel/Web"规则依据"列。
# 编码规范无明文章节的工程/数据质量类规则统一为通用说明，不虚构条款。
STANDARD_DEFAULT = "huodian"
STANDARD_NAMES = {"huodian": "火电导则", "wind": "风电规范", "solar": "光伏规范"}
# 与 国标/ 目录磁盘文件一一对应的完整文档名（展示时直接可对上文件名）
STANDARD_FILES = {
    "huodian": "电力火电厂标识系统（KKS）编码导则-0813",
    "wind": "中国电力投资集团公司风电场标识系统（KKS）编码规范",
    "solar": "中国电力投资集团公司光伏发电站标识系统（KKS）编码规范",
}
# 报告头"规则依据体系"行：完整磁盘文件名（含序号前缀与扩展名，可直接在 国标/ 目录找到）
STANDARD_FULL = {
    "huodian": "国标/电力火电厂标识系统（KKS）编码导则-0813.doc（火电导则）",
    "wind": "国标/3.中国电力投资集团公司风电场标识系统（KKS）编码规范.docx（风电规范）",
    "solar": "国标/4.中国电力投资集团公司光伏发电站标识系统（KKS）编码规范.docx（光伏规范）",
}
GENERIC_SOURCE = "通用审计要求，无专条依据"
_ST_STRUCT = ("火电导则 5.1.3 KKS编码构成（字符类型 N/A，禁用 I/O）", "风电规范 4.1 编码构成（字符类型 N/A，禁用 I/O）", "光伏规范 4.1 编码构成（字符类型 N/A，禁用 I/O）")
_ST_SEG = ("火电导则 5.2.1~5.2.3 码段结构（F0/F1F2F3/FN/A1A2/AN）", "风电规范 4.2.1~4.2.3 码段结构", "光伏规范 4.2.1~4.2.3 码段结构")
_ST_G = ("火电导则 5.1.4 全厂标识（表2）", "风电规范 4.1 全厂码（表1）", "光伏规范 4.1 全厂码（表1）")
_ST_UNIQ = ("火电导则 3 术语和定义·标识（唯一记号）", "风电规范 3 定义与术语·标识（唯一记号）", "光伏规范 3 定义与术语·标识（唯一记号）")
_ST_DEV = ("火电导则 6.2 设备索引（表29~37）", "风电规范 5.2 设备索引（表18~24）", "光伏规范 5.2 设备索引（表19~24）")
_ST_PART = ("火电导则 5.2.3 部件码", "风电规范 4.2.3 部件码", "光伏规范 4.2.3 部件码")
_ST_COMBO = ("火电导则 6.1+6.2 系统/设备索引", "风电规范 5.1+5.2 系统/设备索引", "光伏规范 5.1+5.2 系统/设备索引")

RULE_SOURCES: dict[str, dict[str, str]] = {}
for _rid, _vals in {
    "KKS-04": _ST_STRUCT, "KKS-04b": _ST_SEG, "KKS-05": _ST_STRUCT,
    "KKS-06": _ST_G, "KKS-07": _ST_UNIQ,
    "KKS-08": _ST_STRUCT, "KKS-09": _ST_STRUCT, "KKS-10": _ST_STRUCT,
    "KKS-11": _ST_STRUCT, "KKS-12": _ST_STRUCT,
    "KKS-25": _ST_UNIQ, "KKS-25c": _ST_UNIQ, "KKS-26": _ST_PART,
    "KKS-28": _ST_G, "KKS-29": _ST_STRUCT,
    "KKS-A": _ST_DEV, "KKS-G": _ST_COMBO,
}.items():
    RULE_SOURCES[_rid] = dict(zip(("huodian", "wind", "solar"), _vals))
# 出处列展示为"完整文件名 + 条款"，把元组中的短名前缀替换成 STANDARD_FILES 全名
for _rid, _tbl in RULE_SOURCES.items():
    for _std, _short in STANDARD_NAMES.items():
        _full = STANDARD_FILES[_std]
        _tbl[_std] = _tbl[_std].replace(_short, _full, 1)


def rule_source(rule_id: str, standard: str) -> str:
    """按审核标准体系取规则出处；支持组键回退（如 KKS-13 命中 KKS-13/16）。"""
    table = RULE_SOURCES.get(rule_id)
    if not table:
        table = next((v for k, v in RULE_SOURCES.items() if rule_id in k.split("/")), None)
    if not table:
        return GENERIC_SOURCE
    return table.get(standard) or table.get(STANDARD_DEFAULT) or GENERIC_SOURCE


def detect_standard(records: list[dict[str, Any]]) -> str:
    """按数据内容推断适用的标准文档体系：火电导则/风电规范/光伏规范。

    依据两类信号打分：
      1) KKS 码系统分类字母（第 3 位 F1）——对照三份文档的系统索引字母表；
      2) 设备名称特征词（用无歧义词，如"锅炉/汽轮机"vs"风电机组"vs"光伏/汇流箱"，
         单字"风机"不采用，暖通风机会误判为风电）。
    平分或全 0 时回退火电导则（当前主要业务口径）。
    """
    huodian_w = ("锅炉", "汽轮机", "汽机", "磨煤机", "给水泵", "凝汽器", "除氧器", "暖通", "热网", "送风机", "引风机", "主蒸汽", "再热器")
    wind_w = ("风电机组", "风力发电", "轮毂", "塔筒", "机舱", "桨叶")
    solar_w = ("光伏", "逆变器", "汇流箱", "方阵", "太阳能", "光伏组件")
    score = {"huodian": 0, "wind": 0, "solar": 0}
    f1_letters: set[str] = set()
    seen = 0
    for rec in records:
        if seen >= 800:
            break
        code = text(rec.get("kks_code"))
        name = text(rec.get("name"))
        seen += 1
        if len(code) >= 6 and code[2].isalpha():
            f1_letters.add(code[2].upper())
        for kw in huodian_w:
            if kw in name:
                score["huodian"] += 2
                break
        for kw in wind_w:
            if kw in name:
                score["wind"] += 2
                break
        for kw in solar_w:
            if kw in name:
                score["solar"] += 2
                break
    if "W" in f1_letters:
        score["solar"] += 2
    for ch in ("H", "L", "X"):
        if ch in f1_letters:
            score["huodian"] += 1
    best = max(score, key=lambda k: (score[k], k != STANDARD_DEFAULT))
    return best if score[best] else STANDARD_DEFAULT


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
        "definition": RULE_DEFINITIONS.get(rule_id, ""),
        "excel_row": rec.get("excel_row", ""),
        "kks_code": rec.get("kks_code", ""),
        "parent_code": rec.get("parent_code", ""),
        "old_code": rec.get("old_code", ""),
        "name": rec.get("name", ""),
        "message": message,
        "suggestion": suggestion,
        "evidence": evidence,
    })


def code_unit(g_char: str, common_units: set[str] | None = None) -> int | None:
    """全厂码 G（1 位）→ 机组号。

    - 1~9 → 1~9 号机组；A~G → 10~16 号机组
    - J~R/S~V/Y（公用）与 H/W/X/Z（自由使用）→ None（跳过，不误报）
    - 其他取值（0、I/O、数字外字符）→ None（由 G 取值合规检查另行报错）
    """
    if not g_char:
        return None
    if g_char in (common_units or COMMON_UNITS):
        return None
    if g_char in FREE_UNIT_LETTERS:
        return None
    if g_char.isdigit():
        value = int(g_char)
        return value if 1 <= value <= 9 else None
    return UNIT_LETTER_MAP.get(g_char.upper())


def valid_g_char(g_char: str) -> bool:
    """全厂码 G 合法取值：1-9 / A-G / J-R / S-V / Y / 自由字母 H/W/X/Z。"""
    if not g_char:
        return False
    if g_char.isdigit():
        return 1 <= int(g_char) <= 9
    return g_char.upper() in (set(UNIT_LETTER_MAP) | COMMON_UNITS | FREE_UNIT_LETTERS)


def classify_extension(code: str) -> str:
    if re.search(r"-KF\d{2}$", code):
        return "DCS 点号"
    if re.search(r"-[A-Z]{2}\d{2}$", code):
        return "部件级附加码"
    if len(code) == 13 and code[-1] in "AB":
        return "A/B 位置扩展"
    if len(code) == 13 and code[-1] in "C":
        return "末位 A/B/C 扩展"
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


# 设备索引字母（A 码）——10 版 VGB 旧字母（QM/MA/BT/FT/LS 等）不再用于名实一致性判定
DEVICE_LETTER_SEMANTICS = {
    "泵": {"AP", "BN"},
    "阀": {"AA"},
    "门": {"AA", "AB"},
    "执行机构": {"AS"},
    "测温": {"CT"},
    "温度": {"CT"},
    "流量": {"CF"},
    "液位": {"CL"},
    "风机": {"AN"},
    "压缩机": {"AN"},
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
                issues, seen, records_by_index, rule_id="KKS-16", priority="P2", category="旧版前缀历史提示", index=index,
                message=f"原码/新码前缀命中 10 版旧前缀表：{old[:2]} → {new[:2]}。", suggestion="编码体系不定义版本迁移；如属项目约定请核对迁移表，历史原码保留。", status="needs_review",
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


def apply_ai_review_stage(
    result: dict[str, Any],
    *,
    progress_callback: ProgressCallback | None = None,
) -> dict[str, Any]:
    """Run the optional semantic stage and update the audit conclusion."""
    result["ai_review"] = review_issue_candidates(result, progress_callback=progress_callback)
    result["final_decision_counts"] = final_decision_counts(result)
    final_counts = result["final_decision_counts"]
    final_actionable = final_counts.get("confirmed_issue", 0) + final_counts.get("needs_human", 0)
    if final_actionable:
        result["conclusion"] = (
            f"本次识别出 {final_actionable} 个需要处理或人工确认的问题；"
            f"另有 {final_counts.get('likely_false_positive', 0)} 个疑似误报保留复核记录。源 Excel 未修改。"
        )
    elif final_counts.get("likely_false_positive", 0):
        result["conclusion"] = (
            f"本次规则命中项均经 AI 归并为疑似误报，共 {final_counts['likely_false_positive']} 个；"
            "仍建议人工抽查。源 Excel 未修改。"
        )
    return result


def write_audit_artifacts(
    input_path: Path,
    output_dir: Path,
    result: dict[str, Any],
    *,
    progress_callback: ProgressCallback | None = None,
) -> dict[str, Any]:
    """Write only the two user-facing audit artifacts."""
    _emit_progress(progress_callback, {"phase": "audit", "stage": "report", "percent": 95, "message": "正在生成审核报告", "hint": "整理 HTML 报告和问题 Excel"})
    output_dir.mkdir(parents=True, exist_ok=True)
    html_name, xlsx_name = output_artifact_names(input_path)
    write_html(output_dir / html_name, result)
    write_issue_workbook_xlsx(output_dir / xlsx_name, result)
    LOGGER.info(
        "audit_done file=%s rows=%s issues=%s ai_status=%s ai_reviewed=%s",
        input_path.name,
        result.get("data_rows", 0),
        result.get("issue_count", 0),
        result.get("ai_review", {}).get("status", "pending"),
        result.get("ai_review", {}).get("reviewed_count", 0),
    )
    return result


def audit_tree_collection(
    input_path: Path,
    output_dir: Path,
    sheet: str,
    rows: list[tuple[Any, ...]],
    header_idx: int,
    tree_cols: dict[str, int],
    comparison_path: Path | None = None,
    progress_callback: ProgressCallback | None = None,
    *,
    run_ai_review: bool = True,
    write_outputs: bool = True,
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
    accepted_lengths = {1, 3, 4, 5, 6, 7, 11, 12, 13, 16, 17}

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
            for start, end, label, want in ((0, 1, "全厂码G", "ALNUM"), (1, 2, "系统前缀号F0", "DIGIT"), (2, 5, "系统分类码F1F2F3", "ALPHA"),
                                            (5, 7, "系统编号FN", "DIGIT"), (7, 9, "设备分类码A1A2", "ALPHA"), (9, 12, "设备编号AN", "DIGIT")):
                part = child_code[start:end]
                if want == "ALNUM":
                    ok = part.isalnum()
                elif want == "DIGIT":
                    ok = part.isdigit()
                else:
                    ok = part.isalpha()
                if not ok:
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
    standard = detect_standard(child_records)
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
        "standard": standard,
        "standard_display": STANDARD_FULL.get(standard, STANDARD_FULL[STANDARD_DEFAULT]),
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
        "_ai_context": {"records": child_records},
    }
    _emit_progress(progress_callback, {"phase": "audit", "stage": "rules_completed", "percent": 62, "message": "本地规则审核完成", "hint": f"已发现 {len(issues)} 个规则问题，准备进行 AI 候选复核"})
    if run_ai_review:
        apply_ai_review_stage(result, progress_callback=progress_callback)
    if write_outputs:
        write_audit_artifacts(input_path, output_dir, result, progress_callback=progress_callback)
    return result


def audit_file(
    input_path: Path,
    output_dir: Path,
    requested_sheet: str | None = None,
    comparison_path: Path | None = None,
    progress_callback: ProgressCallback | None = None,
    *,
    run_ai_review: bool = True,
    write_outputs: bool = True,
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
            return audit_tree_collection(
                input_path,
                output_dir,
                sheet,
                rows,
                header_idx,
                tree_cols,
                comparison_path,
                progress_callback,
                run_ai_review=run_ai_review,
                write_outputs=write_outputs,
            )
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

    standard = detect_standard(records)

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
            # 分段结构：全厂码G(1) + 系统前缀号F0(1) + 系统分类码F1F2F3(3) + 系统编号FN(2) + 设备分类码A1A2(2) + 设备编号AN(3)
            segments = ((0, 1, "全厂码G", "ALNUM"), (1, 2, "系统前缀号F0", "DIGIT"), (2, 5, "系统分类码F1F2F3", "ALPHA"),
                        (5, 7, "系统编号FN", "DIGIT"), (7, 9, "设备分类码A1A2", "ALPHA"), (9, 12, "设备编号AN", "DIGIT"))
            seg_has_issue = False
            for start, end, label, want in segments:
                part = code[start:end]
                if want == "ALNUM":
                    ok = part.isalnum()
                elif want == "DIGIT":
                    ok = part.isdigit()
                else:
                    ok = part.isalpha()
                if not ok:
                    seg_has_issue = True
                    suggestion = code
                    if label == "设备编号AN":
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

        expected = code_unit(code[0], common_units)
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
        # 全厂码 G 取值合规：G 必须在 1-9 / A-G / J-R / S-V / Y / 自由字母内
        g_char = code[0] if code else ""
        if g_char and not valid_g_char(g_char):
            add_issue(issues, priority="P0", rule_id="KKS-06", category="全厂码 G 取值非法", rec=rec,
                      message=f"全厂码 G={g_char!r} 不在合法取值范围内（1-9/A-G/J-R/S-V/Y，自由字母 H/W/X/Z）。", suggestion="核对是否为旧版前缀（50/60/L0/J0 现行体系下 G=5/6/L/J 合法）或录入错误。")
        # 10 版旧前缀迁移：G=5/6/L/J 本身合法，迁移 50→05 等属项目约定，仅作 P2 历史提示
        if old and new and old[:2] != new[:2]:
            expected_prefix = old_prefix_migrations.get(old[:2])
            if expected_prefix and new[:2] != expected_prefix:
                add_issue(issues, priority="P2", rule_id="KKS-13/16", category="旧码迁移历史提示", rec=rec,
                          message=f"原码前缀 {old[:2]} 若按项目 10 版→20 版约定应迁移为 {expected_prefix}，当前新码为 {new[:2]}。", suggestion="编码体系不定义版本迁移；如属项目约定请核对迁移表，历史原码保留。", status="needs_review")
            else:
                add_issue(issues, priority="P2", rule_id="KKS-18", category="历史原码与新码前缀变化", rec=rec,
                          message=f"原 KKS 前缀 {old[:2]} 与新 KKS 前缀 {new[:2]} 不同。", suggestion="确认历史映射；若新码、父级、机组三方一致，可作为历史提示保留。", status="needs_review")

        if code[:2] in old_prefix_migrations:
            add_issue(issues, priority="P2", rule_id="KKS-16", category="旧版前缀历史提示", rec=rec,
                      message=f"当前 KKS 编码前两位 {code[:2]} 命中 10 版旧前缀表（{code[:2]}）。", suggestion="G=5/6/L/J 本身合法（分别为 5/6 号机与期别公用），仅提示核对；历史原码可保留。", status="needs_review")

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
        "standard": standard,
        "standard_display": STANDARD_FULL.get(standard, STANDARD_FULL[STANDARD_DEFAULT]),
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
        "_ai_context": {"records": list(template_records.values())},
    }
    _emit_progress(progress_callback, {"phase": "audit", "stage": "rules_completed", "percent": 62, "message": "本地规则审核完成", "hint": f"已发现 {len(issues)} 个规则问题，准备进行 AI 候选复核"})
    if run_ai_review:
        apply_ai_review_stage(result, progress_callback=progress_callback)
    if write_outputs:
        write_audit_artifacts(input_path, output_dir, result, progress_callback=progress_callback)
    return result


def write_issue_workbook_xlsx(path: Path, result: dict[str, Any]) -> None:
    """生成面向业务交付的问题清单，技术 AI 依据放入隐藏页。"""
    wb = Workbook()
    summary = wb.active
    summary.title = "概览"
    issues_ws = wb.create_sheet("问题清单")
    ai_ws = wb.create_sheet("AI复核（技术）")
    resolved_ws = wb.create_sheet("已澄清项")
    structure_ws = wb.create_sheet("文件结构")
    coverage_ws = wb.create_sheet("规则覆盖")
    blue, pale_blue, gray, pale_red, pale_amber, pale_green = "1D4ED8", "EFF6FF", "F8FAFC", "FEF2F2", "FFFBEB", "F0FDF4"

    def header_style(ws, cell_range: str, color: str = blue) -> None:
        for row in ws[cell_range]:
            for cell in row:
                cell.fill = PatternFill("solid", fgColor=color)
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

    def add_table(ws, name: str, end_column: int) -> None:
        end_row = max(1, ws.max_row)
        ref = f"A1:{get_column_letter(end_column)}{end_row}"
        table = Table(displayName=name, ref=ref)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        ws.add_table(table)

    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False

    metrics = quality_metrics(result)
    counts = result.get("final_decision_counts", final_decision_counts(result))
    dimensions = report_quality_dimensions(result)
    standard = result.get("standard", STANDARD_DEFAULT)
    report_issues = sorted(_report_issues(result), key=lambda item: ({"P0": 0, "P1": 1, "P2": 2}.get(str(item.get("priority")), 9), int(item.get("excel_row", 0) or 0)))

    summary.merge_cells("A1:H1")
    summary["A1"] = "KKS 编码质量审核报告"
    summary["A1"].fill = PatternFill("solid", fgColor=blue)
    summary["A1"].font = Font(bold=True, color="FFFFFF", size=16)
    summary["A1"].alignment = Alignment(horizontal="center", vertical="center")
    summary.row_dimensions[1].height = 28
    summary.merge_cells("A2:H2")
    summary["A2"] = f"源文件：{result['source_file']}；主表：{result['sheet']}；表头行：{result['header_row']}；规则依据体系：{result.get('standard_display') or STANDARD_FULL[standard]}（自动识别）"
    summary["A2"].fill = PatternFill("solid", fgColor=pale_blue)
    summary["A2"].alignment = Alignment(wrap_text=True)
    summary.merge_cells("A4:H4")
    summary["A4"] = "一、总体结论"
    header_style(summary, "A4:H4")
    summary.merge_cells("A5:H6")
    summary["A5"] = report_import_conclusion(result)
    summary["A5"].fill = PatternFill("solid", fgColor=pale_green if not metrics["p0"] and not metrics["p1"] else pale_amber)
    summary["A5"].alignment = Alignment(vertical="center", wrap_text=True)

    summary.merge_cells("A8:H8")
    summary["A8"] = "二、关键质量指标"
    header_style(summary, "A8:H8")
    summary.append(["指标", "结果", "指标", "结果", "指标", "结果", "指标", "结果"])
    header_style(summary, f"A{summary.max_row}:H{summary.max_row}")
    metric_pairs = [("有效 KKS 数量", metrics["effective_kks"]), ("设备级编码", metrics["device_level_codes"]), ("重复编码", metrics["duplicate_codes"]), ("父级错误", metrics["parent_errors"]), ("高风险问题（P0）", metrics["p0"]), ("待整改问题（P1）", metrics["p1"]), ("提示问题（P2）", metrics["p2"])]
    for index in range(0, len(metric_pairs), 4):
        row = []
        for label, value in metric_pairs[index:index + 4]:
            row.extend([label, value])
        summary.append(row)
    body_style(summary, f"A{summary.max_row - 1}:H{summary.max_row}")

    summary.merge_cells("A12:H12")
    summary["A12"] = "三、KKS 八维审核结果"
    header_style(summary, "A12:H12")
    summary.append(["维度", "检查项", "结果", "说明"])
    header_style(summary, f"A{summary.max_row}:D{summary.max_row}")
    dimension_header_row = summary.max_row
    summary.merge_cells(start_row=dimension_header_row, start_column=4, end_row=dimension_header_row, end_column=8)
    for item in dimensions:
        summary.append([item["dimension"], item["check"], item["result"], item["detail"]])
        dimension_row = summary.max_row
        summary.merge_cells(start_row=dimension_row, start_column=4, end_row=dimension_row, end_column=8)
    body_style(summary, f"A{summary.max_row - len(dimensions) + 1}:D{summary.max_row}")

    summary.merge_cells("A22:H22")
    summary["A22"] = "四、P0/P1/P2 问题分析"
    header_style(summary, "A22:H22")
    summary.append(["等级", "含义", "数量", "处置要求"])
    header_style(summary, f"A{summary.max_row}:D{summary.max_row}")
    priority_header_row = summary.max_row
    summary.merge_cells(start_row=priority_header_row, start_column=4, end_row=priority_header_row, end_column=8)
    for level, meaning, requirement in (("P0", "阻断问题", "必须修改并复核后才能导入"), ("P1", "结构治理问题", "完成治理或人工确认后再导入"), ("P2", "历史迁移问题", "保留原码证据，按迁移策略处理")):
        summary.append([level, meaning, metrics[level.lower()], requirement])
        priority_row = summary.max_row
        summary.merge_cells(start_row=priority_row, start_column=4, end_row=priority_row, end_column=8)
    body_style(summary, f"A{summary.max_row - 2}:D{summary.max_row}")

    summary.merge_cells("A27:H27")
    summary["A27"] = "五、问题整改建议"
    header_style(summary, "A27:H27")
    summary.merge_cells("A28:H29")
    summary["A28"] = "优先处理 P0 阻断问题；再治理 P1 扩展编码；P2 历史迁移保留原始证据，不直接覆盖源 Excel。其余规则提示仍保留在详细问题清单，整改完成后应重新审核并做目标库导入前验证。"
    summary["A28"].fill = PatternFill("solid", fgColor=gray)
    summary["A28"].alignment = Alignment(wrap_text=True, vertical="center")

    summary.merge_cells("A31:H31")
    summary["A31"] = "六、导入评估结论"
    header_style(summary, "A31:H31")
    summary.merge_cells("A32:H33")
    summary["A32"] = report_import_conclusion(result)
    summary["A32"].fill = PatternFill("solid", fgColor=pale_amber)
    summary["A32"].alignment = Alignment(wrap_text=True, vertical="center")
    summary.merge_cells("A35:H35")
    summary["A35"] = "审核方法"
    header_style(summary, "A35:H35")
    for method in REPORT_METHODS:
        summary.append([f"• {method}"])
        summary.merge_cells(start_row=summary.max_row, start_column=1, end_row=summary.max_row, end_column=8)
        summary.cell(summary.max_row, 1).fill = PatternFill("solid", fgColor=gray)
    summary.append(["源 Excel 始终只读；本次未连接真实 DM8/LOCATIONS 做导入验证。"])
    summary.merge_cells(start_row=summary.max_row, start_column=1, end_row=summary.max_row, end_column=8)
    summary.cell(summary.max_row, 1).fill = PatternFill("solid", fgColor=gray)
    summary.freeze_panes = "A4"
    autosize(summary, 58)
    for column, width in {"A": 19, "B": 13, "C": 19, "D": 13, "E": 19, "F": 13, "G": 19, "H": 13}.items():
        summary.column_dimensions[column].width = width

    issue_headers = ["等级", "规则", "Excel 行号", "KKS", "问题定义", "规则依据", "问题", "整改建议"]
    issues_ws.append(issue_headers)
    for item in report_issues:
        decision = str(item.get("final_decision", ""))
        action = {"confirmed_issue": "需整改", "needs_human": "需确认", "likely_false_positive": "建议抽查"}.get(decision, "待确认")
        message = item.get("final_summary") or item.get("message") or item.get("category") or "待核问题"
        suggestion = item.get("final_suggestion") or item.get("suggestion") or "请结合原始 Excel 和业务资料确认。"
        definition = item.get("definition") or item.get("category") or ""
        source = rule_source(str(item.get("rule_id", "")), standard)
        issues_ws.append([item.get("priority", "P2"), item.get("rule_id", ""), item.get("excel_row", ""), item.get("kks_code", ""), definition, source, f"{action}：{message}", suggestion])
    header_style(issues_ws, "A1:H1")
    body_style(issues_ws, f"A1:H{max(1, issues_ws.max_row)}")
    issues_ws.freeze_panes = "A2"
    issues_ws.auto_filter.ref = f"A1:H{max(1, issues_ws.max_row)}"
    autosize(issues_ws, 60)
    issues_ws.column_dimensions["E"].width = 34
    issues_ws.column_dimensions["F"].width = 54
    add_table(issues_ws, "KKSIssues", 8)

    ai_headers = ["内部状态", "等级", "规则", "Excel 行号", "KKS", "原始问题", "规则证据", "AI 初审", "AI 二次复核", "AI 置信度", "AI 证据", "AI 原因", "AI 建议", "最终判断", "最终证据", "最终原因", "最终建议"]
    ai_ws.append(ai_headers)
    for item in report_issues:
        ai_ws.append([item.get("status", ""), item.get("priority", "P2"), item.get("rule_id", ""), item.get("excel_row", ""), item.get("kks_code", ""), item.get("message", ""), item.get("evidence", ""), item.get("ai_initial_decision", ""), item.get("ai_verification_decision", ""), item.get("ai_confidence", ""), item.get("ai_evidence", ""), item.get("ai_reason", ""), item.get("ai_suggestion", ""), final_decision_label(item.get("final_decision", "")), item.get("final_evidence", ""), item.get("final_reason", ""), item.get("final_suggestion", "")])
    header_style(ai_ws, f"A1:{get_column_letter(len(ai_headers))}1")
    body_style(ai_ws, f"A1:{get_column_letter(len(ai_headers))}{max(1, ai_ws.max_row)}")
    ai_ws.freeze_panes = "A2"
    ai_ws.auto_filter.ref = f"A1:{get_column_letter(len(ai_headers))}{max(1, ai_ws.max_row)}"
    autosize(ai_ws, 55)
    ai_ws.sheet_state = "hidden"

    resolved_ws.append(["Excel 行号", "KKS", "复核结论", "证据"])
    for item in result.get("resolved_reviews", []):
        resolved_ws.append([item.get("excel_row", ""), item.get("kks_code", ""), item.get("message", ""), item.get("evidence", "")])
    header_style(resolved_ws, "A1:D1")
    body_style(resolved_ws, f"A1:D{max(1, resolved_ws.max_row)}")
    resolved_ws.freeze_panes = "A2"
    resolved_ws.auto_filter.ref = f"A1:D{max(1, resolved_ws.max_row)}"
    autosize(resolved_ws)
    add_table(resolved_ws, "KKSResolved", 4)

    structure_ws.append(["Sheet", "最大行", "最大列", "非空行", "首个非空行", "最后非空行", "表头预览"])
    for item in result.get("sheet_info", []):
        preview = " | ".join(str(x) for x in item.get("header_preview", []))
        structure_ws.append([item.get(key, "") for key in ("sheet", "max_row", "max_col", "nonempty_rows", "first_nonempty_row", "last_nonempty_row")] + [preview])
    header_style(structure_ws, "A1:G1")
    body_style(structure_ws, f"A1:G{max(1, structure_ws.max_row)}")
    structure_ws.freeze_panes = "A2"
    structure_ws.auto_filter.ref = f"A1:G{max(1, structure_ws.max_row)}"
    autosize(structure_ws)
    add_table(structure_ws, "KKSSheets", 7)

    coverage_ws.append(["维度", "规则", "执行状态", "命中数", "执行/证据来源"])
    issue_counts = Counter(item.get("rule_id", "") for item in _report_issues(result))
    coverage_items = [
        ("文件与结构", "01-03", "已执行", "动态 Sheet、表头和列定位", "kks-audit/SKILL.md"),
        ("语法与格式", "KKS-04/04b/05/06/29/30", "已执行", "字符、长度、OCR 和文本卫生", "kks-audit/SKILL.md + audit_template.py"),
        ("唯一性", "KKS-07", "已执行", "重复码和跨设备碰撞诊断", "audit_template.py"),
        ("层级树", "KKS-08~KKS-12", "已执行", "孤儿、父子关系和父级链", "kks-audit/SKILL.md"),
        ("历史迁移", "KKS-13/14/15/16/18/25c", "已执行", "旧码、新码和一对多映射", "kks-audit/SKILL.md + config/kks_rules.json"),
        ("语义与命名", "KKS-17/22/24~31/A/C/E/F/G", "已执行", "扩展码、名称和语义一致性", "audit_template.py + SKILL.md"),
        ("跨范围对比", "KKS-19/20/21", "已执行" if result.get("scope_comparison") else "未提供对比文件", "专业覆盖、编码深度和数量瀑布", "--compare-file"),
        ("导入评估", "KKS-22/23", "已执行", "LOCATIONS 约束和历史双码策略", "kks-audit/SKILL.md"),
    ]
    for dimension, rule_id, status, hit_note, source in coverage_items:
        hit_value = sum(issue_counts.get(rule, 0) for rule in re.split(r"[/~]", rule_id) if issue_counts.get(rule))
        coverage_ws.append([dimension, rule_id, status, hit_value, f"{hit_note}；{source}"])
    header_style(coverage_ws, "A1:E1")
    body_style(coverage_ws, f"A1:E{max(1, coverage_ws.max_row)}")
    coverage_ws.freeze_panes = "A2"
    coverage_ws.auto_filter.ref = f"A1:E{max(1, coverage_ws.max_row)}"
    autosize(coverage_ws)
    add_table(coverage_ws, "KKSCoverage", 5)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def write_html(path: Path, result: dict[str, Any]) -> None:
    metrics = quality_metrics(result)
    final_counts = result.get("final_decision_counts", final_decision_counts(result))
    rows = sorted(_report_issues(result), key=lambda item: ({"P0": 0, "P1": 1, "P2": 2}.get(str(item.get("priority")), 9), int(item.get("excel_row", 0) or 0)))
    ai = result.get("ai_review", {})
    standard = result.get("standard", STANDARD_DEFAULT)

    def source_text(item: dict[str, Any]) -> str:
        return rule_source(str(item.get("rule_id", "")), standard)

    rows_html = "".join(
        f"<tr><td><span class='priority p{html.escape(str(item.get('priority', 'P2'))[-1])}'>{html.escape(str(item.get('priority', 'P2')))}</span></td>"
        f"<td>{html.escape(str(item.get('rule_id', '')))}</td><td>{html.escape(str(item.get('excel_row', '')))}</td>"
        f"<td><code>{html.escape(str(item.get('kks_code', '')))}</code></td>"
        f"<td class='def-cell'>{html.escape(str(item.get('definition') or item.get('category') or ''))}</td>"
        f"<td class='src-cell'>{html.escape(source_text(item))}</td>"
        f"<td><b>{html.escape(action_label(item))}</b>：{html.escape(detail_text(item))}</td>"
        f"<td>{html.escape(suggestion_text(item))}</td></tr>"
        for item in rows
    ) or '<tr><td colspan="8" class="empty">未发现需要列入清单的问题</td></tr>'

    tree_html = rule_tree_html(issue_rule_tree(result)) or "<p class='empty'>未发现需要列入清单的问题</p>"
    tree_flat = f"<details class='technical'><summary>查看平铺问题清单（{len(rows)} 条，按 KKS/行号平铺对照）</summary><table><thead><tr><th>等级</th><th>规则</th><th>行号</th><th>KKS</th><th>问题定义</th><th>规则依据</th><th>问题</th><th>整改建议</th></tr></thead><tbody>{rows_html}</tbody></table></details>"

    dimension_rows = "".join(
        f"<tr><td><b>{html.escape(item['dimension'])}</b></td><td>{html.escape(item['check'])}</td>"
        f"<td><span class='result-pill {'ok' if item['result'] == '通过' else 'warn'}'>{html.escape(item['result'])}</span></td>"
        f"<td>{html.escape(item['detail'])}</td></tr>"
        for item in report_quality_dimensions(result)
    )
    priority_rows = "".join(
        f"<div class='priority-card p{level[-1]}'><strong>{level} · {meaning}</strong><b>{metrics[level.lower()]}</b><span>{requirement}</span></div>"
        for level, meaning, requirement in (("P0", "阻断问题", "必须修改并复核后才能导入"), ("P1", "需整改问题", "完成治理或人工确认后再导入"), ("P2", "提示问题", "需人工确认或按治理策略处理"))
    )
    methods_html = "".join(f"<li>{html.escape(method)}</li>" for method in REPORT_METHODS)
    notes_html = "".join(f"<li>{html.escape(str(note))}</li>" for note in result.get("structural_notes", []))

    def ai_detail(item: dict[str, Any]) -> str:
        try:
            confidence = f"{float(item.get('ai_confidence', 0)):.0%}"
        except (TypeError, ValueError):
            confidence = "—"
        return (
            f"<details class='ai-detail'><summary>查看 AI 依据</summary>"
            f"<div class='ai-grid'><span>初审</span><b>{html.escape(str(item.get('ai_initial_decision') or item.get('ai_decision') or '未返回'))}</b>"
            f"<span>二次复核</span><b>{html.escape(str(item.get('ai_verification_decision') or '未触发'))}</b>"
            f"<span>置信度</span><b>{confidence}</b><span>证据</span><p>{html.escape(str(item.get('ai_evidence') or item.get('final_evidence') or item.get('evidence') or '未提供'))}</p>"
            f"<span>原因</span><p>{html.escape(str(item.get('ai_reason') or item.get('final_reason') or '未提供'))}</p>"
            f"<span>建议</span><p>{html.escape(str(item.get('ai_suggestion') or item.get('final_suggestion') or '未提供'))}</p></div></details>"
        )

    ai_rows = "".join(
        f"<tr><td>{html.escape(str(item.get('priority', 'P2')))}</td><td>{html.escape(str(item.get('excel_row', '')))}</td><td>{html.escape(str(item.get('kks_code', '')))}</td><td>{ai_detail(item)}</td></tr>"
        for item in rows if item.get("ai_decision") or item.get("ai_initial_decision") or item.get("ai_evidence")
    ) or '<tr><td colspan="4" class="empty">本次未产生 AI 复核明细</td></tr>'
    comparison = result.get("scope_comparison")
    comparison_html = ""
    if comparison:
        waterfall = comparison.get("quantity_waterfall", {})
        comparison_html = f"<details class='technical'><summary>查看跨范围对比结果</summary><p>对比文件：{html.escape(str(comparison.get('comparison', {}).get('file', '')))}；有效编码数量差：{waterfall.get('primary_minus_comparison', '—')}；瀑布闭合：{html.escape(str(waterfall.get('closed', '—')))}。</p><pre>{html.escape(json.dumps(comparison, ensure_ascii=False, indent=2))}</pre></details>"
    ai_meta = f"状态：{html.escape(str(ai.get('status', 'disabled')))}；候选项 {ai.get('candidate_count', 0)}；已复核 {ai.get('reviewed_count', 0)}；疑似误报二次核验 {ai.get('verified_count', 0)}。"
    body = f"""<!doctype html><html lang='zh-CN'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>KKS 编码质量审核报告</title>
<style>
:root{{--blue:#1d4ed8;--navy:#102a43;--ink:#243b53;--muted:#627d98;--line:#d9e2ec;--soft:#f7faff;--green:#16845b;--amber:#b45309;--red:#c2413b}}
*{{box-sizing:border-box}}body{{margin:0;background:#f5f8fc;color:var(--ink);font:14px/1.65 "Segoe UI","Microsoft YaHei",Arial,sans-serif}}.page{{max-width:1440px;margin:0 auto;padding:32px 38px 56px}}.top{{display:flex;justify-content:space-between;gap:24px;align-items:flex-start;padding:22px 26px;background:#fff;border:1px solid var(--line);border-radius:18px;box-shadow:0 8px 24px rgba(16,42,67,.05)}}h1{{margin:0;color:var(--navy);font-size:30px;letter-spacing:-.5px}}.eyebrow{{margin-bottom:5px;color:#315dcc;font-size:11px;font-weight:800;letter-spacing:1.6px}}.meta{{margin:10px 0 0;color:var(--muted);font-size:13px}}.status{{padding:8px 13px;border:1px solid #bbf7d0;border-radius:999px;background:#f0fdf4;color:var(--green);font-weight:700;white-space:nowrap}}.section{{margin-top:22px;padding:24px 26px;background:#fff;border:1px solid var(--line);border-radius:16px;box-shadow:0 6px 20px rgba(16,42,67,.04)}}h2{{margin:0 0 15px;color:var(--navy);font-size:20px;border-left:4px solid var(--blue);padding-left:11px}}h3{{margin:0 0 12px;color:var(--navy);font-size:16px}}.conclusion{{padding:16px 18px;border:1px solid #bfdbfe;border-radius:12px;background:#eff6ff;font-size:15px;font-weight:650;color:#173f7a}}.kpis{{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}}.kpi{{min-height:94px;padding:15px 16px;border:1px solid var(--line);border-radius:12px;background:linear-gradient(145deg,#fff,#f8fbff)}}.kpi label{{display:block;color:var(--muted);font-size:12px;font-weight:650}}.kpi b{{display:block;margin-top:6px;color:var(--navy);font-size:27px;line-height:1.1}}table{{width:100%;border-collapse:separate;border-spacing:0;overflow:hidden;border:1px solid var(--line);border-radius:12px;font-size:13px}}th,td{{padding:10px 11px;text-align:left;vertical-align:top;border-bottom:1px solid #e8eef5}}th{{background:#eff6ff;color:#173f7a;font-weight:800}}tr:last-child td{{border-bottom:0}}tbody tr:hover{{background:#fbfdff}}.def-cell{{color:#5b7590;font-size:12.5px;line-height:1.55;min-width:180px;max-width:320px}}.src-cell{{color:#0e5a44;font-size:12.5px;line-height:1.55;min-width:150px;max-width:230px;white-space:normal}}.issue-tree{{margin:6px 0 8px}}.it-node{{border:1px solid #dbe5ef;border-radius:9px;margin:5px 0;overflow:hidden}}.it-node>summary{{list-style:none;display:flex;align-items:center;gap:10px;flex-wrap:wrap;padding:7px 13px;cursor:pointer;font-weight:650;color:#173f7a;background:linear-gradient(180deg,#f2f7fd,#eaf2fa)}}.it-node>summary::-webkit-details-marker{{display:none}}.it-node>summary:before{{content:"▸";color:#8ba3bf;font-size:12px;width:14px}}.it-node[open]>summary:before{{content:"▾"}}.it-node[open]>summary{{background:#eaf2fa;border-bottom:1px solid #dbe5ef}}.it-name{{font-family:Consolas,Menlo,monospace;font-size:12.5px;color:#0b5394;font-weight:800}}.it-title{{color:#5b7590;font-size:12.5px;max-width:380px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}.it-count{{color:#8aa0b8;font-size:12px;margin-left:auto}}.tg-rule>summary .it-count{{margin-left:0}}.tb{{min-width:32px;padding:2px 7px;border-radius:999px;font-size:11.5px;font-weight:800;text-align:center}}.tb-p0{{background:#fef2f2;color:#b91c1c;border:1px solid #fecaca}}.tb-p1{{background:#fffbeb;color:#b45309;border:1px solid #fde68a}}.tb-p2{{background:#eff6ff;color:#1d4ed8;border:1px solid #bfdbfe}}.it-kids{{padding:2px 8px 8px 20px}}.it-issues{{padding:2px 14px 6px;background:#fff}}.it-issue{{display:flex;align-items:flex-start;gap:8px;flex-wrap:wrap;padding:9px 2px 6px;border-top:1px dashed #e3ebf3;font-size:13px;line-height:1.55}}.it-issue .priority{{min-width:auto;padding:2px 7px}}.it-issue code{{margin-right:6px}}.it-issue .it-def{{color:#5b7590;font-size:12px}}.it-issue b{{margin-right:4px}}.it-suggest{{flex-basis:100%;color:#627d98;font-size:12px;padding-left:2px}}.it-leaf{{display:flex;align-items:center;gap:10px;flex-wrap:wrap;padding:5px 13px;color:#627d98;font-size:12.5px}}.tg-rule{{border-left:3px solid #1d4ed8;background:#f3f7fc}}.tg-rule>summary{{background:linear-gradient(180deg,#edf4fc,#e0ebf7);font-size:13px;padding:8px 13px}}.tg-rule[open]>summary{{background:#e0ebf7}}.tg-rid{{font-size:13px;letter-spacing:.2px}}.tg-def{{flex:0 1 auto;min-width:0;color:#3e5871;font-size:12px;max-width:360px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}.tg-act{{display:inline-flex;align-items:center;justify-content:center;line-height:1;flex:0 0 auto;margin-left:auto;min-width:64px;padding:2px 9px;border-radius:999px;font-size:11px;font-weight:800;white-space:nowrap}}.tg-act.act-fix{{background:#fee2e2;color:#b91c1c;border:1px solid #fecaca}}.tg-act.act-hm{{background:#fef3c7;color:#a16207;border:1px solid #fde68a}}.tg-act.act-ck{{background:#dbeafe;color:#1d4ed8;border:1px solid #bfdbfe}}.tg-src{{padding:3px 14px 9px;color:#7a93ac;font-size:11.5px;background:#fff}}.it-code{{margin-left:3px;border-color:#e7eef6;background:#fff}}.it-code>summary{{background:#fff;border-bottom:1px dashed #e2e8f0;padding:5px 12px;font-size:12.5px;font-weight:700}}.it-code[open]>summary{{background:#f7fafd;border-bottom:1px solid #e2e8f0}}.it-row{{display:inline-flex;flex:0 0 auto;align-items:center;padding:1px 9px;border-radius:999px;background:#eef2f7;color:#40556b;font-size:11px;font-weight:800;white-space:nowrap}}.it-issue .it-msg{{flex:1 1 340px;color:#334155}}.result-pill{{display:inline-flex;padding:3px 8px;border-radius:999px;font-size:12px;font-weight:750}}.result-pill.ok{{background:#ecfdf5;color:#047857}}.result-pill.warn{{background:#fff7ed;color:#b45309}}.priority-grid{{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px}}.priority-card{{display:grid;grid-template-columns:1fr auto;gap:2px 10px;padding:15px 16px;border:1px solid var(--line);border-left:4px solid var(--blue);border-radius:12px;background:#fbfdff}}.priority-card strong{{font-size:14px}}.priority-card b{{grid-row:span 2;color:var(--navy);font-size:28px}}.priority-card span{{color:var(--muted);font-size:12px}}.priority-card.p0{{border-left-color:var(--red);background:#fffafa}}.priority-card.p1{{border-left-color:#f59e0b;background:#fffdf7}}.priority-card.p2{{border-left-color:var(--blue)}}.priority{{display:inline-flex;min-width:34px;justify-content:center;padding:3px 8px;border-radius:6px;font-weight:800}}.priority.p0{{background:#fee2e2;color:#b91c1c}}.priority.p1{{background:#fef3c7;color:#a16207}}.priority.p2{{background:#dbeafe;color:#1d4ed8}}code{{padding:2px 5px;border-radius:5px;background:#f1f5f9;color:#19324d;font-family:Consolas,monospace;font-size:12px}}.methods{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:7px 24px;margin:0;padding-left:22px}}.methods li{{color:#3e5871}}.advice{{padding:15px 17px;border-radius:12px;background:#f8fafc;border:1px solid var(--line)}}.import{{padding:16px 18px;border-radius:12px;background:#fff7ed;border:1px solid #fed7aa;color:#92400e;font-weight:650}}.empty{{padding:18px;text-align:center;color:#94a3b8}}details{{margin-top:12px;border:1px solid var(--line);border-radius:10px;background:#fbfdff}}summary{{cursor:pointer;padding:10px 12px;color:#2454ad;font-weight:750}}.ai-detail{{margin:0;border:0;background:transparent}}.ai-detail summary{{padding:0 0 7px;font-size:12px}}.ai-grid{{display:grid;grid-template-columns:80px 1fr;gap:5px 10px;padding:0 12px 12px;color:var(--muted);font-size:12px}}.ai-grid b{{color:var(--ink)}}.ai-grid p{{margin:0;color:var(--ink)}}.technical pre{{max-height:360px;overflow:auto;padding:12px;background:#172033;color:#dbeafe;border-radius:8px;font-size:12px}}.footer{{margin-top:22px;color:#718096;font-size:12px}}@media(max-width:950px){{.page{{padding:20px 16px 40px}}.kpis{{grid-template-columns:repeat(2,minmax(0,1fr))}}.priority-grid{{grid-template-columns:1fr}}.top{{display:block}}.status{{display:inline-flex;margin-top:12px}}}}@media(max-width:560px){{.kpis{{grid-template-columns:1fr 1fr;gap:8px}}.section{{padding:18px 15px}}h1{{font-size:24px}}table{{font-size:12px}}th,td{{padding:8px}}.methods{{grid-template-columns:1fr}}}}
@media print{{body{{background:#fff}}.page{{max-width:none;padding:0}}.section,.top{{box-shadow:none;break-inside:avoid}}details:not(.it-node){{display:none}}details.it-node:not([open])>*{{display:block!important}}details.it-node{{display:block}}details.it-node>summary{{border-bottom:1px solid #dbe5ef}}details.it-node>summary:before{{content:none}}}}
</style></head><body><main class='page'>
<header class='top'><div><div class='eyebrow'>KKS QUALITY CONTROL</div><h1>KKS 编码质量审核报告</h1><p class='meta'>源文件：{html.escape(str(result['source_file']))}<br>主表：{html.escape(str(result['sheet']))} · 表头行：{result['header_row']} · 审核范围：{metrics['effective_kks']} 条有效 KKS<br>规则依据体系：{html.escape(str(result.get('standard_display') or STANDARD_FULL.get(standard, standard)))}（按输入内容自动识别）</p></div><span class='status'>质量审核结论</span></header>
<section class='section'><h2>一、总体结论</h2><div class='conclusion'>{html.escape(report_import_conclusion(result))}</div></section>
<section class='section'><h2>二、关键质量指标</h2><div class='kpis'><div class='kpi'><label>有效 KKS 数量</label><b>{metrics['effective_kks']}</b></div><div class='kpi'><label>设备级编码</label><b>{metrics['device_level_codes']}</b></div><div class='kpi'><label>重复编码</label><b>{metrics['duplicate_codes']}</b></div><div class='kpi'><label>父级错误</label><b>{metrics['parent_errors']}</b></div><div class='kpi'><label>高风险问题（P0）</label><b>{metrics['p0']}</b></div><div class='kpi'><label>待治理问题（P1）</label><b>{metrics['p1']}</b></div><div class='kpi'><label>历史迁移问题（P2）</label><b>{metrics['p2']}</b></div></div></section>
<section class='section'><h2>三、KKS 八维审核结果</h2><table><thead><tr><th>维度</th><th>检查项</th><th>结果</th><th>说明</th></tr></thead><tbody>{dimension_rows}</tbody></table></section>
<section class='section'><h2>四、P0/P1/P2 问题分析</h2><div class='priority-grid'>{priority_rows}</div></section>
<section class='section'><h2>五、问题整改建议</h2><div class='advice'>优先处理 P0 阻断问题；再治理 P1 扩展编码；P2 历史迁移保留原始证据，不直接覆盖源 Excel。其余规则提示仍保留在详细问题清单，整改完成后应重新审核，并在目标库做导入前验证。</div><h3 style='margin-top:18px'>审核方法</h3><ol class='methods'>{methods_html}</ol></section>
<section class='section'><h2>六、导入评估结论</h2><div class='import'>{html.escape(report_import_conclusion(result))}<br><span style='font-weight:400'>源 Excel 始终只读；本次未连接真实 DM8/LOCATIONS 做导入验证。</span></div>{comparison_html}</section>
<section class='section'><h2>七、详细问题清单（按规则分组）</h2><p class='meta'>按 规则 → KKS → 行号明细 组织：每条规则一行（含问题定义/处置建议/规则依据），展开后列出该规则涉及的 KKS 编码（每个编码只出现一次）；再点击编码展开该编码下的行号明细。</p><div class='issue-tree'>{tree_html}</div>{tree_flat}</section>
<section class='section'><details class='technical'><summary>查看 AI 语义复核依据（技术人员）</summary><p class='meta'>{ai_meta} AI 仅对规则筛出的不确定项提供辅助判断，正式结论仍保留规则证据和人工确认入口。</p><table><thead><tr><th>等级</th><th>行号</th><th>KKS</th><th>AI 依据</th></tr></thead><tbody>{ai_rows}</tbody></table></details><details class='technical'><summary>查看结构扫描备注</summary><ul>{notes_html}</ul></details></section>
<p class='footer'>本报告定位为 KKS 编码质量审核验收报告。源 Excel 未修改；AI 复核过程和详细证据已保留在折叠区域及问题 Excel 的“AI复核（技术）”隐藏页。</p></main></body></html>"""
    path.write_text(body, encoding="utf-8")


def main() -> int:
    configure_logging(clear=True)
    parser = argparse.ArgumentParser(description="审核 KKS Excel 并生成 HTML 报告和 Excel 问题清单")
    parser.add_argument("input", type=Path)
    parser.add_argument("--output-dir", type=Path, default=Path("outputs/kks-audit"))
    parser.add_argument("--sheet", default=None)
    parser.add_argument("--compare-file", type=Path, default=None, help="可选：用于专业覆盖、编码深度和数量闭合对比的另一份 Excel")
    args = parser.parse_args()
    # Use the same workflow-first Agent entry point as the web service.  The
    # deterministic audit function remains the bounded workbook tool.
    from agent_runtime import KksAuditAgent

    result = KksAuditAgent().run(
        args.input,
        args.output_dir,
        requested_sheet=args.sheet,
        comparison_path=args.compare_file,
    )
    print(json.dumps({"conclusion": result["conclusion"], "issue_count": result.get("issue_count", 0), "resolved_count": result.get("resolved_count", 0), "output_dir": str(args.output_dir)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
