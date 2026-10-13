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
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Callable
from audit_metrics import confirmed_priority_counts, summary_metrics

import openpyxl
from kks_gen.compliance import assess, unit_conflicts, plant_code
from kks_gen.row_policy import classify_row, is_table_instruction
from ai_review import review_issue_candidates, summarize_audit
from kks_gen.audit_policy import findings as device_findings, STAGES, subject as device_subject
from audit_history import record_and_render
from app_runtime import APP_ROOT, configure_logging
from report_business import build_overview, render_html, render_xlsx
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
# GB/T 50549-2020 表4.3.1-2，全厂码H2H3为两位数字。
COMMON_UNITS = {"00"} | {f"{n:02d}" for n in range(61, 100)}
FREE_UNIT_LETTERS = set()
UNIT_LETTER_MAP = {}
ENTERPRISE_UNIT_PREFIXES = {f"{n:02d}": f"{n}号" for n in range(1, 61)}
ENTERPRISE_UNIT_PREFIXES.update({"00": "全厂公用", **{f"{n:02d}": "两机共用" for n in range(61, 91)}})
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
    standard = str(result.get("standard", STANDARD_KEY))
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
    code_node_index: dict[tuple[str, str], dict[str, Any]] = {}

    def bump(node: dict[str, Any], pr: str) -> None:
        # P-level badges in the tree follow the same confirmed-only meaning as
        # the dashboard and KPI cards; pending/false-positive entries remain in
        # the detailed hit list and its total.
        if str(current_item.get("final_decision") or "confirmed_issue") != "confirmed_issue":
            return
        if pr == "P0":
            node["p0"] += 1
        elif pr == "P1":
            node["p1"] += 1
        else:
            node["p2"] += 1

    for item in issues:
        current_item = item
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
        node = code_node_index.get((rid, code))
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
            code_node_index[(rid, code)] = node

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
            "final_decision": item.get("final_decision", "confirmed_issue"),
        })
        bump(rule, pr)
        bump(node, pr)
        rule["count"] += 1
        node["count"] += 1

    for rule in rules:
        rule["note"] = f"{len(rule['children'])} 个编码 · {rule['count']} 条命中"
        for node in rule["children"]:
            node["note"] = f"{node['count']} 条命中"
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

    P0/P1/P2 只统计 AI/规则最终确认为问题的条目；待核和疑似误报分别统计。
    另保留两个业务治理重点的聚焦计数（扩展码治理 KKS-17/22、历史迁移 KKS-18）
    供报告明细作为治理提示，不作为首页大数字口径。
    """
    metrics = summary_metrics(result)
    return {
        **metrics,
        "rule_p1": int(metrics["p1"]),
        "rule_p2": int(metrics["p2"]),
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
    "KKS-39": "设备交付编码严格十二位，数字和字母各段类型必须正确。",
    "KKS-40": "设备主体类别与附录B设备分类代码核对；主体有歧义时待核。",
    "KKS-41": "明确系统字段与附录A分类代码核对；缺少工程证据时待核。",
    "KKS-42": "名称功能、用途、介质和位置与分类含义的语义一致性审查。",
    "KKS-43": "设备流水编号有效性检查，特殊用途须确认工程约定。",
    "KKS-44": "机组、区域及位置归属辅助核查，不能凭名称推测归属。",
    "KKS-45": "缺少工程编号分配与备用设备约定时，编号合理性不能判定通过。",
    "KKS-46": "缺少工程位置、图纸及全厂台账时，位置与全厂唯一性不能判定通过。",
    "KKS-38": "项目设备表中指令/测点的设备KKS须留空；国标信号标识与设备标识分开管理。",
    "KKS-37": "分类、层级及工程约定待核：自动判据无法确认，须补充工程索引、图纸或约定；具体适用依据见问题说明。",
    "KKS-00": "数据行缺少 KKS 编码：非空记录应含设备 KKS 码，缺失则无法入库定位。",
    "KKS-01": "多 Sheet 定位主表：应区分主表与导出/草稿表，防止重复或漏审。",
    "KKS-02": "表头与数据起始行定位：表头不总在第 1 行，错读将导致整列错位。",
    "KKS-03": "关键列动态定位：按表头名匹配 KKS/父级/名称/类型列，不依赖固定列序。",
    "KKS-04": "字符合法性：仅允许大写字母 A-Z、数字 0-9 与分隔符 -；小写、中文、全角及 ~、/ 等未定义字符判非法。",
    "KKS-04b": "分段字符类型校验：12 位骨架逐段定类型——G/F1F2F3/A1A2 为字母段，F0/FN/AN 为数字段，数字段出现字母或字母段出现数字即错。",
    "KKS-05": "编码长度合规：设备级 12 位、系统级 5/6/7 位、根节点 1 位；超过 12 位不立即判错，转编码深度定性。",
    "KKS-06": "结构位取值合规：全厂码H₂H₃按表4.3.1-2取两位数字；I/O约束按标识类型和字母段判定。",
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
    "KKS-28": "机组三方一致：机组列、名称中的机组指代、全厂码H₂H₃三方须一致；共用范围依据工程约定复核。",
    "KKS-29": "OCR 易混字符待核：编码含 I/i/L/l/O/o 等易混字符，人工核对是否为数字 1/0 的识别错误。",
    "KKS-30": "文本卫生：码与名称的首尾空白、换行、制表符、全角空格等脏数据，标记并附清洗。",
    "KKS-31": "父子名称语义一致性：直接父子一级、同类设备的名称标识 token 应一致，矛盾即报。",
    "KKS-32": "系统分类码合法性：系统分类码 S1S2S3 应为 GB/T 50549-2020 附录A 收录的三位字母码（4.4.2 / 4.4.4-1）；未收录按码聚合提示待核。",
    "KKS-33": "设备分类码合法性：设备分类码 T1T2 应为 GB/T 50549-2020 附录B 收录的两位字母码（4.4.3 / 4.4.5-1）；含禁用字母 I/O 为硬错误（4.4.1 注），逐行报 P0。",
    "KKS-35": "阀门/管道编号段语义：QM（截止阀）、WP（管道）的设备编号按功能分段——001-299 通用、301-399 测量、401-499 疏水排气、601-699 取样；编号段与名称功能词矛盾即提示。",
    "KKS-A": "名称语义与设备字母一致：名称含泵/阀等词应对应泵类/阀类字母，不符即名实疑点。",
    "KKS-C": "命名风格归一：1号机/1#机/#1机/一号机 等风格混用，提示统一风格。",
    "KKS-D": "空/缺名称：名称列为空、-、无、None 等，设为必填字段时按 P0 阻断。",
    "KKS-E": "设备类型列与名称一致：类型列值（如 截止阀）与名称推导类型（如 闸阀）矛盾。",
    "KKS-F": "同父级编号断号：同父设备编号序列缺号（001/002/004），可能漏编。",
    "KKS-G": "字母组合合法性：系统字母与设备字母的标准配对校验。",
}


# ---- 规则依据出处（唯一依据：GB/T 50549-2020 + 中煤细则）----
#
# ⚠️ 版本纪律（2026-10-08 定稿）：规则依据**只允许**引用现行有效的标准。
#   GB/T 50549-2020 自 2021-03-01 实施，前言明确「原《电厂标识系统编码标准》
#   （GB/T 50549-2010）同时废止」。
#   此前曾把继承 2010 体系的两份文档一并作为依据——
#     · 《火电厂标识系统（KKS）编码导则》Q/WL（2015）
#     · 风电/光伏《标识系统（KKS）编码规范》CPI（旧版）
#   结果同一分类码出现两种含义（GP：2020 版是泵，2010 版体系是照明检修配电箱），
#   报告里刷出无实际意义的"标准分歧"。经实测（舟山暖通 1190 行）：只按 2020 版判定
#   与两套并集判定的告警数完全一致，故已全部移除，不再保留分歧披露逻辑。
#
# 结构对照（GB/T 50549-2020 4.4.1 功能面标识三级编码；火电项目 H1 缺省）：
#   全厂码 H2H3(2) + 系统分类码 S1S2S3(3) + 系统编号 S4S5(2) + 设备分类码 T1T2(2) + 设备编号 T3T4T5(3) = 12 位
#   与项目沿用的写法 G|F0 | F1F2F3 | F4F5 | A1A2 | AN1AN2AN3 逐位对应。
STANDARD_KEY = "gb50549"
STANDARD_BASIS = "GB/T 50549-2020《电厂标识系统编码标准》"
STANDARD_DISPLAY = (
    "国标/GBT 50549-2020 电厂标识系统编码标准.pdf（GB/T 50549-2020，现行国家标准，2021-03-01 实施）"
    " ＋ 国标/中国中煤能源集团电厂标识系统编码应用细则（试行）A3.2.docx（集团落地细则）"
)
GENERIC_SOURCE = "通用审计要求，无专条依据"

_GB = "GB/T 50549-2020《电厂标识系统编码标准》"
_ZM = "中煤集团《电厂标识系统编码应用细则（试行）》"
_ST_STRUCT = f"{_GB} 4.4.1 功能面标识三级编码（字符类型 N 为阿拉伯数字，A 为大写字母，禁用 I、O）"
_ST_SEG = f"{_GB} 4.4.1 功能面标识三级编码 + 4.4.4 系统码（S1S2S3/S4S5）+ 4.4.5 设备码（T1T2/T3T4T5）"
_ST_G = f"{_GB} 4.3.1 全厂码（火电项目 H1 缺省，取 H2H3，见表4.3.1-2）"
_ST_LEN = f"{_GB} 4.4.1 功能面标识三级编码（H1 缺省时 2+3+2+2+3=12 位）；4.2.2 联合标识各级不超过 12 位"
_ST_UNIQ = f"{_GB} 3.1.1 每一个被标识对象的标识应符合全厂唯一的原则"
_ST_PART = f"{_GB} 4.5 产品面标识（产品分类码按附录B 选用，见 4.5.3）"
_ST_DEV = f"{_GB} 附录B 设备及产品分类索引（4.4.3 设备分类码按附录B 选用）"
_ST_COMBO = f"{_GB} 附录A 系统分类索引 + 附录B 设备及产品分类索引（3.3.4 按附录摘取本项目采用的分类码）"
_ST_CODE_SYS = f"{_GB} 4.4.2 系统分类码应由 3 个大写英文字母组成，按附录A 选用（4.4.4-1 S1/S2/S3）"
_ST_CODE_DEV = f"{_GB} 4.4.3 设备分类码应由 2 个大写英文字母组成，按附录B 选用；4.4.1 注 禁用 I、O"
_ST_VALVE = f"{_GB} 表5.1.4 阀门设备分类码；表5.1.5 QM/WP 设备编号分段（{_ZM} 8.1.5、8.1.6）"

RULE_SOURCES: dict[str, str] = {
    "KKS-39": _ST_LEN,
    "KKS-40": _ST_CODE_DEV,
    "KKS-41": _ST_CODE_SYS,
    "KKS-42": _ST_COMBO,
    "KKS-43": _ST_SEG,
    "KKS-44": _ST_G,
    "KKS-45": f"{_GB} 第3.3节及第4.4.5节；工程编号约定证据要求",
    "KKS-46": f"{_GB} 第3.3节、第3.1.1节；工程位置及全厂唯一性证据要求",
    "KKS-04": _ST_STRUCT, "KKS-04b": _ST_SEG,
    "KKS-05": _ST_LEN, "KKS-06": _ST_G, "KKS-07": _ST_UNIQ,
    "KKS-08": _ST_STRUCT, "KKS-09": _ST_STRUCT, "KKS-10": _ST_STRUCT,
    "KKS-11": _ST_STRUCT, "KKS-12": _ST_STRUCT,
    "KKS-25": _ST_UNIQ, "KKS-25c": _ST_UNIQ, "KKS-26": _ST_PART,
    "KKS-28": _ST_G, "KKS-29": _ST_STRUCT,
    "KKS-A": _ST_DEV, "KKS-G": _ST_COMBO,
    "KKS-32": _ST_CODE_SYS, "KKS-33": _ST_CODE_DEV,
    "KKS-35": _ST_VALVE,
    "KKS-37": "GB/T 50549-2020 第3.3节及表4.3.1-2；具体分类、编号或项目约定依据见问题说明。",
    "KKS-38": "项目设备表交付约定：指令/测点不占设备KKS列；GB/T 50549-2020 第4.7节规定独立的信号标识。",
}


def rule_source(rule_id: str, standard: str | None = None) -> str:
    """取规则出处。

    2026-10-08 起依据不再按厂型分叉——统一引用 GB/T 50549-2020（国标本身覆盖火电、
    水电、核电、风电、光伏、太阳能热发电、厂外供热冷各章）+ 中煤细则。
    standard 形参保留只为不动历史调用点，已无实际作用。
    支持组键回退（如 KKS-13 命中 KKS-13/16）。
    """
    if rule_id in RULE_SOURCES:
        return RULE_SOURCES[rule_id]
    for key, value in RULE_SOURCES.items():
        if rule_id in key.split("/"):
            return value
    return GENERIC_SOURCE


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
        preferred = first(lambda s: "KKS" in s and any(x in s for x in ("新", "修编后", "最终")) and not any(x in s for x in ("原", "旧", "上级", "父级", "设计")))
        if preferred is not None:
            return preferred
        return first(lambda s: "KKS" in s and not any(x in s for x in ("原", "旧", "上级", "父级", "设计")))
    if kind == "original_name":
        return first(lambda s: "名称" in s and any(x in s for x in ("原", "设计")))
    if kind == "old":
        return first(lambda s: "KKS" in s and any(x in s for x in ("原", "旧")))
    if kind == "parent":
        preferred = first(lambda s: any(x in s for x in ("父级", "上级")) and "KKS" in s)
        if preferred is not None:
            return preferred
        return first(lambda s: "父级" in s or "上级" in s)
    if kind == "name":
        preferred = first(lambda s: ("修改后" in s or "新名称" in s or "新设备名称" in s) and "原" not in s)
        if preferred is not None:
            return preferred
        return first(lambda s: "名称" in s and not any(x in s for x in ("原", "机组", "系统", "上级", "父级")))
    if kind == "unit":
        return first(lambda s: ("机组" in s or "机组号" in s) and "名称" not in s)
    if kind == "dtype":
        return first(lambda s: "设备类型" in s or s == "类型")
    if kind == "profession":
        return first(lambda s: "专业" in s or "所属专业" in s)
    if kind == "system":
        # 系统/系统名称列（工艺系统归属，用于报告表格的"系统"列展示）
        return first(lambda s: "系统" in s and not any(x in s for x in ("KKS", "编码", "设备", "上级", "父级", "字母", "名称表")))
    return None


def audit_columns(header):
    columns = {kind: find_column(header, kind) for kind in
               ("kks", "old", "parent", "name", "original_name", "unit", "dtype", "profession", "system")}
    index = columns["kks"]
    finished = index is not None and any(term in text(header[index]) for term in ("新", "修编后", "最终"))
    columns["finished"] = finished
    if finished:
        for kind in ("old", "original_name", "unit", "dtype", "profession", "system"):
            columns[kind] = None
    return columns


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
        "finished_product": bool(cols.get("finished")),
        "excel_row": excel_row,
        "kks_code": get("kks"),
        "old_code": get("old"),
        "parent_code": get("parent"),
        "name": get("name") if cols.get("finished") else get("name") or get("original_name"),
        "unit": get("unit"),
        "device_type": get("dtype"),
        "profession": get("profession"),
        "system": get("system"),
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
        "sheet": rec.get("sheet", ""),
        "kks_code": rec.get("kks_code", ""),
        "parent_code": rec.get("parent_code", ""),
        "old_code": rec.get("old_code", ""),
        "name": rec.get("name", ""),
        "unit": rec.get("unit", ""),
        "profession": rec.get("profession", ""),
        "system": rec.get("system", ""),
        "message": message,
        "suggestion": suggestion,
        "evidence": evidence,
    })


def record_role(rec):
    kwargs = {"include_measurement": not rec.get("finished_product", False)}
    role = classify_row(rec.get("old_code"), rec.get("name"), rec.get("device_type"), **kwargs)
    return role if role.is_point else classify_row(rec.get("kks_code"), rec.get("name"), rec.get("device_type"), **kwargs)


def _add_compliance_issues(issues, rec):
    role = record_role(rec)
    if role.is_point and rec["kks_code"]:
        add_issue(issues, priority="P0", rule_id="KKS-38", category="指令/测点占用设备KKS", rec=rec,
                  message="指令/测点的设备KKS列应留空；" + role.reason,
                  suggestion="清空设备KKS，保留原信号标识与设备归属信息。")
        issues[-1]["compliance_locked"] = True
    assessment = assess(rec["kks_code"])
    for message in assessment.errors:
        add_issue(issues, priority="P0", rule_id="KKS-04b", category="国标及项目结构校验", rec=rec,
                  message=message, suggestion="按适用条款及工程约定整改后重新审核。")
        issues[-1]["compliance_locked"] = True
    for message in assessment.pending:
        add_issue(issues, priority="P1", rule_id="KKS-37", category="工程约定待核", rec=rec,
                  message=message, suggestion="补充工程编码索引、图纸或分类码约定。", status="needs_review")
        issues[-1]["definition"] = "自动检查无法确认的分类、编号及工程约定；具体适用条款见问题说明。"
        issues[-1]["evidence"] = message


def _set_compliance_result(result):
    hard = sum(bool(i.get("compliance_locked")) or i.get("rule_id") == "KKS-07" for i in result.get("issues", []))
    result["compliance"] = {"standard": "GB/T 50549-2020", "status": "blocked" if hard else "pending",
                            "hard_findings": hard, "scope": "工作簿自动校验；工程约定、图纸符合性及全厂唯一性仍须确认"}
    caveat = "国标审核范围限于已实现的自动判据，缺少工程约定、图纸及全厂台账时不能认定全面合规。"
    if caveat not in result.get("conclusion", ""):
        result["conclusion"] = result.get("conclusion", "") + " " + caveat


def code_unit(code: str, common_units=None) -> int | None:
    prefix = plant_code(code)
    return int(prefix) if prefix and 1 <= int(prefix) <= 60 else None


def valid_g_char(value: str) -> bool:
    """Compatibility name: validate complete two-digit H2H3, never one char."""
    return bool(re.fullmatch(r"[0-9]{2}", value))


def enterprise_unit_ok(code: str, enterprise_prefixes=None) -> bool:
    return bool(plant_code(code))


def enterprise_mode(code_prefixes: set[str]) -> bool:
    return bool(code_prefixes) and all(re.fullmatch(r"[0-9]{2}", p) for p in code_prefixes)


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


@contextmanager
def load_workbook_rows(input_path: Path):
    """打开工作簿并**保证关闭**，供调用方立即取完数据。

    ⚠️ 为什么必须有这个包装：openpyxl 的 read_only 模式会一直持有文件句柄，直到
    `wb.close()`。此前审核完从没关过，导致 Windows 上 `runs/<run_id>/input/*.xlsx`
    一直处于"被占用"状态——台账删除产物、移动/重命名 run 目录都会报
    `WinError 32 另一个程序正在使用此文件`。项目里所有 `iter_rows` 都会 `list(...)`
    物化数据，所以取完就能安全关闭。
    """
    wb = openpyxl.load_workbook(input_path, read_only=True, data_only=True)
    try:
        yield wb
    finally:
        try:
            wb.close()
        except Exception:  # 关闭失败不能影响审核结果
            LOGGER.warning("workbook_close_failed file=%s", getattr(input_path, "name", input_path))


def summarize_scope(input_path: Path) -> dict[str, Any]:
    """Summarize one workbook for the Skill's cross-scope comparison rules."""
    with load_workbook_rows(input_path) as wb:
        sheet, rows, header_idx = choose_sheet(wb, None)
        header = rows[header_idx]
        cols = {kind: find_column(header, kind) for kind in ("kks", "old", "parent", "name", "unit", "dtype", "profession", "system", "original_name")}
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
        if cols.get("finished") and idx not in {cols.get("kks"), cols.get("parent"), cols.get("name")}:
            continue
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

    # 同名多码：按聚簇"每条一组"上报（不再按簇内每个编码逐条报），
    # 避免同一簇被重复计数成 N 条把清单撑爆；组内全部编码在消息中列出。
    for system_key, normalized_name, codes in rules.near_dup_by_name(template_cols, template_rows):
        index = next((i for i, rec in records_by_index.items() if rec.get("kks_code") == codes[0]), None)
        shown = ", ".join(codes[:6]) + ("…" if len(codes) > 6 else "")
        _append_skill_issue(
            issues, seen, records_by_index, rule_id="KKS-25", priority="P1", category="同名对应多个 KKS 码", index=index,
            message=f"同一系统 {system_key} 内，归一名称 {normalized_name!r} 对应 {len(codes)} 个编码：{shown}。",
            suggestion="仅作为人工抽检线索；结合原KKS/设计编码、位置和图纸证据确认，不自动合并。",
            status="needs_review", evidence=f"系统前9位={system_key}；编码数={len(codes)}",
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
        first_code = codes[0]
        index = next((i for i, rec in records_by_index.items() if rec.get("kks_code") == first_code), None)
        names = sorted({item[2] for item in items if item[2]})
        shown = ", ".join(codes[:6]) + ("…" if len(codes) > 6 else "")
        _append_skill_issue(
            issues, seen, records_by_index, rule_id="KKS-25c", priority="P1", category="原KKS一对多", index=index,
            message=f"原KKS {identity_key!r} 对应 {len(codes)} 个新码：{shown}。",
            suggestion="以原KKS/旧版设计编码作为 1:1 身份键人工裁定，不自动合并。",
            evidence=f"身份键={identity_key}; 名称={'/'.join(names[:4])}",
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

    # ---- 国标分类码语义（GB/T 50549-2020 附录A/B + 表5.1.4/5.1.5）----
    # 字典来自 tools/build_code_index.py；未收录一律作待核线索（附录分页可能有极少数码未提取），
    # 只有"含禁用字母 I/O"是硬错误。按分类码聚合上报，避免同一码刷屏。
    for item in rules.system_class_code_legality(template_cols, template_rows):
        _append_skill_issue(
            issues, seen, records_by_index, rule_id="KKS-32", priority=item["severity"],
            category="系统分类码未收录", index=item["index"],
            message=f'{item["message"]}（本表 {item["count"]} 处，示例 {"、".join(item["rows"][:3])}）',
            suggestion=item["suggestion"], status="needs_review",
            evidence=f'分类码={item["segment"]}; 出现 {item["count"]} 次',
        )
    diagnostics["executed_rules"].append("KKS-32")

    for item in rules.device_class_code_legality(template_cols, template_rows):
        hard = item["severity"] == "P0"
        _append_skill_issue(
            issues, seen, records_by_index, rule_id="KKS-33", priority=item["severity"],
            category="设备分类码含禁用字母" if hard else "设备分类码未收录", index=item["index"],
            message=item["message"] if hard
                    else f'{item["message"]}（本表 {item["count"]} 处，示例 {"、".join(item["rows"][:3])}）',
            suggestion=item["suggestion"], status="open" if hard else "needs_review",
            evidence=f'分类码={item["segment"]}; 出现 {item["count"]} 次',
        )
    diagnostics["executed_rules"].append("KKS-33")

    for index, code, reason, suggestion, priority in rules.valve_pipe_number_segment_check(template_cols, template_rows):
        _append_skill_issue(
            issues, seen, records_by_index, rule_id="KKS-35", priority=priority,
            category="阀门/管道编号段语义", index=index,
            message=reason, suggestion=suggestion, status="needs_review",
        )
    diagnostics["executed_rules"].append("KKS-35")

    return diagnostics


# 设备索引字母（A 码）——10 版 VGB 旧字母（QM/MA/BT/FT/LS 等）不再用于名实一致性判定
DEVICE_LETTER_SEMANTICS = {
    "泵": {"GP"}, "阀": {"QM", "QN", "RM", "FL", "FM", "QA"},
    "测温": {"BT"}, "温度": {"BT"}, "流量": {"BF"}, "液位": {"BL"},
    "风机": {"GQ"}, "压缩机": {"GQ"},
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
        from kks_gen.identifiers import parse
        letters = parse(code).segments.get("dev", "")
        if letters:
            matched = device_subject(name)
            if matched:
                term, allowed = matched
                # “供给水泵管道/泵系统”中的泵是工艺上下文，不是本行设备主体；
                # 只有最后一个明确设备词才进入名称-设备字母一致性检查。
                if letters not in allowed:
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
        dtype_subject, name_subject = device_subject(dtype), device_subject(name)
        if dtype_subject and name_subject and not dtype_subject[1].intersection(name_subject[1]):
            _append_skill_issue(
                issues, seen, records_by_index, rule_id="KKS-E", priority="P2", category="设备类型与名称矛盾", index=index,
                message=f"设备类型列为“{dtype}”，名称主体为“{name_subject[0]}”。", suggestion="人工核对设备类型列和名称，不自动覆盖任一字段。", status="needs_review",
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
    converged_old_lens: Counter[int] = Counter()
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
            # 扩展码/旧码 → 12 位主码的"收敛"是增量新增/标准化业务本身（本次文件的
            # 新旧长度差异绝大多数源于此），不逐条报 P1；改计聚合数供"范围与结构说明"引用。
            if len(new) == 12 and len(old) != 12:
                converged_old_lens[len(old)] += 1
                continue
            _append_skill_issue(
                issues, seen, records_by_index, rule_id="KKS-15", priority="P1", category="层级位置变化需人工核", index=index,
                message=f"原码长度 {len(old)} 与新码长度 {len(new)} 不同，且新码不是 12 位主码，存在层级重排或扩展建模变化。", suggestion="结合设备语义确认层级是否重排，不自动判为错误。", status="needs_review",
            )
    diagnostics["converged_to_master_old_len"] = dict(sorted(converged_old_lens.items()))
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
    """Review existing findings only; never request replacement codes."""
    result.pop("code_suggestions", None)
    try:
        result["ai_review"] = review_issue_candidates(result, progress_callback=progress_callback)
    except Exception as exc:
        LOGGER.warning("ai_review_stage_failed error=%s", exc)
        result["ai_review"] = {"status": "failed", "error": str(exc)[:200]}
    for issue in result.get("issues", []):
        if issue.get("compliance_locked") or issue.get("rule_id") == "KKS-07":
            issue["final_decision"] = "confirmed_issue"
            issue["status"] = "open"
        elif issue.get("evidence_missing"):
            issue["final_decision"] = "needs_human"
            issue["status"] = "needs_review"
            issue["final_action"] = "需要人工确认"
            issue["final_reason"] = "缺少关联资料，智能判断不能替代缺失的证据。"
    _set_compliance_result(result)
    result["ai_summary"] = summarize_audit(result, progress_callback=progress_callback)
    result["final_decision_counts"] = final_decision_counts(result)
    # The dashboard severity distribution must match confirmed P0/P1/P2 counts.
    # Keep issue_count as the full rule-hit workload for detail/history views.
    result["priority_counts"] = confirmed_priority_counts(result)
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
    result.pop("_ai_evidence_index", None)
    def strip_proposed_codes(value):
        if isinstance(value, dict):
            for field, child in value.items():
                if isinstance(child, str) and (field == "suggestion" or field.endswith("_suggestion")):
                    value[field] = re.sub(r"(?<![A-Za-z0-9])[0-9]{2}[A-Z]{1,3}(?:[0-9]{2}(?:[A-Z]{2}[0-9]{3})?)?(?:[-.][A-Z0-9]+)*(?![A-Za-z0-9])", "〔编码值省略〕", child)
                else:
                    strip_proposed_codes(child)
        elif isinstance(value, list):
            for child in value:
                strip_proposed_codes(child)
    strip_proposed_codes(result)
    return result


def write_audit_artifacts(
    input_path: Path,
    output_dir: Path,
    result: dict[str, Any],
    *,
    progress_callback: ProgressCallback | None = None,
) -> dict[str, Any]:
    result["quality_metrics"] = quality_metrics(result)
    """Write only the two user-facing audit artifacts."""
    _emit_progress(progress_callback, {"phase": "audit", "stage": "report", "percent": 95, "message": "正在生成审核报告", "hint": "整理 HTML 报告和问题 Excel"})
    output_dir.mkdir(parents=True, exist_ok=True)
    html_name, xlsx_name = output_artifact_names(input_path)
    overview = build_overview(result)
    write_html(output_dir / html_name, result, overview)
    write_issue_workbook_xlsx(output_dir / xlsx_name, result, overview)
    # 审计台账：追加本次记录并更新跨文件汇总报告（含 AI 趋势总结）
    try:
        summary_path = record_and_render(input_path, output_dir, result, progress_callback=progress_callback)
        result["history_summary"] = str(summary_path)
    except Exception as exc:  # 台账失败不阻断主流程
        LOGGER.warning("audit_history_record_failed error=%s", exc)
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
        with load_workbook_rows(input_path) as workbook:
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
        return code

    for excel_row, row in enumerate(rows[header_idx + 1:], start=header_idx + 2):
        parent_code = _tree_value(row, tree_cols["parent_code"])
        child_code = _tree_value(row, tree_cols["child_code"])
        parent_name = _tree_value(row, tree_cols["parent_name"])
        child_name = _tree_value(row, tree_cols["child_name"])
        if not child_code and classify_row(name=child_name).is_point:
            continue
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
        _add_compliance_issues(issues, rec)
        if parent_code:
            _add_compliance_issues(issues, _tree_record(excel_row, parent_code, "", parent_name))
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
        "standard": STANDARD_KEY,
        "standard_display": STANDARD_DISPLAY,
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
    _set_compliance_result(result)
    if write_outputs:
        write_audit_artifacts(input_path, output_dir, result, progress_callback=progress_callback)
    return result


def _audit_sheet(
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
    rule_config = _load_rule_config()
    root_parents = {text(item) for item in rule_config.get("root_parent_values", ROOT_PARENTS)} or ROOT_PARENTS
    old_prefix_migrations = rule_config.get("old_prefix_migrations", OLD_PREFIX)
    if not isinstance(old_prefix_migrations, dict):
        old_prefix_migrations = OLD_PREFIX
    common_units = {text(item) for item in rule_config.get("common_unit_prefixes", COMMON_UNITS)} or COMMON_UNITS
    accepted_lengths = {int(item) for item in rule_config.get("accepted_code_lengths", [2, 3, 4, 5, 7, 12, 13])}
    # 工作簿在这里读完就关掉（见 load_workbook_rows）：句柄不释放会让后续
    # "移除产物 / 移动 run 目录"在 Windows 上失败（WinError 32）
    with load_workbook_rows(input_path) as wb:
        sheet_info = inspect_sheets(wb)
        sheet, rows, header_idx = choose_sheet(wb, requested_sheet)
        workbook_codes = defaultdict(list)
        other_points = []
        for candidate_sheet in wb.worksheets:
            candidate_rows = rows if candidate_sheet.title == sheet else list(candidate_sheet.iter_rows(values_only=True))
            try:
                candidate_header = detect_header(candidate_rows)
            except ValueError:
                continue
            candidate_cols = audit_columns(candidate_rows[candidate_header])
            if candidate_cols["kks"] is None:
                continue
            for row_number, candidate_row in enumerate(candidate_rows[candidate_header + 1:], candidate_header + 2):
                candidate = make_record(row_number, candidate_row, candidate_cols)
                if is_table_instruction(candidate["kks_code"], candidate["name"], candidate["old_code"]) or not candidate["kks_code"]:
                    continue
                candidate["sheet"] = candidate_sheet.title
                workbook_codes[candidate["kks_code"].upper()].append(candidate)
                if candidate_sheet.title != sheet and record_role(candidate).is_point:
                    other_points.append(candidate)
    _emit_progress(progress_callback, {"phase": "audit", "stage": "structure", "percent": 25, "message": "正在识别工作表结构", "hint": f"当前工作表：{sheet}；数据行：{max(0, len(rows) - header_idx - 1)}"})
    header = rows[header_idx]
    cols = audit_columns(header)
    missing = [kind for kind in ("kks", "name") if cols[kind] is None and not (kind == "name" and cols.get("original_name") is not None)]
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
    for context_record in template_records.values():
        context_record["sheet"] = sheet
    for i, row in enumerate(bounded_rows, start=header_idx + 2):
        if not nonempty(row):
            blank_rows.append(i)
            continue
        rec = make_record(i, row, cols)
        rec["sheet"] = sheet
        if rec.get("finished_product") and not any(rec.get(k) for k in ("kks_code", "parent_code", "name")):
            continue
        if is_table_instruction(rec["kks_code"], rec["name"], rec["old_code"]):
            continue
        if rec["kks_code"]:
            records.append(rec)
        elif not classify_row(rec.get("old_code"), rec["name"], rec.get("device_type")).is_point:
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

    # Include cross-sheet and case-normalized collisions without repeating
    # the exact duplicates already reported for the selected sheet.
    for code, group in workbook_codes.items():
        if len(group) < 2 or (all(rec["sheet"] == sheet for rec in group)
                              and len({rec["kks_code"] for rec in group}) == 1):
            continue
        locations = "、".join(f"{rec['sheet']}!{rec['excel_row']}" for rec in group)
        for rec in group:
            add_issue(issues, priority="P0", rule_id="KKS-07", category="工作簿重复KKS", rec=rec,
                      message=f"编码 {code} 在工作簿重复：{locations}", suggestion="核对所有冲突对象，编码须唯一。")
            issues[-1]["sheet"] = rec["sheet"]
            issues[-1]["compliance_locked"] = True
    for rec in other_points:
        add_issue(issues, priority="P0", rule_id="KKS-38", category="指令/测点占用设备KKS", rec=rec,
                  message=f"{rec['sheet']}!{rec['excel_row']} 指令/测点的设备KKS列应留空。",
                  suggestion="清空设备KKS，保留原信号标识与设备归属信息。")
        issues[-1]["sheet"] = rec["sheet"]
        issues[-1]["compliance_locked"] = True

    # 模式判定：增量新增批次 / 企业两位数字前缀惯例
    # 增量批次（父级系统码在主库、本文件几乎全为设备级码）里，"父级不在本文件"
    # 是预期而非缺陷，应转为范围验证项（与广元主库闭合校验），不报 P0 孤儿。
    code_prefixes = {code[:2] for code in code_set if code}
    enterprise = enterprise_mode(code_prefixes)
    total_parent_refs = 0
    external_parent_refs = 0
    external_parent_lens: Counter[int] = Counter()
    for rec in records:
        parent = rec["parent_code"]
        if parent and parent not in root_parents:
            total_parent_refs += 1
            if parent not in code_set:
                external_parent_refs += 1
                external_parent_lens[len(parent)] += 1
    incremental_batch = (
        total_parent_refs > 0
        and external_parent_refs / total_parent_refs >= 0.9
        and all(length < 12 for length in external_parent_lens)
    )
    scope_external_parents: Counter[str] = Counter()

    prefix_bad: list[dict[str, Any]] = []
    orphan_recs: list[dict[str, Any]] = []
    for rec in records:
        code, parent = rec["kks_code"], rec["parent_code"]
        if not code:
            continue
        external_parent = bool(parent) and parent not in root_parents and parent not in code_set
        if external_parent and incremental_batch:
            # 增量批次：父级系统/子系统节点在主库（或本批次新增子系统之外），
            # 记录到范围验证；孤儿、前缀、长度类不再按本文件内部缺陷报 P0。
            scope_external_parents[parent] += 1
            # KKS-15 的层级收敛提示、KKS-25 的名称聚簇也归入范围说明，避免 8 千级假阳性
        _add_compliance_issues(issues, rec)
        if parent and parent not in root_parents and parent not in code_set and not incremental_batch:
            orphan_recs.append(rec)
            add_issue(issues, priority="P0", rule_id="KKS-08", category="父级孤儿", rec=rec,
                      message=f"父级 {parent} 不存在于主表编码集合。", suggestion="补齐父级节点或确认是否引用了未迁移的旧码。")
        if len(code) == 2 and parent not in root_parents and parent in code_set:
            add_issue(issues, priority="P1", rule_id="KKS-12", category="根节点父级非哨兵", rec=rec,
                      message=f"根节点 {code} 的父级为 {parent}，不是空值或 -1 哨兵。", suggestion="核对根节点层级；根节点父级应使用空值或 -1。")
        if parent and parent not in root_parents and parent in code_set and not code.startswith(parent):
            prefix_bad.append(rec)
            add_issue(issues, priority="P0", rule_id="KKS-09", category="父子前缀不一致", rec=rec,
                      message=f"子码 {code} 不以父码 {parent} 开头。", suggestion="核对显式父级指针；舟山变长层级不要按固定长度推导。")
        if parent == code:
            add_issue(issues, priority="P0", rule_id="KKS-10", category="自引用", rec=rec,
                      message="父级编码与本行编码相同。", suggestion="修正父级指针。")
        if parent and parent not in root_parents and len(parent) >= len(code):
            add_issue(issues, priority="P0", rule_id="KKS-11", category="父级长度异常", rec=rec,
                      message=f"父级长度 {len(parent)} 不小于子码长度 {len(code)}。", suggestion="核对层级关系。")

        if rec["name"] == "":
            add_issue(issues, priority="P0", rule_id="KKS-D", category="有码无名称", rec=rec,
                      message="有 KKS 码但设备名称为空（含空格占位）。", suggestion="补齐设备名称后再审核。")
        raw = rec["raw_kks"]
        if raw != raw.strip() or any(ch in raw for ch in "\n\r\t\u3000\xa0"):
            add_issue(issues, priority="P1", rule_id="KKS-30", category="编码文本卫生", rec=rec,
                      message="编码含首尾空白、换行、制表符或全角空格。", suggestion="建议清洗文本后重新审核，不直接覆盖源表。")
        raw_name = rec["raw_name"]
        whitespace_hits = rec["name"] and any(ch in raw_name for ch in "\n\r\t\u3000\xa0")
        if whitespace_hits:
            # 名称为空的行已由 KKS-D(空名) 报告，此处不再重复报名称卫生
            add_issue(issues, priority="P1", rule_id="KKS-30", category="名称文本卫生", rec=rec,
                      message="名称含首尾空白、换行、制表符或全角空格。", suggestion="建议清洗文本后重新审核。")
        elif rec["name"] and re.search(r"[\uff01-\uff5e\u3000-\u303f]", raw_name):
            # 全角符号（（）、：、，等）→ P2 半角化清洗提示（与广元口径一致）
            add_issue(issues, priority="P2", rule_id="KKS-30", category="名称文本卫生(全角)", rec=rec,
                      message="名称含全角符号（括号/冒号/逗号等），建议半角化统一。", suggestion="全角→半角统一、去除首尾空格后再合入主库。")
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
        if rec["name"] and (len(rec["name"]) < 2 or not re.search(r"[\u4e00-\u9fffA-Za-z]", rec["name"])):
            # 空名行已由 KKS-D 报告，这里只对"非空但语义可疑"的短名/纯符号名提示人工核
            add_issue(issues, priority="P2", rule_id="KKS-27", category="名称语义需人核", rec=rec,
                      message="名称过短或无法提取中文/字母语义。", suggestion="人工核对名称，不按此项自动判错。", status="needs_review")

        for message in unit_conflicts(code, rec["name"], rec["unit"]):
            add_issue(issues, priority="P1", rule_id="KKS-28", category="机组三方不一致", rec=rec,
                      message=message, suggestion="核对图纸及工程约定，不自动改归属。", status="needs_review")

        old, new = rec["old_code"], rec["kks_code"]
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
        [tuple(None for _ in row) if is_table_instruction(template_records[index]["kks_code"], template_records[index]["name"], template_records[index]["old_code"]) or classify_row(template_records[index].get("old_code") or template_records[index]["kks_code"],
                                                template_records[index]["name"], template_records[index].get("device_type")).is_point else row
         for index, row in enumerate(bounded_rows)],
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
        if rec.get("finished_product") and not device_subject(rec.get("name", "")):
            add_issue(issues, priority="P1", rule_id="KKS-00", category="成品字段缺失待核", rec=rec,
                      message="新编码为空，成品名称不足以确认设备主体或是否属于应留空的测点。",
                      suggestion="补充成品名称或确认留空属性。", status="needs_review")
            issues[-1]["evidence_missing"] = True
            continue
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
    # 原KKS 一对多已由模板规则 identity_dup_by_key(KKS-25c) 按身份键逐组上报，
    # 此处不再重复逐行添加，避免同源问题双份计数。

    # 范围与结构说明：增量新增批次(父级在主库)、旧扩展码收敛为 12 位主码、
    # 企业两位前缀惯例，都属于"范围/约定"而非文件内部缺陷，汇总成结构化说明供报告引用。
    converged_old_lens = additional_skill_diagnostics.get("converged_to_master_old_len", {}) or {}
    converged_rows = sum(converged_old_lens.values())
    parent_rows_external = sum(scope_external_parents.values())
    root_rows = sum(1 for r in records if r["parent_code"] in root_parents)
    scope_notes: list[str] = []
    if incremental_batch:
        scope_notes.append(
            f"本文件为增量新增批次：{parent_rows_external} 行引用的父级系统/子系统节点在主库、不在本文件内"
            f"（去重 {len(scope_external_parents)} 个），按主库闭合校验处理，不视为孤儿缺陷。"
        )
        if converged_rows:
            scope_notes.append(
                f"共 {converged_rows} 行原扩展码/旧码（长度 {', '.join(f'{k}位' for k in sorted(converged_old_lens))}）"
                "收敛为 12 位主码，属本次新增的标准化业务，未逐条列入问题。"
            )
    elif converged_rows:
        scope_notes.append(
            f"共 {converged_rows} 行旧码非 12 位而新码为 12 位主码，按标准化收敛业务处理，未逐条列入问题。"
        )
    if enterprise:
        scope_notes.append(
            "识别为国标两位数字全厂码（00=全厂公用 / 01=1号 / 02=2号 / 61=1,2号公用），"
            "按 GB/T 50549-2020 表4.3.1-2 校验全厂码 H₂H₃。"
        )
    scope_validation = {
        "mode": "incremental" if incremental_batch else "full",
        "incremental_batch": incremental_batch,
        "enterprise": enterprise,
        "data_rows": len(records),
        "root_rows": root_rows,
        "external_parent_rows": parent_rows_external,
        "external_parent_units": len(scope_external_parents),
        "external_parent_samples": [{"parent": p, "children": n} for p, n in scope_external_parents.most_common(15)],
        "converged_to_master_rows": converged_rows,
        "converged_old_lengths": dict(sorted(converged_old_lens.items())),
        "notes": scope_notes,
    }

    # Preserve a structural note rather than treating intentionally blank separator rows as records.
    notes = list(scope_notes)
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
        "standard": STANDARD_KEY,
        "standard_display": STANDARD_DISPLAY,
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
        "scope_validation": scope_validation,
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
    _set_compliance_result(result)
    if write_outputs:
        write_audit_artifacts(input_path, output_dir, result, progress_callback=progress_callback)
    return result


def audit_file(input_path: Path, output_dir: Path, requested_sheet=None, comparison_path=None,
               progress_callback=None, *, run_ai_review=True, write_outputs=True) -> dict[str, Any]:
    """Audit every recognizable device sheet before semantic review and export."""
    input_path, output_dir = Path(input_path), Path(output_dir)
    eligible, excluded, all_records = [], [], []
    with load_workbook_rows(input_path) as workbook:
        if requested_sheet and requested_sheet not in workbook.sheetnames:
            raise ValueError(f"工作表不存在：{requested_sheet}")
        for ws in workbook.worksheets:
            if requested_sheet and ws.title != requested_sheet:
                excluded.append({"sheet": ws.title, "reason": "用户指定审查工作表，未纳入本次完整审查"})
                continue
            rows = list(ws.iter_rows(values_only=True))
            try:
                header_index = detect_header(rows)
                columns = audit_columns(rows[header_index])
                tree_columns = detect_tree_collection_columns(rows[header_index])
            except (ValueError, IndexError):
                excluded.append({"sheet": ws.title, "reason": "未识别设备表头或工作表为空"})
                continue
            if columns["kks"] is None or (columns["name"] is None and columns["original_name"] is None and not tree_columns):
                excluded.append({"sheet": ws.title, "reason": "缺少可识别的编码列或名称列"})
                continue
            eligible.append(ws.title)
            for number, row in enumerate(rows[header_index + 1:], header_index + 2):
                if not nonempty(row):
                    continue
                rec = make_record(number, row, columns)
                if tree_columns:
                    rec = _tree_record(number, _tree_value(row, tree_columns["child_code"]),
                                       _tree_value(row, tree_columns["parent_code"]),
                                       _tree_value(row, tree_columns["child_name"]))
                rec["sheet"] = ws.title
                if rec.get("finished_product") and not any(rec.get(k) for k in ("kks_code", "parent_code", "name")):
                    continue
                if not is_table_instruction(rec["kks_code"], rec["name"], rec["old_code"]):
                    all_records.append(rec)
    if not eligible:
        raise ValueError("没有可执行完整审查的设备工作表；请检查编码列和设备名称列")
    results = []
    for sheet in eligible:
        part = _audit_sheet(input_path, output_dir, sheet, comparison_path,
                            progress_callback, run_ai_review=False, write_outputs=False)
        for issue in part["issues"]:
            if not issue.get("sheet"):
                issue["sheet"] = sheet
        for rec in part.get("_ai_context", {}).get("records", []):
            rec["sheet"] = sheet
        results.append(part)
    result = results[0]
    issues, seen = [], set()
    for part in results:
        for issue in part["issues"]:
            key = tuple(str(issue.get(field, "")) for field in
                        ("sheet", "excel_row", "kks_code", "rule_id", "category", "message"))
            if key not in seen:
                seen.add(key)
                issues.append(issue)
    referenced = {r["parent_code"] for r in all_records if r["parent_code"]}
    global_codes = {r["kks_code"] for r in all_records if r["kks_code"]}
    for rec in all_records:
        parent = rec.get("parent_code", "")
        if rec.get("finished_product") and parent and parent not in ROOT_PARENTS:
            for parent_error in assess(parent).errors:
                add_issue(issues, priority="P0", rule_id="KKS-04b", category="父级编码结构错误", rec=rec,
                          message="父级编码：" + parent_error, suggestion="核对父级编码的适用层级和结构。",
                          evidence="父级KKS编码=" + parent)
                issues[-1].update(compliance_locked=True, audit_stage=0)
        row_findings = device_findings(rec, referenced)
        missing_subject = any(stage == 1 and "不能唯一确认设备主体" in message for stage, message, _, _ in row_findings)
        missing_system = any(stage == 2 and not confirmed for stage, _, _, confirmed in row_findings)
        for stage, message, basis, confirmed in row_findings:
            add_issue(issues, priority="P0" if confirmed else "P1", rule_id=("KKS-39", "KKS-40", "KKS-41", "KKS-42", "KKS-43", "KKS-44")[stage],
                      category=STAGES[stage], rec=rec, message=message,
                      suggestion="核对原始资料和适用依据后重新审查。", evidence=basis,
                      status="open" if confirmed else "needs_review")
            issues[-1]["audit_stage"] = stage
            issues[-1]["definition"] = basis
            issues[-1]["compliance_locked"] = confirmed
            issues[-1]["evidence_missing"] = ((stage == 2 and not confirmed)
                or (stage == 1 and missing_subject)
                or (stage == 3 and not confirmed and (missing_subject or missing_system)))
            if issues[-1]["evidence_missing"]:
                issues[-1]["review_scope_note"] = "设备主体或工艺归属缺少确定依据；直接列为待核，AI不得补造工程事实。"
    # A parent in another audited sheet is present; absent external parents
    # remain unverified regardless of an automatically inferred batch mode.
    issues = [i for i in issues if not (i.get("rule_id") == "KKS-08"
                                      and i.get("parent_code") in global_codes)]
    existing_parent_findings = {(i.get("sheet"), i.get("excel_row")) for i in issues if i.get("rule_id") == "KKS-08"}
    for rec in all_records:
        parent = rec.get("parent_code", "")
        if parent and parent not in ROOT_PARENTS and parent not in global_codes:
            if (rec["sheet"], rec["excel_row"]) not in existing_parent_findings:
                add_issue(issues, priority="P1", rule_id="KKS-08", category="外部父级待核", rec=rec,
                          message="父级未出现在本次审查台账，尚未核实外部台账中的存在性。",
                          suggestion="提供对应父级台账或工程编码索引。", status="needs_review",
                          evidence="GB/T 50549-2020 第3.3节；工作簿父级闭合校验")
                issues[-1]["evidence_missing"] = True
        elif parent in global_codes and len(parent) == 7 and len(rec["kks_code"]) == 12 and not rec["kks_code"].startswith(parent):
            if not any(i.get("sheet") == rec["sheet"] and i.get("excel_row") == rec["excel_row"]
                       and i.get("rule_id") == "KKS-09" for i in issues):
                add_issue(issues, priority="P0", rule_id="KKS-09", category="跨表系统归属冲突", rec=rec,
                          message="设备编码与明确的父级系统标识不一致。", suggestion="核对显式父级和工程系统归属。")
                issues[-1]["compliance_locked"] = True
    grouped = defaultdict(list)
    for rec in all_records:
        if rec["kks_code"]:
            grouped[rec["kks_code"].upper()].append(rec)
    duplicate_findings = {(i.get("sheet"), i.get("excel_row")) for i in issues if i.get("rule_id") == "KKS-07"}
    for code, group in grouped.items():
        if len(group) < 2:
            continue
        locations = "、".join(f"{r['sheet']}!{r['excel_row']}" for r in group)
        for rec in group:
            if (rec["sheet"], rec["excel_row"]) in duplicate_findings:
                continue
            add_issue(issues, priority="P0", rule_id="KKS-07", category="工作簿重复KKS", rec=rec,
                      message=f"编码{code}在工作簿重复：{locations}。", suggestion="核对全部冲突对象。")
            issues[-1]["compliance_locked"] = True
    identities = defaultdict(list)
    for rec in all_records:
        if rec.get("old_code") and rec["kks_code"] and not record_role(rec).is_point:
            identities[rec["old_code"]].append(rec)
    for old, group in identities.items():
        if len({r["kks_code"] for r in group}) > 1 and len({r["sheet"] for r in group}) > 1:
            for rec in group:
                add_issue(issues, priority="P1", rule_id="KKS-25c", category="跨表设备身份一对多", rec=rec,
                          message="同一原始设备标识跨工作表对应多个当前编码，设备身份一致性待核。",
                          suggestion="结合设备台账及历史映射确认物理设备身份。", status="needs_review",
                          evidence="原始设备标识=" + old)
    unique = {code: group[0] for code, group in grouped.items() if len(group) == 1}
    for code in cycle_nodes(unique):
        rec = unique[code]
        if not any(i.get("sheet") == rec["sheet"] and i.get("excel_row") == rec["excel_row"]
                   and i.get("rule_id") == "KKS-10" for i in issues):
            add_issue(issues, priority="P0", rule_id="KKS-10", category="跨表父级环", rec=rec,
                      message="工作簿父级链存在循环引用。", suggestion="核对并修正层级关系。")
            issues[-1]["compliance_locked"] = True
    # Stable order: format, type, system, semantics, numbering, location.
    stage_by_rule = {"KKS-00": 0, "KKS-04": 0, "KKS-04b": 0, "KKS-05": 0, "KKS-06": 0,
                     "KKS-33": 1, "KKS-A": 1, "KKS-E": 1, "KKS-32": 2, "KKS-09": 2,
                     "KKS-35": 4, "KKS-F": 4, "KKS-28": 5}
    for issue in issues:
        issue.setdefault("audit_stage", stage_by_rule.get(issue["rule_id"], 3))
        if issue["rule_id"] in {"KKS-32", "KKS-37", "KKS-F", "KKS-28", "KKS-44"} or (issue["rule_id"] == "KKS-33" and issue.get("status") == "needs_review"):
            issue["evidence_missing"] = True
            issue["status"] = "needs_review"
        if issue["rule_id"] in {"KKS-07", "KKS-10", "KKS-11", "KKS-09", "KKS-D", "KKS-00"} and issue.get("status") == "open":
            issue["compliance_locked"] = True
    issues.sort(key=lambda i: (i["audit_stage"], i.get("sheet", ""), _excel_row_num(i), i["rule_id"]))
    result["issues"] = issues
    result["sheet"] = eligible[0] if len(eligible) == 1 else "、".join(eligible)
    result["audited_sheets"] = eligible
    result["reviewed_rows"] = len(all_records)
    result["excluded_sheets"] = excluded
    result["sheet_results"] = [{"sheet": p.get("_ai_context", {}).get("records", [{}])[0].get("sheet", eligible[n])
                                if p.get("_ai_context", {}).get("records") else eligible[n],
                                "data_rows": p["data_rows"]} for n, p in enumerate(results)]
    result["data_rows"] = sum(p["data_rows"] for p in results)
    result["nonempty_rows_without_code"] = sum(p.get("nonempty_rows_without_code", 0) for p in results)
    result["code_length_distribution"] = dict(sorted(Counter(len(r["kks_code"]) for r in all_records if r["kks_code"]).items()))
    result["metrics"]["unique_codes"] = len(grouped)
    result["metrics"]["duplicate_code_groups"] = sum(len(g) > 1 for g in grouped.values())
    for metric, rules in (("orphan_rows", {"KKS-08"}), ("prefix_mismatch_rows", {"KKS-09"}),
                          ("self_ref_rows", {"KKS-10"}), ("parent_longer_rows", {"KKS-11"})):
        result["metrics"][metric] = len({(i.get("sheet"), i.get("excel_row")) for i in issues if i["rule_id"] in rules})
    result["scope_validation"]["notes"] = [note for p in results for note in p.get("structural_notes", [])
                                            if "父级系统/子系统节点在主库" not in note and "不视为孤儿" not in note]
    result["_ai_context"] = {"records": all_records}
    result["audit_policy"] = {"device_length": 12, "order": list(STAGES), "replacement_codes": False,
                              "standard": "GB/T 50549-2020", "scope": "全部已识别设备工作表"}
    result["audit_policy"]["finished_product_sheets"] = sorted({r["sheet"] for r in all_records if r.get("finished_product")})
    result["audit_policy"]["finished_product_fields"] = ["父级KKS编码", "新KKS编码", "新设备名称"]
    result["structural_notes"] = [note for p in results for note in p.get("structural_notes", [])
                                  if "父级系统/子系统节点在主库" not in note and "不视为孤儿" not in note]
    result["scope_validation"]["incremental_batch"] = False
    result["scope_validation"]["mode"] = "workbook"
    result["structural_notes"] += [f"完整审查工作表：{'、'.join(eligible)}", "编号分配、备用关系、安装位置及全厂唯一性须以工程约定、图纸和全厂台账确认。"]
    result["structural_notes"] += [f"未审查工作表 {item['sheet']}：{item['reason']}" for item in excluded]
    for sheet in eligible:
        for stage, message in ((4, "未提供工程编号分配、预留和备用设备约定，相关编号合理性尚不能确认。"),
                               (5, "安装位置、区域归属及全厂范围唯一性需关联图纸、位置台账和全厂设备台账确认。")):
            add_issue(issues, priority="P1", rule_id="KKS-45" if stage == 4 else "KKS-46", category=STAGES[stage],
                      rec={"sheet": sheet}, message=message, suggestion="补充关联资料完成核验。",
                      status="needs_review", evidence="GB/T 50549-2020 第3.3节；关联资料核验范围")
            issues[-1].update(audit_stage=stage, evidence_missing=True)
    issues.sort(key=lambda i: (i["audit_stage"], i.get("sheet", ""), _excel_row_num(i), i["rule_id"]))
    for issue in issues:
        issue.setdefault("final_decision", "needs_human" if issue.get("status") == "needs_review" else "confirmed_issue")
    result["issue_count"] = sum(i.get("status") != "resolved" for i in issues)
    result["priority_counts"] = dict(Counter(i["priority"] for i in issues if i.get("status") != "resolved"))
    result["conclusion"] = f"本次完整审查 {len(eligible)} 个工作表，发现 {result['issue_count']} 项问题或待核事项；源Excel未修改。"
    if run_ai_review:
        apply_ai_review_stage(result, progress_callback=progress_callback)
    result.pop("stage_scope_note", None)
    _set_compliance_result(result)
    result.setdefault("import_governance", {})["direct_import_ready"] = False
    result["compliance"]["excluded_sheets"] = excluded
    result["quality_metrics"] = quality_metrics(result)
    if write_outputs:
        write_audit_artifacts(input_path, output_dir, result, progress_callback=progress_callback)
    return result


def write_issue_workbook_xlsx(path: Path, result: dict[str, Any], overview: dict[str, Any] | None = None) -> None:
    """生成面向业务交付的问题清单（广元风格：概览 + 每类问题定制 sheet）。"""
    if overview is None:
        overview = build_overview(result)
    render_xlsx(path, result, overview)


def write_html(path: Path, result: dict[str, Any], overview: dict[str, Any] | None = None) -> None:
    """生成广元风格 HTML 审核报告（业务分组 + 样例截断，体积可控）。"""
    if overview is None:
        overview = build_overview(result)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(render_html(result, overview), encoding="utf-8")


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
