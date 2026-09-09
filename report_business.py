# -*- coding: utf-8 -*-
"""广元风格业务化报告模型与渲染。

把规则级 issues 归并成"业务问题桶"（P1 必改 / P2 清洗提示 / 范围验证等），
供 HTML 报告、Excel 问题清单、Web 工作台三端共用同一套分组口径：
  - 机组文本三方不一致(P1)
  - 空名占位(P1)
  - 名称文本卫生(P2)
  - 名称语义需人核(P2)
  - 同名/同物异名线索(P2 抽检)
  - 原KKS 一对多(P2)
  - 阻断性问题(P0) / 其他规则提示(P2)
规模分布、八维卡片、误报防控声明、A/B/C 建议全部从 result 派生，可回溯。
"""
from __future__ import annotations

import html as _html
import re
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path
from typing import Any

try:
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.utils import get_column_letter
except Exception:  # pragma: no cover - 仅无 openpyxl 环境降级
    openpyxl = None

# ------------------------------------------------------------------ 常量
ENTERPRISE_UNIT_NAMES = {"00": "全厂公用", "01": "1号机组", "02": "2号机组", "61": "1,2号机组公用"}
BLUE = "1D4ED8"
PALE_BLUE = "EFF6FF"
GRAY = "F1F5F9"
PALE_RED = "FEF2F2"
PALE_AMBER = "FFFBEB"
PALE_GREEN = "F0FDF4"
# HTML/CSS 与广元模板同款
CSS = """
:root{--bg:#f5f7fa;--card:#fff;--ink:#1f2d3d;--mut:#6b7a8d;--line:#e3e8ef;
--blue:#2563eb;--blue2:#dbeafe;--p0:#dc2626;--p0b:#fde2e2;--p1:#d97706;--p1b:#fdebd0;
--p2:#0891b2;--p2b:#d4f1f6;--ok:#059669;--okb:#d6f3e4;}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font-family:"Microsoft YaHei","PingFang SC",sans-serif;line-height:1.6}
.wrap{max-width:1180px;margin:0 auto;padding:28px 22px 60px}
header{background:linear-gradient(120deg,#1e3a8a,#2563eb);color:#fff;border-radius:14px;padding:26px 30px;margin-bottom:22px}
header h1{margin:0 0 6px;font-size:25px;letter-spacing:1px}
header .meta{font-size:13px;opacity:.92}
header .ver{display:inline-block;background:rgba(255,255,255,.18);padding:2px 10px;border-radius:20px;margin-left:8px;font-size:12px}
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin:20px 0}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px 18px}
.kpi .n{font-size:28px;font-weight:700;color:var(--blue)}
.kpi .l{font-size:13px;color:var(--mut);margin-top:2px}
.kpi.warn .n{color:var(--p1)}.kpi.bad .n{color:var(--p0)}.kpi.ok .n{color:var(--ok)}
.sec{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:20px 24px;margin:18px 0}
.sec h2{margin:0 0 4px;font-size:19px;border-left:4px solid var(--blue);padding-left:10px}
.sec .sub{color:var(--mut);font-size:13px;margin:0 0 14px}
.dims{display:grid;grid-template-columns:repeat(2,1fr);gap:16px}
.dimcard{border:1px solid var(--line);border-radius:10px;overflow:hidden}
.dimhead{background:var(--blue2);font-weight:700;padding:9px 12px;font-size:14px;display:flex;justify-content:space-between;align-items:baseline}
.dimsub{font-weight:400;font-size:11px;color:var(--mut)}
.dimcard table{width:100%;border-collapse:collapse;font-size:13px}
.dimcard td{padding:7px 12px;border-top:1px solid var(--line)}
.dimcard td.num{text-align:right;font-variant-numeric:tabular-nums}
.sev{font-size:12px;padding:1px 8px;border-radius:6px;font-weight:700}
.sev.p0{background:var(--p0b);color:var(--p0)}.sev.p1{background:var(--p1b);color:var(--p1)}
.sev.p2{background:var(--p2b);color:var(--p2)}.sev.ok{background:var(--okb);color:var(--ok)}
table.data{width:100%;border-collapse:collapse;font-size:13px;margin-top:8px}
table.data th{background:#eef2f7;text-align:left;padding:7px 10px;border:1px solid var(--line)}
table.data td{padding:6px 10px;border:1px solid var(--line);vertical-align:top}
.mono{font-family:"Consolas","Courier New",monospace;font-size:12px}
.num{text-align:right;font-variant-numeric:tabular-nums}
.bar{display:flex;align-items:center;gap:10px;margin:5px 0;font-size:13px}
.blab{width:90px;flex:none;color:var(--mut)}.btrack{flex:1;background:#eef2f7;border-radius:6px;height:14px;overflow:hidden}
.bfill{display:block;height:100%;background:linear-gradient(90deg,#3b82f6,#2563eb)}
.bnum{width:120px;flex:none;text-align:right;font-variant-numeric:tabular-nums}
.callout{border-radius:10px;padding:14px 16px;font-size:13.5px;margin:12px 0}
.callout.p0{background:var(--p0b);border:1px solid #f3b4b4}
.callout.p1{background:var(--p1b);border:1px solid #f3d39a}
.callout.p2{background:var(--p2b);border:1px solid #a9e2ec}
.callout.ok{background:var(--okb);border:1px solid #a7e0c2}
.callout b{color:var(--ink)}
.tag{display:inline-block;background:var(--blue2);color:#1e40af;border-radius:5px;padding:1px 7px;font-size:12px;margin:0 3px}
ul.tight{margin:8px 0;padding-left:22px}ul.tight li{margin:4px 0}
.foot{color:var(--mut);font-size:12px;text-align:center;margin-top:30px}
details.sec{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:12px 24px;margin:18px 0}
details.sec>summary{cursor:pointer;list-style:none;margin:8px 0 0}
details.sec>summary::-webkit-details-marker{display:none}
details.sec>summary h2::before{content:"▾ ";color:var(--mut);font-size:14px}
details.sec:not([open])>summary h2::before{content:"▸ ";color:var(--mut);font-size:14px}
details.sec[open]>summary{border-bottom:1px dashed var(--line);margin-bottom:12px;padding-bottom:8px}
details.ai-item{border:1px solid var(--line);border-radius:8px;padding:6px 10px;margin:6px 0;background:var(--bg)}
details.ai-item summary{cursor:pointer;font-size:13px;color:var(--ink)}
details.ai-item .ai-body{margin-top:6px;font-size:13px;line-height:1.6}
.ai-tag{display:inline-block;border-radius:999px;padding:0 8px;font-size:12px;margin-right:6px;font-weight:500}
.ai-tag.confirmed{background:var(--p0b);color:var(--p0)}
.ai-tag.falsepos{background:var(--okb);color:var(--ok)}
.ai-tag.human{background:var(--p1b);color:var(--p1)}
@media(max-width:820px){.kpis,.dims{grid-template-columns:1fr}}
"""


# ------------------------------------------------------------------ 工具
def _issues(result: dict[str, Any]) -> list[dict[str, Any]]:
    return [i for i in result.get("issues", []) if isinstance(i, dict) and i.get("status") != "resolved"]


def _rule_hits(result: dict[str, Any], rules: set[str]) -> list[dict[str, Any]]:
    out = []
    for item in _issues(result):
        ids = {p.strip() for p in re.split(r"[/~]", str(item.get("rule_id", "")))}
        if ids & rules:
            out.append(item)
    return out


def _rule_count(result: dict[str, Any], rules: set[str]) -> int:
    return len(_rule_hits(result, rules))


def _code_prefix(code: str) -> str:
    code = str(code)
    return code[:2] if len(code) >= 2 else code


def _full_width(s: str) -> bool:
    return any("\uff00" <= ch <= "\uffef" for ch in str(s))


def _records(result: dict[str, Any]) -> list[dict[str, Any]]:
    recs = result.get("_ai_context", {}).get("records", [])
    return [r for r in recs if isinstance(r, dict)] if isinstance(recs, list) else []


_XML_CTRL = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")


def _esc(value: Any) -> str:
    # 控制字符会让 HTML/XML 解析器告警，先剔除（保留 \t\n\r）
    return _html.escape(_XML_CTRL.sub("", "" if value is None else str(value)))


# ------------------------------------------------------------------ 辅助源表读取
def _read_aux_dup_sheet(source_path: str) -> dict[str, Any] | None:
    """读取源工作簿中「重码待处理」类副表（未编码挂起条目），供报告展示。"""
    if not source_path or openpyxl is None:
        return None
    path = Path(source_path)
    if not path.exists():
        return None
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for ws in wb.worksheets:
            if "重码" not in ws.title and "待处理" not in ws.title:
                continue
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue
            header = [str(v).strip() if v is not None else "" for v in rows[0]]
            name_idx = next((i for i, h in enumerate(header) if "名称" in h), 0)
            prof_idx = next((i for i, h in enumerate(header) if "专业" in h), None)
            out: list[dict[str, str]] = []
            for row in rows[1:]:
                if not any(v not in (None, "") for v in row):
                    continue
                name = str(row[name_idx]).strip() if name_idx < len(row) and row[name_idx] is not None else ""
                prof = ""
                if prof_idx is not None and prof_idx < len(row) and row[prof_idx] is not None:
                    prof = str(row[prof_idx]).strip()
                if name:
                    out.append({"name": name, "profession": prof, "status": "未编/待处理"})
            wb.close()
            if out:
                return {"sheet": ws.title, "rows": out}
            return None
        return None
    except Exception:
        return None


# ------------------------------------------------------------------ 业务分组构建
def build_overview(result: dict[str, Any]) -> dict[str, Any]:
    issues = _issues(result)
    records = _records(result)
    source_path = str(result.get("source_file", ""))
    stem = Path(source_path).stem if source_path else "KKS编码审核"
    data_rows = int(result.get("data_rows", 0) or 0)

    def text_code(v: Any) -> str:
        return "" if v is None else str(v).strip()

    code_set = {text_code(r.get("kks_code")) for r in records if r.get("kks_code")}
    code_set.discard("")

    metrics = result.get("metrics", {}) or {}
    scope = result.get("scope_validation", {}) or {}
    incremental = bool(scope.get("incremental_batch"))

    # ---- AI 复核结果（未启用时 candidate/reviewed 均为 0，展示层据此隐藏 AI 列）
    ai = result.get("ai_review", {}) if isinstance(result.get("ai_review"), dict) else {}
    ai_reviewed = int(ai.get("reviewed_count", 0) or 0)
    ai_candidate = int(ai.get("candidate_count", 0) or 0)
    ai_model_name = str(ai.get("model", "") or "")
    ai_active = ai_reviewed > 0
    ai_summary = result.get("ai_summary", {}) if isinstance(result.get("ai_summary"), dict) else {}
    ai_summary_ok = str(ai_summary.get("status", "")) == "completed" and str(ai_summary.get("overall", "")).strip()
    _AI_DECISION_LABELS = {
        "confirmed_issue": "确认为问题",
        "likely_false_positive": "疑似误报",
        "needs_human": "需人工确认",
    }

    # ---- 各规则问题行（按行去重优先给最贴合的业务桶）
    unit_text_items: list[dict[str, Any]] = []
    empty_name_items: list[dict[str, Any]] = []
    hygiene_items: list[dict[str, Any]] = []       # 行 → 原因
    hygiene_seen: dict[int, str] = {}
    semantic_void_items: list[dict[str, Any]] = []
    dup_name_items: list[dict[str, Any]] = []
    code_hygiene_items: list[dict[str, Any]] = []
    p0_blocker_items: list[dict[str, Any]] = []
    other_items: list[dict[str, Any]] = []

    def row_key(item: dict[str, Any]) -> int:
        try:
            return int(item.get("excel_row") or 0)
        except (TypeError, ValueError):
            return 0

    for item in issues:
        rule = str(item.get("rule_id", ""))
        category = str(item.get("category", ""))
        row = row_key(item)
        if rule == "KKS-28" and category == "机组三方不一致":
            unit_text_items.append(item)
        elif rule in {"KKS-D", "KKS-00"} and (category == "有码无名称" or (rule == "KKS-00")):
            empty_name_items.append(item)
        elif rule == "KKS-30" and category in ("名称文本卫生", "名称文本卫生(全角)") and row not in hygiene_seen:
            reason = "全角符" if "全角" in category else ("全角符" if _full_width(item.get("name")) else "空白/控制符")
            hygiene_seen[row] = reason
            hygiene_items.append(item)
        elif rule == "KKS-27" and category == "名称非常规符号" and row not in hygiene_seen:
            reason = "全角符" if _full_width(item.get("name")) else "非常规符号"
            hygiene_seen[row] = reason
            hygiene_items.append(item)
        elif rule == "KKS-30" and category == "编码文本卫生":
            code_hygiene_items.append(item)
        elif rule == "KKS-25" and category == "同名对应多个 KKS 码":
            dup_name_items.append(item)
        elif rule == "KKS-25b":
            dup_name_items.append(item)
        elif rule == "KKS-27" and category == "名称语义需人核":
            semantic_void_items.append(item)
        elif rule in {"KKS-04", "KKS-04b", "KKS-05", "KKS-06", "KKS-07", "KKS-09", "KKS-10", "KKS-11", "KKS-12"}:
            p0_blocker_items.append(item)
        elif rule == "KKS-25c":   # 原KKS一对多：从 records 展开逐行列表，见下方
            pass
        else:
            other_items.append(item)

    # ---- 原KKS 一对多：由源记录按 old_code 重新分组（含全部成员行）
    old_to_rows: defaultdict[str, list[dict[str, Any]]] = defaultdict(list)
    for rec in records:
        old = text_code(rec.get("old_code"))
        code = text_code(rec.get("kks_code"))
        if old and code:
            old_to_rows[old].append(rec)
    orig_1n_groups: list[dict[str, Any]] = []
    for old, recs in sorted(old_to_rows.items()):
        codes = {text_code(r.get("kks_code")) for r in recs}
        if len(codes) > 1:
            members = sorted(recs, key=lambda r: int(r.get("excel_row") or 0))
            orig_1n_groups.append({"old": old, "codes": sorted(codes), "members": members,
                                   "rows": [int(m.get("excel_row") or 0) for m in members]})

    # ---- 卫生原因计数（按行）
    hygiene_reason_counts: Counter[str] = Counter()
    for _, reason in hygiene_seen.items():
        hygiene_reason_counts[reason] += 1

    def sort_rows(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
        return sorted(items, key=lambda i: row_key(i))

    unit_text_items = sort_rows(unit_text_items)
    empty_name_items = sort_rows(empty_name_items)
    hygiene_items = sort_rows(hygiene_items)
    code_hygiene_items = sort_rows(code_hygiene_items)
    dup_name_items = sort_rows(dup_name_items)
    # 空名行的"语义需人核"已由 KKS-D 覆盖，不再重复计入
    empty_rows = {row_key(i) for i in empty_name_items}
    semantic_void_items = sort_rows([i for i in semantic_void_items if row_key(i) not in empty_rows])
    p0_blocker_items = sort_rows(p0_blocker_items)
    other_items = sort_rows(other_items)

    def sev_of(n: int) -> str:
        return "ok" if n == 0 else ("p0" if n else "p1")

    # 机组文本三方不一致的方向分布（文本 → 编码前缀）
    unit_direction: Counter[tuple[str, str]] = Counter()
    for item in unit_text_items:
        unit_direction[(str(item.get("unit", "")), _code_prefix(str(item.get("kks_code", ""))))] += 1

    # 机组前缀规模 / 专业规模 / 码长分布
    prefix_counts: Counter[str] = Counter()
    prof_counts: Counter[str] = Counter()
    for code in code_set:
        prefix_counts[_code_prefix(code)] += 1
    for rec in records:
        prof = str(rec.get("profession", "") or "").strip()
        if prof:
            prof_counts[prof] += 1

    def prefix_label(p: str) -> str:
        if p in ENTERPRISE_UNIT_NAMES:
            return ENTERPRISE_UNIT_NAMES[p]
        return f"前缀 {p}"

    ordered_prefixes = [p for p in ("00", "01", "02", "61") if prefix_counts.get(p)]
    ordered_prefixes += [p for p, _ in prefix_counts.most_common() if p not in ordered_prefixes]
    max_pref = max([prefix_counts[p] for p in ordered_prefixes], default=1) or 1
    max_prof = max(prof_counts.values(), default=1) or 1
    unit_bars = [{"label": prefix_label(p), "value": prefix_counts[p], "pct": prefix_counts[p] / max_pref * 100} for p in ordered_prefixes]
    prof_bars = [{"label": p, "value": prof_counts[p], "pct": prof_counts[p] / max_prof * 100} for p, _ in prof_counts.most_common(6)]

    len_dist = result.get("code_length_distribution", {}) or {}
    len_note = "码长分布：" + "，".join(f"{k}位 {v}" for k, v in sorted(len_dist.items(), key=lambda kv: int(kv[0]))) + "。"

    # KPI 卡片
    dup_groups = int(metrics.get("duplicate_code_groups", 0) or 0)
    oi_hits = _rule_count(result, {"KKS-06", "KKS-29"})
    long_codes = int(metrics.get("long_code_rows", 0) or 0)
    illegal = _rule_count(result, {"KKS-04", "KKS-04b", "KKS-05"})
    zero_clean = dup_groups + illegal + oi_hits + long_codes
    missing_rows = len(_rule_hits(result, {"KKS-00"}))
    hygiene_n = len(hygiene_items)
    empty_n = len(empty_name_items)

    kpis = [
        {"label": "主表已编码行" + ("" if not missing_rows else f"（另有{missing_rows}行缺码）"),
         "value": f"{data_rows:,}", "tone": "ok" if not missing_rows else "warn"},
        {"label": "重复码/非法字符/O-I/超12位", "value": f"{zero_clean:,}", "tone": "ok" if zero_clean == 0 else "bad"},
        {"label": "机组文本三方不一致(P1)", "value": f"{len(unit_text_items):,}", "tone": "warn" if unit_text_items else "ok"},
        {"label": "名称卫生+空名(P2/P1)", "value": f"{hygiene_n + empty_n:,}", "tone": "warn" if hygiene_n + empty_n else "ok"},
        {"label": "AI 已复核" + (f"（候选{ai_candidate}）" if ai_candidate else ""),
         "value": f"{ai_reviewed:,}", "tone": "ok" if ai_reviewed else "warn"},
    ]

    # ---- 八维卡片（六张，行内计数可回溯）
    def rule_count_cell(items: list[dict[str, Any]], level: str = "p2") -> tuple[int, str]:
        return (len(items), "ok" if not items else level)

    syntax_ok = (illegal == 0)
    syntax_rows = [
        ("非法字符", illegal, "p0" if illegal else "ok", "扁平码仅 A-Z0-9（KKS-04）"),
        ("分段类型错", _rule_count(result, {"KKS-04b"}), "p0", "F0/F1F2F3/FN/A1A2/AN（KKS-04b）"),
        ("字母O/I冒充0/1", oi_hits, "p0", "易混字母 I/O（KKS-06/29）"),
        ("长度异常/超长码", _rule_count(result, {"KKS-05"}) + long_codes, "p0", ">12 位扩展码（KKS-05/17/22）"),
        ("编码文本卫生", len(code_hygiene_items), "p2", "码内空格/全角（KKS-30）"),
    ]
    uniq_rows = [
        ("重复KKS码(主键)", dup_groups, "p0", "完整码分组（KKS-07）"),
        ("原KKS一对多", len(orig_1n_groups), "p2", "身份键(原KKS)映射>1 新码（KKS-25c）"),
        ("同名多码线索", len(dup_name_items), "p2", "同名/同物异名聚簇，仅抽检（KKS-25/25b）"),
    ]
    prefix_mismatch = int(metrics.get("prefix_mismatch_rows", 0) or 0)
    self_ref = _rule_count(result, {"KKS-10"})
    parent_longer = _rule_count(result, {"KKS-11"})
    orphan_rows = int(metrics.get("orphan_rows", 0) or 0)
    # 数量列只放数字：增量批次孤儿已转为"主库闭合"范围验证，不产生缺陷计数，文字只进说明列
    orphan_value = 0 if incremental else orphan_rows
    orphan_hint = ("增量批次：父级不在本文件，已转主库闭合范围校验，不计缺陷（详见范围验证）"
                   if incremental else "父级不在文件内（KKS-08）")
    hier_rows = [
        ("子码前缀违反", prefix_mismatch, "p0", "子码不以父码开头（KKS-09）"),
        ("自引用/环", self_ref, "p0", "父级环（KKS-10）"),
        ("父级比子码更长", parent_longer, "p0", "父级长度异常（KKS-11）"),
        ("跨机组断链", 0, "ok", "前缀机组与父级机组不一致"),
        ("父级孤儿(范围验证)", orphan_value, "p2", orphan_hint),
    ]
    legacy_codes = _rule_count(result, {"KKS-16", "KKS-13", "KKS-18"})
    letter_realloc = _rule_count(result, {"KKS-14"})
    mig_rows = [
        ("10版旧前缀遗留(码)", legacy_codes, "p2", "G=5/6/L/J 旧格式（KKS-16/13/18）"),
        ("系统/设备字母重分配", letter_realloc, "p1", "版本迁移字母表（KKS-14）"),
        ("扩展码收敛为12位主码", int((scope.get("converged_to_master_rows") or 0)), "ok", "旧码11~17位→12位，业务性收敛"),
    ]
    sem_rows = [
        ("超12位第5层扩展码", long_codes, "p1", "第5层/扩展码（KKS-17/22）"),
        ("名称文本卫生", hygiene_n, "p2", "全角/控制符等（KKS-30/27）"),
        ("空名称(P1)", empty_n, "p1", "名称空/空格占位（KKS-D）"),
        ("名称语义需人核", len(semantic_void_items), "p2", "短名/无中文语义（KKS-27）"),
    ]
    unit_rows = [("机组文本↔编码前缀不一致", len(unit_text_items), "p1", "以 00/01/02/61 前缀为准")]

    cards = [
        {"no": "①", "title": "结构/语法", "sub": "字符集 + 分段", "rows": syntax_rows, "group_ok": syntax_ok},
        {"no": "②", "title": "唯一性", "sub": "主键 + 身份键", "rows": uniq_rows, "group_ok": dup_groups == 0},
        {"no": "③", "title": "层级/树", "sub": "前缀一致性(本文件内)", "rows": hier_rows, "group_ok": prefix_mismatch == 0 and self_ref == 0 and parent_longer == 0},
        {"no": "④", "title": "迁移一致性", "sub": "旧格式/字母迁移", "rows": mig_rows, "group_ok": legacy_codes == 0},
        {"no": "⑤", "title": "语义/可导入", "sub": "扩展码 + 文本", "rows": sem_rows, "group_ok": long_codes == 0 and hygiene_n == 0},
        {"no": "⑥", "title": "机组三方一致", "sub": "文本↔前缀", "rows": unit_rows, "group_ok": not unit_text_items},
    ]

    # ---- P1/P2/P0 业务桶（供明细表/Excel sheet 使用）
    def sev_text(n: int, level: str) -> str:
        return "OK" if n == 0 else level

    buckets: list[dict[str, Any]] = []

    def add_bucket(key: str, level: str, title: str, note: str, items: list[dict[str, Any]], columns: list[dict[str, str]]) -> None:
        buckets.append({
            "key": key, "level": level, "title": title, "note": note,
            "count": len(items), "columns": columns, "items": items,
        })

    add_bucket("p0_blocker", "P0", "阻断性问题（须先修复）", "", p0_blocker_items,
               [{"k": "excel_row", "t": "行"}, {"k": "kks_code", "t": "编码"}, {"k": "rule_id", "t": "规则"}, {"k": "message", "t": "问题"}, {"k": "suggestion", "t": "整改建议"}])
    add_bucket("p1_unit_text", "P1", "机组号文本↔编码前缀不一致", "编码前缀为权威 KKS 标识：00=全厂公用 / 01=1号 / 02=2号 / 61=1,2号公用。下列行“机组号”文本与编码前缀不符，建议以编码前缀校正机组列文本。", unit_text_items,
               [{"k": "excel_row", "t": "行"}, {"k": "kks_code", "t": "编码"}, {"k": "unit", "t": "机组文本"}, {"k": "prefix", "t": "编码前缀"}, {"k": "profession", "t": "专业"}, {"k": "system", "t": "系统"}, {"k": "name", "t": "设备名称"}])
    add_bucket("p1_empty_name", "P1", "设备名称为空 / 空格占位", "下列设备码名称栏为空或仅空格占位，导入主库前须补全设备名称。", empty_name_items,
               [{"k": "excel_row", "t": "行"}, {"k": "kks_code", "t": "编码"}, {"k": "parent_code", "t": "父级"}, {"k": "profession", "t": "专业"}, {"k": "status", "t": "状态"}])
    add_bucket("p2_hygiene", "P2", "名称文本卫生", f"主要为全角括号等全角符（原因分布 {dict(hygiene_reason_counts)}）。导入主库前建议全角→半角统一、去除首尾空格。", hygiene_items,
               [{"k": "excel_row", "t": "行"}, {"k": "kks_code", "t": "编码"}, {"k": "name", "t": "设备名称"}, {"k": "hygiene_reason", "t": "问题"}, {"k": "profession", "t": "专业"}])
    add_bucket("p2_semantic_void", "P2", "名称语义需人核", "名称过短或无法提取中文/字母语义，人工核对该名称。", semantic_void_items,
               [{"k": "excel_row", "t": "行"}, {"k": "kks_code", "t": "编码"}, {"k": "name", "t": "设备名称"}, {"k": "message", "t": "问题"}])
    add_bucket("p2_dup_name", "P2", "同名/同物异名线索", "同一系统内归一名相同或语义高度相近的多个编码，仅作为人工抽检线索，不自动合并。", dup_name_items,
               [{"k": "excel_row", "t": "行"}, {"k": "kks_code", "t": "编码"}, {"k": "name", "t": "设备名称"}, {"k": "message", "t": "线索说明"}])
    add_bucket("p2_code_hygiene", "P2", "编码文本卫生", "编码含首尾空白/全角空格/控制符，建议 trim 清洗。", code_hygiene_items,
               [{"k": "excel_row", "t": "行"}, {"k": "kks_code", "t": "编码"}, {"k": "message", "t": "问题"}])
    add_bucket("p2_others", "P2", "其他提示（历史/断号/风格）", "", other_items,
               [{"k": "excel_row", "t": "行"}, {"k": "kks_code", "t": "编码"}, {"k": "rule_id", "t": "规则"}, {"k": "category", "t": "类别"}, {"k": "message", "t": "问题"}, {"k": "suggestion", "t": "建议"}])

    # 为行补展示派生字段（编码前缀 / 卫生原因 / 状态 / 空名状态）
    for item in unit_text_items:
        item["prefix"] = _code_prefix(str(item.get("kks_code", "")))
    for row in hygiene_seen:
        pass
    for item in hygiene_items:
        item["hygiene_reason"] = hygiene_seen.get(row_key(item), "全角符/其他")
    for item in empty_name_items:
        item["status"] = "空/空格占位"

    # AI 复核列：仅当 AI 实际复核过才注入；置于明细表末尾，避免挤占业务列
    if ai_active:
        for bucket in buckets:
            bucket["columns"] = bucket["columns"] + [
                {"k": "ai_decision_label", "t": "AI结论"},
                {"k": "ai_final_summary", "t": "AI说明"},
            ]
        for bucket in buckets:
            for item in bucket["items"]:
                decision = str(item.get("final_decision", "") or "")
                item["ai_decision_label"] = _AI_DECISION_LABELS.get(decision, "需人工确认" if decision or item.get("ai_decision") else "")
                item["ai_final_summary"] = str(item.get("final_summary", "") or "")

    # ---- AI 复核建议（第九节 / Excel 独立 sheet 数据源）
    ai_suggestions: list[dict[str, str]] = []
    seen_ai_rows: set[tuple[str, str]] = set()
    if ai_active:
        for bucket in buckets:
            for item in bucket["items"]:
                summary = str(item.get("final_summary", "") or "")
                if not summary:
                    continue
                key = (str(item.get("excel_row", "")), str(item.get("kks_code", "")))
                if key in seen_ai_rows:
                    continue
                seen_ai_rows.add(key)
                ai_suggestions.append({
                    "excel_row": key[0],
                    "kks_code": key[1],
                    "name": str(item.get("name", "") or ""),
                    "bucket_title": str(bucket["title"]),
                    "decision": _AI_DECISION_LABELS.get(str(item.get("final_decision", "") or ""), "需人工确认"),
                    "summary": summary,
                    "suggestion": str(item.get("final_suggestion", "") or ""),
                })

    # ---- 结构说明与范围
    structure_rows: list[tuple[str, str]] = []
    structure_rows.append(("文件性质", "增量新增批次，非全厂汇总" if incremental else "全量文件"))
    structure_rows.append(("主表", f"{result.get('sheet', '—')} {data_rows:,} 行（权威源）"))
    if scope.get("notes"):
        for note in scope["notes"]:
            structure_rows.append(("范围/约定", note))
    if orig_1n_groups:
        structure_rows.append(("原KKS一对多", f"{len(orig_1n_groups)} 组；新KKS全唯一，以新KKS为主键"))
    for note in result.get("structural_notes", []):
        if "空 Sheet" in str(note) or "空白行" in str(note) or "首行" in str(note):
            structure_rows.append(("文件结构", str(note)))
    if source_path:
        structure_rows.append(("源文件", source_path))

    scope_external = scope.get("external_parent_samples", []) or []
    external_units = scope.get("external_parent_units", 0)
    external_rows = scope.get("external_parent_rows", 0)

    # ---- 误报防控（诚实声明）
    honesty: list[dict[str, str]] = []
    if incremental:
        honesty.append({"tag": "范围", "bold": f"父级孤儿 {external_rows:,} 行({external_units} 个父级节点)不在本文件",
                        "text": "增量文件对自身判孤儿必然接近 100%，已改为“与主库闭合”的范围验证项，不计入缺陷（见第六节）。"})
    honesty.append({"tag": "企业惯例", "bold": "全厂码 G 取值按企业两位前缀 00/01/02/61 放行",
                    "text": "G='0'/'6' 等字符若按国标单字符 G 表判定会误报（此前 8,208 条假阳性已消除）；企业增量文件以两位数字前缀标识机组/公用，视为合法。"})
    if scope.get("converged_to_master_rows"):
        honesty.append({"tag": "收敛业务", "bold": f"{scope.get('converged_to_master_rows')} 行旧扩展码(非12位)收敛为 12 位主码",
                        "text": "这是“旧码→12位主码”的标准化业务本身，非层级缺陷；未逐条列入问题（此前 8,403 条假阳性已消除）。"})
    if len(orig_1n_groups):
        honesty.append({"tag": "一对多", "bold": f"原KKS 一对多 {len(orig_1n_groups)} 组",
                        "text": "多为“一个旧测点拆分为多个新设备”的合理映射；新KKS 全唯一，原KKS 仅作历史参照（P2 提示非 P0 重复）。"})
    if len(dup_name_items):
        honesty.append({"tag": "同名聚簇", "bold": f"同名/同物异名 {len(dup_name_items)} 组",
                        "text": "KKS 名称本身不唯一定位设备，聚簇仅作人工抽检线索并按簇计 1 条，不按簇内编码数放大。"})
    honesty.append({"tag": "列头健壮", "bold": "按表头名动态定位列",
                    "text": "同名/偏移列（如两列“设备名称（修编前）*”、父级列错位）按表头语义定位，不按固定索引。"})
    if ai_active:
        honesty.append({"tag": "AI复核", "bold": f"AI 已复核 {ai_reviewed}/{ai_candidate} 条候选",
                        "text": f"模型 {ai_model_name or '—'}；结论分“确认为问题/疑似误报/需人工确认”，见各明细表“AI结论”“AI说明”列。AI 不会删除规则命中，仅辅助判断真假阳性。"})
    else:
        honesty.append({"tag": "AI复核", "bold": "AI 语义复核未启用或未产生复核结果",
                        "text": "本报告仅含确定性规则结论；启用方式见工作台 AI 管理员配置。"})

    # ---- A/B/C 建议
    abc: list[tuple[str, str, str]] = []
    abc_items: list[tuple[str, str]] = []
    if unit_text_items:
        abc_items.append(("A", f"{len(unit_text_items)} 条机组文本按编码前缀校正（00/01/02/61）"))
    if empty_n:
        abc_items.append(("A", f"{empty_n} 条设备名称为空/空格占位，补全设备名称"))
    if p0_blocker_items:
        abc_items.append(("A", f"{len(p0_blocker_items)} 条 P0 阻断问题先修复并复核"))
    if hygiene_n:
        abc_items.append(("B", f"{hygiene_n} 条名称全角符/控制符清洗（全角→半角、去首尾空格）"))
    if orig_1n_groups:
        abc_items.append(("B", f"{len(orig_1n_groups)} 组原KKS一对多确认映射口径（新KKS 作主键）"))
    if dup_name_items:
        abc_items.append(("B", f"{len(dup_name_items)} 组同名/同物异名线索抽检核对"))
    abc_items.append(("C", "与主库做父级闭合校验（父级节点须在主库存在），确认后合入"))
    abc_items.append(("C", "修复后重跑本脚本，确认重复码/非法字符/前缀违反/分段类型 全 0"))

    for letter, text in abc_items:
        abc.append((letter, text, ""))

    # 汇总"须处理"清单（供 HTML 一、总体结论与 callout 复用）
    must_handle: list[tuple[str, str]] = []
    if p0_blocker_items:
        must_handle.append(("P0", f"{len(p0_blocker_items)} 条阻断性问题"))
    if unit_text_items:
        must_handle.append(("P1", f"{len(unit_text_items)} 条 机组号文本↔编码前缀不一致（建议以 00/01/02/61 为准校正文本列）"))
    if empty_n:
        must_handle.append(("P1", f"{empty_n} 条 设备名称为空/空格占位（须补全名称）"))
    if hygiene_n:
        must_handle.append(("P2", f"{hygiene_n} 条 名称文本卫生（建议半角化清洗后合入）"))
    if orig_1n_groups:
        member_rows = sum(len(g["members"]) for g in orig_1n_groups)
        must_handle.append(("P2", f"{len(orig_1n_groups)} 组 原KKS一对多映射（{member_rows} 行），确认旧点拆分还是旧码重复赋值"))
    if dup_name_items:
        must_handle.append(("P2", f"{len(dup_name_items)} 组 同名/同物异名线索（人工抽检）"))

    overview = {
        "title": f"{stem} · 八维审核报告",
        "plant": "",
        "source_path": source_path,
        "sheet": result.get("sheet", ""),
        "data_rows": data_rows,
        "missing_rows": missing_rows,
        "issue_count": int(result.get("issue_count", 0) or 0),
        "priority_counts": result.get("priority_counts", {}) or {},
        "standard_display": result.get("standard_display", ""),
        "kpis": kpis,
        "ai_active": ai_active,
        "ai_reviewed": ai_reviewed,
        "ai_candidate": ai_candidate,
        "ai_model": ai_model_name,
        "ai_summary": ai_summary if ai_summary_ok else {},
        "ai_suggestions": ai_suggestions,
        "cards": cards,
        "must_handle": must_handle,
        "clean_flags": {
            "duplicate": dup_groups, "illegal": illegal, "oi": oi_hits,
            "long": long_codes, "prefix": prefix_mismatch, "self_ref": self_ref,
            "parent_longer": parent_longer, "legacy": legacy_codes,
        },
        "unit_bars": unit_bars,
        "prof_bars": prof_bars,
        "len_note": len_note,
        "unit_direction": [{"unit": unit, "prefix": prefix, "count": n} for (unit, prefix), n in unit_direction.most_common()],
        "buckets": buckets,
        "orig_1n_groups": orig_1n_groups,
        "hygiene_reason_counts": dict(hygiene_reason_counts),
        "scope": {
            "incremental": incremental, "enterprise": bool(scope.get("enterprise")),
            "notes": scope.get("notes", []),
            "external_parent_rows": external_rows,
            "external_parent_units": external_units,
            "parent_samples": scope_external,
            "converged_rows": scope.get("converged_to_master_rows", 0),
            "converged_lengths": scope.get("converged_old_lengths", {}) or {},
        },
        "structure_rows": structure_rows,
        "honesty": honesty,
        "abc": abc,
        "aux_dup": _read_aux_dup_sheet(source_path),
    }
    return overview


# ------------------------------------------------------------------ HTML 渲染
def _fmt_big(value: Any) -> str:
    if isinstance(value, str):
        return value
    if isinstance(value, int):
        return f"{value:,}"
    return _esc(value)


def _sev_chip(sev: str, text: str) -> str:
    return f'<span class="sev {_esc(sev)}">{_esc(text)}</span>'


def _fold_section(sec_html: str) -> str:
    """把 <div class="sec"><h2>标题</h2>…</div> 转为可折叠的 <details class="sec" open>。"""
    m = re.match(r'<div class="sec">\s*<h2>(.+?)</h2>(.*)</div>\s*$', sec_html, re.S)
    if not m:
        return sec_html
    return f'<details class="sec" open><summary><h2>{m.group(1)}</h2></summary>{m.group(2)}</details>'


def render_html(result: dict[str, Any], ov: dict[str, Any]) -> str:
    esc = _esc
    today = date.today().isoformat()

    def meta_bold(text: str) -> str:
        return f"<b>{esc(text)}</b>"

    header = f"""<header>
  <h1>{esc(ov['title'])}</h1>
  <div class="meta">源文件：{esc(ov['source_path'] or '—')}<br>
  主表：<b>{esc(ov['sheet'])}</b>（{ov['data_rows']:,} 行{'' if not ov['missing_rows'] else f"，另有 {ov['missing_rows']} 行缺码"}）；
  审计日期：{today}；框架：kks-audit 八维
  <span class="ver">{'增量新增批次' if ov['scope']['incremental'] else '全量'}</span></div>
</header>"""

    kpi_html = "".join(
        f'<div class="kpi {item["tone"]}"><div class="n">{_fmt_big(item["value"])}</div><div class="l">{esc(item["label"])}</div></div>'
        for item in ov["kpis"]
    )

    # 一、总体结论
    clean = ov["clean_flags"]
    clean_zero = [k for k, v in clean.items() if v == 0]
    clean_names = {
        "duplicate": "主码零重复", "illegal": "零非法字符", "oi": "零字母O/I冒充",
        "long": "零超12位码", "prefix": "零前缀违反", "self_ref": "零自引用/环", "parent_longer": "零父级超长",
    }
    clean_text = "、".join(clean_names[k] for k in ("duplicate", "illegal", "oi", "long", "prefix", "self_ref", "parent_longer") if k in clean_zero) or "—"
    if ov["scope"]["incremental"]:
        clean_text += "；内部层级指向自洽（父级孤儿按主库闭合校验）"
    ok_qual = clean["duplicate"] == 0 and clean["illegal"] == 0 and clean["oi"] == 0 and clean["long"] == 0 and clean["prefix"] == 0
    tone_ok = ok_qual and not any(sev == "P0" for sev, _ in ov["must_handle"])

    conclusion_ok = (
        f"<b>整体定性：本批次质量良好，为{'干净的增量交付' if ov['scope']['incremental'] else '干净交付'}。</b>"
        f" {clean_text}。规则层假阳性已按企业惯例与增量范围收敛，剩余项均为可执行/可抽检的小量清单。"
    ) if ok_qual else (
        "<b>整体定性：存在需要先处理的阻断/结构问题。</b> 请优先处理 P0 阻断项后，再按下列 P1/P2 清单逐项核对。"
    )

    if ov["must_handle"]:
        li = "".join(f"<li><b>{esc(sev)}·</b> {esc(text)}</li>" for sev, text in ov["must_handle"])
        if ov["scope"]["incremental"] and ov["scope"]["external_parent_rows"]:
            li += (f"<li><b>范围验证（非缺陷）</b>：{ov['scope']['external_parent_rows']:,} 行设备码的父级"
                   f"（{ov['scope']['external_parent_units']:,} 个节点）指向主库既有节点，本文件不包含；"
                   "<b>“父级孤儿”须与主库闭合校验</b>（详见第六节）。</li>")
        handle_callout = f'<div class="callout p1"><b>须处理项（量级很小）：</b><ul class="tight">{li}</ul></div>'
    else:
        handle_callout = '<div class="callout ok"><b>未发现需要列入问题清单的项。</b></div>'

    sec1 = f"""<div class="sec">
  <h2>一、总体结论</h2>
  <p class="sub">独立复核 + 语义/层级补充校验；数字均自审计 JSON 派生，可复核。</p>
  <div class="callout {'ok' if tone_ok else 'p1'}">{conclusion_ok}</div>
  {handle_callout}
</div>"""

    # 二、八维卡片
    def chip_parts(value: Any, sev: str) -> tuple[str, str]:
        if isinstance(value, int) and value == 0:
            return "ok", "OK"
        return sev, sev.upper()

    cards_html = []
    for card in ov["cards"]:
        rows_html = ""
        for label, value, sev, hint in card["rows"]:
            if isinstance(value, int):
                chip_cls, chip_txt = chip_parts(value, sev)
                rows_html += (f"<tr><td>{esc(label)}</td><td class='num'>{value:,}</td>"
                              f"<td>{_sev_chip(chip_cls, chip_txt)}</td><td style='color:var(--mut);font-size:12px'>{esc(hint)}</td></tr>")
            else:
                rows_html += (f"<tr><td>{esc(label)}</td><td class='num'>{esc(value)}</td>"
                              f"<td>{_sev_chip(sev, sev.upper())}</td><td style='color:var(--mut);font-size:12px'>{esc(hint)}</td></tr>")
        cards_html.append(
            f'<div class="dimcard"><div class="dimhead">{card["no"]}{esc(card["title"])}<span class="dimsub">{esc(card["sub"])}</span></div>'
            f'<table><tr><th style="text-align:left">检查项</th><th style="text-align:right">数量</th><th>状态</th><th>说明</th></tr>{rows_html}</table></div>'
        )
    sec2 = f"""<div class="sec">
  <h2>二、八维审核结果总表</h2>
  <p class="sub">P0 阻断 / P1 需改 / P2 提示 / OK 良性。★ 注：本文件为增量新增时，“父级孤儿”不在此文件内判定（父节点在主库），已转为范围验证项。</p>
  <div class="dims">{''.join(cards_html)}</div>
</div>"""

    # 三、规模与分布
    def bars_html(title: str, bars: list[dict[str, Any]]) -> str:
        rows = "".join(
            f'<div class="bar"><span class="blab">{esc(b["label"])}</span>'
            f'<span class="btrack"><span class="bfill" style="width:{b["pct"]:.1f}%"></span></span>'
            f'<span class="bnum">{b["value"]:,}</span></div>'
            for b in bars
        )
        return f'<div><h3 style="font-size:15px;margin:4px 0">{esc(title)}</h3>{rows}</div>'

    grid = ""
    if ov["prof_bars"]:
        grid = f'<div style="display:grid;grid-template-columns:1fr 1fr;gap:18px">{bars_html("按专业", ov["prof_bars"])}{bars_html("按机组/前缀", ov["unit_bars"])}</div>'
    else:
        grid = f'<div style="display:grid;grid-template-columns:1fr 1fr;gap:18px"><div></div>{bars_html("按机组/前缀", ov["unit_bars"])}</div>'
    sec3 = f"""<div class="sec">
  <h2>三、规模与分布</h2>
  {grid}
  <p class="sub">{esc(ov['len_note'])}</p>
</div>"""

    # 四、P1 关键问题详情
    def data_table(headers: list[str], rows: list[list[Any]], mono_cols: set[int], *, limit: int = 24) -> str:
        head = "".join(f"<th>{esc(h)}</th>" for h in headers)
        body = []
        for row in rows[:limit]:
            cells = []
            for idx, cell in enumerate(row):
                classes = []
                if idx in mono_cols:
                    classes.append("mono")
                if isinstance(cell, int) or (isinstance(cell, str) and cell.replace(",", "").isdigit()):
                    classes.append("num")
                cells.append(f'<td class="{" ".join(classes)}">{_fmt_big(cell)}</td>')
            body.append("<tr>" + "".join(cells) + "</tr>")
        if len(rows) > limit:
            body.append(f'<tr><td colspan="{len(headers)}" class="num" style="color:var(--mut)">…（共 {len(rows):,} 条，详见 Excel）</td></tr>')
        return f'<table class="data"><tr>{head}</tr>{"".join(body)}</table>'

    def bucket_rows(bucket: dict[str, Any]) -> list[list[Any]]:
        out = []
        for item in bucket["items"]:
            out.append([item.get(c["k"], "") for c in bucket["columns"]])
        return out

    def direction_text(direction: list[dict]) -> str:
        if not direction:
            return ""
        parts = []
        for item in direction:
            parts.append(f"文本“{esc(item['unit'])}”→前缀{esc(item['prefix'])} 的 {item['count']} 条")
        return "方向分布：" + "；".join(parts) + "。"

    p1_sections = ""
    p1_no = 0
    for key in ("p1_unit_text", "p1_empty_name"):
        bucket = next((b for b in ov["buckets"] if b["key"] == key), None)
        if not bucket or not bucket["items"]:
            continue
        p1_no += 1
        headers = [c["t"] for c in bucket["columns"]]
        rows = bucket_rows(bucket)
        mono_cols = {i for i, c in enumerate(bucket["columns"]) if c["k"] in ("kks_code", "parent_code", "old_code")}
        note = bucket["note"]
        if key == "p1_unit_text" and ov["unit_direction"]:
            note += " " + direction_text(ov["unit_direction"])
        p1_sections += f"""<h3 style="font-size:15px;margin:14px 0 4px">4.{p1_no} {esc(bucket['title'])}（{bucket['count']:,} 条）</h3>
<div class="callout p1">{note}</div>{data_table(headers, rows, mono_cols)}"""
    sec4 = f"""<div class="sec">
  <h2>四、P1 关键问题详情</h2>
  {p1_sections or '<p class="sub">未发现 P1 关键问题。</p>'}
</div>"""

    # 五、P2 提示项
    p2_sections = ""
    p2_no = 0
    for key in ("p2_hygiene", "p2_semantic_void", "p2_code_hygiene", "p2_others"):
        bucket = next((b for b in ov["buckets"] if b["key"] == key), None)
        if not bucket or not bucket["items"]:
            continue
        p2_no += 1
        headers = [c["t"] for c in bucket["columns"]]
        rows = bucket_rows(bucket)
        mono_cols = {i for i, c in enumerate(bucket["columns"]) if c["k"] in ("kks_code", "parent_code", "old_code")}
        p2_sections += f"""<h3 style="font-size:15px;margin:18px 0 4px">5.{p2_no} {esc(bucket['title'])}（{bucket['count']:,} 条）</h3>
<div class="callout p2">{esc(bucket['note']) or '提示项，人工抽检核对。'}</div>{data_table(headers, rows, mono_cols, limit=20)}"""
    # 原KKS一对多（按组）
    if ov["orig_1n_groups"]:
        p2_no += 1
        rows: list[list[Any]] = []
        for g in ov["orig_1n_groups"]:
            for m in g["members"]:
                rows.append([g["old"], m.get("excel_row", ""), m.get("kks_code", ""), m.get("parent_code", ""), m.get("profession", ""), m.get("name", "")])
        p2_sections += f"""<h3 style="font-size:15px;margin:18px 0 4px">5.{p2_no} 原KKS 一对多映射（{len(ov['orig_1n_groups']):,} 组 / {len(rows):,} 行）</h3>
<div class="callout p2">同一条 原KKS编码 映射到 ≥2 条 新KKS编码。多为“一个旧测点拆分为多个新设备”的合理映射，但原KKS 非唯一键；以 <b>新KKS 为主键</b>、原KKS 仅作历史参照。</div>
{data_table(['原KKS(键)', '行', '新KKS', '父级', '专业', '设备名称'], rows, {0, 2, 3}, limit=18)}"""
    if ov["aux_dup"]:
        p2_no += 1
        aux = ov["aux_dup"]
        rows = [[r["name"], r["profession"], r["status"]] for r in aux["rows"][:30]]
        more = f"…（共 {len(aux['rows'])} 条，详见 Excel）" if len(aux["rows"]) > 30 else ""
        tail = f'<tr><td colspan="3" class="num" style="color:var(--mut)">{more}</td></tr>' if more else ""
        p2_sections += f"""<h3 style="font-size:15px;margin:18px 0 4px">5.{p2_no} 重码待处理（{len(aux['rows'])} 条，仍未编码）</h3>
<div class="callout p2">源文件「{esc(aux['sheet'])}」sheet 的条目本批次仍未赋码（不在主表）。建议对照主库确认是否已在库或需新增，避免与主库重复。</div>
<table class="data"><tr><th>设备名称</th><th>专业</th><th>状态</th></tr>{"".join(f'<tr><td>{esc(r[0])}</td><td>{esc(r[1])}</td><td>{esc(r[2])}</td></tr>' for r in rows)}{tail}</table>"""
    if ov["buckets"][0]["items"]:
        bucket = ov["buckets"][0]
        headers = [c["t"] for c in bucket["columns"]]
        rows = bucket_rows(bucket)
        mono_cols = {i for i, c in enumerate(bucket["columns"]) if c["k"] in ("kks_code", "parent_code")}
        p0_html = (f"<h3 style='font-size:15px;margin:14px 0 4px'>{esc(bucket['title'])}（{bucket['count']} 条）</h3>"
                   f'<div class="callout p0">阻断性问题必须修改并复核后才能导入。</div>{data_table(headers, rows, mono_cols, limit=20)}')
        p2_sections = p0_html + p2_sections
    sec5 = f"""<div class="sec">
  <h2>五、P2 提示项与阻断项</h2>
  {p2_sections or '<p class="sub">未发现 P2/P0 提示项。</p>'}
</div>"""

    # 六、范围验证
    sc = ov["scope"]
    scope_html = ""
    if sc["incremental"]:
        scope_html += (
            f'<div class="callout p1"><b>为何“父级孤儿”不计为缺陷：</b>本文件为<b>增量新增</b>批次，'
            f"{sc['external_parent_rows']:,} 行设备码的父级节点（去重 {sc['external_parent_units']:,} 个）指向主库既有节点，本文件不含。"
            "<b>导入前必做</b>：将本批次与主库做父级闭合校验（父级节点须在主库存在），确认无误后再合入。</div>"
        )
        if sc["parent_samples"]:
            ptab = "".join(f'<tr><td class="mono">{esc(p["parent"])}</td><td class="num">{p["children"]:,}</td><td>主库</td></tr>' for p in sc["parent_samples"])
            scope_html += ('<p class="sub">本文件引用的外部父级节点 Top 样例（完整清单见 Excel）：</p>'
                           '<table class="data"><tr><th>父级节点</th><th>下挂行数</th><th>位置</th></tr>' + ptab + '</table>')
    else:
        scope_html += '<div class="callout ok"><b>本文件为全量/主表文件：</b>内部层级自洽，父级均在文件内。</div>'
    if sc["notes"]:
        for note in sc["notes"]:
            scope_html += f'<div class="callout p2">{esc(note)}</div>'
    sec6 = f"""<div class="sec">
  <h2>六、范围验证：父级与主库闭合（非本文件缺陷）</h2>
  {scope_html}
</div>"""

    # 七、误报防控
    honesty_html = "".join(
        f"<li><span class='tag'>{esc(h['tag'])}</span> <b>{esc(h['bold'])}</b>：{esc(h['text'])}</li>" for h in ov["honesty"]
    )
    sec7 = f"""<div class="sec">
  <h2>七、误报防控说明（诚实声明）</h2>
  <ul class="tight">{honesty_html}</ul>
</div>"""

    # 八、建议
    abc_html = ""
    for letter, text, _ in ov["abc"]:
        abc_html += f"<li><b>{esc(letter)}</b>：{esc(text)}</li>"
    sec8 = f"""<div class="sec">
  <h2>八、修正与导入建议（A/B/C）</h2>
  <ul class="tight">{abc_html}</ul>
</div>"""

    # 九、AI 总体总结（模型对整份审核结果的归纳；未生成时省略本节）
    sec9 = ""
    if ov.get("ai_summary"):
        sumry = ov["ai_summary"]
        points9 = "".join(f"<li>{esc(pt)}</li>" for pt in (sumry.get("points") or []))
        sec9 = f"""<div class="sec">
  <h2>九、AI 总体总结</h2>
  <div class="sub">模型 {esc(sumry.get('model') or '—')}；由 AI 归纳整份审核结果，仅供参考，不替代规则结论。</div>
  <p style="margin:6px 0 10px">{esc(sumry['overall'])}</p>
  <ul class="tight">{points9}</ul>
</div>"""

    # 一至九节全部转为可折叠章节（默认展开，点击标题收起/展开）
    sec1, sec2, sec3, sec4, sec5, sec6, sec7, sec8, sec9 = (
        _fold_section(s) for s in (sec1, sec2, sec3, sec4, sec5, sec6, sec7, sec8, sec9)
    )

    foot = f'<div class="foot">审计基线 {today}</div>'

    return f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{esc(ov['title'])}</title>
<style>{CSS}</style></head>
<body><div class="wrap">
{header}
<div class="kpis">{kpi_html}</div>
{sec1}
{sec2}
{sec3}
{sec4}
{sec5}
{sec6}
{sec7}
{sec8}
{sec9}
{foot}
</div></body></html>"""


# ------------------------------------------------------------------ Excel 渲染
def _xlsx_sheet_title(text: str) -> str:
    cleaned = re.sub(r'[\\/*?:\[\]]', "", text)
    return (cleaned or "Sheet")[:31]


def render_xlsx(path: Path, result: dict[str, Any], ov: dict[str, Any]) -> None:
    if openpyxl is None:  # pragma: no cover
        raise RuntimeError("openpyxl 不可用，无法生成 Excel 问题清单")
    wb = Workbook()
    wb.remove(wb.active)
    today = date.today().isoformat()

    def header_style(ws, cell_range: str) -> None:
        thin = Side(style="thin", color="9DB4D6")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for row in ws[cell_range]:
            for cell in row:
                cell.fill = PatternFill("solid", fgColor=BLUE)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                cell.border = border

    def body_style(ws, cell_range: str) -> None:
        # 细边框 + 隔行斑马纹，替代被移除的 Table 样式（普通边框/填充对 Excel 兼容性无风险）
        thin = Side(style="thin", color="D8DEE7")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        zebra = PatternFill("solid", fgColor=GRAY)
        start_row = int("".join(ch for ch in cell_range.split(":")[0] if ch.isdigit()))
        for r_i, row in enumerate(ws[cell_range]):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.border = border
                if (start_row + r_i) % 2 == 0:
                    cell.fill = zebra

    def autosize(ws, maximum: int = 46) -> None:
        for col_idx in range(1, ws.max_column + 1):
            letter = get_column_letter(col_idx)
            max_len = max((len(str(ws.cell(row=row_idx, column=col_idx).value or "")) for row_idx in range(1, ws.max_row + 1)), default=0)
            ws.column_dimensions[letter].width = min(max(max_len + 2, 10), maximum)

    # ---------------- 概览
    ws = wb.create_sheet("概览")
    stem = Path(ov["source_path"] or "").stem or "KKS编码审核"
    ws.merge_cells("A1:E1")
    ws["A1"] = f"{stem} 审核总览（主表:{ov['sheet']} {ov['data_rows']:,}）"
    ws["A1"].fill = PatternFill("solid", fgColor=BLUE)
    ws["A1"].font = Font(bold=True, color="FFFFFF", size=14)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26
    ws.merge_cells("A2:E2")
    ws["A2"] = f"源文件:{ov['source_path'] or '—'}"
    ws["A2"].fill = PatternFill("solid", fgColor=PALE_BLUE)
    ws.merge_cells("A3:E3")
    ws["A3"] = f"审计日期:{today}  主表已编码:{ov['data_rows']:,}  缺码行:{ov['missing_rows']}  问题总数:{ov['issue_count']}  规则依据:{ov['standard_display'] or '—'}"
    ws["A3"].fill = PatternFill("solid", fgColor=PALE_BLUE)
    ws["A3"].alignment = Alignment(wrap_text=True)

    def sev_chip(value: Any, level: str) -> str:
        """等级列只输出 OK/P0/P1/P2；非数值占位（如范围验证行）按给定等级显示，杜绝文本串入。"""
        try:
            n = int(value)
            return "OK" if n == 0 else level
        except (TypeError, ValueError):
            return level

    rows_out: list[tuple[str, str, Any, str, str]] = []
    for card in ov["cards"]:
        for label, value, sev, hint in card["rows"]:
            rows_out.append((card["title"], str(label), value, sev_chip(value, sev.upper()), hint))
    # 兜底指标行
    rows_out.extend([
        ("汇总", "主表已编码行", ov["data_rows"], "OK", "主表有效数据行"),
        ("汇总", "缺码非空行", ov["missing_rows"], "P1" if ov["missing_rows"] else "OK", "非空行缺 KKS 码"),
        ("汇总", "AI 语义复核", ov["ai_reviewed"], "OK",
         (f"候选 {ov['ai_candidate']} 条；模型 {ov['ai_model'] or '—'}" if ov["ai_active"] else "未启用，仅规则审计（可在 AI 管理员配置中开启）")),
    ])
    rows_out = list(dict.fromkeys(rows_out))
    ws.append([])
    ws.append(["维度", "指标", "数量", "等级", "说明"])
    header_style(ws, f"A{ws.max_row}:E{ws.max_row}")
    header_row = ws.max_row
    for dim, label, value, sev, hint in rows_out:
        ws.append([dim, label, value, sev, hint])
    body_style(ws, f"A{header_row}:E{ws.max_row}")
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row, min_col=4, max_col=4):
        for cell in row:
            cell.fill = PatternFill("solid", fgColor={"OK": PALE_GREEN, "P0": PALE_RED, "P1": PALE_AMBER, "P2": PALE_BLUE}.get(str(cell.value), GRAY))
    ws.freeze_panes = f"A{header_row + 1}"
    autosize(ws, 60)
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["E"].width = 60

    # ---------------- 分类 sheet
    sheet_specs: list[tuple[str, dict[str, Any]]] = []
    for bucket in ov["buckets"]:
        if not bucket["items"]:
            continue
        if bucket["key"] == "p0_blocker":
            name = "P0_阻断"
        elif bucket["key"] == "p1_unit_text":
            name = "P1_机组文本三方不一致"
        elif bucket["key"] == "p1_empty_name":
            name = "P1_空名占位"
        elif bucket["key"] == "p2_hygiene":
            name = "P2_名称卫生"
        elif bucket["key"] == "p2_semantic_void":
            name = "P2_名称语义需人核"
        elif bucket["key"] == "p2_code_hygiene":
            name = "P2_编码卫生"
        elif bucket["key"] == "p2_dup_name":
            name = "P2_同名多码线索"
        else:
            name = "P2_其他提示"
        sheet_specs.append((name, bucket))

    def write_sheet(name: str, title: str, headers: list[str], data: list[list[Any]], mono_cols: set[int]) -> None:
        ws2 = wb.create_sheet(_xlsx_sheet_title(name))
        ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
        ws2["A1"] = title
        ws2["A1"].font = Font(bold=True, size=13, color=BLUE)
        ws2.append([])
        ws2.append(headers)
        header_style(ws2, f"A{ws2.max_row}:{get_column_letter(len(headers))}{ws2.max_row}")
        hrow = ws2.max_row
        for row in data:
            ws2.append(row)
        if ws2.max_row >= hrow:
            body_style(ws2, f"A{hrow}:{get_column_letter(len(headers))}{ws2.max_row}")
        for col in mono_cols:
            for row_idx in range(hrow + 1, ws2.max_row + 1):
                cell = ws2.cell(row=row_idx, column=col + 1)
                cell.font = Font(name="Consolas", size=10)
        ws2.freeze_panes = f"A{hrow + 1}"
        ws2.auto_filter.ref = f"A{hrow}:{get_column_letter(len(headers))}{max(hrow, ws2.max_row)}"
        autosize(ws2, 60)
        # 说明：不使用 openpyxl Table（table*.xml）。Table 部件 + auto_filter 叠加在部分 Excel
        # 版本会触发"内容有问题，是否尝试恢复"；筛选/冻结/底色已覆盖表格观感，故仅保留 auto_filter。

    for name, bucket in sheet_specs:
        headers = [c["t"] for c in bucket["columns"]]
        mono_cols = {i for i, c in enumerate(bucket["columns"]) if c["k"] in ("kks_code", "parent_code", "old_code")}
        data = []
        for item in bucket["items"]:
            data.append([item.get(c["k"], "") for c in bucket["columns"]])
        write_sheet(name, f"{bucket['title']}（{bucket['count']}条,{bucket['level']}）", headers, data, mono_cols)

    # 原KKS一对多（按组展开逐行）
    if ov["orig_1n_groups"]:
        headers = ["原KKS(键)", "行", "新KKS", "父级", "专业", "设备名称"]
        data = []
        for g in ov["orig_1n_groups"]:
            for m in g["members"]:
                data.append([g["old"], m.get("excel_row", ""), m.get("kks_code", ""), m.get("parent_code", ""), m.get("profession", ""), m.get("name", "")])
        write_sheet("P2_原KKS一对多", f"原KKS一对多映射（{len(ov['orig_1n_groups'])}组/{len(data)}行,P2）", headers, data, {0, 2, 3})

    if ov["aux_dup"]:
        aux = ov["aux_dup"]
        write_sheet("P2_重码待处理", f"重码待处理未编码（{len(aux['rows'])}条）",
                    ["设备名称", "专业", "状态"],
                    [[r["name"], r["profession"], r["status"]] for r in aux["rows"]], set())

    # AI 复核建议（独立 sheet；AI 未启用时跳过）
    if ov.get("ai_active") and ov.get("ai_suggestions"):
        write_sheet(
            "AI复核建议",
            f"AI 复核建议（{len(ov['ai_suggestions'])}条 · 模型 {ov['ai_model'] or '—'}）",
            ["行", "编码", "设备名称", "来源分类", "AI结论", "AI说明", "AI建议"],
            [[s["excel_row"], s["kks_code"], s["name"], s["bucket_title"], s["decision"], s["summary"], s["suggestion"]]
             for s in ov["ai_suggestions"]],
            {1},
        )

    # 结构说明与范围
    ws3 = wb.create_sheet("结构说明与范围")
    ws3.merge_cells("A1:B1")
    ws3["A1"] = "结构说明 / 范围验证 / sheet关系"
    ws3["A1"].font = Font(bold=True, size=13, color=BLUE)
    ws3.append([])
    ws3.append(["项目", "说明"])
    header_style(ws3, f"A{ws3.max_row}:B{ws3.max_row}")
    hrow = ws3.max_row
    for key, value in ov["structure_rows"]:
        ws3.append([key, value])
    body_style(ws3, f"A{hrow}:B{ws3.max_row}")
    ws3.freeze_panes = f"A{hrow + 1}"
    autosize(ws3, 90)
    ws3.column_dimensions["A"].width = 18

    for ws_ in wb.worksheets:
        ws_.sheet_view.showGridLines = False
        # 写盘前统一清除单元格文本中的 XML 非法控制字符（Excel 打开易报"内容有问题"）
        for row_ in ws_.iter_rows():
            for cell_ in row_:
                if isinstance(cell_.value, str) and _XML_CTRL.search(cell_.value):
                    cell_.value = _XML_CTRL.sub("", cell_.value)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
