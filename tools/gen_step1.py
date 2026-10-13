# -*- coding: utf-8 -*-
"""第一步智能体 · 按设计编码生成设备树，并把结果**回填到原文件**。

**主要交付物：`<原名>_已填写.xlsx`，另附设备编码待核清单** —— 保留用户原表的格式与内容，只把
`设备名称 / KKS编码 / 上级设备名称 / 上级KKS编码` 四个空列填上（**已有内容不覆盖**），
并附一列「修编说明」说明编码哪里不符国标、依据哪条。

单一行为（**不提供选项**）：AI 补名称默认开启——规则能给出的名称直接用规则，
规则给不出的交由 AI 补全。

命令行：`--no-ai` 可关掉 AI（纯规则、秒级，供测试与离线复现用）。

变更记录（2026-10-10，按业务要求）：
  · 移除「与竣工版逐列比对」整个功能（前后端），不再产出对比列/准确率；
  · 产出收敛为**只保留回填后的原表**，不再落「干净表」「对照表」两份 xlsx；
  · 节点全量仍写 `第一步_生成结果.json`——**仅供后台统计与复现，不在下载列表**。
"""
from __future__ import annotations

import json
import pathlib
import re
import sys
import time
import unicodedata
from contextvars import ContextVar
from collections import Counter

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "kks_gen"))

from tree_rules import build_tree, order_nodes  # noqa: E402

# ------------------------------------------------------------------ 进度上报
# Web 端要用进度条，所以这里提供两个模块级钩子：
#   PROGRESS_HOOK(event: dict)      —— 粗粒度阶段（读取/建树/比对/落盘）
#   PROGRESS_AI_HOOK(event: dict)   —— 细粒度（AI 每一批完成）
# 默认 None，命令行跑时完全不触发，行为与以前一致。
PROGRESS_HOOK = None
PROGRESS_AI_HOOK = None
_RUN_PROGRESS = ContextVar("kks_run_progress", default=None)


def _emit(event: dict) -> None:
    hook = _RUN_PROGRESS.get() or PROGRESS_HOOK
    if hook is None:
        return
    try:
        hook(event)
    except Exception:      # 进度上报绝不能影响生成
        pass

# ------------------------------------------------------------------ 列定位（模糊匹配）
# 各厂/各设计院的表头写法不统一，必须按语义匹配，不能靠固定列名。
# 实测：把「设计编码」改成「设计码」、「原名称」改成「图纸名称」时，
# 固定列名版本会把这两列当空 → 整个流程失效。
COLUMN_RULES: list[tuple[str, tuple[str, ...], tuple[str, ...]]] = [
    # (内部字段, 命中任一即算, 命中即排除)
    # 实际见过的表头写法差异很大："原KKS编码"（现场汇总表）vs "设计编码"（竣工版）
    ("design_code", ("设计编码", "设计码", "设计编号", "设计kks", "设计院编码",
                     "原kks编码", "原kks码", "原编码", "原设计编码", "design"),
     ("修编", "修改后", "新", "最终")),
    ("original_name", ("原名称", "原始名称", "设计名称", "图纸名称", "原名", "原设备名",
                       "原设备名称",
                       # 业主模板写法：`设备名称（原设计图内）` / `设备名称（设计图内）`。
                       # 用 `原设计图` 而不是 `设计图内`——后者在 `原KKS编码（设计图内）` 里也有，
                       # 会把编码列误判成名称列（实测：舟山暖通两版都是这个表头）。
                       "原设计图"),
     ("或图号", "图号", "专业")),
    ("kks", ("kks编码", "kks码", "编码", "kks"), ("原", "旧", "父", "上级", "设计", "位置标识")),
    ("device_name", ("设备名称", "设备名", "名称"), ("原", "设计", "父", "上级", "系统", "图纸")),
    ("parent_kks", ("上级kks", "父级编码", "父kks", "父级kks", "上级编码"), ()),
    ("parent_name", ("上级设备名称", "父级名称", "上级名称", "父设备名称"), ()),
    ("unit_label", ("机组号", "机组号列"), ()),
    ("system_name", ("系统",), ("专业", "分类", "图")),
    ("device_type", ("设备类型", "设备种类"), ()),
    ("profession", ("专业",), ("图纸",)),
    ("drawing", ("图纸名称或图号", "图号", "图纸", "图纸名称"), ()),
]

# 关键列：缺任何一个都无法生成
KEY_COLUMNS = ("design_code", "original_name")

# ---- 「修编说明」里不该出现的两类文字（都是实测出来的噪声）----
# ① **状态回显**：没改任何东西的行写「已符合国标」等于没说 —— 默认就是不用改。
_NOOP_NOTE = re.compile(r"^已(?:符合国标|通过自动结构校验)")
# ② **待办提示**：只在对应目标列**确实为空**时才有意义。
#    竣工版实测：设备名称列本来就有数据的行，还提示"需按编码语义推断"，是噪声。
_PROMPT_NOTE_COLUMN = {
    "原名称为空，设备名称需按编码语义+图号推断": "device_name",
}


def head_name(key: str) -> str:
    return {
        "design_code": "设计编码", "original_name": "原名称", "kks": "KKS编码",
        "device_name": "设备名称", "parent_kks": "上级KKS编码",
        "parent_name": "上级设备名称", "profession": "专业", "drawing": "图纸名称或图号",
        "unit_label": "机组号", "system_name": "系统", "device_type": "设备类型",
    }.get(key, key)


def resolve_columns(header: list[str]) -> dict[str, int | None]:
    """按表头语义定位列。返回 {内部字段: 列下标 or None}。"""
    lowered = [h.strip().lower() for h in header]
    out: dict[str, int | None] = {key: None for key, _hit, _not in COLUMN_RULES}
    used: set[int] = set()
    for key, hits, negatives in COLUMN_RULES:
        best = None
        for i, text in enumerate(lowered):
            if i in used or not text:
                continue
            if any(neg in text for neg in negatives):
                continue
            if any(hit in text for hit in hits):
                # 命中越"专有"越好：取最长命中词的长度作为匹配度
                score = max(len(hit) for hit in hits if hit in text)
                if best is None or score > best[0]:
                    best = (score, i)
        if best is not None:
            out[key] = best[1]
            used.add(best[1])
    return out


# 数据行开始前，真实表格常有一行"逐列说明"（合并单元格里的长文字）。
# 判据：该行任一单元格是长文本（>40 字）→ 说明行，不是数据。
HEADER_SCAN_ROWS = 25
DESC_CELL_LIMIT = 40
CJK = re.compile(r"[\u4e00-\u9fff]")


def _row_score(row: list) -> tuple[int, int]:
    """给一行打"像不像表头"的分：命中列越多越像；有超长文本则扣分。"""
    cells = [str(v or "").strip() for v in row]
    col = resolve_columns(cells)
    hits = sum(1 for key in col if col[key] is not None)
    long_text = sum(1 for c in cells if len(c) > DESC_CELL_LIMIT)
    return hits - long_text * 2, hits


def detect_header_row(rows: list[list]) -> int | None:
    """在前若干行里找出真表头行（返回行下标，0 基）。

    真实文件常见：第 1 行是"表格说明"长文本（合并单元格），表头在第 2 行甚至更后。
    判据：命中列关键字最多、且没有超长说明文本；且必须能找到关键列。
    """
    best_index, best_score = None, -999
    for i, row in enumerate(rows[:HEADER_SCAN_ROWS]):
        score, hits = _row_score(row)
        if hits == 0:
            continue
        col = resolve_columns([str(v or "").strip() for v in row])
        if any(col.get(key) is None for key in KEY_COLUMNS):
            continue                      # 关键列不全的行不算表头
        if score > best_score:
            best_index, best_score = i, score
    return best_index


def normalize_design_code(raw: str) -> str:
    """清洗设计编码：去内部空格、全角转半角；含中文的直接判为无效（返回空串）。

    · `07GQJ40 BT301` → `07GQJ40BT301`（真实表格里编码被空格断开）
    · `7号炉辅汽至锅炉71-54m区域产品伴热用汽管` → 空（这是名称不是编码）
    """
    text = unicodedata.normalize("NFKC", str(raw or "")).strip()
    if not text or CJK.search(text):
        return ""
    return re.sub(r"\s+", "", text)



def _first_sheet(path: pathlib.Path) -> str:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    return names[0]



def load(path: pathlib.Path):
    """读输入文件。

    三处按真实文件放宽（都是踩过的坑）：
      1. **列定位按表头语义模糊匹配**，不依赖固定列名（"原KKS编码" vs "设计编码"）；
      2. **自动探测表头行**：真实汇总表第 1 行常是"表格说明"长文本，表头在第 2 行；
      3. **设计编码做清洗**：去内部空格（`07GQJ40 BT301`→`07GQJ40BT301`），
         含中文的判为无效（那是名称不是编码）。
    """
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True)[_first_sheet(path)]
    try:
        rows = list(ws.iter_rows(values_only=True))
    finally:
        # ⚠️ 必须关闭：openpyxl 的 read_only 模式会一直持有文件句柄，
        #    不关的话上传的输入文件在 Windows 上会一直"被占用"，
        #    保留策略清理、删除产物、pytest 清理临时目录都会报 WinError 32。
        ws.parent.close()

    header_index = detect_header_row(rows)
    if header_index is None:
        preview = [h for h in (str(v or "").strip() for v in (rows[0] if rows else [])) if h]
        raise ValueError(
            f"表头识别失败：在前 {HEADER_SCAN_ROWS} 行里没找到同时含"
            "「设计编码 / 原KKS编码」与「原名称 / 原设备名称」的表头行。"
            f"第 1 行内容={preview[:2]}。请确认表格有列标题行。"
        )
    header = [str(v or "").strip() for v in rows[header_index]]
    col = resolve_columns(header)
    missing = [head_name(k) for k in KEY_COLUMNS if col.get(k) is None]
    if missing:
        raise ValueError(
            f"表头识别失败，缺少关键列 {missing}。"
            f"识别到的表头行=第 {header_index + 1} 行，内容={[h for h in header if h]}。"
            "请确认文件含「设计编码」或「原KKS编码」，以及「原名称」或「原设备名称」。"
        )
    print(f"表头行: 第 {header_index + 1} 行")
    print("列定位结果:", {head_name(k): header[v] for k, v in col.items() if v is not None})

    def cell(r, key):
        i = col.get(key)
        return str(r[i] or "").strip() if i is not None and i < len(r) else ""

    seeds = []
    skipped_bad_code = 0
    final_code_rows = 0
    for idx, r in enumerate(rows[header_index + 1:], start=header_index + 2):
        raw_code = cell(r, "design_code")
        code = normalize_design_code(raw_code) if raw_code else ""
        is_final = False
        if not code:
            # ---- 兜底：这一行**没有可用的设计编码**（空格，或填的是名称串），但已有最终编码 ----
            # 实测（鄂州四期竣工版）：113,413 行里 47,984 行属于这一类——
            # 只有 KKS 编码和设备名称，没有设计编码。以前这些行整行"未判定"、
            # 上级编码/系统/设备类全填不出来。它们的信息是**现成的**，
            # 所以用已有编码当种子参与建树：**只审查、绝不改写**（见 main 的 code_is_final）。
            fallback = normalize_design_code(cell(r, "kks"))
            if fallback:
                code, is_final = fallback, True
            else:
                if raw_code:
                    skipped_bad_code += 1     # 说明行/名称串/表头残留，且没有最终编码可兜底
                continue
        if is_final:
            final_code_rows += 1
        seeds.append({
            "excel_row": idx,
            "design_code": code,
            "code_is_final": is_final,   # True：编码是原表已有的最终码，不得改写
            # 无设计编码的行，「原名称」列常为空 —— 用已有的「设备名称」当名称来源
            "original_name": cell(r, "original_name") or (cell(r, "device_name") if is_final else ""),
            "profession": cell(r, "profession"),
            "drawing": cell(r, "drawing"),
            # 修编用：机组号（定全厂码）、系统（定系统分类码）、设备类型（定设备分类码）
            "unit_label": cell(r, "unit_label"),
            "system_name": cell(r, "system_name"),
            "device_type": cell(r, "device_type"),
        })
    if skipped_bad_code:
        print(f"已跳过 {skipped_bad_code} 行：设计编码列不是有效编码（说明行/名称串/表头残留）")
    if final_code_rows:
        print(f"已有最终编码、无设计编码的行: {final_code_rows:,}（按「审查」处理，绝不改写）")
    layout = {"header_index": header_index, "columns": col, "header": header}
    return seeds, layout


def _collect_ai_rows(nodes: dict) -> list[dict]:
    """挑出需要 AI 生成名称的行，并附上判断依据（分相/同类名称/系统与设备码语义）。

    只挑两类：
      ① 名称为空（原名称为空 → 只能靠语义推）
      ② 原名称与规则结果差异较大（需要语义重写的行）
    规则已给对的不送模型，省时间省钱。
    """
    import json as _json

    idx_path = ROOT / "config" / "kks_code_index.json"
    try:
        idx = _json.loads(idx_path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        idx = {}
    sys_codes = idx.get("system_codes") or {}
    dev_codes = idx.get("device_codes") or {}

    # 同类设备已有名称（同设备分类码且名称非空），作为"同厂惯例"证据
    by_class: dict[str, list[str]] = {}
    for kks, node in nodes.items():
        base = kks.split(".", 1)[0]
        if len(base) >= 9 and node.name:
            by_class.setdefault(base[7:9], []).append(node.name)
    # 同父级兄弟名称
    by_parent: dict[str, list[str]] = {}
    for kks, node in nodes.items():
        if node.parent_kks and node.name:
            by_parent.setdefault(node.parent_kks, []).append(node.name)

    out: list[dict] = []
    for kks, node in nodes.items():
        base = kks.split(".", 1)[0]
        needs = (not node.name) or node.need_human
        if not needs or not node.design_code:
            continue
        phase = ""
        if len(base) == 13 and base[-1] in ("A", "B", "C"):
            phase = f"{base[-1]}相"
        out.append({
            "row": len(out) + 1,
            "kks": kks,
            "original_name": node.original_name,
            "parent_name": node.parent_name,
            "parent_kks": node.parent_kks,
            "profession": node.profession,
            "drawing": node.drawing,
            "system_meaning": str(sys_codes.get(base[2:5], ""))[:60] if len(base) >= 5 else "",
            "device_meaning": str(dev_codes.get(base[7:9], ""))[:60] if len(base) >= 9 else "",
            "phase_suffix": phase,
            "same_class_names": [n for n in by_class.get(base[7:9], []) if n != node.name][:6],
            "sibling_names": [n for n in by_parent.get(node.parent_kks, []) if n != node.name][:6],
        })
    return out


def _tree_order(nodes: dict) -> list[str]:
    """深度优先的树序：每个节点**紧接**在其子节点之前（与竣工版一致）。

    ⚠️ 不能用 `order_nodes()`——那是**层序**（先按 专业/机组，再按层级长度排序：
       所有机组行 → 所有主组行 → 所有系统行 → …），父子不相邻，
       呈现出来不是竣工版那种"一棵树"的样子。
    同层兄弟按 KKS 升序（与竣工版一致）。
    """
    children: dict[str, list[str]] = {}
    for kks, node in nodes.items():
        parent = getattr(node, "parent_kks", "") or ""
        if parent and parent in nodes and parent != kks:
            children.setdefault(parent, []).append(kks)
    for lst in children.values():
        lst.sort()

    def _parent_of(k: str) -> str:
        return getattr(nodes[k], "parent_kks", "") or ""

    # 根 = 上级不在节点集里的（含机组/全厂码这类没有上级的）
    roots = sorted(k for k in nodes if _parent_of(k) not in nodes)
    out: list[str] = []
    stack = list(reversed(roots))
    seen: set[str] = set()
    while stack:
        k = stack.pop()
        if k in seen:                     # 防环（异常数据兜底）
            continue
        seen.add(k)
        out.append(k)
        stack.extend(reversed(children.get(k) or []))
    out.extend(sorted(k for k in nodes if k not in seen))    # 兜底：环上未访问到的节点
    return out


def _write_cell(ws, row: int, col: int, value, style) -> bool:
    """写一格（值 + 样式），返回是否真写了。

    ⚠️ **合并区内的非锚点格是只读的 `MergedCell`**，硬写会 `AttributeError: 'MergedCell'
       object attribute 'value' is read-only`（实测踩过：原表说明行 `A1:H1` 右移成 `B1:I1` 后，
       写 C1..I1 就炸）。Excel 只显示锚点格内容，所以**跳过**不影响显示效果。
    """
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, openpyxl.cell.cell.MergedCell):
        return False
    cell.value, cell._style = value, style
    return True


def _remap_layout(ws, delta: int, *, defer_merges: bool = False) -> None:
    """把工作表的**版式**整体右移 delta 列：列宽、合并区、自动筛选、冻结窗格。

    ⚠️ 为什么必须做：左插一列（「原文件序号」）后，若只挪单元格内容而不挪这些，会出现
       **列宽错位**（A 列的宽度用在了 B 列上）、**合并区压在新列上**（原 `A1:J1` 的说明行
       会盖住新序号列）、**筛选范围少一列**。三者都是"打开文件才看得出来"的坑，故一并处理。
    不做的事：行高（行号不动）、页边距/打印设置（与列无关）。
    """
    # ① 列宽：逐个右移（含"默认宽度"那条跨到 16384 列的定义）
    # ⚠️ 必须**先把值记下来、清空、再重建**：`ws.column_dimensions['B']` 返回的是**同一个对象**，
    #    边遍历边写会自我覆盖（A 的宽度写进 B 后，下一轮从 B 读到的已是 A 的值 → 全部退化成同一个宽度），
    #    而且随后删旧键会把刚建好的新键一起删掉。实测过：结果所有列宽都变成默认 13.0。
    _old_dims = {k: (dim.min, dim.max, dim.width, dim.hidden, dim.bestFit)
                 for k, dim in ws.column_dimensions.items()}
    for _k in list(ws.column_dimensions.keys()):
        del ws.column_dimensions[_k]
    for key, (mn, mx, width, hidden, best) in _old_dims.items():
        new_min = (mn or column_index_from_string(key)) + delta
        new_max = (mx or new_min) + delta
        target = ws.column_dimensions[get_column_letter(new_min)]
        target.width = width
        target.hidden = hidden
        target.bestFit = best
        # ⚠️ 只搬"列本身的属性"，**不搬列级样式**：openpyxl 的 `customWidth`/`style_id` 是
        #    只读属性，赋值直接 AttributeError。单元格样式另有 `_style` 逐格搬运。
        target.min, target.max = new_min, new_max
    # ② 合并区
    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))
        if not defer_merges:
            ws.merge_cells(start_row=rng.min_row, start_column=rng.min_col + delta,
                           end_row=rng.max_row, end_column=rng.max_col + delta)
    # ③ 自动筛选：右移后**左边界拉到 A** —— 把新列也纳入筛选，方便按原文件序号筛行
    if ws.auto_filter and ws.auto_filter.ref:
        from openpyxl.worksheet.cell_range import CellRange

        a = CellRange(ws.auto_filter.ref)
        ws.auto_filter.ref = (f"A{a.min_row}:"
                              f"{get_column_letter(a.max_col + delta)}{a.max_row}")
    # ④ 冻结窗格：只冻行时列部分是 A（不涉及列）→ 不动；冻了列则右移（否则冻错列）
    m = re.match(r"^([A-Z]+)(\d+)$", str(ws.freeze_panes or ""))
    if m and m.group(1) != "A":
        ws.freeze_panes = (get_column_letter(column_index_from_string(m.group(1)) + delta)
                           + m.group(2))


def fill_original_table(src: pathlib.Path, out_path: pathlib.Path, nodes: dict,
                        layout: dict, code_map: dict | None = None, *,
                        ordered: list[str] | None = None,
                        seeds: list[dict] | None = None) -> dict:
    """把生成结果写回原表，另存为新文件（**不动源文件**）。

    ## 行序（2026-10-10 晚按用户口径改，**这是本项目第二次改行序，改回"不换顺序"**）

    · **原表每一行按 Excel 行序逐行往下**（第 1 行起），**不重排** —— 输出与原文件可以一行行
      对下去做人工审核（此前是"全表树序"，用户拿输出找不到原文件对应行）。
    · **新增的父级行插在「它的第一个子行」上方**（自上而下：全厂码 → 主组 → 系统 → 子系统 → 设备），
      即用户说的"和竣工版一样放在对应位置"；这些行**不参与原文件序号的连续性**。
    · `ordered` 参数**已不再影响行序**（保留形参只为兼容既有调用），排序规则见上。

    ## 列（新增第一列「原文件序号」）

    · 原表所有列**整体右移一列**（列宽 / 合并区 / 自动筛选 / 冻结窗格一并右移，见 `_remap_layout`）；
    · 第 1 列 = **原文件序号 = 该行在原文件里的 Excel 实际行号**（表头行写列名）；
    · **新增的父级行该格留空** → 口径："有号 = 原文件来的行、空号 = 新增行"，1:1 可对；
    · 原表自带的「序号」列（人为写的 1/2/3 或竣工版那种 `总1000001`）**原样保留**、不是一回事。

    目标列：设备名称 / KKS编码 / 上级设备名称 / 上级KKS编码（+ 列规则表推导的系统/设备类），
    末尾另加一列「修编说明」。

    ⚠️ 两条经业务确认的口径（2026-10-10，**以业务铁律为准**）：
      ① 🔴 **只填原本为空的格**：原表一开始有数据的格不回填；2026-10-12 最新要求例外：最终列指令/测点和重复编码必须清空并注明原值。
         此前的"引擎结果覆盖原值"写法已在竣工版上覆盖 24.7 万格，**已撤销**。
      ② **只生成父级行，不生成部件行**：部件级自动生成实测精确率 0%（交接文档 §6.9），不进交付文件。

    实现方式：**在加载出来的原工作簿里原地重写**（不是新建工作簿再套样式），
    这样页边距、行高、字体（含红色字体）等都天然继承，最大程度保住用户原表的样子。
    """
    import column_rules as CR                    # 列规则表（新增列 = 加一条规则）
    from kks_gen.compliance import assess, plant_code

    def can_fill_code(value):
        assessment = assess(value)
        return assessment.writable and (not assessment.pending or
            (assessment.kind == "hierarchy" and int(plant_code(value) or "99") < 91))

    col = layout.get("columns") or {}
    header_index = layout.get("header_index", 0)        # 0 基
    targets = [(key, col.get(key)) for key in ("device_name", "kks", "parent_kks", "parent_name")]
    targets = [(key, idx) for key, idx in targets if idx is not None]
    tgt_idx = dict(targets)                      # 目标列下标（写"待办提示"前要查该列空不空）
    # 额外列（系统 / 设备类…）：按**列规则表**从编码推导，同样只填空格。
    # 规则表见 kks_gen/column_rules.py —— 新增一列只在那里加一行即可。
    _used = {idx for _k, idx in targets}
    for _i, _h in enumerate(layout.get("header") or []):
        _name = str(_h or "").strip()
        _rule = CR.COLUMN_RULES.get(_name)
        if _rule and _i not in _used:
            targets.append((f"@derive:{_rule}:{_name}", _i))
            _used.add(_i)
    if "kks" not in tgt_idx:
        raise ValueError("原表没有最终KKS编码列，无法生成可审核的交付文件")
    notes = layout.get("notes") or {}
    pending = layout.get("pending") or {}
    # 指令/测点行的**原KKS编码**集合 → 这些行不写「KKS编码（修编后）」列。
    # 口径（2026-10-12 用户）：**只有设备才有 KKS 编码**；指令/测点是设备发出的信号，
    # 编码由设备本体行给出，不在这类行上重复（重复 = 交付物报废，实测 18,088 行）。
    not_device = set(layout.get("not_device") or ())

    wb = openpyxl.load_workbook(src)                    # 非 read_only：保留样式与合并单元格
    ws = wb[wb.sheetnames[0]]
    ncol = ws.max_column
    last_row = ws.max_row

    # ---------- 0) 行号归属：表头 + 表头与首个数据行之间的"图例/说明行"原地保留 ----------
    row_orig: dict[int, str] = {}                       # excel 行号 → 原始设计编码
    seed_by_row = {int(seed["excel_row"]): seed for seed in (seeds or []) if seed.get("excel_row")}
    for s in seeds or []:
        er = s.get("excel_row")
        if er:
            row_orig[int(er)] = s.get("original_code") or ""
    first_data = (min(row_orig) if row_orig else last_row + 1)
    # 表头行之前 + 表头行与首个数据行之间（图例行、稀疏说明行）→ 保持不动
    lead_last = first_data - 1

    # ---------- 0b) 新布局：**第一列 = 原文件序号**，原表所有列右移一列 ----------
    # 🔴 用户口径（2026-10-10）：① **不改变原文件行序**；② 第一列给「原文件序号」（**Excel 实际行号**），
    #    否则拿输出没法回原文件逐行人工核对。左插一列会连带版式 → 见 `_remap_layout`。
    SEQ_COL = 0                     # 0 基：新的第一列（原文件序号）
    SHIFT = 1                       # 原表各列右移一列
    src_cols = ncol                 # 原表列数
    note_col = src_cols + SHIFT     # 0 基：修编说明列（紧跟在原表各列之后）
    total_cols = note_col + 1
    ncol = total_cols

    # ---------- 1) 快照**整张表**（值 + 样式）----------
    # 以前只快照数据区（说明行/表头行原地不动）；现在整表右移一列，连它们也要重写。
    snap: dict[int, list[tuple]] = {}
    for r in range(1, last_row + 1):
        snap[r] = [(ws.cell(row=r, column=c + 1).value,
                    ws.cell(row=r, column=c + 1)._style) for c in range(src_cols)]
    source_merges = list(ws.merged_cells.ranges)
    _remap_layout(ws, SHIFT, defer_merges=True)
    from kks_gen.row_policy import classify_row, is_table_instruction
    def source_role(source_row):
        if not source_row:
            return classify_row()
        def original(key):
            index = col.get(key)
            return snap[source_row][index][0] if index is not None else ""
        role = classify_row(original("design_code"), original("original_name") or original("device_name"), original("device_type"))
        return role if role.is_point else classify_row(original("kks"), original("original_name") or original("device_name"), original("device_type"))

    # ---------- 2) 组装输出行序 ----------
    # **原表每一行按 Excel 行序逐行往下（不重排）**；新增的父级行插在「它的第一个子行」上方。
    # 这样输出与原文件可以一行行对下去做人工审核；父级行就地可见（与竣工版的呈现呼应）。
    rows_by_final: dict[str, list[int]] = {}
    for s in seeds or []:
        er, fin = s.get("excel_row"), s.get("design_code") or ""
        if er:
            rows_by_final.setdefault(fin, []).append(int(er))
    row_final: dict[int, str] = {}
    for kks, ers in rows_by_final.items():
        for er in ers:
            row_final[er] = kks

    def _ancestor_chain(kks: str) -> list[str]:
        """根 → 直接父级的祖先链（只含树里存在的节点）。"""
        chain: list[str] = []
        seen: set[str] = set()
        cur = nodes[kks].parent_kks if kks in nodes else ""
        while cur and cur in nodes and cur not in seen:
            chain.append(cur)
            seen.add(cur)
            cur = nodes[cur].parent_kks
        chain.reverse()                                   # 根在前 → 自上而下插入
        return chain

    plan: list[tuple[str, int | None, object]] = []        # (kind, 原表行号, node)
    emitted: set[str] = set()
    for er in range(lead_last + 1, last_row + 1):
        kks = row_final.get(er)
        if not kks or kks not in nodes:
            plan.append(("orig", er, None))                # 说明行/未判定行 → 原地原样（一行不丢）
            continue
        for anc in _ancestor_chain(kks):
            # 只把**生成出来的**父级节点当新增行；原表里本来就有的父级行会走到它自己的位置
            if anc not in emitted and not getattr(nodes[anc], "design_code", ""):
                emitted.add(anc)
                plan.append(("gen", None, nodes[anc]))
        emitted.add(kks)
        plan.append(("orig", er, nodes[kks]))
    # 兜底：没有任何子行的生成节点（理论上不该有）→ 追加末尾，绝不丢节点
    for kks in sorted(k for k, n in nodes.items()
                      if not getattr(n, "design_code", "") and k not in emitted):
        plan.append(("gen", None, nodes[kks]))

    # ---------- 2b) 「KKS编码（修编后）」列**全表唯一**（国标 3.1.1）----------
    # 🔴 交付列里**一条重复都不许有** —— 这是写出时的**最后一道硬闸门**，判据是
    #    **实际要写出去的值**，不是引擎内部结论。为什么必须判在这里：
    #      · 引擎内部那道唯一性校验用的是 `revised_code`，早于本函数的**结构闸门**与
    #        「指令/测点不写」两条规则 —— 被它们拦下的行在内部看是"没撞"，实际却很可能是撞的；
    #      · **源表自带**的重复（两个不同设备被给了同一个原KKS编码）内部逻辑根本发现不了，
    #        因为它按"原码是否相同"分组，同码两行只能算出同一个交付码。
    #    命中撞码的行**一律留空** + 写明冲突（绝不静默产出重复码，也绝不替谁挑一个留下）。
    dup_rows: dict[int, tuple[str, list[int]]] = {}      # 原表行号 → (要写的码, 同码的其它行号)
    kks_idx = next((idx for key, idx in targets if key == "kks"), None)
    if kks_idx is not None:
        planned: dict[str, list[int]] = {}
        for pos, (_kind, _er, _node) in enumerate(plan):
            identity = int(_er) if _er else -(pos + 1)
            existing = snap[_er][kks_idx][0] if _er else None
            if source_role(_er).is_point:
                value = ""
            elif existing not in (None, ""):
                value = str(existing)
            elif (_node is not None and not ((row_orig.get(_er) or "") in not_device)
                  and seed_by_row.get(_er, {}).get("revise_status") != "unresolved"):
                value = str(getattr(_node, "kks", "") or "")
                if not can_fill_code(value):
                    value = ""
            else:
                value = ""
            if value:
                planned.setdefault(value, []).append(identity)
        for value, identities in planned.items():
            if len(identities) > 1:
                for identity in identities:
                    dup_rows[identity] = (value, [x for x in identities if x != identity])

    # ---------- 3) 原地写回 ----------
    layout["note_column"] = note_col
    hc = ws.cell(row=header_index + 1, column=note_col + 1)
    hc.value = "修编说明"
    hc.font = openpyxl.styles.Font(bold=True)
    first_data_row = min(row_final) if row_final else (lead_last + 1)
    base_style = (snap.get(first_data_row) or [(None, None)])[0][1]   # 生成行沿用首个数据行样式

    # 数据区之前的说明行/表头行：原地保留 → 右移一列写回；新第 1 列写行号（表头行写列名）
    for src_r in range(1, lead_last + 1):
        for c in range(min(len(snap.get(src_r) or []), src_cols)):
            _write_cell(ws, src_r, c + 1 + SHIFT, snap[src_r][c][0], snap[src_r][c][1])
        seq = ws.cell(row=src_r, column=SEQ_COL + 1)
        if src_r == header_index + 1:
            seq.value = "原文件序号"
            seq.font = openpyxl.styles.Font(bold=True)
        else:
            seq.value = src_r          # 口径：有号 = 原文件来的行，空号 = 新增行（1:1 可对）

    cells_written = cells_kept = notes_written = gen_rows = blocked_rows = 0
    skipped_kks = 0                                  # 指令/测点行：按口径**有意不填**编码列
    dup_blocked = 0                                  # 撞码行：国标 3.1.1，全表唯一 → 留空
    notes_emptied = 0
    cleared_points = cleared_duplicates = 0
    r = lead_last
    for plan_pos, (kind, er, node) in enumerate(plan):
        r += 1
        if kind == "orig":
            cells = snap[er]
            for c in range(min(len(cells), src_cols)):
                _write_cell(ws, r, c + 1 + SHIFT, cells[c][0], cells[c][1])
            seq = ws.cell(row=r, column=SEQ_COL + 1)
            seq.value = er                                     # ← 原文件序号 = 该行在原文件的 Excel 行号
        else:
            gen_rows += 1
            for c in range(total_cols):
                _write_cell(ws, r, c + 1, None, base_style)     # 新增父级行：序号格留空
        role = source_role(er)
        if role.is_point:
            previous = ws.cell(r, tgt_idx["kks"] + SHIFT + 1).value
            ws.cell(r, tgt_idx["kks"] + SHIFT + 1).value = None
            explanation = "指令/测点：设备KKS编码留空；" + role.reason
            explanation += f"；设备本体码 {role.body}" if role.body else "；设备归属待图纸确认"
            if previous not in (None, ""):
                cleared_points += 1
                explanation += f"；原最终编码 {previous} 已清空，原设计列保留"
            ws.cell(r, note_col + 1).value = explanation
        if node is None:
            continue                                     # 未判定行：只保留原内容 + 写说明
        values = {"device_name": node.name, "kks": node.kks,
                  "parent_kks": node.parent_kks, "parent_name": node.parent_name}
        oc = row_orig.get(er) if er else ""
        skip_kks = role.is_point or (bool(oc) and oc in not_device)
        clash = dup_rows.get(int(er) if er else -(plan_pos + 1))         # 本行要写的码与别的行撞了
        gated = []
        derived = []                                     # 列推导**取不到值**的原因说明
        for key, idx in targets:
            if key == "kks" and seed_by_row.get(er, {}).get("revise_status") == "unresolved":
                derived.append("修编证据不足或存在冲突；空白编码不填，待人工确认。" + seed_by_row[er].get("revise_note", ""))
                continue
            if key == "kks" and skip_kks:
                # 指令/测点行不编码（业务口径 2026-10-12：「只有设备才有 KKS 编码」）。
                # 这行**不是设备**，设备的码在设备本体行上；此处留空 + 修编说明讲清本体码。
                skipped_kks += 1
                continue
            if key == "kks" and clash:
                # 国标 3.1.1：标识在全厂范围内唯一 → 撞码的行**一律留空**，
                # 谁也不留下（两个不同设备被给了同一个原码时，引擎无从判断该给谁）。
                _v, _others = clash
                dup_note = (f"编码 {_v} 与原文件第 {'、'.join(str(x) if x > 0 else '新增父级' for x in _others)} 行"
                            f"撞成同一编码（GB/T 50549-2020 3.1.1：标识在全厂范围内唯一）"
                            f"→ 所有参与行的最终编码清空，待人工确认")
                if ws.cell(r, idx + 1 + SHIFT).value not in (None, ""):
                    cleared_duplicates += 1
                ws.cell(r, idx + 1 + SHIFT).value = None
                if dup_note not in derived:
                    derived.append(dup_note)
                dup_blocked += 1
                continue
            if key.startswith("@derive:"):            # 列规则表推导（系统 / 设备类…）
                _rule = key.split(":", 2)[1]
                val = CR.derive(_rule, node.kks or "")
                if not val:
                    # ⚠️ 只在**这一格确实是空的**时才出"留空原因"说明：
                    #    原表这一列若本来就有数据（铁律：不动），说"留空"是错的（实测踩过）。
                    if ws.cell(row=r, column=idx + 1 + SHIFT).value in (None, ""):
                        _why = CR.derive_note(_rule, node.kks or "")
                        if _why and _why not in derived:
                            derived.append(_why)         # 取不到 → 留空 + 说明原因
                    continue
            else:
                val = values.get(key) or ""
            if not val:
                continue                                 # 引擎没有结论 → 不动这一格
            # 🔴 **最后一道闸门**：不合国标结构的编码**绝不写进交付文件**。
            #    只在 revise 层设闸不够——源表里畸形的原码本身会成为树节点，
            #    回填时又被写出去（实测漏了 8 行：`016`、`0161BHE01ABC`、`00BE02QA03`…）。
            if key == "parent_kks":
                final_value = ws.cell(r, tgt_idx["kks"] + SHIFT + 1).value
                if final_value not in (None, "") and (not str(final_value).startswith(str(val)) or len(str(val)) >= len(str(final_value))):
                    derived.append(f"实际最终编码 {final_value} 与候选父级 {val} 不一致；父级不填，原码与最终码归属冲突待图纸确认")
                    continue
            if key in ("kks", "parent_kks") and not can_fill_code(val):
                assessment = assess(val)
                gated.append(f"{val} 未通过结构或约定检查，未写入：" + "；".join(assessment.errors + assessment.pending))
                continue
            cell = ws.cell(row=r, column=idx + 1 + SHIFT)
            if isinstance(cell, openpyxl.cell.cell.MergedCell):
                continue                                     # 合并区内非锚点格不可写（见 _write_cell）
            # 🔴 业务铁律（2026-10-10 确认）：**原表一开始有数据的格，绝对不回填**。
            #    引擎只负责"把空的补上"，不覆盖用户/设计院已有的任何数据。
            #    （此前一度按"覆盖"实现过，在竣工版上覆盖了 24.7 万格 —— 已撤销。）
            if cell.value not in (None, ""):
                cells_kept += 1
                continue
            cell.value = val
            cells_written += 1
        if gated:
            blocked_rows += 1
        note = notes.get(oc or "", "") if oc else ""
        if not note and node is not None and getattr(node, "design_code", ""):
            note = notes.get(node.design_code, "")
        # 把节点的 notes（如"与其他行编码重复，需人工确认"）也带出来 ——
        # 否则这类**唯一性/需人核**的警告在交付文件里**完全看不到**（踩过）。
        extra = "；".join(str(x) for x in (getattr(node, "notes", []) or []) if x) if node is not None else ""
        if extra:
            note = f"{note}；{extra}" if note else extra
        for _why in derived:
            note = f"{note}；{_why}" if note else _why
        if gated:
            g = "；".join(gated)
            note = f"{note}；{g}" if note else g
        if note:
            # 只留**真问题**：状态回显（「已符合国标」）不算说明；
            # 「待推断」这类**待办提示**在对应列**已有数据**时也没有意义
            # （竣工版实测：这两类制造了 17,652 条噪声，占说明总数的 19%）。
            _parts = [p.strip() for p in note.split("；") if p.strip()]
            _keep = []
            for _p in _parts:
                if _NOOP_NOTE.match(_p):
                    continue
                _col_key = _PROMPT_NOTE_COLUMN.get(_p)
                if _col_key is not None:
                    _j = tgt_idx.get(_col_key)
                    if _j is not None and ws.cell(row=r, column=_j + 1 + SHIFT).value not in (None, ""):
                        continue          # 该列已有数据 → 这条提示是噪声
                _keep.append(_p)
            note = "；".join(_keep)
            if not note:
                notes_emptied += 1        # 原本全是"状态回显/无意义提示" → 不写
        if note:
            ncell = ws.cell(row=r, column=note_col + 1)
            # 同样遵守铁律：**已有内容不覆盖**。该列是本工具新建的（新建时为空），
            # 但若用户表里本来就有同名列并填了内容，也必须保留。
            if ncell.value in (None, ""):
                ncell.value = note[:32767]
                notes_written += 1
    # 输出行数可能少于原数据行（理论上不会；保险起见清掉尾部残留）
    for extra in range(r + 1, last_row + 1):
        for c in range(1, total_cols + 1):
            ws.cell(row=extra, column=c).value = None

    from kks_gen.compliance import assess, unit_conflicts
    checks = []
    written_codes = {}
    if kks_idx is not None:
        for output_row in range(lead_last + 1, r + 1):
            value = ws.cell(output_row, kks_idx + 1 + SHIFT).value
            origin = ws.cell(output_row, SEQ_COL + 1).value
            seed = seed_by_row.get(origin, {})
            errors, waiting = [], []
            if value not in (None, ""):
                code = str(value)
                written_codes.setdefault(code, []).append(output_row)
                assessment = assess(code)
                errors.extend(assessment.errors)
                waiting.extend(assessment.pending)
                parent_idx = tgt_idx.get("parent_kks")
                parent_value = ws.cell(output_row, parent_idx + SHIFT + 1).value if parent_idx is not None else None
                if parent_value not in (None, "", "-1"):
                    parent_value = str(parent_value)
                    errors.extend("父级：" + message for message in assess(parent_value).errors)
                    if not code.startswith(parent_value) or len(parent_value) >= len(code):
                        errors.append(f"实际最终编码 {code} 与父级 {parent_value} 前缀或层级不一致")
                elif seed and seed.get("revise_status") == "unresolved":
                    waiting.append("源编码与已有最终编码存在证据冲突；未补父级，须核对图纸及工程约定")
                waiting.extend(unit_conflicts(code, seed.get("original_name", ""), seed.get("unit_label", "")))
                from revise import signal_body
                if signal_body(code) != code:
                    errors.append("本列为设备编码，已有指令/信号码须人工拆分；原有内容保留")
            elif seed and not source_role(origin).is_point and (seed.get("original_code") or "") not in not_device:
                waiting.append("设备编码尚未填写，不能认定本行通过审核")
            if errors or waiting:
                checks.append({"row": output_row, "source_row": origin, "errors": errors, "pending": waiting})
        for code, duplicate_rows in written_codes.items():
            if len(duplicate_rows) > 1:
                checks.append({"rows": duplicate_rows, "errors": [f"交付列编码 {code} 重复；原有值须人工处理"], "pending": []})
        for check in checks:
            for output_row in check.get("rows", [check.get("row")]):
                note_cell = ws.cell(output_row, note_col + 1)
                explanation = "；".join(check["errors"] + check["pending"])
                note_cell.value = (str(note_cell.value or "") + "；" + explanation).strip("；")[:32767]
    compliance = {"standard": "GB/T 50549-2020", "status": "blocked" if any(c["errors"] for c in checks) else "pending",
                  "issues": checks, "scope": "自动校验不代替图纸、工程约定及全厂范围唯一性确认"}
    row_map = {source_row: source_row for source_row in range(1, lead_last + 1)}
    row_map.update({source_row: lead_last + position + 1
                    for position, (_, source_row, _) in enumerate(plan) if source_row})
    from kks_gen.formula_refs import remap_formula
    from openpyxl.workbook.properties import CalcProperties
    formulas_remapped = 0
    for source_row, cells in snap.items():
        output_row = row_map.get(source_row)
        if output_row is None:
            continue
        for source_column, (original_value, _) in enumerate(cells):
            if isinstance(original_value, str) and original_value.startswith("="):
                target = ws.cell(output_row, source_column + SHIFT + 1)
                if target.value != original_value:
                    continue  # authorized clearing of a final-code formula
                target.value = remap_formula(original_value, row_map, SHIFT, ws.title, ws.title)
                formulas_remapped += 1
    for other_sheet in wb.worksheets:
        if other_sheet is ws:
            continue
        for row in other_sheet.iter_rows():
            for cell in row:
                if cell.data_type == "f":
                    translated = remap_formula(cell.value, row_map, SHIFT, ws.title, other_sheet.title)
                    if translated != cell.value:
                        cell.value = translated
                        formulas_remapped += 1
    wb.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)
    compliance["formulas_remapped"] = formulas_remapped
    for merged in source_merges:
        mapped = [row_map[row] for row in range(merged.min_row, merged.max_row + 1)]
        if mapped[-1] - mapped[0] + 1 == len(mapped):
            ws.merge_cells(start_row=mapped[0], end_row=mapped[-1],
                           start_column=merged.min_col + SHIFT, end_column=merged.max_col + SHIFT)
        else:
            compliance["issues"].append({"pending": [f"原合并区{merged}跨新增父级行，已展开以保留内容，须核对版式"], "errors": []})
    # Final enforcement includes preserved/orphan values and all sheets with
    # a final device-KKS column. Original/design columns are evidence only.
    final_occurrences = {}
    final_slots, point_slots = [], []
    note_headers = {}
    for target_sheet in wb.worksheets:
        candidates = list(target_sheet.iter_rows(min_row=1, max_row=min(HEADER_SCAN_ROWS, target_sheet.max_row), values_only=True))
        selected = [(i, resolve_columns([str(x or "") for x in row])) for i, row in enumerate(candidates)]
        selected = [(i, columns) for i, columns in selected if columns.get("kks") is not None]
        if not selected:
            continue
        header_row, columns = max(selected, key=lambda item: sum(v is not None for v in item[1].values()))
        final_column = columns["kks"] + 1
        note_column = next((i + 1 for i, value in enumerate(candidates[header_row]) if str(value or "") == "修编说明"), target_sheet.max_column + 1)
        note_headers[target_sheet.title] = header_row + 1
        def annotate(row, message):
            target_sheet.cell(header_row + 1, note_column).value = "修编说明"
            cell = target_sheet.cell(row, note_column)
            cell.value = (str(cell.value or "") + "；" + message).strip("；")[:32767]
        for output_row in range(header_row + 2, target_sheet.max_row + 1):
            def value(key):
                index = columns.get(key)
                return target_sheet.cell(output_row, index + 1).value if index is not None else ""
            if is_table_instruction(value("kks"), value("device_name"), value("design_code")):
                continue
            role = classify_row(value("design_code") or value("kks"), value("original_name") or value("device_name"), value("device_type"))
            if not role.is_point:
                role = classify_row(value("kks"), value("original_name") or value("device_name"), value("device_type"))
            cell = target_sheet.cell(output_row, final_column)
            final_slots.append(cell)
            if role.is_point:
                point_slots.append(cell)
                if cell.value not in (None, ""):
                    annotate(output_row, f"指令/测点，原最终编码 {cell.value} 清空；{role.reason}")
                    cell.value = None
                    cleared_points += 1
                continue
            if cell.value not in (None, ""):
                canonical = str(cell.value).strip().upper()
                final_occurrences.setdefault(canonical, []).append((target_sheet, output_row, final_column, note_column))
    for code, locations in final_occurrences.items():
        if len(locations) < 2:
            continue
        references = "、".join(f"{sheet.title}!{row}" for sheet, row, _, _ in locations)
        for sheet, row, column, note_column in locations:
            sheet.cell(row, column).value = None
            sheet.cell(note_headers[sheet.title], note_column).value = "修编说明"
            note = sheet.cell(row, note_column)
            note.value = (str(note.value or "") + f"；重复编码 {code} 清空，冲突位置 {references}；GB/T 50549-2020 3.1.1，全厂范围内唯一").strip("；")[:32767]
            cleared_duplicates += 1
    actual = Counter(str(cell.value).strip().upper() for cell in final_slots if cell.value not in (None, ""))
    duplicate_count = sum(count > 1 for count in actual.values())
    point_count = sum(cell.value not in (None, "") for cell in point_slots)
    if duplicate_count or point_count:
        raise RuntimeError("交付硬约束检查失败：仍有重复编码或指令/测点编码")
    compliance["delivery_policy"] = {"verified": True, "signal_point_codes": point_count, "duplicate_codes": duplicate_count,
                                      "cleared_points": cleared_points, "cleared_duplicates": cleared_duplicates}
    compliance["delivery_policy"]["final_codes"] = sum(actual.values())
    compliance["delivery_policy"]["point_rows"] = len(point_slots)
    # Explain each original row's blank by the delivered value and role, not
    # just revision status (a point row may be syntactically valid).
    blank_reasons = Counter()
    for output_row in range(lead_last + 1, r + 1):
        origin = ws.cell(output_row, SEQ_COL + 1).value
        if not origin:
            continue
        final_value = ws.cell(output_row, kks_idx + SHIFT + 1).value
        if final_value not in (None, ""):
            continue
        if source_role(origin).is_point:
            reason = "指令/测点按要求留空"
        elif origin in dup_rows:
            reason = "重复编码待核"
        elif seed_by_row.get(origin, {}).get("revise_status") == "unresolved":
            reason = "修编证据冲突或结构无法修复"
        else:
            reason = "分类/编号/工程约定待核"
        blank_reasons[reason] += 1
    compliance["blank_reasons"] = dict(blank_reasons)
    wb.save(out_path)
    wb.close()
    return {
        "compliance": compliance,
        "filled_rows": len([p for p in plan if p[0] == "orig"]),
        "generated_rows": gen_rows,
        "cells_written": cells_written,
        "cells_kept": cells_kept,
        "notes_written": notes_written,
        "notes_emptied": notes_emptied,
        "orphan_rows": len([p for p in plan if p[0] == "orig" and p[2] is None]),
        "blocked_rows": blocked_rows,
        "skipped_kks_rows": skipped_kks,
        "dup_blocked_rows": dup_blocked,
        "target_columns": [(k.split(":", 2)[-1] if k.startswith("@derive:") else head_name(k))
                           for k, _ in targets],
    }


def main(argv=None, *, progress_callback=None) -> int:
    token = _RUN_PROGRESS.set(progress_callback)
    try:
        return _main(list(sys.argv[1:] if argv is None else argv))
    finally:
        _RUN_PROGRESS.reset(token)


def _main(argv) -> int:
    src = pathlib.Path(argv[0])
    out_dir = pathlib.Path(argv[1] if len(argv) > 1 else "outputs/step1")
    use_ai = "--no-ai" not in argv
    ai_limit = 0
    for arg in argv[2:]:
        if arg.startswith("--ai-limit="):
            ai_limit = int(arg.split("=", 1)[1])
    out_dir.mkdir(parents=True, exist_ok=True)

    seeds, layout = load(src)
    _emit({"stage": "loaded", "percent": 8,
           "message": f"已读取 {len(seeds):,} 行原始数据",
           "hint": "正在定位列并准备生成"})
    # ---------- 修编：把设计编码换成符合国标的编码（依据 GB/T 50549-2020）----------
    # 先修编、再用修编后的编码建树 —— 这样上级链也是基于正确编码算出来的。
    from revise import infer_unit_labels, revise_code
    inferred_units = infer_unit_labels(seeds)
    for seed, unit in zip(seeds, inferred_units):
        if unit and not seed.get("unit_label"):
            seed["unit_label"] = unit
            seed["unit_inferred"] = True
    rev_stat = Counter()
    final_reviewed = 0
    for seed in seeds:
        rev = revise_code(seed["design_code"], unit_label=seed.get("unit_label", ""),
                          description=seed.get("original_name", ""),
                          device_type=seed.get("device_type", ""),
                          system_name=seed.get("system_name", ""))
        seed["original_code"] = seed["design_code"]
        if seed.get("code_is_final"):
            # ---- 原表已有最终编码（无设计编码）→ **只审查，绝不改写** ----
            # 用户铁律：不许改原表已有的编码。这里连"修编"都不做：
            # 把审查结论落进修编说明，编码原样保留。
            # 审查**通过的行不写说明**——否则 47,984 行全是"已审查"的噪声（实测）。
            seed["revised_code"] = seed["design_code"]
            seed["revise_status"] = rev.status
            seed["revise_needs_review"] = rev.needs_review
            seed["revise_level"] = rev.level
            note = ""
            if rev.status == "not_device":
                # 这类行按口径本不该编码，但铁律是"原表已有编码绝不改写" → 原样保留，
                # 说明里讲清两者，别让审核人以为引擎漏了（旧口径会写"不填编码"，与实际矛盾）。
                note = ("本行是指令/测点，按口径不单独编码（只有设备才有 KKS 编码）；"
                        f"其所服务的设备本体码为 {rev.device_body}。"
                        f"原表已有最终编码按最新指令/测点规则清空，并保留原值及原因")
            elif rev.status != "ok":
                note = "原表已有最终编码，只审查不改写（本行无设计编码）"
                if rev.status == "revised" and rev.revised:
                    note += f"；按国标应修编为 {rev.revised}，需人工确认"
                if rev.note:
                    note += f"；{rev.note}"
            seed["revise_note"] = note
            final_reviewed += 1
            if rev.status != "ok":
                rev_stat[rev.status] += 1
            continue
        seed["revised_code"] = rev.revised
        seed["revise_status"] = rev.status
        seed["revise_note"] = rev.note
        # needs_review 是**正交**标记：含「待核」（码未收录/号段未定义）或未判定
        seed["revise_needs_review"] = rev.needs_review
        seed["revise_level"] = rev.level
        # 指令/测点行（status=not_device）**不写编码**（业务口径 2026-10-12：
        # 「只有设备才有 KKS 编码」）。但**建树与父级仍按设备本体码**走 ——
        # 否则这些行会脱离所属设备，「父级KKS编码」列也会空掉。
        seed["revise_not_device"] = rev.status == "not_device"
        rev_stat[rev.status] += 1
        if rev.status == "revised" and rev.revised:
            seed["design_code"] = rev.revised
        elif rev.status == "not_device" and rev.device_body:
            seed["design_code"] = rev.device_body

    # ---------- 修编后的**唯一性**校验（国标 3.1.1：标识在全厂范围内唯一）----------
    # 🔴 修编**不得破坏唯一性**。实测：`01DCS01XA105`（发电机实发功率）与
    #    `01DCS01YA105`（机组功率给定）描述都是"功率"，T₂ 都被改成 `BJ` → 撞成同一个码。
    #    `XA`/`YA` 在国标里是**信号标识**的 X₁（信号源/信号应用），不该被抹平成同一个设备码。
    #    撞码的行**一律不改动**：保留原码、写明冲突，交人工 —— 绝不静默产出重复码。
    #
    # 这一遍只处理"**不同原码**收敛成同一个交付码"（引擎自己造成的撞码）——对这类
    # **保留原码**比留空更好：唯一性和信息都保住了。
    # ⚠️ 它**不是**唯一性保障的全部：判据是 `revised_code`，早于回填时的结构闸门与
    #    「指令/测点不写」两条规则；**源表自带**的同码重复（同码两行）它也发现不了
    #    —— 那些由 `fill_original_table` 的**写出级撞码闸门**兜底（判实际写出去的值）。
    by_final: dict[str, list[dict]] = {}
    for s in seeds:
        fin = s.get("revised_code") or ""
        if fin:
            by_final.setdefault(fin, []).append(s)
    unit_conflicts = 0
    for fin, group in by_final.items():
        origs = sorted({(s.get("original_code") or "") for s in group})
        if len(origs) <= 1:
            continue
        for s in group:
            mine = s.get("original_code") or ""
            others = "、".join(o for o in origs if o != mine)
            s["revised_code"] = mine                 # 保留原码，不改
            # ⚠️ 必须**同时**回退 design_code：树是用它建的（`build_tree(seeds)`），
            #    只改 revised_code 的话，树里仍是撞码的修编值 → 交付文件照样重复（踩过）。
            s["design_code"] = mine
            s["revise_status"] = "unresolved"
            s["revise_level"] = "error"
            s["revise_needs_review"] = True
            s["revise_note"] = (f"修编后与 {others} 撞成同一编码 {fin}"
                               f"（国标 3.1.1 要求标识全厂唯一）→ 保留原码，待人工确认")
            unit_conflicts += 1
    if unit_conflicts:
        print(f"[注意] 唯一性校验：{unit_conflicts} 行因修编后会与其它行撞码，已保留原码并标记待人工"
              f"（国标 3.1.1）")
    rev_stat = Counter(s.get("revise_status") or "unresolved" for s in seeds)
    revise_needs_review = sum(1 for s in seeds if s.get("revise_needs_review"))
    revise_pending_only = sum(1 for s in seeds
                              if s.get("revise_status") == "ok" and s.get("revise_level") == "pending")
    stats_extra = {"revise_ok": rev_stat["ok"], "revise_changed": rev_stat["revised"],
                   "revise_unresolved": rev_stat["unresolved"],
                   "revise_not_device": rev_stat["not_device"],
                   "revise_needs_review": revise_needs_review,
                   "revise_pending_only": revise_pending_only,
                   "unit_inferred": sum(1 for s in seeds if s.get("unit_inferred"))}
    print(f"编码修编：结构检查通过 {rev_stat['ok']:,} ｜ 已修编 {rev_stat['revised']:,} ｜ "
          f"指令/测点不编码 {rev_stat['not_device']:,} ｜ 待人工 {rev_stat['unresolved']:,} ｜ "
          f"待核（未收录等，原码保留）{revise_pending_only:,} ｜ "
          f"机组号上下文推断补全 {stats_extra['unit_inferred']:,}")
    if final_reviewed:
        print(f"其中原表已有最终编码、只审查不改写的行: {final_reviewed:,}")
    _emit({"stage": "revise", "percent": 18,
           "message": f"编码修编：改动 {rev_stat['revised']:,} 行，待人工 {rev_stat['unresolved']:,} 行",
           "hint": f"指令/测点行不编码 {rev_stat['not_device']:,} 行"
                   f"（只有设备才有 KKS 编码）；需人工核对合计 {revise_needs_review:,} 行"
                   f"（其中「码未收录」类待核 {revise_pending_only:,} 行，原码保留）"})

    result = build_tree(seeds, progress_callback=_emit)
    nodes = result["nodes"]
    stats = result["stats"]
    stats.update(stats_extra)
    _emit({"stage": "tree", "percent": 25,
           "message": f"设备树构建完成：{len(nodes):,} 个节点",
           "hint": f"推导出父级节点 {sum(1 for n in nodes.values() if not n.design_code):,} 个"})

    # ---------- AI 名称层（可选，只处理规则给不出的行）----------
    ai_result: dict = {}
    if use_ai:
        from ai_names import resolve_names

        targets = _collect_ai_rows(nodes)
        if ai_limit:
            targets = targets[:ai_limit]
        if targets:
            print(f"\n=== AI 名称层：{len(targets)} 行需要语义生成 ===")
            _emit({"stage": "ai_start", "percent": 30,
                   "message": f"AI 名称生成：{len(targets):,} 行需要语义处理",
                   "hint": "规则已给对名称的行不送模型"})
            t0 = time.time()
            ai_result = resolve_names(
                targets,
                progress_callback=_RUN_PROGRESS.get() or PROGRESS_AI_HOOK,
            )
            print(f"  耗时 {time.time()-t0:.1f}s | status={ai_result.get('status')} "
                  f"returned={ai_result.get('returned')} valid={ai_result.get('valid')} "
                  f"rejected={ai_result.get('rejected')} errors={len(ai_result.get('errors') or [])}")
            by_row = {str(t["row"]): t for t in targets}
            applied = skipped_has_name = 0
            for row_no, item in (ai_result.get("items") or {}).items():
                # ⚠️ 行号一律按字符串匹配：AI 返回的 row 可能是 int 也可能是 str，
                # 混用会让绝大部分结果被静默丢弃（实测 186 条只应用了 22 条）。
                target = by_row.get(str(row_no).strip())
                if not target or not item.get("valid") or not item.get("name"):
                    continue
                node = nodes.get(target["kks"])
                if node is None:
                    continue
                # ⚠️ 只在"规则给不出名称"时用 AI，绝不覆盖规则已经给出的名称。
                # 实测教训：让 AI 覆盖规则结果时，名称准确率反而从 20,047 掉到 19,194
                # （改对 830、改错 1,683）——规则在它擅长的场景比模型稳。
                if str(node.name or "").strip() and node.source == "原始":
                    skipped_has_name += 1
                    continue
                node.name = item["name"]
                node.source = "AI生成"
                node.confidence = min(item.get("confidence", 0.7), 0.9)
                node.need_human = True          # AI 产出仍标记需人核
                node.notes.append(f"AI生成：{item.get('reason', '')[:80]}")
                applied += 1
            stats["ai_names_applied"] = applied
            stats["ai_names_rejected"] = ai_result.get("rejected", 0)
            stats["ai_names_skipped_rule_ok"] = skipped_has_name
            print(f"  采纳 {applied} 条；退回 {ai_result.get('rejected', 0)} 条；"
                  f"跳过（规则已给出名称，不让 AI 覆盖）{skipped_has_name} 条")
            _emit({"stage": "ai_done", "percent": 75,
                   "message": f"AI 名称生成完成：采纳 {applied:,} 条",
                   "hint": f"退回 {ai_result.get('rejected', 0)} 条；跳过 {skipped_has_name:,} 条"})
    # 修编说明：按"最终采用的编码"索引，供回填时写入「修编说明」列
    revise_notes: dict[str, str] = {}
    revise_pending: dict[str, bool] = {}
    # 指令/测点行（原KKS编码）→ 回填时**不写「KKS编码（修编后）」列**（业务口径 2026-10-12）
    revise_not_device: set[str] = set()
    for seed in seeds:
        final_code = seed.get("design_code") or ""
        note = seed.get("revise_note") or ""
        key = seed.get("original_code") or final_code
        # 「待核」行的 status 是 ok（原码保留），但说明必须带出去，否则用户看不到"码未收录"这类线索
        if note and note not in ("已符合国标", "已通过自动结构校验"):
            revise_notes[key] = note
        if seed.get("revise_needs_review"):
            revise_pending[key] = True
        if seed.get("revise_not_device") and seed.get("original_code"):
            revise_not_device.add(seed["original_code"])

    ordered = order_nodes(nodes)
    _emit({"stage": "write", "percent": 90, "message": "正在回填原表", "hint": ""})

    # 原表为主要交付物；按最新要求另附待核设备编码清单。
    # 节点全量仍写进JSON供后台统计，不作为下载项。
    filled: dict = {}
    try:
        layout["notes"] = revise_notes
        layout["pending"] = revise_pending
        layout["not_device"] = revise_not_device
        original_to_node = {s["original_code"]: s["design_code"]
                            for s in seeds if s.get("original_code")}
        filled = fill_original_table(src, out_dir / f"{src.stem}_已填写.xlsx", nodes,
                                     layout, original_to_node,
                                     ordered=_tree_order(nodes), seeds=seeds)
        print(f"回填原表（原序+父级就近）: 原表行 {filled.get('filled_rows', 0):,} ｜ "
              f"新增父级行 {filled.get('generated_rows', 0):,} ｜ "
              f"写入 {filled.get('cells_written', 0):,} 格 / 保留原有 {filled.get('cells_kept', 0):,} 格 ｜ "
              f"指令/测点行不编码 {filled.get('skipped_kks_rows', 0):,} 行 ｜ "
              f"撞码留空 {filled.get('dup_blocked_rows', 0):,} 行 ｜ "
              f"修编说明 {filled.get('notes_written', 0):,} 条（另有 {filled.get('notes_emptied', 0):,} 行无需说明）｜ 未判定行 {filled.get('orphan_rows', 0):,} ｜ "
              f"结构闸门拦下 {filled.get('blocked_rows', 0):,} 行"
              f" → {src.stem}_已填写.xlsx")
    except Exception as exc:
        raise RuntimeError(f"回填交付文件失败：{exc}") from exc

    from kks_gen.pending_list import PENDING_FILENAME, write_pending_list
    filled["pending_list"] = write_pending_list(out_dir / f"{src.stem}_已填写.xlsx", out_dir / PENDING_FILENAME, resolve_columns)

    # ---------- 落盘 JSON（仅供后台统计 / 复现，**不在下载列表**）----------
    payload = {
        "stats": stats,
        "fill": filled,
        "nodes": [
            {"kks": nodes[k].kks, "name": nodes[k].name,
             "parent_kks": nodes[k].parent_kks, "parent_name": nodes[k].parent_name,
             "source": nodes[k].source, "confidence": nodes[k].confidence,
             "need_human": nodes[k].need_human, "level": nodes[k].level,
             "profession": nodes[k].profession, "design_code": nodes[k].design_code,
             "original_name": nodes[k].original_name, "drawing": nodes[k].drawing,
             "src_rows": nodes[k].src_rows, "notes": nodes[k].notes}
            for k in ordered
        ],
    }
    json_path = out_dir / "第一步_生成结果.json"
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")

    print(f"\n已写出:\n  {src.stem}_已填写.xlsx")
    print(f"\n生成 {len(ordered)} 行")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
