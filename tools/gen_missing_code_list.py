# -*- coding: utf-8 -*-
"""生成「应编未编 / 无设计编码」核查清单。

背景（实测，鄂州四期）：
  竣工版 113,413 行里，有设计编码的只有 65,429 行。剩下的 47,984 行分两类：
    · 可由输入推导      → 由 gen_step1 生成（父级链、部件、树结构必需节点）
    · 推导不出来的      → 本清单：交人工确认是"图纸漏摘"还是"应编未编"

D 类（12位设备级、设计编码列没有）22,646 行的三分法：
  P1 树结构必需（有下级挂着）            536 行  → 生成
  P2 支路支撑（上级与同类兄弟在种子集）  4,693 行  → 判为漏摘，生成
  P3 整段无码（上级也不在种子集）       17,417 行  → 本清单

清单列（供业务直接核对）：
  Excel行号 / KKS编码 / 设备名称 / 上级KKS编码 / 上级设备名称 / 专业 /
  判定类别 / 依据 / 建议动作 / 图纸名称或图号
"""
from __future__ import annotations

import pathlib
import sys
from collections import Counter, defaultdict

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

HEADER = ["Excel行号", "KKS编码", "设备名称", "上级KKS编码", "上级设备名称",
          "专业", "判定类别", "判定依据", "建议动作", "图纸名称或图号"]


def load(path: pathlib.Path):
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True)["KKS竣工版"]
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(v or "").strip() for v in rows[0]]
    col = {h: i for i, h in enumerate(hdr) if h}

    def cell(r, key):
        i = col.get(key)
        return str(r[i] or "").strip() if i is not None and i < len(r) else ""

    return rows[1:], cell


def main() -> int:
    src = pathlib.Path(sys.argv[1])
    out = pathlib.Path(sys.argv[2] if len(sys.argv) > 2 else "outputs/step1/应编未编核查清单.xlsx")
    out.parent.mkdir(parents=True, exist_ok=True)

    body, cell = load(src)
    seed_bases = {cell(r, "设计编码").split(".")[0] for r in body if cell(r, "设计编码")}
    seed_prefix: set[str] = set()
    for d in seed_bases:
        for cut in range(1, len(d)):
            seed_prefix.add(d[:cut])

    parent_of_count: Counter[str] = Counter()
    for r in body:
        par = cell(r, "上级KKS编码")
        if par:
            parent_of_count[par] += 1

    seed_by_parent_cls: dict[tuple[str, str], int] = defaultdict(int)
    for d in seed_bases:
        if len(d) >= 9:
            seed_by_parent_cls[(d[:7], d[7:9])] += 1

    buckets: list[list] = []
    for idx, r in enumerate(body, start=2):
        kks = cell(r, "KKS编码")
        if not kks or cell(r, "设计编码"):
            continue
        base = kks.split(".")[0]
        if base in seed_prefix or any(base[:cut] in seed_bases for cut in range(len(base), 0, -1)):
            continue                                   # 可由输入推导 → 不列
        if len(base) != 12:
            continue                                   # 非设备级，另行处理
        parent7, cls = base[:7], base[7:9]
        children = parent_of_count.get(kks, 0)
        parent_in_seed = parent7 in seed_prefix or parent7 in seed_bases
        siblings = seed_by_parent_cls.get((parent7, cls), 0)
        if children:
            continue                                   # P1 树结构必需 → 生成，不列
        if parent_in_seed and siblings:
            continue                                   # P2 支路支撑 → 生成，不列
        if not parent_in_seed:
            kind = "P3 整段无编码"
            why = f"上级 {parent7} 本身不在有设计编码的集合内（整段无码）"
            action = "确认该专业是否本就无设计编码；若是，应由设计出码后重跑；不是则作为应编未编上报"
        else:
            kind = "P2b 有上级无同类"
            why = f"上级 {parent7} 在种子集内，但同父级同类型({cls})没有其他有码设备可对照"
            action = "人工核对图纸：该设备是否漏摘设计编码"
        buckets.append([
            idx, kks, cell(r, "设备名称"), cell(r, "上级KKS编码"), cell(r, "上级设备名称"),
            cell(r, "专业"), kind, why, action, cell(r, "图纸名称或图号"),
        ])

    wb = Workbook()
    ws = wb.active
    ws.title = "应编未编核查清单"
    ws.append(HEADER)
    for row in buckets:
        ws.append(row)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="FDECEA")
        c.alignment = Alignment(vertical="center")
    for col, width in zip("ABCDEFGHIJ", (10, 18, 38, 16, 30, 10, 18, 46, 46, 34)):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:J{max(1, ws.max_row)}"
    wb.save(out)

    print(f"清单行数: {len(buckets)}")
    print("判定类别分布:", dict(Counter(b[6] for b in buckets)))
    print("专业分布 TOP8:", dict(Counter(b[5] or "(空)" for b in buckets).most_common(8)))
    print(f"\n已写出: {out}")
    print("\n样例:")
    for b in buckets[:8]:
        print(f"  行{b[0]:>7} {b[1]:<18} {str(b[2])[:30]:<30} {b[6]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
