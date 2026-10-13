# -*- coding: utf-8 -*-
"""第一步智能体 · 规则层跑通 + 用竣工版做验证。

用法:
  python -B tools/gen_step1_rules.py <竣工版.xlsx> <输出目录>

做的事:
  1) 从竣工版取"有设计编码"的行当种子（模拟图纸摘抄的原始数据）
  2) 跑规则生成完整设备树
  3) 用竣工版（校正机组号后）当参考答案，逐列比对：
     KKS编码 / 上级KKS编码 / 设备名称 / 上级设备名称
  4) 输出 stats 与差异样例，便于定位规则缺口
"""
from __future__ import annotations

import json
import pathlib
import re
import sys
from collections import Counter

import openpyxl

ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "kks_gen"))

from tree_rules import (  # noqa: E402
    build_tree, fix_unit_reference, normalize_name, order_nodes,
)

UNIT_REF = re.compile(r"([1-9]|1[0-6])\s*号\s*(机组|机)")


def load_completion(path: pathlib.Path):
    """读竣工版：返回 (种子行, 参考全表)。参考表按 07=7号机组/08=8号机组 校正机组号。"""
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True)["KKS竣工版"]
    rows = list(ws.iter_rows(values_only=True))
    header = [str(v or "").strip() for v in rows[0]]
    col = {h: i for i, h in enumerate(header)}

    def cell(r, key):
        i = col.get(key)
        return str(r[i] or "").strip() if i is not None and i < len(r) else ""

    seeds, reference = [], {}
    for idx, r in enumerate(rows[1:], start=2):
        kks = cell(r, "KKS编码")
        design = cell(r, "设计编码")
        name = cell(r, "设备名称")
        # 参考答案：机组号按编码前缀校正（07→7号机、08→8号机）
        fixed_name, _ = fix_unit_reference(normalize_name(name), kks)
        entry = {
            "excel_row": idx,
            "kks": kks,
            "name": fixed_name,
            "raw_name": name,
            "parent_kks": cell(r, "上级KKS编码"),
            "parent_name": cell(r, "上级设备名称"),
            "profession": cell(r, "专业"),
            "design": design,
            "original_name": cell(r, "原名称"),
            "drawing": cell(r, "图纸名称或图号"),
        }
        if kks:
            reference[kks] = entry
        if design:
            seeds.append({
                "excel_row": idx,
                "design_code": design,
                "original_name": cell(r, "原名称"),
                "profession": cell(r, "专业"),
                "drawing": cell(r, "图纸名称或图号"),
            })
    return seeds, reference


def main() -> int:
    src = pathlib.Path(sys.argv[1])
    out_dir = pathlib.Path(sys.argv[2] if len(sys.argv) > 2 else "outputs/step1_rules")
    out_dir.mkdir(parents=True, exist_ok=True)

    seeds, reference = load_completion(src)
    print(f"种子行（有设计编码）: {len(seeds)}")
    print(f"竣工版参考行        : {len(reference)}")

    result = build_tree(seeds)
    nodes = result["nodes"]
    stats = result["stats"]
    print("\n=== 规则层生成统计 ===")
    for k, v in stats.items():
        print(f"  {k}: {v}")
    print(f"  生成节点总数: {len(nodes)}")
    print(f"  其中父节点（无设计编码）: {sum(1 for n in nodes.values() if not n.design_code)}")

    # ---- 与参考答案比对 ----
    hit_kks = sum(1 for k in nodes if k in reference)
    print("\n=== 与竣工版比对（按 KKS 编码对齐）===")
    print(f"  生成 {len(nodes)} 个节点，其中能在竣工版里找到同名编码的: {hit_kks}")
    print(f"  竣工版有但规则层没生成的: {len(set(reference) - set(nodes))}")

    same_parent = diff_parent = 0
    parent_examples = []
    for kks, node in nodes.items():
        ref = reference.get(kks)
        if not ref:
            continue
        if node.parent_kks == ref["parent_kks"]:
            same_parent += 1
        else:
            diff_parent += 1
            if len(parent_examples) < 12:
                parent_examples.append((kks, node.parent_kks, ref["parent_kks"], node.level))
    total = same_parent + diff_parent
    if total:
        print(f"\n  上级KKS编码：完全一致 {same_parent}/{total} = {same_parent/total*100:.1f}%")
        print("  不一致样例（编码 / 我生成 / 竣工版 / 层级）:")
        for k, mine, ref_p, lvl in parent_examples:
            print(f"    {k:<20} 我={mine or '(空)':<14} 参考={ref_p or '(空)':<14} {lvl}")

    same_name = diff_name = 0
    name_examples = []
    for kks, node in nodes.items():
        ref = reference.get(kks)
        if not ref or not node.design_code:      # 父节点名称还没生成，暂不比对
            continue
        if node.name == ref["name"]:
            same_name += 1
        else:
            diff_name += 1
            if len(name_examples) < 15:
                name_examples.append((kks, node.name, ref["name"], ref["raw_name"]))
    total_name = same_name + diff_name
    if total_name:
        print(f"\n  设备名称（仅种子行）：完全一致 {same_name}/{total_name} = {same_name/total_name*100:.1f}%")
        print("  不一致样例（编码 / 我生成 / 参考(已校正) / 竣工版原文）:")
        for k, mine, ref_n, raw in name_examples:
            print(f"    {k:<16}\n      我  ={mine[:46]!r}\n      参考={ref_n[:46]!r}\n      原文={raw[:46]!r}")

    # ---- 落盘 ----
    ordered = order_nodes(nodes)
    payload = {
        "stats": stats,
        "nodes": [
            {
                "kks": nodes[k].kks, "name": nodes[k].name,
                "parent_kks": nodes[k].parent_kks, "parent_name": nodes[k].parent_name,
                "source": nodes[k].source, "confidence": nodes[k].confidence,
                "need_human": nodes[k].need_human, "level": nodes[k].level,
                "profession": nodes[k].profession, "design_code": nodes[k].design_code,
                "original_name": nodes[k].original_name, "drawing": nodes[k].drawing,
                "src_rows": nodes[k].src_rows, "notes": nodes[k].notes,
            }
            for k in ordered
        ],
    }
    out_json = out_dir / "step1_rules.json"
    out_json.write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"\n已写出: {out_json}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
