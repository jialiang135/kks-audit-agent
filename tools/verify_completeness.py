# -*- coding: utf-8 -*-
"""核验目标硬约束：每个"有设计编码的行"必须产出完整四列（不得有空）。

判据：从竣工版取全部有设计编码的行，逐个去生成结果里核对四个字段。
同时统计整体完整性与可追溯性。
"""
from __future__ import annotations

import json
import pathlib
import re
import sys
from collections import Counter

import openpyxl

src = pathlib.Path(sys.argv[1])
gen = json.loads(pathlib.Path(sys.argv[2]).read_text(encoding="utf-8"))
nodes = {n["kks"]: n for n in gen["nodes"]}

ws = openpyxl.load_workbook(src, read_only=True, data_only=True)["KKS竣工版"]
rows = ws.iter_rows(values_only=True)
hdr = [str(v or "").strip() for v in next(rows)]
col = {h: i for i, h in enumerate(hdr) if h}

# ⚠️ 不能先 .upper() 再匹配字面量：那样中文会被当成"字母"放过
#    （实测踩过：`7号炉辅汽至锅炉71-54m区域产品伴热用汽管` 被判成合法编码）
KKS_SHAPE = re.compile(r"^[0-9A-Za-z][0-9A-Za-z.\-]*$")

seeds = []
malformed = []
for idx, r in enumerate(rows, start=2):
    design = str(r[col["设计编码"]] or "").strip()
    if not design:
        continue
    # ⚠️ 竣工版有数据缺陷：个别行的"设计编码"列装的是名称文本
    #    （实测行 25143 = '7号炉辅汽至锅炉71-54m区域产品伴热用汽管'，真实码在 KKS编码 列）
    if not KKS_SHAPE.match(design.upper()):
        malformed.append((idx, design, str(r[col["KKS编码"]] or "").strip()))
        continue
    seeds.append((idx, design, str(r[col["原名称"]] or "").strip()))

print(f"有设计编码的行: {len(seeds)}")
if malformed:
    print(f"[注意] 设计编码列填了非编码内容（已剔除，属竣工版数据缺陷）: {len(malformed)} 行")
    for idx, bad, real in malformed[:5]:
        print(f"    行{idx}: 设计编码列={bad[:40]!r} → 真实 KKS编码={real!r}")

missing_kks = missing_name = missing_parent = 0
no_parent_expected = 0
samples: dict[str, list] = {}
for idx, design, old_name in seeds:
    node = nodes.get(design)
    if node is None:
        missing_kks += 1
        if len(samples.setdefault("缺 KKS 编码", [])) < 5:
            samples["缺 KKS 编码"].append((idx, design, old_name))
        continue
    if not str(node.get("name") or "").strip():
        missing_name += 1
        if len(samples.setdefault("缺设备名称", [])) < 5:
            samples["缺设备名称"].append((idx, design, old_name))
    pk = str(node.get("parent_kks") or "").strip()
    if not pk:
        # 2 位机组码没有上级是正常的
        if len(design.split(".")[0]) <= 2:
            no_parent_expected += 1
        else:
            missing_parent += 1
            if len(samples.setdefault("缺上级KKS", [])) < 5:
                samples["缺上级KKS"].append((idx, design, old_name))
    elif pk not in nodes:
        missing_parent += 1
        if len(samples.setdefault("上级不在树内", [])) < 5:
            samples["上级不在树内"].append((idx, design, pk))

print(f"\n=== 四列完整性核验 ===")
print(f"  缺 KKS 编码      : {missing_kks}")
print(f"  缺设备名称       : {missing_name}")
print(f"  缺上级 KKS 编码  : {missing_parent}（另有 {no_parent_expected} 行是机组码，本就无上级）")
print(f"  合计缺口         : {missing_kks + missing_name + missing_parent}")

for key, items in samples.items():
    print(f"\n=== {key} 样例 ===")
    for item in items:
        print("  ", item)

print(f"\n=== 可追溯性 ===")
with_rows = sum(1 for n in gen["nodes"] if n.get("src_rows"))
print(f"  带回溯源 Excel 行号的节点: {with_rows} / {len(gen['nodes'])}")
print(f"  带来源标记的节点        : {sum(1 for n in gen['nodes'] if n.get('source'))} / {len(gen['nodes'])}")
print(f"  需人核标记              : {sum(1 for n in gen['nodes'] if n.get('need_human'))}")

print(f"\n=== 结论 ===")
ok = (missing_kks == 0 and missing_name == 0 and missing_parent == 0)
print("  " + ("[通过] 每个有设计编码的行都产出了完整数据" if ok else "[失败] 存在缺口，见上"))
sys.exit(0 if ok else 1)
