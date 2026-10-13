# -*- coding: utf-8 -*-
"""国标合规自检：对**生成/修编后的交付表**逐行核验是否满足 GB/T 50549-2020。

用途：给"生成编码 / 审查编码都符合国标"这句话一个**可复现的凭据**，而不是口头保证。
它**独立重算**（只读交付文件，复用与引擎同一套规则层 `kks_gen/identifiers.py`），
因此能抓出"引擎自己写错"的情况。

核验项与依据（逐条对应 `kks_gen/项目编码规则总表.md`）：

| 项 | 依据 | 判定 |
|---|---|---|
| 1 结构可解析（段位/位数/字符集） | 图4.4.1、4.4.2~4.4.5、总表 §2/§4 | 判错 |
| 2 系统分类码须命中附录A | 4.4.2、附录A（总表 `4-07`） | 3 位未命中 → **待核**（§10.13 未收录不判错）|
| 3 设备分类码须命中附录B | 4.4.3、附录B（总表 `4-09`） | 未命中 → **待核** |
| 4 核电码 J/K 须排除 | 中煤细则 T25（总表 §10.1） | 判错 |
| 5 测点 T₃T₄T₅ 号段（701~899 锁定） | 编制说明 5.2.7 表5（总表 §8.3） | 判错 |
| 6 全厂码须落表4.3.1-2 号段 | 表4.3.1-2（总表 §3） | 判错 |
| 7 **标识全厂唯一** | **3.1.1**（总表 §1.1 规则 1.5，L1 强制） | 判错 |
| 8 上级编码须是本级的直接父级 | 图4.4.1 层级（总表 §2 / README §1） | 判错 |
| 9 禁用字母 I/O（段位化，不整串扫） | 图4.4.1 注1 等 5 处图注（总表 §5.2） | 判错 |

用法：
  python -B tools/verify_gb_compliance.py <已填写.xlsx>
  python -B tools/verify_gb_compliance.py <已填写.xlsx> --json out.json
"""
from __future__ import annotations

import json
import pathlib
import re
import sys
from collections import Counter, defaultdict

ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "kks_gen"))

import openpyxl  # noqa: E402
import identifiers as ids  # noqa: E402

# 直接父级：位长 → 上级位长（README §1，竣工版实测）
PARENT_LEN = {2: None, 3: 2, 5: 3, 6: 5, 7: 5, 9: 7, 12: 7, 13: 7, 14: 7}


def core(code: str) -> str:
    return str(code or "").strip().split("-", 1)[0].split(".", 1)[0]


def check_row(code: str, parent: str, seen: dict[str, int]) -> list[str]:
    """核验一行编码，返回问题列表（空 = 通过）。"""
    from kks_gen.compliance import assess
    checked = assess(code)
    canonical = code.strip().upper()
    seen[canonical] = seen.get(canonical, 0) + 1
    bad = list(checked.errors) + ["待核：" + message for message in checked.pending]
    if parent and parent != "-1":
        parent_result = assess(parent)
        bad.extend("上级：" + message for message in parent_result.errors)
        if not code.startswith(parent) or len(parent) >= len(code):
            bad.append("父子编码前缀或层级不一致；须核对功能树及工程约定")
    return bad



def main() -> int:
    src = pathlib.Path(sys.argv[1])
    out_json = None
    if "--json" in sys.argv:
        out_json = pathlib.Path(sys.argv[sys.argv.index("--json") + 1])

    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    # 表头行：前 30 行里同时含「KKS」与「名称」的那行
    hdr_i = next((i for i, r in enumerate(rows[:30])
                  if any("KKS" in str(c or "") for c in r)
                  and any("名称" in str(c or "") for c in r)), None)
    if hdr_i is None:
        print("找不到表头行"); return 1
    hdr = [str(c or "").strip() for c in rows[hdr_i]]

    # 定位列：**先按最具体的列名找**，再逐级放宽。
    # ⚠️ 不能"按表头顺序取第一个含 KKS 的列"——那会匹配到 `原KKS编码`（用户的原始编码，
    #    本来就是非标的，正是我们要修的对象），核验结果会完全错（踩过）。
    def col(keys):
        for key in keys:                       # 先最具体
            for i, h in enumerate(hdr):
                if key in h:
                    return i
        return None
    ki = col(("KKS编码（修编后", "KKS码（修编后", "修编后"))
    if ki is None:
        ki = next((i for i, h in enumerate(hdr) if "KKS" in h.upper() and not any(word in h for word in ("原", "旧", "父", "上级", "设计"))), None)
    pi = col(("父级KKS编码", "上级KKS编码", "父级编码", "上级编码"))
    print(f"核验文件: {src.name}")
    if ki is None:
        print("未找到最终编码列"); return 1
    print(f"编码列=「{hdr[ki]}」  上级列={'「'+hdr[pi]+'」' if pi is not None else '（无）'}")
    print(f"数据行: {len(rows) - hdr_i - 1:,}")

    seen: dict[str, int] = {}
    problems = Counter()
    samples: dict[str, list[str]] = defaultdict(list)
    from kks_gen.row_policy import is_table_instruction
    old_i = next((i for i, h in enumerate(hdr) if "KKS" in h and any(x in h for x in ("原", "设计"))), None)
    name_i = next((i for i, h in enumerate(hdr) if "设备名称" in h and "原" not in h), None)
    checked = blank = 0
    for r in rows[hdr_i + 1:]:
        code = str(r[ki] or "") if ki < len(r) else ""
        if is_table_instruction(code, r[name_i] if name_i is not None else "", r[old_i] if old_i is not None else ""):
            continue
        if not code:
            blank += 1
            continue
        checked += 1
        parent = str(r[pi] or "") if (pi is not None and pi < len(r)) else ""
        for msg in check_row(code, parent, seen):
            problems[msg] += 1
            if len(samples[msg]) < 3:
                samples[msg].append(f"{code}  上级={parent or '（空）'}")
    dups = {k: v for k, v in seen.items() if v > 1}
    for k, v in dups.items():
        problems["⑦标识不唯一（同一编码出现多次）"] += v
        if len(samples["⑦标识不唯一（同一编码出现多次）"]) < 3:
            samples["⑦标识不唯一（同一编码出现多次）"].append(f"{k}  出现 {v} 次")

    hard = Counter({k: v for k, v in problems.items() if "待核" not in k})
    soft = Counter({k: v for k, v in problems.items() if "待核" in k})
    print(f"\n已核验 {checked:,} 行；留空（未判定/待人工）{blank:,} 行")
    print("  [范围] 自动规则校验，不等于全厂国标合规认证")
    print(f"  [违反] 自动结构及项目规则: {sum(hard.values()):,}")
    for k, v in hard.most_common():
        print(f"       {v:>6,}  {k}")
        for s in samples[k]:
            print(f"              {s}")
    print(f"  [待核] 分类、层级及工程约定: {sum(soft.values()):,}")
    for k, v in soft.most_common():
        print(f"       {v:>6,}  {k}")

    if out_json:
        out_json.write_text(json.dumps(
            {"file": src.name, "checked": checked, "blank": blank,
             "hard": hard, "soft": soft, "dups": len(dups)},
            ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"\n报告已写出: {out_json}")
    return 2 if hard else (3 if soft or blank else 0)


if __name__ == "__main__":
    raise SystemExit(main())
