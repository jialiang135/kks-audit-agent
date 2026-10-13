# -*- coding: utf-8 -*-
"""第一步智能体 · 交付准确率报告。

对竣工版（已按 07=7号机组 校正机组号）逐列比对，输出：
  ① 四列准确率（KKS编码 / 上级KKS编码 / 上级设备名称 / 设备名称）
  ② 名称差异归因（完全一致 / 前缀差异 / 标点差异 / 语义重写）
  ③ 完整性（每个有设计编码的行是否产出四列）
  ④ 来源与置信度分布、需人核规模
  ⑤ 已知缺口与人工确认指引

用法:
  python -B tools/report_step1_accuracy.py <竣工版.xlsx> <生成结果.json> [输出.md]
"""
from __future__ import annotations

import json
import pathlib
import re
import sys
from collections import Counter

import openpyxl

# 机组号口径：07=7号机组、08=8号机组（不是 1/2 号）
UNIT_LONG = re.compile(r"([1-9]|1[0-6])\s*号\s*机组")
UNIT_ANY = re.compile(r"([1-9]|1[0-6])\s*号\s*(机组|机)")
UNIT_BY_PREFIX = {"07": "7", "08": "8", "01": "1", "02": "2", "03": "3",
                  "04": "4", "05": "5", "06": "6"}
KKS_SHAPE = re.compile(r"^[0-9A-Za-z][0-9A-Za-z.\-]*$")
# 设备编号里的"1号"不是机组指代，不能改
KEEP_EQ = ("主变", "高厂变", "启备变")


def fix_unit(name: str, kks: str) -> str:
    """把名称里的机组指代纠正为编码前缀对应的机组号（保留"1号主变"这类设备编号）。"""
    if not name or len(kks) < 2:
        return name
    want = UNIT_BY_PREFIX.get(kks[:2], "")
    if not want:
        return name
    out, pos = [], 0
    for m in UNIT_ANY.finditer(name):
        tail = name[m.end():m.end() + 3]
        if any(kw.startswith(tail[:len(kw)]) for kw in KEEP_EQ):
            out.append(name[pos:m.end()])
        else:
            out.append(name[pos:m.start()] + f"{want}号{m.group(2)}")
        pos = m.end()
    out.append(name[pos:])
    return "".join(out)


def norm(name: str, kks: str) -> str:
    """参考侧归一：**把错误机组号纠正到编码前缀**（07→7号、08→8号），并折叠多余空格。

    ⚠️ 方向不能反：竣工版里两种写法并存（`1号机NCS测控屏` 是错的、`7号机…` 是对的），
    我方按 07=7号机组 生成。若把参考的 `7号机` 反折成 `1号机`，就会把改对的行判成改错
    （实测因此把名称准确率从 39.4% 压到 33.3%）。
    """
    text = UNIT_LONG.sub(r"\1号机", name or "")
    text = re.sub(r"\s+", " ", text).strip()
    return fix_unit(text, kks)


def strip_punct(text: str) -> str:
    return re.sub(r"[\s（）()·、,，\-/]", "", text or "")


def main() -> int:
    src = pathlib.Path(sys.argv[1])
    gen_path = pathlib.Path(sys.argv[2])
    out_path = pathlib.Path(sys.argv[3] if len(sys.argv) > 3 else "outputs/step1/准确率报告.md")

    gen = json.loads(gen_path.read_text(encoding="utf-8"))
    nodes = {n["kks"]: n for n in gen["nodes"]}

    ws = openpyxl.load_workbook(src, read_only=True, data_only=True)["KKS竣工版"]
    rows = ws.iter_rows(values_only=True)
    hdr = [str(v or "").strip() for v in next(rows)]
    col = {h: i for i, h in enumerate(hdr) if h}

    def cell(r, key):
        i = col.get(key)
        return str(r[i] or "").strip() if i is not None and i < len(r) else ""

    reference, seeds, malformed = {}, [], []
    for idx, r in enumerate(rows, start=2):
        kks = cell(r, "KKS编码")
        if kks:
            reference[kks] = {
                "name": norm(cell(r, "设备名称"), kks),
                "parent_kks": cell(r, "上级KKS编码"),
                "parent_name": norm(cell(r, "上级设备名称"), kks),
            }
        design = cell(r, "设计编码")
        if not design:
            continue
        if not KKS_SHAPE.match(design):
            malformed.append((idx, design, kks))
            continue
        seeds.append((idx, design))

    # ---------- ① 列级准确率 ----------
    cmp_stat = Counter()
    name_bucket = Counter()
    name_diff_samples: list[tuple] = []
    parent_name_examples: list[tuple] = []
    for _idx, design in seeds:
        node = nodes.get(design)
        ref = reference.get(design)
        if node is None:
            cmp_stat["我方未生成"] += 1
            continue
        if ref is None:
            cmp_stat["竣工版无此码"] += 1
            continue
        cmp_stat["可比对"] += 1
        cmp_stat["KKS编码一致" if node["kks"] == design else "KKS编码不一致"] += 1
        cmp_stat["上级KKS一致" if (node["parent_kks"] or "") == (ref["parent_kks"] or "") else "上级KKS不一致"] += 1
        cmp_stat["上级名称一致" if node["parent_name"] == ref["parent_name"] else "上级名称不一致"] += 1
        got, want = node["name"], ref["name"]
        if got == want:
            cmp_stat["设备名称一致"] += 1
        else:
            cmp_stat["设备名称不一致"] += 1
            if got and want and (got in want or want in got):
                name_bucket["前缀/后缀差异"] += 1
            elif strip_punct(got) == strip_punct(want):
                name_bucket["仅标点空格差异"] += 1
            else:
                name_bucket["语义差异（需 AI 或人工）"] += 1
                if len(name_diff_samples) < 10:
                    name_diff_samples.append((design, got, want))

    comparable = max(1, cmp_stat["可比对"])

    # ---------- ③ 完整性 ----------
    miss_kks = miss_name = miss_parent = 0
    for design in (d for _i, d in seeds):
        node = nodes.get(design)
        if node is None:
            miss_kks += 1
            continue
        if not str(node.get("name") or "").strip():
            miss_name += 1
        if not str(node.get("parent_kks") or "").strip() and len(design.split(".")[0]) > 2:
            miss_parent += 1

    # ---------- ④ 来源与置信度 ----------
    source_dist = Counter(n["source"] for n in gen["nodes"])
    conf_dist = Counter(round(float(n.get("confidence", 0)), 1) for n in gen["nodes"])
    need_human = sum(1 for n in gen["nodes"] if n.get("need_human"))
    ai_meta = gen.get("stats", {})

    def pct(a: int, b: int) -> str:
        return f"{a}/{b} = {a / max(1, b) * 100:.1f}%"

    lines = [
        "# 第一步智能体 · 交付准确率报告",
        "",
        f"- 输入：`{src.name}`（有设计编码的行 {len(seeds):,}）",
        f"- 生成：{len(gen['nodes']):,} 个节点（含父级链）",
        f"- 参考：竣工版全表 {len(reference):,} 行，机组号已按 **07=7号机组 / 08=8号机组** 校正",
        "",
        "## 一、逐列准确率（与竣工版比对）",
        "",
        "| 列 | 准确率 |",
        "|---|---|",
        f"| KKS 编码 | {pct(cmp_stat['KKS编码一致'], comparable)} |",
        f"| 上级 KKS 编码 | {pct(cmp_stat['上级KKS一致'], comparable)} |",
        f"| 上级设备名称 | {pct(cmp_stat['上级名称一致'], comparable)} |",
        f"| 设备名称 | {pct(cmp_stat['设备名称一致'], comparable)} |",
        "",
        f"可比对行数：{comparable:,}（我方未生成 {cmp_stat['我方未生成']}，竣工版无此码 {cmp_stat['竣工版无此码']}）",
        "",
        "## 二、设备名称差异归因",
        "",
    ]
    for key in ("前缀/后缀差异", "仅标点空格差异", "语义差异（需 AI 或人工）"):
        lines.append(f"- {key}：**{name_bucket.get(key, 0):,}**")
    if name_diff_samples:
        lines += ["", "差异样例：", "", "| 编码 | 我方生成 | 竣工版（已校正机组号） |", "|---|---|---|"]
        for kks, got, want in name_diff_samples:
            lines.append(f"| `{kks}` | {got[:34] or '（空）'} | {want[:34]} |")

    lines += [
        "",
        "## 三、完整性（目标硬约束）",
        "",
        "> 要求：每个有设计编码的行都必须产出完整四列。",
        "",
        "| 项 | 缺口数 |",
        "|---|---|",
        f"| 缺 KKS 编码 | **{miss_kks}** |",
        f"| 缺设备名称 | **{miss_name:,}** |",
        f"| 缺上级 KKS 编码 | **{miss_parent:,}** |",
        "",
    ]
    if malformed:
        lines += [
            f"⚠️ 竣工版有 **{len(malformed)} 行**「设计编码」列填的不是编码（已从输入剔除）：",
            "",
            "| 行号 | 设计编码列内容 | 真实 KKS 编码 |",
            "|---|---|---|",
        ]
        for idx, bad, real in malformed[:12]:
            lines.append(f"| {idx} | `{bad[:40]}` | `{real}` |")
        lines.append("")

    lines += [
        "## 四、生成来源与可信度",
        "",
        "| 来源 | 数量 |",
        "|---|---|",
    ]
    for key, value in source_dist.most_common():
        lines.append(f"| {key} | {value:,} |")
    lines += [
        "",
        f"**需人工确认：{need_human:,} / {len(gen['nodes']):,}**（{(need_human / max(1, len(gen['nodes'])) * 100):.1f}%）",
        "",
        "置信度分布：" + "，".join(f"{k} → {v:,}" for k, v in sorted(conf_dist.items())),
        "",
    ]
    if ai_meta.get("ai_names_applied"):
        lines += [
            f"AI 名称层：采纳 **{ai_meta.get('ai_names_applied'):,}** 条，退回 {ai_meta.get('ai_names_rejected', 0):,} 条",
            "",
        ]

    lines += [
        "## 五、已知边界（如实声明）",
        "",
        "1. **部件级不做自动生成**：实测两条路（兄弟参照、设备类型推断）精确率均为 0%，",
        "   改为输出「应编未编核查清单」（225 项）交人工。详见 `kks_gen/README.md`。",
        "2. **名称规则上限 43.9%**：超过部分必须靠 AI 语义层或人工，规则堆不出来。",
        "3. **D 类整段无编码**（土建/消防/信息/输煤段 11,220 行）不在生成本体内，",
        "   因输入里就没有设计编码；按「只对有设计编码的行操作」的口径，转清单交业务确认。",
        "4. 以上准确率**对竣工版而言被低估**：竣工版自身带有机组号错误（如 `1号机NCS测控屏`），",
        "   我方按 07=7号机组 生成，比对的参考侧已做校正，但个别行仍可能出现「改对了却被判错」。",
    ]

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    print(f"已写出: {out_path}")
    print(f"  可比对 {comparable:,} 行")
    print(f"  KKS 编码一致 {cmp_stat['KKS编码一致']:,}")
    print(f"  上级 KKS 一致 {cmp_stat['上级KKS一致']:,}")
    print(f"  设备名称一致 {cmp_stat['设备名称一致']:,}")
    print(f"  缺口：KKS {miss_kks} / 名称 {miss_name:,} / 上级 {miss_parent:,}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
