# -*- coding: utf-8 -*-
"""抽取国标 GB/T 50549-2020 附录 A/B 的分类码字典。

附录 A 系统分类索引 → 系统分类码 S1S2S3（以及主组 S1、组 S1S2）
附录 B 设备及产品分类索引 → 设备分类码 T1T2（以及主组 T1）

产物：
  config/gb_system_codes.json   {code: {"cn": 中文名, "en": 英文名, "note": 说明}}
  config/gb_device_codes.json   同上

这些都只来自国标原文表格，不做任何推断。
"""
from __future__ import annotations

import html
import json
import pathlib
import re
import sys

if hasattr(sys.stdout, "reconfigure"):
    # Windows 控制台默认 GBK；显式声明 UTF-8，避免中文路径/特殊字符打印时崩在最后一行
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else None
if SRC is None:
    # 自动定位国标 OCR md（中文路径经 argv 传参会被 ANSI 代码页破坏 → 推荐无参运行）
    _cands = sorted(pathlib.Path("国标").glob("*PaddleOCR*.md"))
    if not _cands:
        raise SystemExit("用法: python -B tools/extract_gb_dictionaries.py [国标OCR.md]"
                         "（未找到 国标/*PaddleOCR*.md）")
    SRC = _cands[0]
print(f"数据源: {SRC}")

OUT = pathlib.Path("config")
OUT.mkdir(parents=True, exist_ok=True)
lines = SRC.read_text(encoding="utf-8", errors="replace").splitlines()


def find_heading(letter: str) -> int:
    for i, l in enumerate(lines):
        s = l.strip()
        if re.match(rf"^#*\s*附录\s*{letter}\s", s) and "..." not in s:
            return i
    raise SystemExit(f"找不到附录 {letter}")


def clean_cell(raw: str) -> str:
    """单元格清洗：去标签 → **解 HTML 实体** → 还原字面转义序列 → 归一空白。

    🔴 `html.unescape` 不能省（总表 §12.1 明列此项）：OCR 的表格单元格里带 `&lt;` / `&gt;`，
    不解开就会把 `525kV ≤ Un &lt; 765kV 输电`（46 条）这类释义原样写进码表，
    进而污染自动生成的父级名称（如 `00ABE` → `全厂525kV ≤ Un &lt`）。
    源文里还有**字面的**两字符转义 `\\n`（不是换行），也要还原成空格。
    """
    text = re.sub(r"<[^>]+>", "", raw)
    text = html.unescape(text)
    text = text.replace("\\n", " ").replace("\\t", " ").replace("\\r", " ")
    return re.sub(r"\s+", " ", text).strip()


def parse_region(start: int, end: int) -> tuple[dict, dict]:
    """解析区间内所有表格。返回 (全部条目, 统计)。

    🔴 踩过的坑（必须按整块解析，不能逐行）：
       OCR 会把**一个表格行拆到多行**——例如表B.0.1 的 `A` 行，其开头 `<tr><td>A</td><td>…`
       与结尾 `…两个或更多目的或任务</td><td>—</td></tr>` 分处 L6597 / L6598。
       逐行跑 `<tr>(.*?)</tr>` 时两行都**不含完整的行**，于是 `A` 整行被**静默丢弃**。
       后果：附录B 只剩 217 条（1 位码缺 `A`、`B`，2 位码缺 `AF`、`AG`、`BA`），
       而 `B`/`BA` 正是测量回路（测点）的 T₁/T₂ 判据码 —— 漏了会直接判错点表。
       正确做法：先把区间合并成一整块再跑正则（`<tr>` 自带边界，跨行是安全的）。
       实测：合并后附录B = 222（1 位 26 + 2 位 196），与标准正文一致；附录A 无变化。
    """
    blob = "".join(lines[start:end])
    entries: dict[str, dict] = {}
    stats = {"tables": 0, "rows": 0, "skipped": 0, "by_len": {}}
    code_re = re.compile(r"^[A-Z][A-Z0-9]{0,2}\*?$")
    stats["tables"] = sum(1 for l in lines[start:end] if "<table" in l)
    for row in re.findall(r"<tr>(.*?)</tr>", blob, re.S):
        cells = [clean_cell(c) for c in re.findall(r"<td[^>]*>(.*?)</td>", row, re.S)]
        if len(cells) < 3:
            stats["skipped"] += 1
            continue
        code = cells[0].replace(" ", "").upper()
        if not code_re.match(code):
            stats["skipped"] += 1          # 表头行或说明行
            continue
        base = code.rstrip("*")
        cn = cells[2]
        note = cells[3] if len(cells) > 3 else ""
        # OCR 会把"说明"并进名称单元格（跨行行的常见现象，如 `A` 行：
        # `两个或更多目的或任务说明:此类仅适于无法识别其使用目的或任务的对象`）。
        # 拆开，避免"名称"里带一长串说明污染下游的匹配与展示。
        m = re.split(r"说明\s*[:：]", cn, maxsplit=1)
        if len(m) == 2:
            cn, extra = m[0].strip(), m[1].strip()
            note = (extra + ("；" + note if note and note not in ("—", "-") else "")).strip()
        entries[base] = {
            "cn": cn,
            "en": cells[1],
            "note": note,
            "starred": code.endswith("*"),
        }
        stats["rows"] += 1
        stats["by_len"][len(base)] = stats["by_len"].get(len(base), 0) + 1
    return entries, stats


a_start, b_start, c_start = find_heading("A"), find_heading("B"), find_heading("C")
sys_entries, sys_stats = parse_region(a_start, b_start)
dev_entries, dev_stats = parse_region(b_start, c_start)

(OUT / "gb_system_codes.json").write_text(
    json.dumps(sys_entries, ensure_ascii=False, indent=1), encoding="utf-8")
(OUT / "gb_device_codes.json").write_text(
    json.dumps(dev_entries, ensure_ascii=False, indent=1), encoding="utf-8")

print("=== 附录 A 系统分类码 ===")
print(f"  表 {sys_stats['tables']} 张，收录 {len(sys_entries)} 条，跳过 {sys_stats['skipped']} 行")
print(f"  按码长分布（1=主组S1, 2=组S1S2, 3=系统分类码S1S2S3）: {dict(sorted(sys_stats['by_len'].items()))}")

print("\n=== 附录 B 设备及产品分类码 ===")
print(f"  表 {dev_stats['tables']} 张，收录 {len(dev_entries)} 条，跳过 {dev_stats['skipped']} 行")
print(f"  按码长分布（1=主组T1, 2=设备分类码T1T2）: {dict(sorted(dev_stats['by_len'].items()))}")

print("\n=== 抽样（3 位系统分类码）===")
n = 0
for code, item in sys_entries.items():
    if len(code) == 3 and n < 8:
        print(f"  {code}  {item['cn']}")
        n += 1
print("\n=== 抽样（2 位设备分类码）===")
n = 0
for code, item in dev_entries.items():
    if len(code) == 2 and n < 10:
        print(f"  {code}  {item['cn']}")
        n += 1
