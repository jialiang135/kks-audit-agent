# -*- coding: utf-8 -*-
"""一次性修复：消除《项目编码规则总表》里的规则编号撞号（只动指定小节块）。

撞号来源：本次续写新增的小节沿用了别处已有的编号段
（§10.3 与 §10.4 都落到 10.13~10.15；§13 的"13.2b"与 §13.2/§13.3 都落到 13.9~13.15）。
本表把编号当**引用词汇**用（正文写"见 10.13"），撞号会让引用有歧义，故必须唯一。

做法：
  1. 把插在 §13.1 之后的"13.2b"小节**移到 §13 末尾**（顺序变连贯）；
  2. 对这几个小节块**按块内顺序**重编号到互不重叠的号段；
  3. 其它节一律不动（正文只引用小节号 §10.1/§10.2/§10.3，不受影响）。

用法: python -B tools/fix_rule_numbering.py [--write]
"""
from __future__ import annotations

import pathlib
import re
import sys

ROOT = pathlib.Path(__file__).resolve().parents[1]
DOC = ROOT / "kks_gen" / "项目编码规则总表.md"

MOVE_HEAD = "### 13.2b"
MOVE_NEW_HEAD = "### 13.6 代码落点与已支持的编码形态（**如实记录代码行为**）"

# (小节标题前缀, 规则编号前缀) —— 前缀互不重叠
BLOCKS: list[tuple[str, str]] = [
    ("### 10.3 修复记录", "10.2"),
    ("### 10.4 ", "10.3"),
    ("### 13.2 名称层", "13.2"),
    ("### 13.3 ", "13.3"),
    ("### 13.4 ", "13.4"),
    ("### 13.5 ", "13.5"),
    ("### 13.6 ", "13.6"),
]

ROW_RE = re.compile(r"^(\|\s*\*{0,2})\d+\.\d+(\*{0,2}\s*\|)")


def move_block(lines: list[str]) -> list[str]:
    """把 MOVE_HEAD 小节整块移到「## 编制说明」之前。"""
    start = next((i for i, l in enumerate(lines) if l.startswith(MOVE_HEAD)), None)
    if start is None:
        return lines
    end = start + 1
    while end < len(lines) and not (lines[end].startswith("### ") or lines[end].startswith("## ")):
        end += 1
    block = lines[start:end]
    block[0] = MOVE_NEW_HEAD
    rest = lines[:start] + lines[end:]
    anchor = next((i for i, l in enumerate(rest) if l.startswith("## 编制说明")), len(rest))
    # 补一个空行分隔
    return rest[:anchor] + block + [""] + rest[anchor:]


def renumber(lines: list[str]) -> tuple[list[str], int]:
    block = counter = None, 0
    block = None
    counter = 0
    changed = 0
    for i, ln in enumerate(lines):
        if ln.startswith("### ") or ln.startswith("## "):
            block, counter = None, 0
            for head, prefix in BLOCKS:
                if ln.startswith(head):
                    block, counter = prefix, 0
                    break
        if block is None:
            continue
        m = ROW_RE.match(ln)
        if not m:
            continue
        counter += 1
        new = f"{m.group(1)}{block}{counter}{m.group(2)}" + ln[m.end():]
        if new != ln:
            lines[i] = new
            changed += 1
    return lines, changed


def main() -> None:
    write = "--write" in sys.argv
    text = DOC.read_text(encoding="utf-8")
    lines = move_block(text.splitlines())
    lines, changed = renumber(lines)
    out = "\n".join(lines) + "\n"
    ids = re.findall(r"^\|\s*\*{0,2}(\d+\.\d+)\*{0,2}\s*\|", out, re.M)
    dups = sorted({x for x in ids if ids.count(x) > 1})
    print(f"重编号 {changed} 行；规则行合计 {len(ids)}；重复编号 {dups or '无'}")
    if write:
        DOC.write_text(out, encoding="utf-8")
        print(f"已写入 {DOC.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
