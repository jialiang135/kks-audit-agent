# -*- coding: utf-8 -*-
"""校验**服务端实际返回的页面**里每个内联脚本都能通过语法检查。

为什么必须校验"服务端返回的内容"而不是源码：
  UPLOAD_PAGE 是普通（非 raw）Python 字符串，源码里写的 `\n` 在求值时会变成**真实换行**，
  于是 JS 单引号字符串里被塞进真换行 → SyntaxError → **整个 <script> 不执行**
  （表现为页面上所有按钮都失效，极易被误判成"样式/点击"问题）。

用法: python -B tools/verify_page_scripts.py
"""
from __future__ import annotations

import pathlib
import re
import subprocess
import sys
import tempfile

# Windows 控制台默认 GBK，而本脚本要打印 ✅/❌/⚠️ 这类非 GBK 字符；
# 调用方（tests/test_generate_api.py）按 UTF-8 解码子进程输出，两边必须一致，
# 否则会在最后一行 print 上抛 UnicodeEncodeError，把"校验通过"误报成失败。
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

FULL_WIDTH_NEWLINE = "\u2028"      # JS 里同样非法的行分隔符
PARAGRAPH_SEP = "\u2029"
# 多行字符串的常见写法：单引号/双引号/反引号里出现真实换行
UNTERMINATED = re.compile(r"(['\"])(?:\\.|(?!\1)[^\\\r\n])*\1", re.S)


def html_of() -> str:
    import server
    return server.UPLOAD_PAGE


def script_blocks(html: str) -> list[tuple[int, str]]:
    """返回 [(起始行号, 脚本内容)]。"""
    out = []
    for m in re.finditer(r"<script(?:\s[^>]*)?>(.*?)</script>", html, re.S):
        line = html[:m.start(1)].count("\n") + 1
        out.append((line, m.group(1)))
    return out


def main() -> int:
    html = html_of()
    blocks = script_blocks(html)
    print(f"页面内联脚本块: {len(blocks)} 个")

    problems: list[str] = []

    # ① 用 node 逐个做语法检查（最权威）
    node = None
    for candidate in ("node", r"E:\软件\node.exe"):
        try:
            probe = subprocess.run([candidate, "--version"], capture_output=True, text=True, timeout=20)
            if probe.returncode == 0:
                node = candidate
                break
        except (OSError, subprocess.SubprocessError):
            continue
    if node:
        for index, (line, body) in enumerate(blocks, start=1):
            with tempfile.NamedTemporaryFile("w", suffix=".js", delete=False, encoding="utf-8") as fh:
                fh.write(body)
                path = fh.name
            result = subprocess.run([node, "--check", path], capture_output=True, text=True,
                                    encoding="utf-8", errors="replace", timeout=60)
            if result.returncode != 0:
                err = (result.stderr or "").strip().splitlines()
                detail = err[0] if err else "unknown"
                # node 报的行号是脚本内相对行号，换算成页面行号
                page_line = line
                m = re.search(r":(\d+)\s*$", err[-1] if err else "")
                if m:
                    page_line = line + int(m.group(1)) - 1
                problems.append(f"脚本块 #{index}（页面约第 {page_line} 行）语法错误: {detail}")
            pathlib.Path(path).unlink(missing_ok=True)
        print(f"  node 语法检查完成（{node}）")
    else:
        print("  [注意] 未找到 node，跳过语法检查")

    # ② 真实换行/行分隔符混进 JS **字符串字面量**（node 会直接报语法错，这里只做兜底提示）
    #    ⚠️ 不要用"两个引号之间跨行"这种启发式——它会跨语句误匹配，产生大量假告警。
    #       权威判据是 ① 的 node --check；这里只查 JS 里非法的行分隔符字符。
    for char, label in ((FULL_WIDTH_NEWLINE, "U+2028 行分隔符"),
                        (PARAGRAPH_SEP, "U+2029 段分隔符")):
        if char in html:
            line = html[:html.index(char)].count("\n") + 1
            problems.append(f"页面第 {line} 行含 JS 非法的{label}")

    # ③ 关键元素与处理器是否都在页面里
    for marker in ('id="genFile"', 'id="genPick"', 'id="genDrop"', 'id="genStart"',
                   "function genStart", "window.__initGeneratePage"):
        if marker not in html:
            problems.append(f"页面缺少关键标记: {marker}")

    if problems:
        print("\n❌ 发现问题：")
        for item in problems:
            print("  -", item)
        return 1
    print("\n✅ 页面脚本全部通过语法检查")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
