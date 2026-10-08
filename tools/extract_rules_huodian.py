# -*- coding: utf-8 -*-
"""把《电力火电厂标识系统（KKS）编码导则-0813》（Q/WL 企业标准，.doc）抽成结构化规则。

为什么要单独抽这一份
--------------------
项目里 KKS 审核的**默认标准**就是这份火电导则（run_audit.STANDARD_DEFAULT = "huodian"），
但它此前只以"章节号"的形式被引用（如"火电导则 6.2 设备索引"），没有可机读的码表。
结果是我们只能用 GB/T 50549-2020 的附录 A/B 去判分类码是否合法——而这两套标准的
分类码体系并不一致（详见 config/kks_rules.json 的 standard_divergences），
会造成误报和误导。

抽取产物
--------
  kks-audit/references/rules_huodian.json   结构化（章节 / 条款 / 表 / 三张索引码表）
  kks-audit/references/rules_huodian.md     人读版
（与 rules_gb50549.* / rules_zhongmei.* 一样属标准正文复制，不进版本库，只本地生成）

依赖
----
.doc 是 OLE2 旧格式，Python 标准库读不了，用 Word COM 逐段落导出带样式名的文本。
导出结果缓存到 _tmp_extract/huodian_list.txt，已存在则直接复用（可加 --refresh 强制重抽）。
"""
from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

PROJECT = Path(__file__).resolve().parent.parent
DOC = PROJECT / "国标" / "电力火电厂标识系统（KKS）编码导则-0813.doc"
WORK = PROJECT / "_tmp_extract"
LIST_TXT = WORK / "huodian_list.txt"
PS1 = WORK / "extract_doc.ps1"

OUT_JSON = PROJECT / "kks-audit" / "references" / "rules_huodian.json"
OUT_MD = PROJECT / "kks-audit" / "references" / "rules_huodian.md"

# 段落行形如：[样式名][列表号] 正文
LINE_RE = re.compile(r"^\[(?P<style>[^\]]*)\]\[(?P<list>[^\]]*)\]\s*(?P<text>.*)$")
CODE_RE = re.compile(r"^([A-Z]{1,3})$")
CLAUSE_RE = re.compile(r"^(\d+(?:\.\d+){1,3})\s+(\S.*)$")
# Word 的自动编号放在「列表号」列（如 [6.1　]），正文段落里没有编号，
# 所以标题要同时看列表号和正文两种写法。
INDEX_HEAD_RE = re.compile(r"^(\d+\.\d+)?\s*[　\s]?\s*(\S+?)索引$")
GROUP_HEAD_RE = re.compile(r"^([A-Z])\s*[\t　]?\s*(\S.*)$")
CAPTION_RE = re.compile(r"^表\s*(\d+)\s*[：:]\s*(\S.*)$")

PS_SCRIPT = r"""param([string]$src,[string]$out)
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $doc = $word.Documents.Open($src, $false, $true)
  $count = $doc.Paragraphs.Count
  $sb = New-Object System.Text.StringBuilder
  foreach ($p in $doc.Paragraphs) {
    $t = $p.Range.Text
    if ($t -and $t.Trim().Length -gt 0) {
      $sn = ''
      try { $sn = [string]$p.Style.NameLocal } catch {}
      $ls = ''
      try { $ls = [string]$p.Range.ListFormat.ListString } catch {}
      [void]$sb.AppendLine('[' + $sn + '][' + $ls + '] ' + $t.Trim())
    }
  }
  [System.IO.File]::WriteAllText($out, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))
  $doc.Close($false)
  Write-Output ("OK paragraphs=" + $count + " -> " + $out)
} catch {
  Write-Output ("ERR " + $_.Exception.Message)
} finally {
  $word.Quit()
}
"""


def ensure_source(refresh: bool = False) -> Path:
    """确认 .doc 已导出为带样式名的逐段落文本。"""
    if LIST_TXT.is_file() and not refresh and LIST_TXT.stat().st_size > 10000:
        return LIST_TXT
    if not DOC.is_file():
        raise SystemExit(f"找不到标准原文：{DOC}")
    WORK.mkdir(parents=True, exist_ok=True)
    PS1.write_text(PS_SCRIPT, encoding="utf-8-sig")
    result = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(PS1),
         "-src", str(DOC), "-out", str(LIST_TXT)],
        capture_output=True, text=True,
    )
    print((result.stdout or "").strip() or (result.stderr or "").strip())
    if not LIST_TXT.is_file():
        raise SystemExit("Word COM 导出失败，请确认本机装有 Word")
    return LIST_TXT


def load_lines(path: Path) -> list[tuple[str, str, str]]:
    """返回 [(样式名, 列表号, 正文), ...]，保持原文顺序。

    Word 的自动编号（如章节号 6.1）不在正文里，而在 ListFormat.ListString，
    导出脚本把它放在第二个方括号中，解析标题时必须一起看。
    """
    out: list[tuple[str, str, str]] = []
    for raw in path.read_text(encoding="utf-8", errors="replace").splitlines():
        m = LINE_RE.match(raw.strip())
        if m and m.group("text").strip():
            out.append((m.group("style"), m.group("list").strip(), m.group("text").strip()))
    return out


def find_index_sections(lines: list[tuple[str, str, str]]) -> dict[str, tuple[int, int]]:
    """按正文里的「6.x xxx索引」一级条标题定位三段索引，返回 {名称: (起, 止)}。"""
    marks: list[tuple[str, int]] = []
    for i, (style, list_no, text) in enumerate(lines):
        if style != "一级条标题":
            continue
        label = (list_no + text).replace("　", " ").strip()
        m = INDEX_HEAD_RE.match(label)
        if m:
            marks.append((m.group(2), i))
    if not marks:
        raise SystemExit("没有找到 6.x 索引标题，段落样式可能变了")
    bounds: dict[str, tuple[int, int]] = {}
    for idx, (name, start) in enumerate(marks):
        end = marks[idx + 1][1] if idx + 1 < len(marks) else len(lines)
        bounds[name] = (start + 1, end)
    return bounds


def parse_index(lines: list[tuple[str, str, str]], start: int, end: int) -> dict:
    """解析索引表：按「主组标题 → 码 → 名称 → 说明」的段落顺序还原码表。"""
    groups: list[dict] = []
    codes: dict[str, str] = {}
    group: dict | None = None
    i = start
    while i < end:
        style, _list_no, text = lines[i]
        cap = CAPTION_RE.match(text)
        if style == "题注" and cap:
            group = {"table": cap.group(1), "title": cap.group(2).strip(), "rows": []}
            groups.append(group)
            i += 1
            continue
        if style in ("标题 3", "标题 4") and group is None:
            group = {"table": "", "title": text.strip(), "rows": []}
            groups.append(group)
            i += 1
            continue
        m = CODE_RE.match(text)
        if m and group is not None:
            code = m.group(1)
            name, desc = "", ""
            j = i + 1
            while j < end:
                nstyle, _nlist, ntext = lines[j]
                if CODE_RE.match(ntext) or nstyle in ("题注", "标题 3", "标题 4", "一级条标题"):
                    break
                if nstyle == "正文" and not name:
                    name = ntext
                elif nstyle == "正文" and name and not desc:
                    desc = ntext
                j += 1
            if code not in codes:
                codes[code] = name
            group["rows"].append({"code": code, "name": name, "desc": desc})
            i = j
            continue
        i += 1
    return {"groups": groups, "codes": codes}


def parse_clauses(lines: list[tuple[str, str, str]], stop: int) -> list[dict]:
    """抽 1~5 章的条款。

    Word 把条款号放在自动编号里（如 [5.1.1　] 正文…），所以条款号取列表号、
    正文取文本；编号缺失时回退到文本里自带编号的写法。
    """
    clauses: list[dict] = []
    section = ""
    for i in range(stop):
        style, list_no, text = lines[i]
        if style in ("章标题", "一级条标题"):
            section = (list_no + text).replace("　", " ").strip()
            continue
        if style in ("二级条标题", "三级条标题"):
            no = list_no.replace("　", " ").strip() or text.split()[0]
            body = text if list_no.strip() else text[len(no):].strip()
            clauses.append({"no": no, "section": section, "text": body})
            continue
        if style != "正文":
            continue
        m = CLAUSE_RE.match(text)
        if m and len(m.group(2)) > 3:
            clauses.append({"no": m.group(1), "section": section, "text": m.group(2)})
        elif clauses and len(text) > 4:
            clauses[-1]["text"] += text
    return clauses


def parse() -> dict:
    lines = load_lines(ensure_source())
    bounds = find_index_sections(lines)
    first_index = min(start for start, _ in bounds.values())
    sections = {name: parse_index(lines, a, b) for name, (a, b) in bounds.items()}
    return {
        "source": DOC.name,
        "standard": "Q/WL 电力有限公司企业标准《火电厂标识系统（KKS）编码导则》",
        "clauses": parse_clauses(lines, first_index),
        "index_sections": sections,
    }


def render(data: dict) -> str:
    L: list[str] = [f"# {data['standard']}", "", f"来源文件：`{data['source']}`", ""]
    L.append(f"- 条款 {len(data['clauses'])} 条")
    for name, sec in data["index_sections"].items():
        L.append(f"- {name}：{len(sec['codes'])} 个码 / {len(sec['groups'])} 张表")
    L.append("")
    L.append("## 编码索引（码 → 名称）")
    L.append("")
    for name, sec in data["index_sections"].items():
        L.append(f"### {name}")
        L.append("")
        L.append("| 码 | 名称 | 说明 |")
        L.append("| --- | --- | --- |")
        for code, label in sec["codes"].items():
            desc = ""
            for g in sec["groups"]:
                for row in g["rows"]:
                    if row["code"] == code:
                        desc = row["desc"]
                        break
                if desc:
                    break
            L.append(f"| {code} | {label} | {desc} |")
        L.append("")
    L.append("## 条款")
    L.append("")
    for c in data["clauses"]:
        L.append(f"- **{c['no']}**（{c['section']}）{c['text']}")
    L.append("")
    return "\n".join(L)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--refresh", action="store_true", help="强制重新用 Word COM 导出")
    args = ap.parse_args()
    if args.refresh:
        ensure_source(refresh=True)
    data = parse()
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    OUT_MD.write_text(render(data), encoding="utf-8")
    print(f"条款 {len(data['clauses'])} 条")
    for name, sec in data["index_sections"].items():
        print(f"  {name}: {len(sec['codes'])} 码 / {len(sec['groups'])} 表")
    print(f"-> {OUT_MD.name} {OUT_MD.stat().st_size} B ; {OUT_JSON.name} {OUT_JSON.stat().st_size} B")
    return 0


if __name__ == "__main__":
    sys.exit(main())
