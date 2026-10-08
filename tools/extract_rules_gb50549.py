# -*- coding: utf-8 -*-
"""从《GB/T 50549-2020 电厂标识系统编码标准》OCR Markdown 抽取结构化规则。

来源文件由 PaddleOCR-VL 从 PDF 转出，特点：
  - 章节是标准 Markdown 标题（## 4 标识结构 / ### 4.4 功能面标识）；
  - 条款是行首编号（4.4.2 xxx），可能跨行续写；
  - **表格是单行 HTML `<table>`**，可精确还原行列（比 docx 好处理得多）；
  - 图是 <img src=...> 外链 + `图 X.Y.Z 标题`，**文字无法从正文取**，只记标题；
  - 文档后半段（约 7161 行起）是重复章节号的"条文说明"，须排除。

输出：
  kks-audit/references/rules_gb50549.md
  kks-audit/references/rules_gb50549.json
"""
import html
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "国标" / "GBT 50549-2020 电厂标识系统编码标准.pdf_by_PaddleOCR-VL-1.6.md"
OUT_MD = ROOT / "kks-audit" / "references" / "rules_gb50549.md"
OUT_JSON = ROOT / "kks-audit" / "references" / "rules_gb50549.json"

HEAD_RE = re.compile(r"^(#{1,4})\s+(\d+(?:\.\d+)?)\s+(\S.*?)\s*$")
APPENDIX_RE = re.compile(r"^#{1,4}\s*附\s*录\s*([A-Z])\s*(.*?)\s*$")
APPENDIX_CLAUSE_RE = re.compile(r"^([A-Z])\.\s*(\d+(?:\.\d+)*)\s+(\S.*)$")
CLAUSE_RE = re.compile(r"^(\d+\.\d+(?:\.\d+)?)\s+(\S.*)$")
ITEM_RE = re.compile(r"^(\d{1,2})\s+(\S.*)$")
TABLE_RE = re.compile(r"<table[^>]*>(.*?)</table>", re.S)
ROW_RE = re.compile(r"<tr[^>]*>(.*?)</tr>", re.S)
CELL_RE = re.compile(r"<t[dh][^>]*>(.*?)</t[dh]>", re.S)
TAG_RE = re.compile(r"<[^>]+>")
IMG_RE = re.compile(r"<img[^>]*>")
CAPTION_RE = re.compile(r"^\s*图\s*(\d+(?:\.\d+)*)\s*(.*)$")


def clean(text: str) -> str:
    text = TAG_RE.sub("", text)
    text = html.unescape(text)
    text = text.replace("\\n", " ")
    text = re.sub(r"\$\s*\\mathrm\{([^}]*)\}\s*\$", r"\1", text)
    text = re.sub(r"\$\s*([^$]{1,40}?)\s*\$", r"\1", text)
    text = text.replace("{", "").replace("}", "")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def parse_tables(line: str) -> list[list[list[str]]]:
    out = []
    for table_html in TABLE_RE.findall(line):
        rows = []
        for row_html in ROW_RE.findall(table_html):
            cells = [clean(c) for c in CELL_RE.findall(row_html)]
            if any(cells):
                rows.append(cells)
        if rows:
            out.append(rows)
    return out


def body_range(lines: list[str]) -> tuple[int, int]:
    """返回正文区间：第一个 '## 1 总 则' 到第二个同名标题（条文说明）之前。"""
    hits = [i for i, l in enumerate(lines) if re.match(r"^#{2,3}\s+1\s+总\s*则\s*$", l.strip())]
    if len(hits) >= 2:
        return hits[0], hits[1]
    return 0, len(lines)


def parse() -> dict:
    lines = SRC.read_text(encoding="utf-8").splitlines()
    start, end = body_range(lines)
    chapters: list[dict] = []
    appendices: list[dict] = []
    container: dict | None = None   # 当前章 或 当前附录
    section: dict | None = None

    def ensure_section() -> dict:
        nonlocal section
        if section is None:
            section = {
                "num": container["num"] if container else "?",
                "title": container["title"] if container else "",
                "clauses": [], "tables": [], "figures": [],
            }
            container["sections"].append(section)  # type: ignore[union-attr]
        return section

    for raw in lines[start:end]:
        line = raw.rstrip()
        if not line.strip():
            continue
        stripped = line.strip()

        app = APPENDIX_RE.match(stripped)
        if app:
            container = {"num": f"附录 {app.group(1)}", "title": clean(app.group(2)), "sections": []}
            appendices.append(container)
            section = None
            continue

        head = HEAD_RE.match(stripped)
        if head:
            level, num, title = len(head.group(1)), head.group(2), clean(head.group(3))
            if level <= 2 and "." not in num:
                container = {"num": num, "title": title, "sections": []}
                chapters.append(container)
                section = None
            elif container is not None:
                section = {"num": num, "title": title, "clauses": [], "tables": [], "figures": []}
                container["sections"].append(section)
            continue

        if container is None:
            continue

        if "<table" in line:
            for table in parse_tables(line):
                ensure_section()["tables"].append(table)
            rest = TABLE_RE.sub("", line)
            if IMG_RE.search(rest):
                ensure_section()["figures"].append(clean(rest))
            continue

        if IMG_RE.search(line):
            ensure_section()["figures"].append(clean(line))
            continue

        cap = CAPTION_RE.match(clean(line))
        if cap and len(clean(line)) <= 40:
            target = ensure_section()
            label = f"图 {cap.group(1)} {cap.group(2)}".strip()
            if target["figures"]:
                target["figures"][-1] = label
            else:
                target["figures"].append(label)
            continue

        text = clean(line)
        if not text:
            continue

        clause = CLAUSE_RE.match(text)
        if clause:
            ensure_section()["clauses"].append({"no": clause.group(1), "text": clause.group(2), "items": []})
            continue
        app_clause = APPENDIX_CLAUSE_RE.match(text)
        if app_clause:
            ensure_section()["clauses"].append({
                "no": f"{app_clause.group(1)}.{app_clause.group(2)}",
                "text": app_clause.group(3), "items": [],
            })
            continue

        item = ITEM_RE.match(text)
        target = ensure_section()
        if item and target["clauses"]:
            target["clauses"][-1]["items"].append(f"{item.group(1)} {item.group(2)}")
        elif target["clauses"]:
            tail = target["clauses"][-1]
            if tail["items"]:
                tail["items"][-1] += text
            else:
                tail["text"] += text
        else:
            target.setdefault("notes", []).append(text)
    return {"source": SRC.name, "chapters": chapters, "appendices": appendices}


def render_md(data: dict) -> str:
    L = ["# GB/T 50549-2020《电厂标识系统编码标准》规则抽取", ""]
    L.append(f"- 来源：`{data['source']}`（PaddleOCR-VL 转换）")
    L.append("- 由脚本逐条抽取**正文条文**，不含条文说明；条款保留原文编号。")
    L.append("- 表格已从 HTML 还原为 Markdown 表；`图` 为 OCR 外链图片，正文无文字，需看图。")
    L.append("")
    L.append("## 目录")
    L.append("")
    for entity in data["chapters"] + data.get("appendices", []):
        n = sum(len(s["clauses"]) for s in entity["sections"])
        t = sum(len(s["tables"]) for s in entity["sections"])
        L.append(f"- **{entity['num']} {entity['title']}** — 条款 {n}，表 {t}")
    L.append("")
    for ch in data["chapters"] + data.get("appendices", []):
        L.append(f"## {ch['num']} {ch['title']}")
        L.append("")
        for s in ch["sections"]:
            if s["num"] != ch["num"]:
                L.append(f"### {s['num']} {s['title']}")
                L.append("")
            for c in s["clauses"]:
                L.append(f"- **{c['no']}** {c['text']}")
                for it in c["items"]:
                    L.append(f"    - {it}")
            if s["clauses"]:
                L.append("")
            for i, t in enumerate(s["tables"], 1):
                width = max(len(r) for r in t)
                L.append(f"**表 {s['num']}-{i}**")
                L.append("")
                L.append("| " + " | ".join(f"列{j+1}" for j in range(width)) + " |")
                L.append("|" + "---|" * width)
                for r in t:
                    row = r + [""] * (width - len(r))
                    L.append("| " + " | ".join(x.replace("|", "\\|") for x in row) + " |")
                L.append("")
            for f in s["figures"]:
                if f:
                    L.append(f"- （图，需看原件）{f}")
            if s["figures"]:
                L.append("")
    return "\n".join(L)


def main() -> None:
    data = parse()
    OUT_MD.parent.mkdir(parents=True, exist_ok=True)
    OUT_MD.write_text(render_md(data), encoding="utf-8")
    OUT_JSON.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    allent = data["chapters"] + data.get("appendices", [])
    print(f"章节 {len(data['chapters'])} 个 + 附录 {len(data.get('appendices', []))} 个；条款 "
          f"{sum(len(s['clauses']) for c in allent for s in c['sections'])} 条；"
          f"表 {sum(len(s['tables']) for c in allent for s in c['sections'])} 张")
    print(f"-> {OUT_MD.name} {OUT_MD.stat().st_size} B ; {OUT_JSON.name} {OUT_JSON.stat().st_size} B")
    print()
    for ch in allent:
        print("  %-8s %-20s 节 %-3d 条款 %-4d 表 %d" % (
            ch["num"], ch["title"][:20], len(ch["sections"]),
            sum(len(s["clauses"]) for s in ch["sections"]),
            sum(len(s["tables"]) for s in ch["sections"])))


if __name__ == "__main__":
    main()
