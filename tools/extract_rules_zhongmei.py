# -*- coding: utf-8 -*-
"""从《中国中煤能源集团电厂标识系统编码应用细则（试行）A3.2》抽取结构化规则。

只做"搬运+归类"：不改写、不推断，每条规则保留原文条款号，便于回查。
docx 表格用 </w:tr> / </w:tc> 显式分隔符解析（先把段落标记换成单元格内空格），
比按换行切分可靠得多。

输出：
  kks-audit/references/rules_zhongmei.md    人读的结构化规则参考
  kks-audit/references/rules_zhongmei.json  机器可校验的子集
"""
import html
import json
import re
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCX = ROOT / "国标" / "中国中煤能源集团电厂标识系统编码应用细则（试行）A3.2.docx"
OUT_MD = ROOT / "kks-audit" / "references" / "rules_zhongmei.md"
OUT_JSON = ROOT / "kks-audit" / "references" / "rules_zhongmei.json"

SECTION_RE = re.compile(r"^(\d+\.\d+)\s+(\S.*)$")
CLAUSE_RE = re.compile(r"^(\d+\.\d+\.\d+)\s+(.*)$")
APPENDIX_RE = re.compile(r"^附录\s*([A-Z])")


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def load_rows() -> list[tuple[str, object]]:
    """按文档顺序返回 [('p', 文本) | ('tr', [单元格, ...]), ...]。

    关键：</w:tr> 只出现在表格行里，正文段落没有行标记。所以按 </w:tr> 切块后，
    块内含 </w:tc>（\\x01）的是表格行，且 \\x01 之前还可能有属于正文的段落。
    """
    raw = zipfile.ZipFile(DOCX).read("word/document.xml").decode("utf-8", "replace")
    raw = re.sub(r"<w:tab/>", "\x04", raw)
    raw = raw.replace("</w:tc>", "\x01").replace("</w:tr>", "\x02").replace("</w:p>", "\x03")
    text = html.unescape(re.sub(r"<[^>]+>", "", raw))

    out: list[tuple[str, object]] = []
    for block in text.split("\x02"):
        if "\x01" not in block:
            for para in block.split("\x03"):
                if _norm(para):
                    out.append(("p", _norm(para)))
            continue
        parts = block.split("\x01")
        for para in parts[0].split("\x03"):
            if _norm(para):
                out.append(("p", _norm(para)))
        cells = [_norm(c.replace("\x03", " ").replace("\x04", " ")) for c in parts[1:]]
        if any(cells):
            out.append(("tr", cells))
    return out


def first(row: list[str]) -> str:
    return row[0].strip() if row else ""


def sole(row: list[str]) -> bool:
    return sum(1 for c in row if c) == 1


def parse() -> dict:
    entries = load_rows()
    sections: list[dict] = []
    appendix: dict[str, dict] = {}
    current: dict | None = None
    current_app: dict | None = None

    for kind, payload in entries:
        if kind == "p":
            text = str(payload)
            app = APPENDIX_RE.match(text)
            if app:
                current_app = {"letter": app.group(1), "title": "", "rows": [], "notes": []}
                appendix[app.group(1)] = current_app
                current = None
                continue
            sec = SECTION_RE.match(text)
            if sec and sec.group(1).split(".")[0] in {"7", "8", "9"} and len(sec.group(2)) <= 30:
                current = {"num": sec.group(1), "title": sec.group(2).strip(), "clauses": [], "tables": []}
                sections.append(current)
                current_app = None
                continue
            if current_app is not None:
                if re.match(r"^[A-Z]\.\d", text):
                    current_app["notes"].append(text)
                elif not current_app["title"] and text != "（规范性附录）" and len(text) <= 20:
                    current_app["title"] = text
                continue
            if current is not None:
                clause = CLAUSE_RE.match(text)
                if clause:
                    current["clauses"].append({"no": clause.group(1), "text": clause.group(2).strip()})
                elif "\x04" in text:
                    # 用制表符排版、而非真表格的规则表（如 8.1.5 / 8.1.6 阀门分类码表）。
                    # 内容按原文顺序保留；行列结构以原文为准，此处不猜测合并单元格。
                    current["tables"].append([c.strip() for c in text.split("\x04") if c.strip()])
                elif current["tables"] and len(text) <= 20 and not text.endswith(("。", "；", "，", "：", ":", "、")):
                    # 表格里用独立段落写的单格（如 FL / 名称 / 001～299），按原文顺序保留
                    current["tables"].append([text])
                elif current["clauses"] and not current["tables"]:
                    current["clauses"][-1]["text"] += text
            continue

        cells = payload  # type: ignore[assignment]
        head = first(cells)  # type: ignore[arg-type]
        clause = CLAUSE_RE.match(head)
        if clause:
            rest = " ".join(c for c in cells[1:] if c)  # type: ignore[index]
            target = current if current is not None else None
            if target is not None:
                target["clauses"].append(
                    {"no": clause.group(1), "text": (clause.group(2) + " " + rest).strip()})
            continue
        if current is not None:
            current["tables"].append([c for c in cells if c])  # type: ignore[union-attr]
        elif current_app is not None:
            current_app["rows"].append([c for c in cells if c])  # type: ignore[union-attr]

    return {"source": DOCX.name, "sections": sections, "appendices": appendix}


def render_md(data: dict) -> str:
    L: list[str] = []
    L.append("# 中煤集团《电厂标识系统编码应用细则（试行）》规则抽取")
    L.append("")
    L.append(f"- 来源：`{data['source']}`（2026-05 版）")
    L.append("- 本文由脚本逐条抽取原文，**未改写、未推断**；条款保留原文编号便于回查。")
    L.append("- 用途：支撑 `kks-audit` 的「设备分类码 / 编号区间 / 系统分类码」语义校验。")
    L.append("")
    L.append("## 目录")
    L.append("")
    L.append("### 正文（火电厂 8.x、可再生能源 9.x、命名规则 7.x）")
    L.append("")
    for s in data["sections"]:
        L.append(f"- **{s['num']} {s['title']}** — 条款 {len(s['clauses'])} 条，表 {len(s['tables'])} 行"
                 f"（[跳转](#sec-{s['num'].replace('.', '-')})）")
    L.append("")
    L.append("### 附录")
    L.append("")
    for letter, app in data["appendices"].items():
        L.append(f"- **附录 {letter} {app['title']}** — {len(app['rows'])} 行"
                 f"（[跳转](#app-{letter})）")
    L.append("")
    L.append("---")
    L.append("")
    for s in data["sections"]:
        L.append(f'<a id="sec-{s["num"].replace(".", "-")}"></a>')
        L.append("")
        L.append(f"## {s['num']} {s['title']}")
        L.append("")
        if s["clauses"]:
            L.append("### 条款")
            L.append("")
            for c in s["clauses"]:
                L.append(f"- **{c['no']}** {c['text']}")
            L.append("")
        if s["tables"]:
            L.append("### 表格原文（按行）")
            L.append("")
            L.append("```text")
            for r in s["tables"]:
                L.append(" | ".join(r))
            L.append("```")
            L.append("")
    for letter, app in data["appendices"].items():
        L.append(f'<a id="app-{letter}"></a>')
        L.append("")
        L.append(f"## 附录 {letter} {app['title']}")
        L.append("")
        for n in app.get("notes", []):
            L.append(f"- {n}")
        if app.get("notes"):
            L.append("")
        L.append("| " + " | ".join(["码", "名称", "说明", "英文"]) + " |")
        L.append("|---|---|---|---|")
        for r in app["rows"]:
            cells = (r + ["", "", "", ""])[:4]
            L.append("| " + " | ".join(c.replace("|", "\\|") for c in cells) + " |")
        L.append("")
    return "\n".join(L)


def main() -> None:
    data = parse()
    OUT_MD.parent.mkdir(parents=True, exist_ok=True)
    OUT_MD.write_text(render_md(data), encoding="utf-8")
    OUT_JSON.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    total_c = sum(len(s["clauses"]) for s in data["sections"])
    total_t = sum(len(s["tables"]) for s in data["sections"])
    print(f"章节 {len(data['sections'])} 个 / 条款 {total_c} 条 / 表行 {total_t} 行")
    print(f"附录 {len(data['appendices'])} 个, 行数 "
          + ", ".join(f"{k}={len(v['rows'])}" for k, v in data["appendices"].items()))
    print(f"-> {OUT_MD.name} {OUT_MD.stat().st_size} B ; {OUT_JSON.name} {OUT_JSON.stat().st_size} B")
    print()
    for s in data["sections"]:
        print("  %-6s %-16s 条款 %-3d 表行 %d" % (s["num"], s["title"][:16], len(s["clauses"]), len(s["tables"])))
    print()
    for k, v in data["appendices"].items():
        print("  附录 %s %-14s %d 行；样例: %s" % (k, v["title"][:14], len(v["rows"]),
              " / ".join(v["rows"][2][:3]) if len(v["rows"]) > 2 else "-"))


if __name__ == "__main__":
    main()
