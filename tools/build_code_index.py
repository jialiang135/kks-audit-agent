# -*- coding: utf-8 -*-
"""从已抽取的标准规则里生成 KKS 分类码索引（config/kks_code_index.json）。

数据来源优先级：
  1. GB/T 50549-2020 附录 A（系统分类索引）、附录 B（设备及产品分类索引）
     —— 法定依据，表格是结构化的 HTML，可靠性最高；
  2. 中煤细则附录 A/B —— 集团落地，作为补充（只补 GB/T 没有的码）。

产出结构：
  system_s1_groups  S1 主组（1 位字母）→ 中文系统名
  system_codes      S1S2S3 系统分类码（3 位字母）→ 中文系统名
  device_t1_groups  T1 主组（1 位字母）→ 中文名称
  device_codes      T1T2 设备/产品分类码（2 位字母）→ 中文名称/说明
"""
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
GB = ROOT / "kks-audit" / "references" / "rules_gb50549.json"
ZMEI = ROOT / "kks-audit" / "references" / "rules_zhongmei.json"
OUT = ROOT / "config" / "kks_code_index.json"

CODE_RE = re.compile(r"^[A-Z]{1,3}$")
HEADER_WORDS = ("系统分类码", "设备分类码", "设备/产品分类码", "设备及产品分类码", "产品分类码", "码")


def norm(text: str) -> str:
    return re.sub(r"\s+", " ", str(text)).strip()


def pick_name(cells: list[str]) -> str:
    """在表行的若干单元格里挑出中文名称/说明（跳过纯英文、页码、符号）。"""
    cands = []
    for c in cells[1:]:
        c = norm(c)
        if not c or c in ("—", "-", "一", "－"):
            continue
        if re.fullmatch(r"[A-Za-z0-9 ,.()/&'’\-]+", c):
            continue  # 纯英文
        cands.append(c)
    return " / ".join(dict.fromkeys(cands))[:160]


def collect(appendix: dict, out_groups: dict, out_codes: dict) -> None:
    for section in appendix.get("sections", []):
        for table in section.get("tables", []):
            for row in table:
                if not row:
                    continue
                code = norm(row[0]).replace("*", "").replace(" ", "")
                if not CODE_RE.fullmatch(code):
                    continue
                if any(w in norm(row[0]) for w in HEADER_WORDS):
                    continue
                name = pick_name(row)
                if not name:
                    continue
                target = out_groups if len(code) == 1 else out_codes
                if len(code) == 2:
                    continue  # 2 位只在设备附录里收集，由调用方分别处理
                target.setdefault(code, name)


def main() -> None:
    gb = json.loads(GB.read_text(encoding="utf-8"))
    zmei = json.loads(ZMEI.read_text(encoding="utf-8"))
    gb_app = {a["num"]: a for a in gb["appendices"]}

    system_s1: dict[str, str] = {}
    system_codes: dict[str, str] = {}
    device_t1: dict[str, str] = {}
    device_codes: dict[str, str] = {}

    def harvest(appendix: dict, groups: dict, codes: dict) -> None:
        for section in appendix.get("sections", []):
            for table in section.get("tables", []):
                for row in table:
                    if not row:
                        continue
                    raw = norm(row[0])
                    code = raw.replace("*", "").replace(" ", "")
                    if not CODE_RE.fullmatch(code):
                        continue
                    if any(w in raw for w in HEADER_WORDS):
                        continue
                    name = pick_name(row)
                    if not name:
                        continue
                    if len(code) == 1:
                        groups.setdefault(code, name)
                    elif len(code) == 2 and codes is device_codes:
                        codes.setdefault(code, name)
                    elif len(code) == 3 and codes is system_codes:
                        codes.setdefault(code, name)

    harvest(gb_app["附录 A"], system_s1, system_codes)
    harvest(gb_app["附录 B"], device_t1, device_codes)

    # 中煤细则作为补充：其附录 A 行格式为 [码, 中文, 说明, 英文]
    z_app = zmei["appendices"]  # {"A": {...}, "B": {...}, "C": {...}}
    if "A" in z_app:
        for row in z_app["A"]["rows"]:
            if len(row) < 2:
                continue
            code = norm(row[0])
            if not CODE_RE.fullmatch(code):
                continue
            name = norm(row[1])
            if not name or name in ("—", "一"):
                continue
            if len(code) == 1:
                system_s1.setdefault(code, name)
            elif len(code) == 3:
                system_codes.setdefault(code, name)
    if "B" in z_app:
        for row in z_app["B"]["rows"]:
            if len(row) < 2:
                continue
            code = norm(row[0])
            if not CODE_RE.fullmatch(code):
                continue
            name = norm(row[1]) or norm(row[2]) if len(row) > 2 else norm(row[1])
            if not name or name in ("—", "一"):
                continue
            if len(code) == 1:
                device_t1.setdefault(code, name)
            elif len(code) == 2:
                device_codes.setdefault(code, name)

    # 补全 T1 主组 A/B/C：GB/T 附录B 的表格分页导致首组缺失，而中煤细则附录B 的条目
    # 顺序与 GB/T 一致（其第 4 个条目"后续标准预留"即 GB/T 的 D=后续标准预留，可作锚点）。
    # 这里只补经过两份文档互相印证的 3 条，不做顺序外推。
    T1_SUPPLEMENT = {
        "A": "两个或更多目的或任务（仅适于无法识别其使用目的或任务的对象）",
        "B": "将输入变量（物理属性、条件或事件）转换为进一步处理的信号",
        "C": "能源、信息或材料的存储",
    }
    for code, name in T1_SUPPLEMENT.items():
        device_t1.setdefault(code, name)

    payload = {
        "coverage_notes": (
            "S1 主组、S1S2S3 系统分类码来自 GB/T 50549-2020 附录A（表A.0.1/表A.0.2）；"
            "T1 主组、T1T2 设备分类码来自附录B（表B.0.1/表B.0.2）。"
            "标准附录在 PDF 中分页，抽取后仍可能有极少数码未收录——"
            "因此「码未收录」只能作为待核线索，不能直接判定为错误。"
        ),
        "source": [
            "GB/T 50549-2020 附录A 系统分类索引",
            "GB/T 50549-2020 附录B 设备及产品分类索引",
            "中煤集团《电厂标识系统编码应用细则（试行）》附录A/B（补充）",
        ],
        "generated_by": "tools/build_code_index.py",
        "note": "分类码取值是否合法以 GB/T 50549-2020 为准；本索引供自动校验使用。",
        "system_s1_groups": dict(sorted(system_s1.items())),
        "system_codes": dict(sorted(system_codes.items())),
        "device_t1_groups": dict(sorted(device_t1.items())),
        "device_codes": dict(sorted(device_codes.items())),
    }
    OUT.write_text(json.dumps(payload, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    print(f"-> {OUT.relative_to(ROOT)} {OUT.stat().st_size} B")
    print(f"  S1 主组 {len(system_s1)} 个；系统分类码 S1S2S3 {len(system_codes)} 个")
    print(f"  T1 主组 {len(device_t1)} 个；设备分类码 T1T2 {len(device_codes)} 个")
    for label, d in (("S1", system_s1), ("S1S2S3", system_codes), ("T1", device_t1), ("T1T2", device_codes)):
        keys = list(d)[:8]
        print(f"  {label} 样例: " + ", ".join(f"{k}={d[k][:18]}" for k in keys))
    print()
    for k in ("LBA", "AXB", "MST", "QM", "WP", "GP", "FL", "RM", "QN"):
        where = []
        for label, d in (("S1", system_s1), ("S1S2S3", system_codes), ("T1", device_t1), ("T1T2", device_codes)):
            if k in d:
                where.append(f"{label}:{d[k][:24]}")
        print(f"  {k:5s} -> " + ("; ".join(where) if where else "（未收录）"))


if __name__ == "__main__":
    main()
