# -*- coding: utf-8 -*-
"""生成 KKS 分类码索引（config/kks_code_index.json）。

数据来源（**只依据现行有效文档**）：
  1. GB/T 50549-2020《电厂标识系统编码标准》附录 A（系统分类索引）、附录 B（设备及产品分类索引）
     —— 现行国家标准，2021-03-01 实施，同时废止 GB/T 50549-2010；
  2. 中煤集团《电厂标识系统编码应用细则（试行）》附录 A/B —— 集团落地，作为补充。

⚠️ 版本纪律（2026-10-08 定的口径）：
  项目的规则依据只能来自 **GB/T 50549-2020**。此前曾把 Q/WL《火电厂标识系统（KKS）编码导则》
  （2015 年企业标准，继承已废止的 GB/T 50549-2010 体系）一并作为字典，导致同一码出现两种含义
  （如 GP：2020 版是泵、2010 版体系是照明检修配电箱），并在报告里刷出无实际意义的"标准分歧"。
  经实测（舟山暖通 1190 行）：只按 2020 版判定与两套并集判定的告警数**完全一致**，因此已移除
  2010 体系字典，不再保留任何分歧披露逻辑。

产出结构：
  system_s1_groups   S1 主组（1 位字母）→ 中文系统名
  system_codes       S1S2S3 系统分类码（3 位字母）→ 中文系统名
  device_t1_groups   T1 主组（1 位字母）→ 中文名称
  device_codes       T1T2 设备/产品分类码（2 位字母）→ 中文名称
  source / coverage_notes / generated_by / note
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

# GB/T 附录B 的表格在 PDF 分页处会拆开，首表只有表头无 A/B/C 三行，
# 用中煤细则附录B 的条目顺序（其第 4 条"后续标准预留"即 GB/T 的 D）互证补全，
# 只补经两份文档互相印证的三条，不做顺序外推。
T1_SUPPLEMENT = {
    "A": "两个或更多目的或任务（仅适于无法识别其使用目的或任务的对象）",
    "B": "将输入变量（物理属性、条件或事件）转换为进一步处理的信号",
    "C": "能源、信息或材料的存储",
}


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


def harvest(appendix: dict, groups: dict, codes: dict, code_lengths: tuple[int, ...]) -> None:
    """把 GB/T 附录表格里的「码 → 中文名称」收进字典。

    code_lengths 指明该附录里哪些位数算"分类码"：
      附录A（系统分类索引）→ (2, 3)（2 位是组级前缀，3 位是正式系统分类码）
      附录B（设备分类索引）→ (2,)  （2 位设备及产品分类码）
    1 位一律进 groups（主组）。
    """
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
                elif len(code) in code_lengths:
                    codes.setdefault(code, name)


def merge_zmei(zmei: dict, system_s1: dict, system_codes: dict,
               device_t1: dict, device_codes: dict) -> None:
    """中煤细则附录作为补充：只补 GB/T 没有的码，不覆盖国标释义。"""
    z_app = zmei.get("appendices") or {}
    for row in (z_app.get("A") or {}).get("rows") or []:
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
        elif len(code) in (2, 3):
            system_codes.setdefault(code, name)
    for row in (z_app.get("B") or {}).get("rows") or []:
        if len(row) < 2:
            continue
        code = norm(row[0])
        if not CODE_RE.fullmatch(code):
            continue
        name = norm(row[1]) or (norm(row[2]) if len(row) > 2 else "")
        if not name or name in ("—", "一"):
            continue
        if len(code) == 1:
            device_t1.setdefault(code, name)
        elif len(code) == 2:
            device_codes.setdefault(code, name)


def main() -> None:
    gb = json.loads(GB.read_text(encoding="utf-8"))
    zmei = json.loads(ZMEI.read_text(encoding="utf-8"))
    gb_app = {a["num"]: a for a in gb["appendices"]}

    system_s1: dict[str, str] = {}
    system_codes: dict[str, str] = {}
    device_t1: dict[str, str] = {}
    device_codes: dict[str, str] = {}

    harvest(gb_app["附录 A"], system_s1, system_codes, (2, 3))
    harvest(gb_app["附录 B"], device_t1, device_codes, (2,))
    merge_zmei(zmei, system_s1, system_codes, device_t1, device_codes)
    for code, name in T1_SUPPLEMENT.items():
        device_t1.setdefault(code, name)

    payload = {
        "standard": "GB/T 50549-2020《电厂标识系统编码标准》（2021-03-01 实施；同时废止 GB/T 50549-2010）",
        "coverage_notes": (
            "S1 主组、S1S2S3 系统分类码来自 GB/T 50549-2020 附录A（表A.0.1/表A.0.2）；"
            "T1 主组、T1T2 设备及产品分类码来自附录B（表B.0.1/表B.0.2）；"
            "中煤集团《电厂标识系统编码应用细则（试行）》附录A/B 作为集团补充（只补国标未收录的码）。"
            "标准附录在 PDF 中分页，抽取后仍可能有极少数码未收录——"
            "因此「码未收录」只能作为待核线索，不能直接判定为错误。"
            "本项目不引用 GB/T 50549-2010 体系（含继承该体系的 2015 年火电导则）的分类码。"
        ),
        "source": [
            "GB/T 50549-2020 附录A 系统分类索引",
            "GB/T 50549-2020 附录B 设备及产品分类索引",
            "中煤集团《电厂标识系统编码应用细则（试行）》附录A/B（补充）",
        ],
        "generated_by": "tools/build_code_index.py",
        "note": "分类码取值是否合法以 GB/T 50549-2020 为准；本索引供自动校验与 AI 建议编码使用。",
        "system_s1_groups": dict(sorted(system_s1.items())),
        "system_codes": dict(sorted(system_codes.items())),
        "device_t1_groups": dict(sorted(device_t1.items())),
        "device_codes": dict(sorted(device_codes.items())),
    }
    OUT.write_text(json.dumps(payload, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    print(f"-> {OUT.relative_to(ROOT)} {OUT.stat().st_size} B")
    print(f"  S1 主组 {len(system_s1)} 个；系统分类码 S1S2S3 {len(system_codes)} 个")
    print(f"  T1 主组 {len(device_t1)} 个；设备分类码 T1T2 {len(device_codes)} 个")
    print()
    for k in ("GP", "GQ", "AP", "AA", "AN", "BR", "FL", "QM", "QN", "RM", "WP", "XA", "LBA", "AXB"):
        where = []
        for label, d in (("S1", system_s1), ("系统码", system_codes), ("T1", device_t1), ("设备码", device_codes)):
            if k in d:
                where.append(f"{label}:{d[k][:26]}")
        print(f"  {k:5s} -> " + ("; ".join(where) if where else "（2020 版未收录）"))


if __name__ == "__main__":
    main()
