# -*- coding: utf-8 -*-
"""从已抽取的标准规则里生成 KKS 分类码索引（config/kks_code_index.json）。

数据来源：
  1. GB/T 50549-2020 附录 A（系统分类索引）、附录 B（设备及产品分类索引）
     —— 国家标准，表格是结构化的 HTML，可靠性最高；
  2. Q/WL《火电厂标识系统（KKS）编码导则》6.1/6.2/6.3 索引
     —— 项目默认采用的企业标准（run_audit.STANDARD_DEFAULT = "huodian"）；
  3. 中煤细则附录 A/B —— 集团落地，作为补充。

⚠️ GB/T 与火电导则的**分类码体系并不相同**（例如阀门：国标 QM、企标 AA；
泵：国标 GP、企标 AP；管道：国标 WP、企标 BR）。因此本索引同时保留两套字典，
并在 cross_standard_divergences 里显式登记"同码不同义"，供审核时按标准选码、
把分歧如实披露给审核员，而不是让程序偷偷按某一方判定。

产出结构：
  system_s1_groups          S1 主组（1 位字母）→ 中文系统名（GB/T）
  system_codes              S1S2S3 系统分类码（3 位字母）→ 中文系统名（GB/T）
  device_t1_groups          T1 主组（1 位字母）→ 中文名称（GB/T）
  device_codes              T1T2 设备/产品分类码（2 位字母）→ 中文名称（GB/T）
  system_codes_huodian      火电导则 6.1 系统索引（码 → 名称）
  device_codes_huodian      火电导则 6.2 设备索引（码 → 名称）
  cross_standard_divergences 同码在两套标准中含义不同的清单
"""
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
GB = ROOT / "kks-audit" / "references" / "rules_gb50549.json"
ZMEI = ROOT / "kks-audit" / "references" / "rules_zhongmei.json"
HUODIAN = ROOT / "kks-audit" / "references" / "rules_huodian.json"
OUT = ROOT / "config" / "kks_code_index.json"

CODE_RE = re.compile(r"^[A-Z]{1,3}$")
HEADER_WORDS = ("系统分类码", "设备分类码", "设备/产品分类码", "设备及产品分类码", "产品分类码", "码")
# 名称归一：去标点后抹掉近义连接词与通用后缀，避免把"电网与配电系统"和
# "电网和配电系统"这类同义异写误判成标准冲突。
FOLD_RE = re.compile(r"[与和及/、，,（）()\s的之]|系统|设备|回路|装置")


def norm(text: str) -> str:
    return re.sub(r"\s+", " ", str(text)).strip()


def fold(text: str) -> str:
    return FOLD_RE.sub("", str(text or "")).lower()


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


def harvest_huodian(data: dict) -> tuple[dict, dict, dict, dict]:
    """从火电导则 6.1/6.2 索引还原（主组, 码表）。"""
    sections = data.get("index_sections", {})
    sys_groups: dict[str, str] = {}
    sys_codes: dict[str, str] = {}
    dev_groups: dict[str, str] = {}
    dev_codes: dict[str, str] = {}
    for name, sec in sections.items():
        is_system = "系统" in name
        groups = sys_groups if is_system else dev_groups
        codes = sys_codes if is_system else dev_codes
        for code, label in (sec.get("codes") or {}).items():
            label = norm(label)
            if not label:
                continue
            if len(code) == 1:
                groups.setdefault(code, label)
            else:
                codes.setdefault(code, label)
    return sys_groups, sys_codes, dev_groups, dev_codes


# 国标把这些位留空（"后续标准预留"/"可用"），企业标准在预留位上做了定义——
# 这属于**正常细化**，不是标准冲突，判定时必须排除，否则会把上百个"企标填了国标预留位"
# 误报成分歧。
RESERVED_WORDS = ("后续标准预留", "预留", "可用", "reserved")


def _head_name(text: str) -> str:
    """取名称首段（去掉 " / " 之后的举例、说明），避免把举例差异误判成分歧。"""
    return norm(str(text or "").split(" / ")[0])[:80]


def divergences(gb_codes: dict, hd_codes: dict, dimension: str, max_len: int) -> list[dict]:
    """找出两套标准里**同一码含义不同**的码。

    只比较到 max_len 位：分歧只会出现在「主组/组」这一层，再往下细分两套标准本就用
    不同的编号分配（那是 DIV-06 登记的差异），逐码列出来只会淹没审核人员，没有价值。

    判定口径（做过误报校准）：
      1. 任一方是国标预留位（后续标准预留/可用）→ 不算分歧，属企标细化；
      2. 名称折叠同义后相等 → 不算；
      3. 短名称是长名称的子串 → 不算分歧，属企标补充限定词（如"给水系统"↔"除氧给水系统"）。
    """
    out: list[dict] = []
    for code in sorted(set(gb_codes) & set(hd_codes)):
        if len(code) > max_len:
            continue
        a, b = _head_name(gb_codes[code]), _head_name(hd_codes[code])
        if not a or not b:
            continue
        if any(w in a or w in b for w in RESERVED_WORDS):
            continue
        fa, fb = fold(a), fold(b)
        if not fa or not fb or fa == fb:
            continue
        if fa in fb or fb in fa:
            continue
        out.append({"code": code, "dimension": dimension, "gb50549": a[:80], "huodian": b[:80]})
    return out


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
                    elif len(code) in (2, 3) and codes is system_codes:
                        codes.setdefault(code, name)   # 2 位是系统码的组级前缀，一并保留

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

    # 火电导则（项目默认标准）的索引码表
    hd_sys_groups: dict[str, str] = {}
    hd_sys_codes: dict[str, str] = {}
    hd_dev_groups: dict[str, str] = {}
    hd_dev_codes: dict[str, str] = {}
    if HUODIAN.is_file():
        hd = json.loads(HUODIAN.read_text(encoding="utf-8"))
        hd_sys_groups, hd_sys_codes, hd_dev_groups, hd_dev_codes = harvest_huodian(hd)

    # 标准间分歧：只登记「主组/组」层（系统码 2 位、设备码 2 位）含义不同的码；
    # 再往下细分两套标准的编号分配本就不同（DIV-06），逐码列出只会淹没审核人员。
    div_sys = divergences(system_codes, hd_sys_codes, "系统分类码 S1S2", 2)
    div_dev = divergences(device_codes, hd_dev_codes, "设备分类码 T1T2", 2)
    # 主组级分歧（1 位）单独登记：它解释"为什么会有这么多码不同义"，但逐行匹配不到，
    # 只作为说明性数据随索引一起留档（对应 config/kks_rules.json 的 DIV-01/DIV-04）。
    div_sys_main = divergences(system_s1, hd_sys_groups, "系统分类码主组 S1", 1)
    div_dev_main = divergences(device_t1, hd_dev_groups, "设备分类码主组 T1", 1)

    payload = {
        "coverage_notes": (
            "S1 主组、S1S2S3 系统分类码来自 GB/T 50549-2020 附录A；T1 主组、T1T2 设备分类码"
            "来自附录B；另收录 Q/WL 火电导则 6.1/6.2 索引（项目默认采用的企业标准）。"
            "两套标准的分类码体系不同，同码不同义的码登记在 cross_standard_divergences；"
            "判定「码是否合法」时按两套标准的并集，未收录只能作为待核线索，不能直接判为错误。"
            "标准附录在 PDF 中分页，抽取后仍可能有极少数码未收录。"
        ),
        "source": [
            "GB/T 50549-2020 附录A 系统分类索引",
            "GB/T 50549-2020 附录B 设备及产品分类索引",
            "Q/WL 火电厂标识系统（KKS）编码导则 6.1/6.2/6.3 索引",
            "中煤集团《电厂标识系统编码应用细则（试行）》附录A/B（补充）",
        ],
        "generated_by": "tools/build_code_index.py",
        "note": "分类码合法性按 GB/T 50549-2020 与火电导则两套标准的并集判定；同码异义以 cross_standard_divergences 为准。",
        "system_s1_groups": dict(sorted(system_s1.items())),
        "system_codes": dict(sorted(system_codes.items())),
        "device_t1_groups": dict(sorted(device_t1.items())),
        "device_codes": dict(sorted(device_codes.items())),
        "system_s1_groups_huodian": dict(sorted(hd_sys_groups.items())),
        "system_codes_huodian": dict(sorted(hd_sys_codes.items())),
        "device_t1_groups_huodian": dict(sorted(hd_dev_groups.items())),
        "device_codes_huodian": dict(sorted(hd_dev_codes.items())),
        "cross_standard_divergences": {
            "note": (
                "只登记「主组/组」层含义不同的码。审核时不得据此判错："
                "系统码按前 2 位、设备码按 2 位精确匹配后披露双方定义，由审核员确认采用哪套标准。"
            ),
            "system": div_sys,
            "device": div_dev,
            "system_main_groups": div_sys_main,
            "device_main_groups": div_dev_main,
        },
    }
    OUT.write_text(json.dumps(payload, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    print(f"-> {OUT.relative_to(ROOT)} {OUT.stat().st_size} B")
    print(f"  GB/T  : S1 主组 {len(system_s1)}；系统码 {len(system_codes)}；T1 主组 {len(device_t1)}；设备码 {len(device_codes)}")
    print(f"  火电导则: S1 主组 {len(hd_sys_groups)}；系统码 {len(hd_sys_codes)}；T1 主组 {len(hd_dev_groups)}；设备码 {len(hd_dev_codes)}")
    print(f"  标准间分歧: 系统码 {len(div_sys)} 个；设备码 {len(div_dev)} 个")
    print()
    for k in ("LBA", "AXB", "MST", "QM", "WP", "GP", "FL", "RM", "QN", "AA", "AP", "AN", "BR", "CP"):
        where = []
        for label, d in (("S1", system_s1), ("S1S2S3", system_codes), ("T1", device_t1), ("T1T2", device_codes),
                         ("企标sys", hd_sys_codes), ("企标dev", hd_dev_codes)):
            if k in d:
                where.append(f"{label}:{d[k][:22]}")
        print(f"  {k:5s} -> " + ("; ".join(where) if where else "（未收录）"))


if __name__ == "__main__":
    main()
