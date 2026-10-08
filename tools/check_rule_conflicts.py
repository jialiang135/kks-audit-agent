# -*- coding: utf-8 -*-
"""规则一致性核查：把「编码规则有没有冲突」变成一条能自动跑、能当门禁用的命令。

为什么需要它
------------
项目的规则来自 4 份文件（国标 GB/T 50549-2020、火电导则 Q/WL、中煤细则、风电/光伏规范），
落地到 5 个地方（config/kks_rules.json、run_audit.py、audit_template.py、SKILL.md、
config/kks_code_index.json）。任何一处改动都可能让"同一件事被两条规则用不同标准判定"。
2026-10-08 就真出过一次：KKS-32/33 用国标附录A/B 判分类码，而项目默认标准是火电导则，
两套分类码体系不同（阀门 国标 QM / 企标 AA；泵 国标 GP / 企标 AP），结果是整片误报。

检查项
------
A 注册一致性  规则号在「定义表 / 出处表 / 代码 / SKILL.md」四处必须对齐，不能有幽灵规则。
B 配置自洽    配置里的常量必须与代码里的常量一致，码表/区间/字典不得互相矛盾。
C 跨标准      两套标准的同码异义必须被登记；已登记的分歧必须与索引里的实际差异一致。
D 同条件重复  同一行被多条规则以不同等级命中时，必须已登记为"互补"，否则算冲突。

退出码：0 = 无 ERROR；1 = 有 ERROR（供 CI / pytest 使用）。
用法：python tools/check_rule_conflicts.py [--json 路径] [--quiet]
"""
from __future__ import annotations

import argparse
import ast
import importlib.util
import json
import re
import sys
from pathlib import Path

PROJECT = Path(__file__).resolve().parents[1]
if str(PROJECT) not in sys.path:
    sys.path.insert(0, str(PROJECT))

RULES_JSON = PROJECT / "config" / "kks_rules.json"
CODE_INDEX = PROJECT / "config" / "kks_code_index.json"
TEMPLATE = PROJECT / "kks-audit" / "scripts" / "audit_template.py"
SKILL_MD = PROJECT / "kks-audit" / "SKILL.md"
RUN_AUDIT = PROJECT / "run_audit.py"
HUODIAN_JSON = PROJECT / "kks-audit" / "references" / "rules_huodian.json"

RULE_ID_RE = re.compile(r"KKS-[0-9A-Za-z]+(?:/\d+[A-Za-z]?)*")

# 已知互补关系：同一行同时命中属于「不同维度」，不是规则冲突。
# 每条都必须写明为什么互补，避免这里变成"遇到告警就加白名单"。
COMPLEMENTARY_PAIRS = {
    ("KKS-29", "KKS-33"): "KKS-29 是 OCR 易混字符待核（P0/P1，提示人工核对是否为数字误识别）；"
                          "KKS-33 是设备分类码含禁用字母 I/O（P0，标准明文禁止）。同一事实两种处置，互补不冲突。",
    ("KKS-06", "KKS-33"): "KKS-06 判全厂码 G 位与整体结构位取值；KKS-33 判第 8-9 位设备分类码，作用位不同。",
    ("KKS-32", "KKS-36"): "KKS-32 判「码是否被两套标准收录」；KKS-36 披露「码在两套标准里含义不同」，维度不同。",
    ("KKS-33", "KKS-36"): "同上：一条管收录、一条管分歧披露。",
    ("KKS-04", "KKS-30"): "KKS-04 判字符合法性；KKS-30 判首尾空白/控制符等文本卫生，症状不同。",
    ("KKS-30", "KKS-33"): "编码文本卫生与设备分类码取值，作用维度不同。",
    ("KKS-05", "KKS-17/22"): "KKS-05 判长度异常；KKS-17/22 对 12 位以上的扩展码做定性分流，阈值相邻不矛盾。",
    ("KKS-04b", "KKS-29"): "KKS-04b 判分段字符类型；KKS-29 判易混字符，前者定类型后者提待核。",
    ("KKS-27", "KKS-30"): "KKS-27 判名称语义可解性；KKS-30 判文本卫生，同列不同维度。",
    ("KKS-10", "KKS-09"): "自引用环与前缀一致性，同一层级关系上的两种缺陷。",
    ("KKS-25", "KKS-25b"): "同名多码与同物异名，聚簇口径不同。",
}


class Finding(dict):
    def __init__(self, level: str, area: str, rule: str, detail: str, evidence: str = "") -> None:
        super().__init__(level=level, area=area, rule=rule, detail=detail, evidence=evidence)


def _load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def _rule_ids_in_source(path: Path) -> set[str]:
    """用 AST 取源码里出现的规则号字面量（字符串常量），避免正则误伤注释。"""
    tree = ast.parse(path.read_text(encoding="utf-8"))
    found: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Constant) and isinstance(node.value, str):
            for m in RULE_ID_RE.finditer(node.value):
                found.add(m.group(0))
    return found


# --------------------------------------------------------------- A 注册一致性
def check_registry(run, cfg, findings: list[Finding]) -> None:
    defs = set(run.RULE_DEFINITIONS)
    sources = set(run.RULE_SOURCES)
    extra_src = sources - defs
    if extra_src:
        findings.append(Finding("ERROR", "A 注册一致性", ", ".join(sorted(extra_src)),
                                "有规则出处但没有规则定义"))

    # 出处是"组键回退"口径：取样时 rule_source() 会做前缀回退，因此不能要求两个表键完全相等。
    # 真正要保证的是——每条规则要么能取到具体条款，要么已在 config 里登记"无专条依据"。
    declared_generic = set(((cfg.get("generic_rules_without_clause") or {}).get("rules")) or [])
    generic = []
    for rid in sorted(defs):
        if rid in declared_generic:
            continue
        for std in ("huodian", "wind", "solar"):
            src = run.rule_source(rid, std)
            if not src or src == run.GENERIC_SOURCE:
                generic.append(rid)
                break
    if generic:
        findings.append(Finding("ERROR", "A 注册一致性", f"{len(generic)} 条规则无专条依据且未登记",
                                "要么补 RULE_SOURCES，要么把规则号加进 config 的 "
                                "generic_rules_without_clause.rules（登记为「标准无对应条款的数据治理要求」）",
                                "、".join(generic)))
    stale = sorted(declared_generic - defs)
    if stale:
        findings.append(Finding("ERROR", "A 注册一致性", "generic_rules_without_clause",
                                "登记了不存在的规则号", "、".join(stale)))

    # 代码里引用的规则号（去掉组键里的成员，如 KKS-13/16 的两个成员单独出现时不算幽灵）
    src_text = RUN_AUDIT.read_text(encoding="utf-8")
    literal_ids = {i for i in _rule_ids_in_source(RUN_AUDIT) | _rule_ids_in_source(TEMPLATE)
                   if not i.startswith("KKS-TREE")}
    known = set(defs)
    members: set[str] = set()
    for key in defs:
        members.update(part for part in key.split("/") if part)
    ghosts = {i for i in literal_ids if i not in known and i not in members}
    if ghosts:
        findings.append(Finding("ERROR", "A 注册一致性", ", ".join(sorted(ghosts)),
                                "代码里引用了未登记的规则号（幽灵规则）"))

    # config 里的 rule_sources 必须是「已实现」的规则，未实现的条款放 reference_clauses
    cfg_rules = set((cfg.get("rule_sources") or {})) - {"note"}
    orphan_cfg = cfg_rules - known
    if orphan_cfg:
        findings.append(Finding("ERROR", "A 注册一致性", ", ".join(sorted(orphan_cfg)),
                                "config/kks_rules.json 的 rule_sources 里登记了未实现的规则号；"
                                "未实现的条款应放在 reference_clauses"))

    # SKILL.md 提到的规则号必须在定义表里
    skill_ids = {m.group(0) for m in RULE_ID_RE.finditer(SKILL_MD.read_text(encoding="utf-8"))}
    skill_ghosts = {i for i in skill_ids if i not in known and i not in members}
    if skill_ghosts:
        findings.append(Finding("WARN", "A 注册一致性", ", ".join(sorted(skill_ghosts)),
                                "SKILL.md 提到但未在 RULE_DEFINITIONS 登记的规则号（文档与代码漂移）"))

    # 反向：定义了但代码里从未出现的规则号（声明未实现）
    unused = sorted(i for i in defs if i not in literal_ids and i not in src_text)
    if unused:
        findings.append(Finding("ERROR", "A 注册一致性", ", ".join(unused),
                                "规则已定义但代码里没有任何地方产出它（声明未实现）"))


# --------------------------------------------------------------- B 配置自洽
def check_config(run, cfg, index, findings: list[Finding]) -> None:
    def err(rule, detail, evidence=""):
        findings.append(Finding("ERROR", "B 配置自洽", rule, detail, evidence))

    def warn(rule, detail, evidence=""):
        findings.append(Finding("WARN", "B 配置自洽", rule, detail, evidence))

    # B1 全厂码取值表 ↔ 代码常量
    table = cfg.get("unit_code_table") or {}
    units = set(table.get("units_1to9") or [])
    letters = set(table.get("units_10to16") or [])
    common = set(table.get("phase_common") or []) | set(table.get("multi_phase_common") or []) | set(table.get("plant_common") or [])
    free = set(table.get("free_use") or [])
    if common != set(run.COMMON_UNITS):
        err("unit_code_table.phase_common/multi_phase_common/plant_common",
            "配置里的公用字母集合与 run_audit.COMMON_UNITS 不一致",
            f"配置={sorted(common)}；代码={sorted(run.COMMON_UNITS)}")
    if free != set(run.FREE_UNIT_LETTERS):
        err("unit_code_table.free_use", "配置与 run_audit.FREE_UNIT_LETTERS 不一致",
            f"配置={sorted(free)}；代码={sorted(run.FREE_UNIT_LETTERS)}")
    mapping = table.get("unit_mapping") or {}
    if {k: v for k, v in mapping.items() if k.isalpha()} != run.UNIT_LETTER_MAP:
        err("unit_code_table.unit_mapping", "配置与 run_audit.UNIT_LETTER_MAP 不一致")
    if not units <= set("123456789"):
        err("unit_code_table.units_1to9", "1~9 号机组取值应只含数字 1~9", str(sorted(units)))

    # B2/B3 根父级与旧前缀
    if set(cfg.get("root_parent_values") or []) != set(run.ROOT_PARENTS):
        err("root_parent_values", "配置与 run_audit.ROOT_PARENTS 不一致",
            f"配置={cfg.get('root_parent_values')}；代码={sorted(run.ROOT_PARENTS)}")
    cfg_old = {k: v for k, v in (cfg.get("old_prefix_migrations") or {}).items() if k != "note"}
    if cfg_old != run.OLD_PREFIX:
        err("old_prefix_migrations", "配置与 run_audit.OLD_PREFIX 不一致",
            f"配置={cfg_old}；代码={run.OLD_PREFIX}")

    # B4 长度与骨架：段名形如 F1F2F3 表示"3 个字符位"，G 表示 1 位，AN1AN2AN3 表示 3 位
    grammar = cfg.get("code_grammar") or {}
    std_len = cfg.get("standard_device_length")
    skeleton = str(grammar.get("skeleton") or "")

    def _seg_width(part: str) -> int:
        tokens = re.findall(r"[A-Z]{1,2}\d+", part)
        return len(tokens) if tokens else len(part)

    expanded = sum(_seg_width(part) for part in skeleton.split("|") if part)
    if expanded and expanded != std_len:
        err("code_grammar.skeleton", f"骨架展开 {expanded} 位与 standard_device_length={std_len} 不符", skeleton)

    # 骨架分段位置必须与 audit_template.SEG_12 一致（AI 建议编码的本地复核也读这一份）
    positions_cfg = grammar.get("segments_positions")
    seg12 = getattr(_load_module("audit_template", TEMPLATE), "SEG_12", None)
    if positions_cfg and seg12:
        cfg_tuples = [tuple(x) for x in positions_cfg]
        tpl_tuples = [tuple(x) for x in seg12]
        if cfg_tuples != tpl_tuples:
            err("code_grammar.segments_positions",
                "配置里的骨架分段位置与 audit_template.SEG_12 不一致",
                f"配置={cfg_tuples}；模板={tpl_tuples}")
    if std_len not in set(cfg.get("accepted_code_lengths") or []):
        err("accepted_code_lengths", "标准 12 位不在可接受码长集合里")

    # B5 编号段语义
    seg_cfg = cfg.get("number_segment_semantics") or {}
    seen: list[tuple[int, int]] = []
    for item in seg_cfg.get("segments") or []:
        lo, hi = (item.get("range") or [None, None])[:2]
        if lo is None or hi is None or not (0 <= lo <= hi <= 999):
            err("number_segment_semantics.segments", f"区间非法：{item.get('range')}")
            continue
        for olo, ohi in seen:
            if not (hi < olo or lo > ohi):
                err("number_segment_semantics.segments", f"区间重叠：{lo}-{hi} 与 {olo}-{ohi}")
        seen.append((lo, hi))
    bands = {f"{lo}-{hi}" for lo, hi in seen}
    for key in (seg_cfg.get("name_keywords") or {}):
        if key not in bands:
            err("number_segment_semantics.name_keywords", f"名称关键词的段 {key} 不在 segments 区间内")

    # B6 特殊码必须在字典里
    known_dev = dict(index.get("device_codes") or {})
    known_dev.update(index.get("device_codes_huodian") or {})
    known_sys = dict(index.get("system_codes") or {})
    known_sys.update(index.get("system_codes_huodian") or {})
    special = []
    special += list((cfg.get("valve_device_codes") or {}).items())
    special += [(k, v) for k, v in (cfg.get("pipe_device_code") or {}).items() if k != "note"]
    sc = cfg.get("special_device_codes") or {}
    special += [(sc.get("coil_device_code"), "线圈设备分类码")]
    for code, _desc in special:
        if not code or code == "note":
            continue
        if code not in known_dev:
            err("特殊设备码", f'配置里用了设备码 "{code}"，但它不在分类码索引里（国标与企标都没有）')
    for key, code in (("power_output_transformer_system", sc.get("power_output_transformer_system")),):
        if code and code not in known_sys:
            err("特殊系统码", f'配置里用了系统码 "{code}"，但它不在分类码索引里')
    # B7 供应系统字母
    letters_cfg = cfg.get("supply_system_letters") or {}
    all_groups = set(index.get("system_s1_groups") or {}) | set(index.get("system_s1_groups_huodian") or {})
    for key in ("independent_s1", "same_group_s2", "same_subgroup_s3"):
        bad = [c for c in (letters_cfg.get(key) or []) if c not in all_groups]
        if bad:
            err("supply_system_letters", f"{key} 里的字母不在系统主组取值内：{bad}")

    # B8 索引文件存在且可解析
    if not CODE_INDEX.is_file():
        err("code_grammar.code_index_file", "分类码索引文件不存在，KKS-32/33/36 会静默失效", str(CODE_INDEX))
    declared = ((cfg.get("code_grammar") or {}).get("code_index_file") or "")
    if declared and not (PROJECT / declared).is_file():
        err("code_grammar.code_index_file", f"配置声明的索引路径不存在：{declared}")

    # B9 禁用字母与允许字符集的关系（这是一条"看似冲突其实互补"的约定，必须写清楚）
    allowed = str(cfg.get("allowed_code_characters") or "")
    forbidden = set(grammar.get("forbidden_letters") or [])
    if forbidden and not forbidden <= set("ABCDEFGHIJKLMNOPQRSTUVWXYZ"):
        err("code_grammar.forbidden_letters", "禁用字母表里出现了非 A-Z 的字符")
    if allowed and forbidden:
        widen = [c for c in forbidden if f"{c}-{c}" in allowed]
        if widen and not grammar.get("forbidden_reason"):
            warn("allowed_code_characters",
                 f"允许字符集 {allowed} 覆盖了禁用字母 {sorted(widen)}，需保证由 KKS-33 单独拦禁用字母",
                 "该关系已在 code_grammar.forbidden_reason 里说明")


# --------------------------------------------------------------- C 跨标准
def check_cross_standard(run, cfg, index, findings: list[Finding]) -> None:
    def err(rule, detail, evidence=""):
        findings.append(Finding("ERROR", "C 跨标准", rule, detail, evidence))

    div = index.get("cross_standard_divergences") or {}
    gb_dev = index.get("device_codes") or {}
    hd_dev = index.get("device_codes_huodian") or {}
    gb_sys = index.get("system_codes") or {}
    hd_sys = index.get("system_codes_huodian") or {}

    if not hd_dev:
        err("分类码索引", "索引里没有火电导则的设备码字典；项目默认标准是火电导则，"
                          "缺了它 KKS-33 会退化成只按国标判（整片误报）")
        return

    # C1 已登记分歧必须与索引里实际算出来的差异一致
    for kind, gb, hd in (("system", gb_sys, hd_sys), ("device", gb_dev, hd_dev)):
        entries = div.get(kind) or []
        for item in entries:
            code = str(item.get("code") or "")
            if code not in gb or code not in hd:
                err(f"cross_standard_divergences.{kind}",
                    f'登记的码 "{code}" 在两套字典里并非都存在')
                continue
            if item.get("gb50549") == item.get("huodian"):
                err(f"cross_standard_divergences.{kind}",
                    f'登记的码 "{code}" 两方定义相同，不应算分歧')
        if not entries:
            findings.append(Finding("WARN", "C 跨标准", f"cross_standard_divergences.{kind}",
                                    "两套标准的这一层没有登记任何分歧；若确认无差异可忽略"))

    # C2 主组级分歧必须已在 config 有说明（DIV-01/04）
    main_dev = div.get("device_main_groups") or []
    main_sys = div.get("system_main_groups") or []
    divergences_cfg = cfg.get("standard_divergences") or {}
    text = json.dumps(divergences_cfg, ensure_ascii=False)
    if main_dev and "设备分类码主组" not in text and "DIV-01" not in text:
        err("standard_divergences", f"索引里有 {len(main_dev)} 个设备分类码主组级分歧，但 config 没有登记", 
            ", ".join(x["code"] for x in main_dev[:8]))
    if main_sys and "DIV-04" not in text:
        err("standard_divergences", f"索引里有 {len(main_sys)} 个系统分类码主组级分歧，但 config 没有登记",
            ", ".join(x["code"] for x in main_sys[:8]))

    # C3 结构层必须两边一致（国标 4.2.1 注 与 火电导则 5.1.3.1 都禁用 I/O）
    if HUODIAN_JSON.is_file():
        hd = json.loads(HUODIAN_JSON.read_text(encoding="utf-8"))
        clauses = {c["no"]: c["text"] for c in hd.get("clauses") or []}
        text513 = " ".join(t for no, t in clauses.items() if no.startswith("5.1.3"))
        for letter in (cfg.get("code_grammar") or {}).get("forbidden_letters") or []:
            if letter not in text513:
                err("code_grammar.forbidden_letters",
                    f"火电导则 5.1.3.x 里没有出现禁用字母 {letter}，结构层「两标准一致」的说法不成立")
    else:
        findings.append(Finding("WARN", "C 跨标准", "rules_huodian.json",
                                "本地没有火电导则抽取结果，跳过结构层互证（本地跑 tools/extract_rules_huodian.py 生成）"))

    # C4 阀门/管道分段规则的适用性必须写明
    valve = cfg.get("valve_device_codes") or {}
    applies = set(((cfg.get("number_segment_semantics") or {}).get("applies_to_device_codes")) or [])
    hd_meanings = {code: hd_dev.get(code, "") for code in applies}
    for code, meaning in hd_meanings.items():
        if meaning and any(w in meaning for w in ("液位", "水分", "压力", "配电", "插座")):
            err("number_segment_semantics",
                f'KKS-35 适用的 "{code}" 在火电导则里含义是「{meaning[:16]}」，与国标阀门语义冲突，'
                "必须在本规则的出处里写明仅适用于按国标/中煤编码的项目")
    if applies and not valve:
        findings.append(Finding("WARN", "C 跨标准", "valve_device_codes", "编号段规则声明了适用码但没有阀门码表"))


# --------------------------------------------------------------- D 同条件重复
PROBES = [
    ("00LAC10AP008", "给水泵", "企标泵码"),
    ("00XAG10GG001", "轴流风机", "两套标准都没有的扩展码"),
    ("00LBA10GP001", "照明检修配电箱", "同码异义（GP）"),
    ("00LBA10QI001", "给水阀", "设备码含禁用字母 I"),
    ("00XAG10AN001", "轴流风机", "企标风机码"),
    ("00LAB10AA001", "电动阀门", "企标阀门码"),
    ("00LBA10WP420", "主蒸汽管道疏水管", "管道编号段"),
    ("00LBA10QM350", "截止阀取样管", "阀门编号段与名称矛盾"),
    ("00QQQ10AP001", "给水泵", "系统码主组不存在"),
    ("00LBA10ap001", "给水泵", "设备码小写"),
]


def check_overlap(template, findings: list[Finding]) -> None:
    cols = {"kks": 0, "name": 1}
    rules = [("KKS-32", template.system_class_code_legality),
             ("KKS-33", template.device_class_code_legality),
             ("KKS-35", template.valve_pipe_number_segment_check),
             ("KKS-36", template.standard_divergence_check)]
    hits: dict[int, set[tuple[str, str]]] = {}
    for idx, (code, name, _note) in enumerate(PROBES):
        row = [code, name]
        for rid, fn in rules:
            for item in fn(cols, [row]):
                sev = item["severity"] if isinstance(item, dict) else item[4]
                hits.setdefault(idx, set()).add((rid, sev))
    allow = set(COMPLEMENTARY_PAIRS) | {(b, a) for a, b in COMPLEMENTARY_PAIRS}
    for idx, found in sorted(hits.items()):
        ids = {rid for rid, _ in found}
        if len(ids) < 2:
            continue
        for a in sorted(ids):
            for b in sorted(ids):
                if a >= b or (a, b) in allow:
                    continue
                findings.append(Finding(
                    "ERROR", "D 同条件重复", f"{a} + {b}",
                    f"探针行 {PROBES[idx][0]}（{PROBES[idx][2]}）同时被两条规则命中，但未登记为互补关系",
                    f"{a} + {b}"))


# --------------------------------------------------------------- 主流程
def run_checks(*, quiet: bool = False) -> tuple[list[Finding], dict]:
    run = _load_module("run_audit", RUN_AUDIT)
    template = _load_module("audit_template", TEMPLATE)
    cfg = json.loads(RULES_JSON.read_text(encoding="utf-8"))
    index = json.loads(CODE_INDEX.read_text(encoding="utf-8"))

    findings: list[Finding] = []
    check_registry(run, cfg, findings)
    check_config(run, cfg, index, findings)
    check_cross_standard(run, cfg, index, findings)
    check_overlap(template, findings)

    stats = {
        "规则定义": len(run.RULE_DEFINITIONS),
        "规则出处键": len(run.RULE_SOURCES),
        "代码里产出的规则号": sum(
            1 for rid in run.RULE_DEFINITIONS if rid in RUN_AUDIT.read_text(encoding="utf-8")),
        "国标系统码": len(index.get("system_codes") or {}),
        "企标系统码": len(index.get("system_codes_huodian") or {}),
        "国标设备码": len(index.get("device_codes") or {}),
        "企标设备码": len(index.get("device_codes_huodian") or {}),
        "已登记码级分歧": len((index.get("cross_standard_divergences") or {}).get("system") or [])
                    + len((index.get("cross_standard_divergences") or {}).get("device") or []),
        "互补关系登记": len(COMPLEMENTARY_PAIRS),
    }
    if not quiet:
        errors = [f for f in findings if f["level"] == "ERROR"]
        warns = [f for f in findings if f["level"] == "WARN"]
        print("=" * 74)
        print("KKS 编码规则一致性核查")
        print("=" * 74)
        for key, value in stats.items():
            print(f"  {key:<14} {value}")
        print()
        if findings:
            for f in findings:
                mark = "✗ ERROR" if f["level"] == "ERROR" else "! WARN "
                print(f"{mark}  [{f['area']}] {f['rule']}")
                print(f"          {f['detail']}")
                if f["evidence"]:
                    print(f"          证据：{f['evidence']}")
        else:
            print("未发现冲突。")
        print()
        print(f"结论：ERROR {len(errors)} 项，WARN {len(warns)} 项。")
    return findings, stats


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--json", default="", help="把核查结果写入指定 JSON 文件")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()
    findings, stats = run_checks(quiet=args.quiet)
    if args.json:
        out = Path(args.json)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps({"stats": stats, "findings": findings}, ensure_ascii=False, indent=1),
                       encoding="utf-8")
        print(f"-> {out}")
    return 1 if any(f["level"] == "ERROR" for f in findings) else 0


if __name__ == "__main__":
    sys.exit(main())
