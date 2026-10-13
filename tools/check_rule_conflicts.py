# -*- coding: utf-8 -*-
"""规则一致性核查：把「编码规则有没有冲突」变成一条能自动跑、能当门禁用的命令。

为什么需要它
------------
规则来自 2 份**现行**标准文档（GB/T 50549-2020、中煤细则），落地到 5 个地方
（config/kks_rules.json、run_audit.py、audit_template.py、SKILL.md、config/kks_code_index.json）。
任何一处改动都可能让"同一件事被两条规则用不同依据判定"。

2026-10-08 就真出过一次：KKS-32/33 的分类码字典里混进了《火电厂标识系统（KKS）编码导则》
（Q/WL 2015，继承已废止的 GB/T 50549-2010 体系）的码表，导致同一码出现两种含义
（GP：2020 版是泵、2010 版体系是照明检修配电箱）。现已按"只用现行标准"收敛。

检查项
------
A 注册一致性  规则号在「定义表 / 出处表 / 代码 / SKILL.md」四处必须对齐，不能有幽灵规则。
B 配置自洽    配置里的常量必须与代码里的常量一致，码表/区间/字典不得互相矛盾。
C 版本纪律    只允许引用现行标准；已废止文档不得出现在依据或字典里（登记说明除外）。
D 条款真实性  引用的条款号必须在抽取出的标准语料里真实存在（防止虚构出处）。
E 同条件重复  同一行被多条规则以不同等级命中时，必须已登记为"互补"，否则算冲突。

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
README = PROJECT / "README.md"
RUN_AUDIT = PROJECT / "run_audit.py"
UPLOAD_PAGE_SRC = PROJECT / "server.py"
GB_JSON = PROJECT / "kks-audit" / "references" / "rules_gb50549.json"
ZMEI_JSON = PROJECT / "kks-audit" / "references" / "rules_zhongmei.json"

RULE_ID_RE = re.compile(r"KKS-[0-9A-Za-z]+(?:/\d+[A-Za-z]?)*")
# 已废止体系的关键词：出现在依据/字典/文档里即为违规（standard_basis.excluded 的登记说明除外）
RETIRED_TOKENS = ("火电导则", "风电场标识系统", "光伏发电站标识系统", "GB/T 50549-2010", "50549-2010")
# "这行是在说明某条规则/某个版本已被删除"，用于排除版本纪律说明本身的误报
EXCUSE_WORDS = ("废止", "移除", "已排除", "不得", "不再", "旧体系", "旧版", "历史", "禁止",
                "删除", "已并入", "幽灵", "excluded", "deprecated")


def _expand_group_ids(ids: set[str]) -> set[str]:
    """展开组键的成员号。

    组键写法是 `KKS-17/22`（另一半省略 "KKS-" 前缀），所以按 "/" 切分后
    **后续片段必须补回前缀**——否则会得到 "22" 这种半截号，成员判定永远不成立。
    """
    out = set(ids)
    for key in list(ids):
        parts = key.split("/")
        if len(parts) < 2:
            continue
        head = parts[0]
        prefix = head[: head.index("-") + 1] if "-" in head else ""
        out.add(head)
        for part in parts[1:]:
            out.add(part if "-" in part else prefix + part)
    return out


def _clean_doc_lines(path: Path) -> str:
    """按**段落**过滤掉"已删除/已废止"的说明，返回剩下可能真正作为依据引用的文本。

    按段落而不是按行：说明一条规则"已删除"时，规则号与"已删除"往往不在同一行
    （例如编号在上一行、结论在下一行），按行过滤会漏掉。
    """
    kept = []
    for block in re.split(r"\n\s*\n", path.read_text(encoding="utf-8")):
        if any(w in block for w in EXCUSE_WORDS):
            continue
        kept.append(block)
    return "\n".join(kept)

# 已知互补关系：同一行同时命中属于「不同维度」，不是规则冲突。
# 每条都必须写明为什么互补，避免这里变成"遇到告警就加白名单"。
COMPLEMENTARY_PAIRS = {
    ("KKS-29", "KKS-33"): "KKS-29 是 OCR 易混字符待核（提示人工核对是否为数字误识别）；"
                          "KKS-33 是设备分类码含禁用字母 I/O（标准明文禁止）。同一事实两种处置，互补不冲突。",
    ("KKS-06", "KKS-33"): "KKS-06 判全厂码与整体结构位取值；KKS-33 判设备分类码，作用位不同。",
    ("KKS-04", "KKS-30"): "KKS-04 判字符合法性；KKS-30 判首尾空白/控制符等文本卫生，症状不同。",
    ("KKS-30", "KKS-33"): "编码文本卫生与设备分类码取值，作用维度不同。",
    ("KKS-05", "KKS-17/22"): "KKS-05 判长度异常；KKS-17/22 对 12 位以上的扩展码做定性分流，阈值相邻不矛盾。",
    ("KKS-04b", "KKS-29"): "KKS-04b 判分段字符类型；KKS-29 判易混字符，前者定类型后者提待核。",
    ("KKS-27", "KKS-30"): "KKS-27 判名称语义可解性；KKS-30 判文本卫生，同列不同维度。",
    ("KKS-10", "KKS-09"): "自引用环与前缀一致性，同一层级关系上的两种缺陷。",
    ("KKS-25", "KKS-25b"): "同名多码与同物异名，聚簇口径不同。",
    ("KKS-32", "KKS-33"): "系统分类码与设备分类码分处不同位段，互不覆盖。",
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
    extra_src = set(run.RULE_SOURCES) - defs
    if extra_src:
        findings.append(Finding("ERROR", "A 注册一致性", ", ".join(sorted(extra_src)),
                                "有规则出处但没有规则定义"))

    # 出处是"规则号→依据"的一对一表；取不到就是漏洞，必须补或登记为无专条依据。
    declared_generic = set(((cfg.get("generic_rules_without_clause") or {}).get("rules")) or [])
    generic = []
    for rid in sorted(defs):
        if rid in declared_generic:
            continue
        src = run.rule_source(rid)
        if not src or src == run.GENERIC_SOURCE:
            generic.append(rid)
    if generic:
        findings.append(Finding("ERROR", "A 注册一致性", f"{len(generic)} 条规则无专条依据且未登记",
                                "要么补 RULE_SOURCES，要么把规则号加进 config 的 "
                                "generic_rules_without_clause.rules（登记为「标准无对应条款的数据治理要求」）",
                                "、".join(generic)))
    stale = sorted(declared_generic - defs)
    if stale:
        findings.append(Finding("ERROR", "A 注册一致性", "generic_rules_without_clause",
                                "登记了不存在的规则号", "、".join(stale)))

    # 代码里引用的规则号（组键的成员单独出现时不算幽灵）
    src_text = RUN_AUDIT.read_text(encoding="utf-8")
    literal_ids = {i for i in _rule_ids_in_source(RUN_AUDIT) | _rule_ids_in_source(TEMPLATE)
                   if not i.startswith("KKS-TREE")}
    members = _expand_group_ids(defs)
    ghosts = {i for i in literal_ids if i not in defs and i not in members}
    if ghosts:
        findings.append(Finding("ERROR", "A 注册一致性", ", ".join(sorted(ghosts)),
                                "代码里引用了未登记的规则号（幽灵规则）"))

    # config 的 rule_sources 必须是「已实现」的规则
    cfg_rules = set((cfg.get("rule_sources") or {})) - {"note"}
    orphan_cfg = cfg_rules - defs
    if orphan_cfg:
        findings.append(Finding("ERROR", "A 注册一致性", ", ".join(sorted(orphan_cfg)),
                                "config/kks_rules.json 的 rule_sources 里登记了未实现的规则号；"
                                "未实现的条款应放在 reference_clauses"))

    # SKILL.md 提到的规则号必须在定义表里（"该规则已删除"的说明行不算）
    skill_ids = {m.group(0) for m in RULE_ID_RE.finditer(_clean_doc_lines(SKILL_MD))}
    skill_ghosts = {i for i in skill_ids if i not in defs and i not in members}
    if skill_ghosts:
        findings.append(Finding("ERROR", "A 注册一致性", ", ".join(sorted(skill_ghosts)),
                                "SKILL.md 把未登记的规则号当作现行规则描述（文档与代码漂移）；"
                                "若只是说明它已删除，请在同一行写明「已删除/已并入」"))

    # 反向：定义了但代码里找不到产出点（幽灵规则）。只有真正出现在 add_issue/_append_skill_issue
    # 调用里的规则号才算"已实现"；组键（KKS-17/22）的成员通过组键产出，也算已实现。
    called = _expand_group_ids({rid for rid in defs
                                if f'rule_id="{rid}"' in src_text or f"rule_id='{rid}'" in src_text})
    # Tuple dispatch and conditional expressions also produce registered rules.
    for node in ast.walk(ast.parse(src_text)):
        if isinstance(node, ast.Call):
            for keyword in node.keywords:
                if keyword.arg == "rule_id":
                    called.update(child.value for child in ast.walk(keyword.value)
                                  if isinstance(child, ast.Constant) and isinstance(child.value, str)
                                  and child.value in defs)
    # 流程口径 / 策略说明类规则本就不产出问题行，已在 config 登记
    no_finding = set(((cfg.get("process_only_rules") or {}).get("rules")) or {})
    no_finding |= set(((cfg.get("policy_only_rules") or {}).get("rules")) or {})
    pure_defined = sorted(i for i in defs if i not in called and i not in no_finding)
    if pure_defined:
        findings.append(Finding("WARN", "A 注册一致性", ", ".join(pure_defined),
                                "已在 RULE_DEFINITIONS 登记，但代码里找不到产出点。"
                                "若确实产出问题行请补 add_issue；若是流程口径/策略说明，"
                                "请登记进 config 的 process_only_rules / policy_only_rules"))
    # 反向的漏洞：登记为"不产出问题行"却仍在产出
    contradictory = sorted(i for i in no_finding if i in called)
    if contradictory:
        findings.append(Finding("ERROR", "A 注册一致性", ", ".join(contradictory),
                                "登记为「不产出问题行」的规则实际有产出点，登记与实现矛盾"))


# --------------------------------------------------------------- B 配置自洽
def check_config(run, cfg, index, findings: list[Finding]) -> None:
    def err(rule, detail, evidence=""):
        findings.append(Finding("ERROR", "B 配置自洽", rule, detail, evidence))

    def warn(rule, detail, evidence=""):
        findings.append(Finding("WARN", "B 配置自洽", rule, detail, evidence))

    # B1 全厂码取值表 ↔ 代码常量
    table = cfg.get("unit_code_table") or {}
    expected = {"single_units": [f"{n:02d}" for n in range(1,61)],
                "pair_common": [f"{n:02d}" for n in range(61,91)],
                "agreed_common": [f"{n:02d}" for n in range(91,100)], "plant_common": ["00"]}
    for key, values in expected.items():
        if table.get(key) != values:
            err("unit_code_table." + key, "未对齐GB/T 50549-2020表4.3.1-2", str(table.get(key)))
    if run.code_unit("07HAC10GP001") != 7 or run.code_unit("61HAC10GP001") is not None:
        err("unit_parser", "机组解析不得将两机共用61读成6号机")

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
    if std_len not in set(cfg.get("accepted_code_lengths") or []):
        err("accepted_code_lengths", "标准 12 位不在可接受码长集合里")

    # B5 骨架分段位置必须与 audit_template.SEG_12 一致（AI 建议编码的本地复核也读这一份）
    positions_cfg = grammar.get("segments_positions")
    seg12 = getattr(_load_module("audit_template", TEMPLATE), "SEG_12", None)
    if positions_cfg and seg12:
        cfg_tuples = [tuple(x) for x in positions_cfg]
        tpl_tuples = [tuple(x) for x in seg12]
        if cfg_tuples != tpl_tuples:
            err("code_grammar.segments_positions",
                "配置里的骨架分段位置与 audit_template.SEG_12 不一致",
                f"配置={cfg_tuples}；模板={tpl_tuples}")

    # B6 编号段语义
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

    # B7 特殊码必须在字典里
    known_dev = dict(index.get("device_codes") or {})
    known_sys = dict(index.get("system_codes") or {})
    special = list((cfg.get("valve_device_codes") or {}).items())
    special += [(k, v) for k, v in (cfg.get("pipe_device_code") or {}).items() if k != "note"]
    sc = cfg.get("special_device_codes") or {}
    special += [(sc.get("coil_device_code"), "线圈设备分类码")]
    for code, _desc in special:
        if not code or code == "note":
            continue
        if code not in known_dev:
            err("特殊设备码", f'配置里用了设备码 "{code}"，但它不在分类码索引里')
    if sc.get("power_output_transformer_system") and sc["power_output_transformer_system"] not in known_sys:
        err("特殊系统码", f'配置里用了系统码 "{sc["power_output_transformer_system"]}"，但它不在分类码索引里')
    letters_cfg = cfg.get("supply_system_letters") or {}
    all_groups = set(index.get("system_s1_groups") or {})
    for key in ("independent_s1", "same_group_s2", "same_subgroup_s3"):
        bad = [c for c in (letters_cfg.get(key) or []) if c not in all_groups]
        if bad:
            err("supply_system_letters", f"{key} 里的字母不在系统主组取值内：{bad}")

    # B8 索引文件存在
    if not CODE_INDEX.is_file():
        err("code_grammar.code_index_file", "分类码索引文件不存在，KKS-32/33 会静默失效", str(CODE_INDEX))
    declared = str(grammar.get("code_index_file") or "")
    if declared and not (PROJECT / declared).is_file():
        err("code_grammar.code_index_file", f"配置声明的索引路径不存在：{declared}")

    # B9 禁用字母与允许字符集的关系（互补约定，必须写清楚）
    allowed = str(cfg.get("allowed_code_characters") or "")
    forbidden = set(grammar.get("forbidden_letters") or [])
    if forbidden and not forbidden <= set("ABCDEFGHIJKLMNOPQRSTUVWXYZ"):
        err("code_grammar.forbidden_letters", "禁用字母表里出现了非 A-Z 的字符")
    if allowed and forbidden:
        widen = [c for c in forbidden if f"{c}-{c}" in allowed]
        if widen and not grammar.get("forbidden_reason"):
            warn("allowed_code_characters",
                 f"允许字符集 {allowed} 覆盖了禁用字母 {sorted(widen)}，需保证由 KKS-33 单独拦禁用字母")


# --------------------------------------------------------------- C 版本纪律
def check_version_discipline(cfg, index, findings: list[Finding]) -> None:
    """只允许引用现行标准：已废止文档不得出现在依据/字典/文档里（登记说明除外）。"""
    def err(rule, detail, evidence=""):
        findings.append(Finding("ERROR", "C 版本纪律", rule, detail, evidence))

    basis = cfg.get("standard_basis") or {}
    if not basis:
        err("standard_basis", "config 缺少 standard_basis（现行依据与已排除版本的登记），"
                              "没有它就无法机器判定「规则是否引用了废止版本」")
    else:
        if not basis.get("current"):
            err("standard_basis.current", "没有登记任何现行依据")
        if not basis.get("excluded"):
            err("standard_basis.excluded", "没有登记已排除的旧版本，无法追溯为什么移除")
        for item in basis.get("current") or []:
            f = str(item.get("file") or "")
            if f and not (PROJECT / f).exists():
                err("standard_basis.current", f'登记的现行依据文件不存在：{f}')

    # 索引里不得残留旧体系字典
    for key in index:
        if key.endswith("_huodian") or "divergence" in key:
            err("config/kks_code_index.json", f'索引里残留旧体系数据：{key}'
                '（旧体系字典与分歧清单已废弃，须重跑 tools/build_code_index.py）')
    if index.get("standard") and "50549-2020" not in str(index.get("standard")):
        err("config/kks_code_index.json", "索引里的 standard 字段不是 GB/T 50549-2020",
            str(index.get("standard"))[:80])

    # 代码/文档里不得把废止版本当作依据。
    # 逐行判断：只有"提到旧版本但没有明确标注它是废止/已移除"的行才算违规，
    # 这样版本纪律说明本身（"GB/T 50549-2010 已被 2020 版废止，故移除"）不会被误报。
    for path in (RUN_AUDIT, TEMPLATE, SKILL_MD, README, UPLOAD_PAGE_SRC):
        if not path.is_file():
            continue
        bad_lines = [line.strip()[:80] for line in path.read_text(encoding="utf-8").splitlines()
                     if any(tok in line for tok in RETIRED_TOKENS)
                     and not any(w in line for w in EXCUSE_WORDS)]
        if bad_lines:
            findings.append(Finding("ERROR", "C 版本纪律", path.name,
                                    "把已废止版本当作依据引用（出现在依据/字典/规则表里，"
                                    "而不是版本纪律说明中）",
                                    " ｜ ".join(bad_lines[:4])))
    # config 里允许出现旧版本（standard_basis 就是登记用途，current 项也常需要说明
    # "本版已废止 2010 版"），但把 standard_basis 整段摘掉之后，剩下任何地方都不该再提到它们。
    outside = json.dumps({key: value for key, value in cfg.items() if key != "standard_basis"}, ensure_ascii=False)
    leaked = [tok for tok in RETIRED_TOKENS if tok in outside]
    if leaked:
        err("config/kks_rules.json",
            "在 standard_basis 之外仍把已废止版本当依据引用",
            "、".join(leaked))


# --------------------------------------------------------------- D 条款真实性
CLAUSE_RE = re.compile(r"\b(\d+(?:\.\d+){1,3})\b")
TABLE_RE = re.compile(r"表\s?(\d+(?:\.\d+){1,3})")
APPENDIX_RE = re.compile(r"附录\s?([A-Z])")


def check_clause_truth(cfg, findings: list[Finding]) -> None:
    """引用的条款号/表号必须在抽取出的标准语料里真实存在——防止虚构出处。"""
    if not GB_JSON.is_file():
        findings.append(Finding("WARN", "D 条款真实性", "rules_gb50549.json",
                                "本地没有国标抽取结果，跳过条款真实性检查"
                                "（本地跑 tools/extract_rules_gb50549.py 生成）"))
        return
    gb = json.loads(GB_JSON.read_text(encoding="utf-8"))
    clauses: set[str] = set()
    corpus_parts: list[str] = []
    for ch in gb.get("chapters", []) + gb.get("appendices", []):
        corpus_parts.append(str(ch.get("title", "")))
        for s in ch.get("sections", []):
            corpus_parts.append(str(s.get("title", "")))
            for cl in s.get("clauses", []):
                clauses.add(str(cl.get("no")))
                corpus_parts.append(str(cl.get("text", "")))
                corpus_parts.extend(str(x) for x in (cl.get("items") or []))
            for table in s.get("tables", []):
                for row in table:
                    corpus_parts.extend(str(c) for c in row)
    corpus = " ".join(corpus_parts)
    appendix_nums = {str(a.get("num")) for a in gb.get("appendices", [])}

    zmei_clauses: set[str] = set()
    if ZMEI_JSON.is_file():
        zm = json.loads(ZMEI_JSON.read_text(encoding="utf-8"))
        for s in zm.get("sections", []):
            for cl in s.get("clauses", []):
                zmei_clauses.add(str(cl.get("no")))

    def probe(where: str, text: str) -> None:
        # 先把"表5.1.5"这类表号整体摘掉，否则 "1.5" 会被误当成条款号
        cleaned = TABLE_RE.sub(" ", text)
        for m in CLAUSE_RE.finditer(cleaned):
            token = m.group(1)
            if token in clauses or token in zmei_clauses:
                continue
            if token.rsplit(".", 1)[0] in clauses:   # 允许"4.4.4-1"这种带条内项号的写法
                continue
            findings.append(Finding("ERROR", "D 条款真实性", where,
                                    f'引用了国标/中煤语料里不存在的条款号 "{token}"', text[:90]))

    for rid, item in (cfg.get("rule_sources") or {}).items():
        if rid == "note":
            continue
        for src in item.get("sources") or []:
            clause = str(src.get("clause") or "")
            if not str(src.get("standard", "")).startswith("GB/T"):
                continue
            probe(f"rule_sources.{rid}", clause)
            for m in TABLE_RE.finditer(clause):
                if m.group(1) not in corpus:
                    findings.append(Finding("ERROR", "D 条款真实性", f"rule_sources.{rid}",
                                            f'引用了国标语料里不存在的表号 "表{m.group(1)}"', clause))
            for m in APPENDIX_RE.finditer(clause):
                if f"附录 {m.group(1)}" not in appendix_nums:
                    findings.append(Finding("ERROR", "D 条款真实性", f"rule_sources.{rid}",
                                            f'引用了不存在的附录 "{m.group(1)}"', clause))

    import run_audit as R
    for rid, text in R.RULE_SOURCES.items():
        probe(f"RULE_SOURCES.{rid}", text)
        for m in TABLE_RE.finditer(text):
            if m.group(1) not in corpus:
                findings.append(Finding("ERROR", "D 条款真实性", f"RULE_SOURCES.{rid}",
                                        f'引用了国标语料里不存在的表号 "表{m.group(1)}"', text[:90]))


# --------------------------------------------------------------- E 同条件重复
PROBES = [
    ("00LAC10AP008", "给水泵", "旧体系泵码 AP"),
    ("00XAG10GG001", "轴流风机", "两套体系都未收录的码"),
    ("00LBA10GP001", "照明检修配电箱", "GP 在 2020 版是泵类"),
    ("00LBA10QI001", "给水阀", "设备码含禁用字母 I"),
    ("00XAB10GQ001", "轴流风机", "GQ 在 2020 版是风机"),
    ("00LBA10AA001", "电动阀门", "旧体系阀门码 AA"),
    ("00LBA10WP420", "主蒸汽管道疏水管", "管道编号段"),
    ("00LBA10QM350", "截止阀取样管", "阀门编号段与名称矛盾"),
    ("00QQQ10AP001", "给水泵", "系统码主组不存在"),
    ("00LBA10ap001", "给水泵", "设备码小写"),
]


def check_overlap(findings: list[Finding]) -> None:
    template = _load_module("audit_template", TEMPLATE)
    cols = {"kks": 0, "name": 1}
    rules = [("KKS-32", template.system_class_code_legality),
             ("KKS-33", template.device_class_code_legality),
             ("KKS-35", template.valve_pipe_number_segment_check)]
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
                    "ERROR", "E 同条件重复", f"{a} + {b}",
                    f"探针行 {PROBES[idx][0]}（{PROBES[idx][2]}）同时被两条规则命中，但未登记为互补关系",
                    f"{a} + {b}"))


# --------------------------------------------------------------- 主流程
def run_checks(*, quiet: bool = False) -> tuple[list[Finding], dict]:
    run = _load_module("run_audit", RUN_AUDIT)
    cfg = json.loads(RULES_JSON.read_text(encoding="utf-8"))
    index = json.loads(CODE_INDEX.read_text(encoding="utf-8"))

    findings: list[Finding] = []
    check_registry(run, cfg, findings)
    check_config(run, cfg, index, findings)
    check_version_discipline(cfg, index, findings)
    check_clause_truth(cfg, findings)
    check_overlap(findings)

    basis = cfg.get("standard_basis") or {}
    stats = {
        "现行依据": len(basis.get("current") or []),
        "已排除旧版本": len(basis.get("excluded") or []),
        "规则定义": len(run.RULE_DEFINITIONS),
        "规则出处": len(run.RULE_SOURCES),
        "无专条依据(已登记)": len((cfg.get("generic_rules_without_clause") or {}).get("rules") or []),
        "系统分类码": len(index.get("system_codes") or {}),
        "设备分类码": len(index.get("device_codes") or {}),
        "互补关系登记": len(COMPLEMENTARY_PAIRS),
    }
    if not quiet:
        errors = [f for f in findings if f["level"] == "ERROR"]
        warns = [f for f in findings if f["level"] == "WARN"]
        print("=" * 74)
        print("KKS 编码规则一致性核查")
        print("=" * 74)
        for key, value in stats.items():
            print(f"  {key:<18} {value}")
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
