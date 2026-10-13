# -*- coding: utf-8 -*-
"""设备名称生成规则（从竣工版实测反推，全部可复现、不调模型）。

生成顺序（越靠前越可靠）：
  R1 原名称存在              → 机组号纠正 + 长写法缩写（准确率最高，实测 31.6% 直接命中，其余是清洗差异）
  R2 部件级（父码 + 后缀）    → 父名称 + 空格 + 部件类型词（按 AA/AA+N 后缀查表）
  R3 有原名称但孤立            → 父名称 + 空格 + 原名称（实测 13280 行走这条）
  R4 柜/箱/屏类（UC/UC001）   → 父名称 + 空格 + 编码末2位 + 柜/箱/屏
  R5 兜底                    → 标 need_human，交由 AI 推断

关键实测依据：
- 部件后缀 → 类型词：-BC→CT、-FA→过电压保护器、-QC→接地刀闸、-QA→开关…
- 名称 = 父名称 + 空格 + X 的形态：长描述 13280、短词 2295、数字+柜/箱/屏 590
"""
from __future__ import annotations

import re

# ---------------- 部件后缀 → 类型词（从竣工版实测，可扩充） ----------------
# 键：连字符后的字母部分（部件分类码）；值：默认类型词
COMPONENT_WORDS: dict[str, str] = {
    "BC": "CT",
    "BF": "熔断器",
    "FA": "过电压保护器",
    "FT": "变送器",
    "KA": "继电器",
    "KF": "信号",
    "MA": "执行机构",
    "MM": "手动",
    "QA": "开关",
    "QB": "隔离开关",
    "QC": "接地刀闸",
    "QM": "阀门",
    "QP": "闸门",
    "RM": "止回阀",
    "BT": "热电偶",
    "BP": "压力变送器",
    "BF_": "过滤器",
    "XA": "端子箱",
    "XG": "端子箱",
}

# 不带连字符时（设备分类码）→ 类型词
DEVICE_WORDS: dict[str, str] = {
    "BT": "热电偶",
    "BP": "压力变送器",
    "BF": "流量计",
    "BL": "液位计",
    "UH": "控制柜",
    "UC": "柜",
    "FA": "避雷器",
    "QB": "刀闸",
    "QC": "地刀",
    "QM": "阀门",
    "QN": "调节阀",
    "WP": "管道",
    "GQ": "风机",
    "GP": "泵",
}

# 柜/箱/屏类设备分类码（名称里通常带柜号）
CABINET_CODES = ("UC", "UH", "UA", "UB", "UD")

# 连字符后缀：-AANN（部件分类码2位 + 编号2位）
DASH_SUFFIX_RE = re.compile(r"^-(?P<cls>[A-Z]{2})(?P<num>\d{2,3})$")
# 字母尾缀：AA001A（分相）→ 名称通常加 " A相/B相/C相"
PHASE_RE = re.compile(r"^(?P<base>.+?)(?P<phase>[ABC])$")

# 去掉机构/专业前缀，用于判断原名称是否已自带父级语义
UNIT_REF_RE = re.compile(r"([1-9]|1[0-6])\s*号\s*(机组|机)")

# ---------------------------------------------------------------- 标签串识别
# 名称栏里被填成**标签串**（不是设备名）的判据。
# 实测（热控点表 58,280 行，全表只命中 6 行 → 精度高）：`汽消 汽电 汽电 汽源`、
# `交通试 交通试 交通 交通试 交通试`、`气 水 水 油量 转`。
# 危害：兄弟行继承会把一个脏格当"基名"传播——实测 `01CSB02XA005` 的一个脏格
#       把整个 XA00x 组（两个机组共 10 行）叫成了 `1号机汽机2号控制电源柜 汽消 汽电 汽电 汽源`。
_DEVICE_NOUN = re.compile(
    r"柜|屏|箱|盘|泵|阀|门|机|器|风机|加热器|电机|开关|变送器|传感器|仪表|"
    r"管道|母管|装置|系统|空调|电缆|变压器|母线|槽|缸|罐|塔|炉|磨|筛|仓|平台|梯|泵组")
_TAG_TOKEN = re.compile(r"^.{1,3}$")


def tag_junk(name: str) -> bool:
    """名称栏里填的是**标签串**而不是设备名？（源表脏数据）

    判据：**含空格**、每个词 **≤3 字**、整串**没有任何设备名词**。
    三条件同时成立才算（宽松到只有真正的串行/标签才中招）。
    """
    text = str(name or "").strip()
    if " " not in text:
        return False
    if _DEVICE_NOUN.search(text):
        return False
    return all(_TAG_TOKEN.match(tok) for tok in text.split())


def component_type_word(code: str) -> str:
    """由编码推出部件/设备类型词。返回空串表示推不出。"""
    base = code.split(".", 1)[0]
    m = DASH_SUFFIX_RE.match(base[-5:]) if "-" in base else None
    if m:
        cls = m.group("cls")
        return COMPONENT_WORDS.get(cls, "")
    if len(base) >= 9:
        return DEVICE_WORDS.get(base[7:9], "")
    return ""


def phase_suffix(code: str) -> str:
    """13 位、末位为 A/B/C → 返回 'A相' 这类后缀（实测：名称里**不加空格**）。"""
    base = code.split(".", 1)[0]
    if len(base) == 13 and base[-1] in ("A", "B", "C"):
        return f"{base[-1]}相"
    return ""


def strip_phase_and_seq(name: str) -> str:
    """去掉名称末尾的分相与序号，得到"基线名"，用于兄弟行继承。"""
    text = re.sub(r"\s*[ABC]相$", "", name or "")
    text = re.sub(r"\s*\d{2,3}$", "", text)
    return text.strip()


def sibling_name(code: str, sibling_names: list[str]) -> tuple[str, bool]:
    """兄弟行继承：用同父级+同设备分类码的兄弟名称，改造成本行名称。

    实测（鄂州四期 12,268 行空名称行）：完全命中 3,815 行、仅空格差异 157 行。
    分相设备（末位 A/B/C）名称**不带序号**：`1号主变 高压侧PT A相`。
    """
    candidates = [n for n in sibling_names if n and not tag_junk(n)]
    if not candidates:
        return "", False
    base_code = code.split(".", 1)[0]
    base_name = strip_phase_and_seq(candidates[0])
    if not base_name:
        return "", False
    if len(base_code) == 13 and base_code[-1] in ("A", "B", "C"):
        return f"{base_name}{base_code[-1]}相", True
    # 非分相：若兄弟名里带序号，用本行序号替换
    m = re.search(r"(\d{2,3})$", candidates[0])
    if m:
        seq = base_code[-2:] if len(base_code) >= 12 else m.group(1)
        return f"{base_name}{seq}", True
    return base_name, True



def dash_tail(code: str) -> str:
    """取连字符后的部分（如 -BC01 → BC01）；没有则空串。"""
    base = code.split(".", 1)[0]
    return base.split("-", 1)[1] if "-" in base else ""


def strip_state_words(name: str) -> tuple[str, bool]:
    """剔除 DCS **状态/指令词**——名称只描述设备本体，不含运行状态与控制指令。

    来源：豆包样本实测的词表（`…已全开`→`…`、`打开X`→`X`）。
    ⚠️ 只在**结尾**剔状态、在**开头**剔命令动词；**不做整串替换**，
       否则会误伤「关闭挡板」这类**真实设备名**（踩过的坑）。
    """
    s = (name or "").strip()
    changed = False
    for w in STATE_HEAD:                      # 开头命令动词：打开X / 关闭X
        rest = s[len(w):].strip()
        if s.startswith(w) and len(rest) > 1 and (INSTANCE_HINT.search(rest) or len(rest) >= 6):
            s = rest
            changed = True
            break
    for w in STATE_TAIL:                      # 结尾状态：…已全开 / …故障
        if s.endswith(w) and len(s) > len(w) + 1:
            s = s[: -len(w)].strip()
            changed = True
            break
    return s, changed


def ensure_unit_prefix(name: str, code: str) -> tuple[str, bool]:
    """名称里**没有机组/分支标识**时，按全厂码补 `N号机` 前缀（表4.3.1-2，短写法）。

    依据：本表列名即「设备名称（**标准修改加机组号**）」，要求带机组号；
    实测竣工版口径是短写法 `1号机`（短:长 = 25,602:258 ≈ 99:1），故用短写法。
    已有标识（`7A…`/`1号机…`/`全厂…`）的一律不动。
    """
    if not name or len(code) < 2:
        return name, False
    if UNIT_MARK.search(name):
        return name, False
    from tree_rules import _unit_prefix          # 复用"全厂码→机组前缀"的国标映射
    prefix = _unit_prefix(code[:2])
    if not prefix:
        return name, False
    return f"{prefix}{name}", True


def generate_name(
    *,
    code: str,
    original_name: str,
    parent_name: str,
    parent_code: str,
) -> tuple[str, str, float, bool]:
    """生成设备名称。

    返回 (名称, 生成规则编号, 置信度, 是否需人核)。
    """
    from tree_rules import abbreviate_unit, normalize_name   # 复用同一套清洗

    original = normalize_name(original_name)
    pname = normalize_name(parent_name)
    # ⚠️ 名称栏里填的若是**标签串**（`汽消 汽电 汽电 汽源`），当作"没有名称"处理 →
    #    落到 R2/R4/R5（交 AI / 人工），而不是把标签串当设备名用（实测踩过）。
    if tag_junk(original):
        original = ""
    base = code.split(".", 1)[0]
    tail = dash_tail(base)
    phase = phase_suffix(base)

    # ---- R1：有原名称 → **清洗后**直接用（+分相后缀）----
    if original:
        name = original
        conf = 1.0
        rule = "R1-原名称"
        # ⚠️ 原名称**本身就是**状态/指令词（`远方控制`、`已全开`、`故障`）→ 没有设备本体信息
        #    → 不能生成（否则会产出 `1号机远方控制` 这种垃圾名）→ 交 AI/人工。
        if original in STATE_TAIL or original in STATE_HEAD or len(original) < 2:
            return "", "R1-原名称仅为状态/指令词，无设备本体信息", 0.2, True
        cleaned, hit = strip_state_words(name)       # 剔除 DCS 状态/指令词
        if hit:
            name, rule, conf = cleaned, "R1-原名称-剔状态词", 0.95
        # 剔完没有设备本体信息 → 同样不生成
        if len(name) < 2:
            return "", "R1-剔状态词后无设备本体信息", 0.2, True
        name = tidy_spaces(name)
        with_unit, added = ensure_unit_prefix(name, base)   # 补机组号（表4.3.1-2）
        if added:
            name, rule, conf = with_unit, rule + "+补机组号", min(conf, 0.95)
        if phase and not name.endswith(phase.strip()):
            name = name + phase
            rule = rule + "+分相"
            conf = min(conf, 0.9)
        return name, rule, conf, False

    # ---- R2：部件级（带连字符后缀）→ 父名称 + 类型词 ----
    if tail:
        word = component_type_word(base)
        if pname and word:
            return f"{pname}{word}", "R2-父名+部件类型词", 0.85, False
        if pname:
            return f"{pname} {tail}", "R2-父名+部件后缀", 0.55, True
        return "", "R2-无法生成", 0.2, True

    # ---- R4：柜类（父名 + 编码末2位 + 柜/箱/屏）----
    if len(base) >= 9 and base[7:9] in CABINET_CODES and pname:
        num = base[-2:]
        word = "柜" if base[7:9] in ("UC", "UH") else "屏"
        return f"{pname} {num}{word}", "R4-父名+柜号", 0.6, True

    # ---- R5：兜底，交 AI ----
    return "", "R5-待AI推断", 0.2, True


# ------------------------------------------------------------------ 文本微调规则
# 实测（竣工版比对，44,200 行）：
#   ① 一方是另一方子串（前缀/后缀差异）4,726 行 —— 多为"父名/系统名前缀"缺失
#   ② 仅标点/空格差异                84 行 —— 全是"机组号/分支号后多了个空格"
#   ③ 仅机组号写法差异                6 行 —— 名称中段缺机组指代
# 下面两条把 ②③ 补掉（纯规则，确定性）。
SPACE_AFTER_MARK = re.compile(r"([0-9]{1,2}[A-Z]?)\s+(?=[\u4e00-\u9fff])")
GEN_WORD = re.compile(r"(?<![0-9号])(发电机|汽轮机|锅炉)")

# ---- DCS 状态/指令词（名称只描述设备本体，不含运行状态与控制指令）----
# 来源：豆包样本实测词表 + 本表实测分布（已全开 1,176 / 已全关 1,170 / 远方控制 1,527 / 故障 2,941）。
# ⚠️ **不含「报警」**：`A送风机调节臂过力矩报警` 这类里「报警」是设备/信号名的一部分，
#    剔掉会改变设备含义（实测 552 行，其中 `方向振动高`/`给煤机超温` 都是被测对象）。
STATE_TAIL = ("已全开", "已全关", "已开", "已关", "远方控制", "故障")
STATE_HEAD = ("打开", "关闭")
# 判"开头的 打开/关闭 是命令行"的守卫：去掉动词后必须是**具体设备实例**
# （含数字/字母限定，或足够长）。否则 `关闭挡板`（真实设备名）会被误剔（实测踩过）。
INSTANCE_HINT = re.compile(r"[0-9A-Za-z]")
# 名称里"已有机组/分支标识"的判据（已有就不再补前缀）
UNIT_MARK = re.compile(r"\d+\s*号\s*机|\d{1,2}[A-D](?![A-Za-z])|公用|全厂")


def tidy_spaces(name: str) -> str:
    """去掉机组号/分支号与中文之间的多余空格：`7A 凝结水泵` → `7A凝结水泵`。"""
    return SPACE_AFTER_MARK.sub(r"\1", name or "")


def fill_unit_before_equipment(name: str, kks: str) -> tuple[str, bool]:
    """名称中段的「发电机/锅炉」缺机组号时按编码前缀补上。

    实测：`凝结水至发电机定子冷却水补水手动门` → `凝结水至7号发电机定子冷却水补水手动门`。
    只补"光杆"的发电机/汽轮机/锅炉；前面已经是「N号」的不动。
    """
    if not name or len(kks) < 2:
        return name, False
    from tree_rules import UNIT_BY_PREFIX

    want = UNIT_BY_PREFIX.get(kks[:2], "")
    if not want:
        return name, False
    changed = False

    def repl(match: re.Match[str]) -> str:
        nonlocal changed
        head = name[max(0, match.start() - 3):match.start()]
        if re.search(r"\d\s*号\s*$", head):
            return match.group(1)
        changed = True
        return f"{want}号{match.group(1)}"

    return GEN_WORD.sub(repl, name), changed
