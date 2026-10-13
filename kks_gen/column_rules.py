# -*- coding: utf-8 -*-
"""**列规则表**：声明"每一列能从哪些列推、按什么规则推、依据哪条国标"。

为什么要单独一个模块：回填原来硬编码了 4 列（设备名称/KKS编码/上级KKS编码/上级设备名称），
新增一列就要改主流程。改成规则表后，**新增列 = 加一条规则**。

## 三条纪律（都是实测踩出来的）

1. **只填原本为空的格**（业务铁律）：原表一开始有数据的格一律不动。
2. **只在有确定性依据时填**：没有依据就留空，并在「修编说明」里写清缺什么。**不猜**。
3. **每个规则都要能拿竣工版量化**（命中率），只保留有提升的。

## 当前规则

| 目标列 | 来源 | 规则 | 依据 |
|---|---|---|---|
| 设备名称 | 原名称 + 编码 | 剔 DCS 状态/指令词 → 补机组号 | 表4.3.1-2 |
| KKS编码（修编后） | 设计编码 | 段位化修编；**指令/测点行留空**（只有设备才有 KKS 编码）；**全表唯一**（撞码留空） | 图4.4.1、5.2.11、4.1.2 |
| 上级KKS编码 | 编码层级 | 取直接父级 | 图4.4.1 |
| **系统** | KKS编码 | S₁S₂S₃ 查附录A → **压成可读短名**；预留段留空 | 4.4.2、附录A |
| **设备类** | KKS编码 | T₁T₂ 查附录B → **概括成上位类目** | 4.4.3、附录B |
"""
from __future__ import annotations

import json
import pathlib
import re

CONFIG = pathlib.Path(__file__).resolve().parents[1] / "config"
_INDEX: dict | None = None


def _index() -> dict:
    """码表索引（模块级缓存）。"""
    global _INDEX
    if _INDEX is None:
        try:
            _INDEX = json.loads((CONFIG / "kks_code_index.json").read_text(encoding="utf-8"))
        except (OSError, ValueError):
            _INDEX = {}
    return _INDEX


# ------------------------------------------------------------------ 可读短名
# 让"系统"列**容易理解**：码表原文常常是一整句（如 LBS = 「用于主凝结水预加热系统的
# 抽汽管道系统 / 从(不包括):汽轮机出口…」），直接照搬没法看。
_RANGE_TAIL = re.compile(r"\s*/\s*从?\(?不?包括?\).*$")     # ` / 从(不包括):…` 的标识范围说明
_PARENS = re.compile(r"[（(【\[][^）)】\]]*[）)】\]]")          # 仅供测试/兼容；实际用 _clean_brackets
_OPEN = {"(": ")", "（": "）", "[": "]", "【": "】"}
_CLOSE = {v: k for k, v in _OPEN.items()}
_CORE_AFTER_DE = re.compile(r"^.*的([^的]{2,10}(?:系统|装置|设备|回路|管道|母管))$")
_LEAD_WORDS = ("用于", "包括", "由", "将", "对", "在")
# 电压等级系统（附录A 里「Un」系列）：`525kV ≤ Un < 765kV 输电`、`Un ≤ 400V 交流配电系统`。
# 直接照搬会超过长度上限被**截断成 `765kV≤Un<1000k`**（实测踩过）：
#   ① 电压式后面还有实质名称（`… 交流配电系统`）→ 去掉电压式，留名称（最易读）；
#   ② 电压式后面只剩用途词（`输电`/`系统`）→ 电压就是它的名字 → `525kV输电系统`。
_VOLT_A = re.compile(r"^(\d+(?:\.\d+)?)\s*(k?V)\s*[≤<≥>]\s*Un(?:\s*[≤<≥>]\s*\d+(?:\.\d+)?\s*k?V)?\s*")
_VOLT_B = re.compile(r"^Un\s*([≤<≥>])\s*(\d+(?:\.\d+)?)\s*(k?V)\s*")
# OCR 合并单元格残留：一格挤了「中压主供电系统1,电压等级1」两条 → 取逗号前的主名
_COMMA_SPLIT = re.compile(r"\s*[,，]\s*")
# 截断后残留的半截英文/数字记号（`…Un<1000k`）
_PARTIAL_TAIL = re.compile(r"[A-Za-z0-9≤<>=~.]+$")


def _clean_brackets(text: str) -> str:
    """剥掉括号里的补充说明，**保留括号外的正文**。

    ⚠️ 不能用正则：OCR 里常有**嵌套**（`输电 [标准电压 150(154)kV]`），半截正则会剥成
       `输电 kV]`（实测踩过）。用栈配对：
       · 正常配对的整段（含嵌套）内部丢弃、外部保留 → `凝结水泵(电动)系统` → `凝结水泵系统`
       · **不闭合**的开括号按尾部元数据丢弃（从它起到结尾）→ `输电 [标准电压…]` → `输电`
       · 无主的闭括号直接丢
    """
    out: list[str] = []
    stack: list[tuple[str, int]] = []        # (开括号, 它在 out 里的截断位置)
    for ch in text:
        if ch in _OPEN:
            stack.append((ch, len(out)))
        elif ch in _CLOSE:
            if stack and _OPEN[stack[-1][0]] == ch:
                stack.pop()
            elif not stack:
                continue                     # 无主的闭括号
        elif not stack:
            out.append(ch)
    if stack:
        del out[stack[0][1]:]
    return "".join(out)


def _kind_word(kind: str | None) -> str:
    """用途词补成「…系统」：`输电` → `输电系统`；空 → `系统`；已带 → 原样。"""
    kind = (kind or "").strip()
    if not kind:
        return "系统"
    return kind if kind.endswith("系统") else kind + "系统"


def _strip_voltage(text: str) -> tuple[str, str]:
    """剥掉开头的**电压等级限定**，返回 (剩余名称, 电压标签)；没有则返回 (原文, "")。

    `525kV ≤ Un < 765kV 输电` → (`输电`, `525kV`)
    `Un ≤ 400V 交流配电系统`   → (`交流配电系统`, `400V以下`)
    """
    m = _VOLT_A.match(text)
    if m:
        return text[m.end():].strip(), f"{m.group(1)}{m.group(2)}"
    m = _VOLT_B.match(text)
    if m:
        op, num, unit = m.group(1), m.group(2), m.group(3)
        word = "以上" if op in ">≥" else ("以下" if op in "<≤" else "")
        return text[m.end():].strip(), f"{num}{unit}{word}"
    return text, ""


def readable_name(cn: str, *, limit: int = 14) -> str:
    """把码表释义压成**可读短名**（保义、去套话、限长）。

    · 去掉 ` / 从(不包括):…至…` 这类**标识范围**说明（它不是名称）
    · 去掉括号里的补充（圆括号、方括号都有）
    · 电压等级段压成「下限 + 用途」：`525kV ≤ Un < 765kV 输电` → `525kV输电系统`
    · 合并单元格残留取主名：`中压主供电系统1,电压等级1` → `中压主供电系统1`
    · `用于…的抽汽管道系统` → 取「的」之后的核心名词 `抽汽管道系统`
    · 去掉引导词（用于/包括/由/将/对/在）

    实测：`LBS` → 「用于主凝结水预加热系统的抽汽管道系统」→ **`抽汽管道系统`**；
          `A` → 「电网与配电系统」（本来就短，不动）；`HAC` → 「省煤器系统」；
          `ABE` → 「525kV ≤ Un < 765kV 输电」→ **`525kV输电系统`**。
    """
    text = str(cn or "").strip()
    if not text:
        return ""
    text = _RANGE_TAIL.sub("", text)
    text, volt = _strip_voltage(text)                # 电压式在括号清洗**之前**剥（AEC 里有 `[标准电压…kV]`）
    if volt:
        rest = re.sub(r"\s+", "", _clean_brackets(text).strip())
        if len(rest) >= 4:                          # ① 后面还有实质名称 → 用名称
            text = rest
        else:                                       # ② 只剩用途词 → 电压就是名字
            return f"{volt}{_kind_word(rest)}"[:limit]
    text = _clean_brackets(text).strip()
    head = _COMMA_SPLIT.split(text, 1)[0].strip()
    if len(head) >= 3:
        text = head
    m = _CORE_AFTER_DE.match(text)
    if m:
        text = m.group(1)
    for w in _LEAD_WORDS:
        if text.startswith(w) and len(text) > len(w) + 1:
            text = text[len(w):]
            break
    text = re.sub(r"\s+", "", text)
    if len(text) > limit:                            # 限长；截断后别留半截记号
        text = _PARTIAL_TAIL.sub("", text[:limit])
    return text


# 附录A 的「后续标准预留…」= 国标**尚未定义**该码段 → 不是名称，不能当系统名写进去
_RESERVED = re.compile(r"^后续标准预留")


def is_reserved_system(s3: str) -> bool:
    """该 S₁S₂S₃ 是否落在国标**预留段**（附录A 未定义其含义）。"""
    cn = str((_index().get("system_codes") or {}).get(str(s3 or "").upper(), ""))
    return bool(_RESERVED.match(cn))


# ------------------------------------------------------------------ 上位类目
# 让"设备类"列容易理解：附录B 的释义是"功能描述 + 应用实例"（如
# QM = 「封闭外壳中可流动物质的流动转换 / 切断阀、截止阀、蝶阀…」），
# 人看的是**实例**（阀门），不是功能描述。所以按实例里的关键词概括成类目。
# 顺序即优先级（先命中先取）。
DEVICE_CLASSES: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("阀门", ("阀", "闸门", "挡板", "蝶阀", "截止阀", "疏水")),
    ("泵", ("泵", "排水器")),
    ("风机", ("风机", "通风机", "鼓风机", "引风机")),
    ("电动机", ("电动机", "电机", "马达")),
    ("变压器", ("变压器", "互感器", "电抗器")),
    ("开关／断路器", ("断路器", "隔离开关", "接触器", "开关", "刀闸")),
    ("柜／屏／箱", ("柜", "屏", "箱", "盘")),
    ("传感器／变送器", ("传感器", "变送器", "探测器", "继电器", "仪表")),
    ("换热器", ("换热器", "加热器", "冷却器", "冷凝器")),
    ("容器／罐", ("箱", "罐", "容器", "储")),
    ("管道／管件", ("管道", "管件", "母管")),
    ("电缆／电缆桥架", ("电缆", "桥架", "导线")),
)


def device_class(t12: str) -> str:
    """由设备分类码 T₁T₂ 概括出**上位类目**（如 `QM` → `阀门`）。

    ⚠️ 判据是**附录B 释义里的「应用实例」列表**（`cn / 实例…` 的右半），不是功能描述
       —— 国标附录B 的分类码是**功能分类**（`QM` = 「封闭外壳中可流动物质的流动转换」），
       人看的是实例（切断阀、截止阀、蝶阀…）。

    ⚠️ **不能"首次命中即取"**：同一实例列表里往往同时出现「泵」「阀」「开关」等词，
       首次命中会把 `GP`（提升泵、清洗泵…）判成「阀门」、把 `BP`（差压传感器、压力变送器、
       压力开关…）判成「开关」。改为**按命中次数取多数**（平手时按优先级）。
    """
    entry = (_index().get("device_codes") or {}).get(str(t12 or "").upper())
    if not entry:
        return ""
    text = str(entry)
    best, best_score = "", 0
    for label, keys in DEVICE_CLASSES:
        score = sum(text.count(k) for k in keys)
        if score > best_score:
            best, best_score = label, score
    return best


# ------------------------------------------------------------------ 列规则注册表
# (目标列的关键词, 推导函数名) —— 新增列只在这里加一行
COLUMN_RULES: dict[str, str] = {
    "系统": "system",
    "设备类": "device_class",
    "设备类型": "device_class",
}


def derive(rule: str, code: str) -> str:
    """按规则名推导一列的值（取不到返回空串）。

    ⚠️ 取不到就**留空**，不猜（纪律 2）。「后续标准预留」段同样留空——
       它是国标说"此码未定义"，不是系统名；该行由 `derive_note()` 出修编说明。
    """
    from kks_gen.identifiers import parse
    segments = parse(code).segments
    system = segments.get("sys", "")
    device = segments.get("dev", "")
    # A classification node may contain a system class but no system number.
    if not system:
        match = re.fullmatch(r"[BCEGKTUW]?[0-9]{2}[0-9]?([A-HJ-NP-Z]{3})", str(code or ""))
        system = match[1] if match else ""
    if rule == "system" and system:
        if is_reserved_system(system):
            return ""
        return readable_name(str((_index().get("system_codes") or {}).get(system, "")))
    if rule == "device_class" and device:
        return device_class(device)
    return ""


def derive_note(rule: str, code: str) -> str:
    """推导取不到值时的**修编说明**（说清缺什么、下一步做什么）。取到值则返回空串。"""
    from kks_gen.identifiers import parse
    system = parse(code).segments.get("sys", "")
    if not system:
        match = re.fullmatch(r"[BCEGKTUW]?[0-9]{2}[0-9]?([A-HJ-NP-Z]{3})", str(code or ""))
        system = match[1] if match else ""
    if rule == "system" and is_reserved_system(system):
        return (f"系统码 {system} 属国标预留段（GB/T 50549-2020 附录A 未定义其含义），"
                f"「系统」列留空，需人工确认")
    return ""
