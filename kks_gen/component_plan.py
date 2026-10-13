# -*- coding: utf-8 -*-
"""部件级（第 5 层扩展码）：生成与"应编未编"判定。

## 实测结论（鄂州四期竣工版，85,754 个设备 / 6,349 个带部件）

**"某设备该不该有部件"由「设备分类码 + 名称关键词」决定，与兄弟设备无关。**

1. 有部件的 6,349 个设备里 **5,088 个只有 1 个部件**，且它们的同父级同类型兄弟**也没有部件**。
   所以"照兄弟补部件"这条路不通——按它跑，一个都补不出来（这是正确的，别硬补）。
2. 真正可预测的是"该不该有"，信号很强：

   设备分类码 → 有部件比例：
     RA 100% / AM 100% / AT 99.7% / RB 95.5% / GP 86.2% / HW 81.9%
     QP 80% / QB 74.3% / QC 73.4% / TA 58.4% / QN 44.8% / GL 42.8%
     UH 26.4% / GM 21.9%

   名称关键词 → 有部件比例：
     调节门 91.4% / 地刀 75.0% / 刀闸 73.6% / 气动门 72.0% / 电动门 69.5%
     手动门 0.1% / 热电阻 0% / 端子箱 0% / CT 0%

3. 因此本模块的产物分两种用途：
   - **能生成**：设备分类码属于"高带部件率"且名称给出部件类型线索时，按类型词生成部件编码
   - **应编未编清单**：高带部件率的分类码下、表里却没有部件的设备 → 交第二步审核报出

⚠️ 绝不给"手动门""热电阻"这类**天然不带部件**的设备硬补部件——那是造假。
"""

from __future__ import annotations

import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field

from component_rules import COMPONENT_TYPE_WORD

# 名称关键词 → 部件分类码（用于在 MA/MM 之间做二次判定）
# 实测误差几乎全部是这两个码互串：电动/气动/调节 → MA，手动/手操 → MM
ACTUATOR_HINTS = {
    "MA": ("电动", "气动", "液动", "调节", "执行", "伺服", "电磁"),
    "MM": ("手动", "手操", "手轮"),
}
# 名称关键词 → 部件分类码（其他常见部件，按名称直接判定）
NAME_HINTS: tuple[tuple[str, str], ...] = (
    ("开关", "QA"),
    ("刀闸", "QB"),
    ("地刀", "QC"),
    ("接地刀", "QC"),
    ("变送器", "BP"),
    ("热电偶", "BT"),
    ("热电阻", "BT"),
    ("温度", "BT"),
    ("端子箱", "XG"),
    ("CT", "BC"),
    ("电流互感器", "BC"),
    ("避雷器", "FA"),
    ("熔断器", "BF"),
    ("滤网", "BF"),
    ("流量", "BF"),
    ("液位", "BL"),
    ("压力", "BP"),
)

# ------------------------------------------------------------------ 实测判据表
# 设备分类码 → 在竣工版中"带部件"的比例（样本 ≥30 才收录；实测 85,754 个设备）
DEVICE_CLASS_RATIO: dict[str, float] = {
    "RA": 1.00, "AM": 1.00, "AT": 0.997, "RB": 0.955, "GP": 0.862, "HW": 0.819,
    "AP": 0.80, "QP": 0.80, "QB": 0.743, "QC": 0.734, "TA": 0.584, "HU": 0.467,
    "QN": 0.448, "GL": 0.428, "VQ": 0.404, "UH": 0.264, "GM": 0.219, "GS": 0.154,
}

# 名称里出现这些词 → 天然不带部件，绝不补（实测：手动门 0.1%、热电阻 0%、端子箱 0%）
NO_COMPONENT_HINTS: tuple[str, ...] = (
    "手动门", "手动阀", "热电阻", "端子箱", "避雷器", "手动", "就地", "基础", "管道",
)

# 设备分类码 → 典型部件分类码（来自实测配对，仅在没有名称线索时使用）
TYPICAL_COMPONENTS: dict[str, list[str]] = {
    "UH": ["KF", "QA", "FC"],
    "UC": ["QA", "BC", "FA"],
    "QM": ["MA", "MM"],
    "QN": ["MA", "MM"],
    "GP": ["MA"],
    "GQ": ["MA"],
    "BF": ["QM"],
    "BP": ["QM"],
    "RB": ["QA"],
    "AT": ["MA"],
    "GL": ["UR"],
    "QB": ["QA"],
    "QC": ["QA"],
    "TA": ["MA"],
}



@dataclass
class ComponentPlan:
    """一个待生成的部件。"""
    parent_kks: str
    parent_name: str
    cls: str
    num: int
    name: str
    source: str
    confidence: float
    need_human: bool = True
    notes: list[str] = field(default_factory=list)

    @property
    def kks(self) -> str:
        return f"{self.parent_kks}-{self.cls}{self.num:02d}"


def _parse_component(kks: str, parent: str) -> tuple[str, int] | None:
    if "-" not in kks:
        return None
    tail = kks.split("-", 1)[1]
    if len(tail) < 4 or not tail[:2].isalpha() or not tail[2:].isdigit():
        return None
    return tail[:2], int(tail[2:])


def build_component_index(rows: list[dict]) -> dict[str, dict[str, int]]:
    """从已有数据建立「设备码 → {部件分类码: 最大编号}」。"""
    index: dict[str, dict[str, int]] = defaultdict(dict)
    for row in rows:
        kks = str(row.get("kks", "") or "")
        parent = str(row.get("parent_code", "") or "")
        parsed = _parse_component(kks, parent)
        if not parsed:
            continue
        cls, num = parsed
        prev = index[parent].get(cls, 0)
        index[parent][cls] = max(prev, num)
    return dict(index)


def _hint_from_name(name: str) -> str | None:
    """按名称关键词判定部件分类码（优先于兄弟签名，因为它直接说明部件是什么）。"""
    for keyword, cls in NAME_HINTS:
        if keyword in (name or ""):
            return cls
    return None


def _actuator_cls(name: str) -> str | None:
    """在 MA/MM 之间判定：名称含"手动"→MM，含"电动/气动/调节"→MA。"""
    for cls, hints in ACTUATOR_HINTS.items():
        if any(h in (name or "") for h in hints):
            return cls
    return None


def plan_components(
    devices: list[dict],
    component_index: dict[str, dict[str, int]],
    *,
    allow_speculative: bool = False,
) -> list[ComponentPlan]:
    """规划要补的部件 —— **默认返回空**（见下方实测结论）。

    ## 🔴 重要负面结论（实测，2026-10-09，两次试验都是 0% 精确率）

    试过两条路，都不成立，**不要再用它们自动补部件**：

    | 假设 | 做法 | 实测结果 |
    |---|---|---|
    | 兄弟参照 | 同父级+同类设备的部件签名套过来 | 留一验证 87.7% 一致，但**全量一个都补不出**：有部件的 6,349 个设备里 5,088 个只有 1 个部件，且兄弟也没有部件 |
    | 类型推断 | 按"设备分类码带部件率 + 名称关键词"补 | 补出 3,830 个，与真值对照**精确率 0%**：`GP` 类里"油站滤网/冷却器"没有部件；名称含"开关/电动"的 1,238 个设备里只有 11 个真有部件 |

    **根因**：某台设备"有没有部件"取决于该台设备的实际构造（控制柜里有开关与测控装置、
    阀门带执行机构，而滤网/冷却器/变压器没有），**从设备名称与分类码判定不了**。
    硬补就是造假 —— 宁可留空。

    ## 因此本模块的产物

    只有 `list_missing_components()` 的**应编未编清单**（225 个设备）交第二步审核与人工确认。
    `component_index` / `parse` / 名称规则仍然保留：**当输入里已经给了部件行时**，
    用它们拼编码、判类型词、算序号，这些是确定性的、可用的。

    若确实需要探索性生成（例如人工确认了某批设备该带部件），传 `allow_speculative=True`。
    """
    if not allow_speculative:
        return []
    plans: list[ComponentPlan] = []
    for dev in devices:
        kks = str(dev["kks"])
        if kks in component_index:
            continue
        cls = str(dev.get("cls", ""))
        ratio = DEVICE_CLASS_RATIO.get(cls, 0.0)
        name = str(dev.get("name", "") or "")
        if any(kw in name for kw in NO_COMPONENT_HINTS):
            continue
        by_name = _hint_from_name(name)
        act = _actuator_cls(name)
        classes = [c for c in (by_name, act) if c]
        for final_cls in dict.fromkeys(classes):
            if final_cls not in COMPONENT_TYPE_WORD:
                continue
            word = COMPONENT_TYPE_WORD[final_cls]
            pname = str(dev.get("parent_name", "") or "")
            plans.append(ComponentPlan(
                parent_kks=kks, parent_name=pname, cls=final_cls, num=1,
                name=f"{pname}{word}" if pname else word,
                source="探索性生成(名称证据)",
                confidence=0.4, need_human=True,
                notes=[f"名称出现部件类型词（{final_cls}）；{cls} 类带部件率 {ratio:.0%}，"
                       f"但实测此类推断精确率为 0%，必须人工确认"],
            ))
    return plans




def list_missing_components(devices: list[dict], component_index: dict[str, dict[str, int]],
                            *, min_ratio: float = 0.5) -> list[dict]:
    """应编未编清单：高带部件率的设备类型下，表里却没有部件的设备。

    交第二步审核报出，不自动补——因为"表里没有"也可能是该台设备真的没有部件。
    """
    out: list[dict] = []
    for dev in devices:
        kks = str(dev["kks"])
        if kks in component_index:
            continue
        cls = str(dev.get("cls", ""))
        ratio = DEVICE_CLASS_RATIO.get(cls, 0.0)
        if ratio < min_ratio:
            continue
        name = str(dev.get("name", "") or "")
        out.append({
            "kks": kks,
            "name": name,
            "cls": cls,
            "parent_kks": dev.get("parent_kks", ""),
            "class_ratio": round(ratio, 3),
            "expected_components": TYPICAL_COMPONENTS.get(cls, []),
            "reason": f"{cls} 类设备在竣工版中 {ratio:.0%} 带部件，本设备未见部件",
        })
    return out
