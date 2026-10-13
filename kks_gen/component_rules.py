# -*- coding: utf-8 -*-
"""部件级（第 5 层扩展码）生成规则。

实测依据（鄂州四期竣工版 21,523 个部件）：
1. 部件分类码与父设备分类码**配对稳定**，可直接查表决定"该设备下通常挂哪些部件"：

   父设备 UH（控制柜/屏）→ KF(信号) 5751、QA(开关) 3130、FC(过流保护) 625
   父设备 UC（配电柜）    → QA 2017、BC(CT) 313、FA(避雷器) 171、QC(接地刀闸) 153
   父设备 QM（截止阀）    → MA(执行机构) 1263、MM(手动) 1124
   父设备 GP（泵）        → MA 624
   父设备 BF（流量计）    → QM(阀门) 464
   父设备 QN（调节阀）    → MA 439、MM 292
   父设备 RB             → QA 395
   父设备 BP（压力变送器）→ QM 308
   父设备 GL（电动机）    → UR(紧固/机架) 234

2. 编号在同父设备下**从 01 连续递增**（实测 5545/6349 组连续）。
3. 部件码形态：`父设备码 + "-" + 部件分类码(2位) + 编号(2位)`，如 `07AAB02UH001-FC01`。

⚠️ 名称无法靠规则完全生成：部件的具体名称（"JK开关"、"1-21ZKK开关"）来自图纸，
   规则只能给出"类型词"（如 KF→信号、QA→开关）。因此：
   - 编码 + 层级 + 编号：**规则生成**，置信度 0.75
   - 名称：用父名称 + 类型词兜底，标记**需人核**，可交 AI 细化
"""
from __future__ import annotations

from collections import Counter, defaultdict

# 父设备分类码 → [(部件分类码, 典型数量权重)]
# 权重来自竣工版实测频次，用于在"该挂几个部件"不确定时给出候选顺序
COMPONENT_MAP: dict[str, list[str]] = {
    "UH": ["KF", "QA", "FC", "BC", "AG"],
    "UC": ["QA", "BC", "FA", "QC", "AG"],
    "QM": ["MA", "MM"],
    "QN": ["MA", "MM"],
    "GP": ["MA"],
    "GQ": ["MA"],
    "BF": ["QM"],
    "BP": ["QM"],
    "BT": ["QM"],
    "BL": ["QM"],
    "RB": ["QA"],
    "AT": ["MA"],
    "GL": ["UR"],
    "HQ": ["MA"],
    "VQ": ["QM"],
}

# 部件分类码 → 名称类型词（用于生成兜底名称）
COMPONENT_TYPE_WORD: dict[str, str] = {
    "KF": "信号",
    "QA": "开关",
    "QB": "隔离开关",
    "QC": "接地刀闸",
    "FC": "过流保护",
    "BC": "CT",
    "FA": "避雷器",
    "AG": "附件",
    "MA": "执行机构",
    "MM": "手动装置",
    "QM": "阀门",
    "UR": "机架",
}


def expected_components(device_code: str) -> list[str]:
    """给出某设备分类码下"通常存在的部件分类码"候选（按实测频次排序）。"""
    base = device_code.split(".", 1)[0]
    if len(base) < 9:
        return []
    return list(COMPONENT_MAP.get(base[7:9], []))


def component_kks(device_kks: str, cls: str, num: int) -> str:
    """拼部件编码：07AAB02UH001 + KF + 1 → 07AAB02UH001-KF01"""
    return f"{device_kks}-{cls}{num:02d}"


def component_name(parent_name: str, cls: str) -> str:
    """部件兜底名称 = 父名称 + 类型词（需人核/AI 细化）。"""
    word = COMPONENT_TYPE_WORD.get(cls, "")
    if not word:
        return ""
    return f"{parent_name}{word}" if parent_name else ""


def observed_component_counts(pairs: list[tuple[str, str]]) -> Counter:
    """由已有数据统计"父设备分类码 → 部件分类码"的真实频次（用于校准 COMPONENT_MAP）。"""
    counts: Counter = Counter()
    for device_code, cls in pairs:
        base = device_code.split(".", 1)[0]
        if len(base) >= 9:
            counts[(base[7:9], cls)] += 1
    return counts


def group_components(rows: list[dict]) -> dict[str, list[tuple[str, int]]]:
    """把已有的部件行按父设备分组，供"照同类设备补齐"使用。

    rows 每项含 kks / parent_code。返回 {父设备码: [(部件分类码, 编号)]}
    """
    grouped: dict[str, list[tuple[str, int]]] = defaultdict(list)
    for row in rows:
        kks = str(row.get("kks", "") or "")
        parent = str(row.get("parent_code", "") or "")
        if not kks or not parent or "-" not in kks:
            continue
        tail = kks[len(parent):] if kks.startswith(parent) else kks.split("-", 1)[-1]
        tail = tail.lstrip("-")
        if len(tail) < 4:
            continue
        cls, num = tail[:2], tail[2:]
        if cls.isalpha() and num.isdigit():
            grouped[parent].append((cls, int(num)))
    return dict(grouped)
