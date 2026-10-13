# -*- coding: utf-8 -*-
"""GB/T 50549-2020 合规校验（按国标原文，不靠反推）。

国标依据（原文条款）：
  4.4.1 功能面标识采用三级编码
  4.4.2 系统分类码应由 3 个大写英文字母组成              → S1S2S3
  4.4.3 设备分类码应由 2 个大写英文字母组成              → T1T2
  4.4.4 系统码 = 系统分类码 S1S2S3 + 系统编号 S4S5       → 5 位（两位数字）
  4.4.5 设备码 = 设备分类码 T1T2 + 设备编号 T3T4T5       → 5 位（三位数字，001~999）
  注 1  字符类型 N 为阿拉伯数字，A 为大写英文字母（**禁用 I、O**）
  4.3.1 全厂码 H1H2H3(H4)；火电等 H1 可缺省 → 实际用 H2H3
  表4.3.1-2 全厂码取值：01~60=1~60号机组；61~90=两机共用（61→1、2号）；
                        91~99=三机及以上共用；00=全厂公用

由此得到完整功能面标识形态（12 位）：
  全厂码(2) + 系统分类码(3字母) + 系统编号(2数字) + 设备分类码(2字母) + 设备编号(3数字)
  = 2 + 5 + 5 = 12
"""
from __future__ import annotations

import re

# 禁用字母（国标注 1）：I 与 O 不得使用
BANNED_LETTERS = ("I", "O")

# 数字类型位置（0 基）与字母类型位置
FULL_DEVICE_PATTERN = re.compile(r"^(?P<plant>\d{2})(?P<syscls>[A-Z]{3})(?P<sysno>\d{2})"
                                 r"(?P<devcls>[A-Z]{2})(?P<devno>\d{3})$")
SYSTEM_PATTERN = re.compile(r"^(?P<plant>\d{2})(?P<syscls>[A-Z]{3})(?P<sysno>\d{2})$")


def unit_label_to_plant_code(label: str) -> str:
    """把「机组号」列的文字按国标表4.3.1-2 换成全厂码。无法判定时返回空串。

    · `1号机组` / `2号机组`      → 01 / 02        （01~60 = 1~60 号机组）
    · `1、2号机组共用`           → 61              （61~90 = 两机共用，61→1、2号）
    · `1、2、3号机组共用`        → 空串            （91~99需工程约定）
    · `全厂公用`（**必须带"全厂"**） → 00            （00 = 全厂公用）
    · 裸 `公用` / `共用`         → 空串（**歧义**：本项目实测它是"1、2号机共用"的简称）
    · `待复核` / `未注明` / 空   → 空串（**不猜**，交人工）
    """
    text = re.sub(r"\s+", "", str(label or ""))
    if not text:
        return ""
    # ⚠️ 机组号可能是「1、2号机组共用」——要抓住"号"之前的**全部**数字。
    #    踩过的坑：用 findall(r"(\d+)号") 只会拿到紧邻"号"的那个（2），漏掉 1，
    #    于是国标表4.3.1-2 里 61=1、2号共用 这条永远匹配不上。
    units: list[int] = []
    matched = re.match(r"^([\d、,，和及\-~～]+)号", text)
    if matched:
        units = [int(n) for n in re.findall(r"\d+", matched.group(1))]
    if not units:
        units = [int(n) for n in re.findall(r"(\d+)号", text)]
    shared = any(word in text for word in ("共用", "公用"))
    if units:
        if not shared and len(units) == 1:
            return f"{units[0]:02d}" if 1 <= units[0] <= 60 else ""
        if shared:
            if len(units) >= 3:
                return ""  # 91～99的共用范围由各方约定，不能自行分配91
            if len(units) == 2 and units[1] == units[0] + 1 and units[0] % 2 == 1:
                return f"{60 + (units[0] + 1) // 2:02d}"   # 1、2→61；3、4→62
        return ""
    # ⚠️ **必须写明"全厂"才判 00**（2026-10-10 实测纠正，原实现是"含公用即 00"）：
    #    热控点表里 `61LAC10`（电动给水泵）**同一系统**下并存两种写法——
    #    一部分标「公用」、一部分标「1、2号机组共用」，全厂仅 3 行真·全厂公用（500kV GIS）用的是 `00`。
    #    可见本项目的裸「公用」是**共用**的简称，据它把 `61`（1、2号机共用）改成 `00`（全厂公用）
    #    会**改变含义**、并让同一系统的设备分到两个分支。故裸「公用」按**歧义**处理（返回空串，
    #    由 `revise` 保留原码并提请人工确认，见 `ambiguous_public_label()`）。
    if "全厂" in text:
        return "00"
    return ""


def ambiguous_public_label(label: str) -> bool:
    """「公用/共用」类标签是否**歧义**到无法定全厂码（表4.3.1-2）。

    自洽定义：**判不出码**（`unit_label_to_plant_code` 返回空串）**且**文字里出现「公用/共用」。
    · `1、2号机组共用` → 判得出 61 → 不算歧义
    · `全厂公用`       → 判得出 00 → 不算歧义
    · `公用`           → 判不出，且是术语 → **歧义**（本项目实测：它就是"1、2号机共用"的简称）
    """
    text = re.sub(r"\s+", "", str(label or ""))
    if not text or unit_label_to_plant_code(text):
        return False
    return "公用" in text or "共用" in text


def check_compliance(code: str) -> dict:
    """按国标校验一个功能面标识。返回 {ok, level, reason}。

    level ∈ {plant, main_group, system_class, system, device_class, device, extended, unknown}
    """
    from kks_gen.compliance import assess
    result = assess(code)
    return {"ok": result.writable and not result.pending, "level": "device" if result.kind == "function" else result.kind,
            "reason": "；".join(result.errors + result.pending) or "已通过本次自动结构校验"}
