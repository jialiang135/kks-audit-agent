# -*- coding: utf-8 -*-
"""生成「AI 建议编码」验证用例表：一份小而可读的 xlsx，专门覆盖两类情形。

设计意图（每行都在 说明 sheet 里写明期望）：
- A 类「码本身写错」→ 应该给出建议码：数字段混入字母 O、小写、禁用字母 I、
  未收录设备分类码、防火阀用了开关码；
- B 类「改码解决不了」→ 不应该给建议码：名称全角符、编号断号；
- C 类「改了就是错」→ 必须被本地语义复核拦住：去改机组前缀（名称里写着 1号机组）。

生成后用真实模型跑一遍，报告第十节会按"有建议码优先"显示。
"""
from __future__ import annotations

import pathlib
import sys

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

OUT = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else "outputs/验证用例_错误编码.xlsx")
OUT.parent.mkdir(parents=True, exist_ok=True)

HEADER = ["KKS编码", "父级KKS编码", "设备名称", "机组号", "专业", "系统", "设备类型", "原KKS编码"]

# (编码, 父级, 名称, 机组号, 专业, 系统, 设备类型, 原KKS, 用例说明)
ROWS: list[tuple[str, ...]] = [
    # A 类：应有建议码
    ("00LBA10GP0O1", "00LBA10", "1号机组凝结水泵", "1号机组", "暖通", "凝结水系统", "离心泵", "00LBA10GP001",
     "A1 数字段第10位是字母 O（应为 0）→ 期望建议 00LBA10GP001"),
    ("00LBA10GP002", "00LBA10", "1号机组凝结水泵电机", "1号机组", "暖通", "凝结水系统", "电动机", "00LBA10GP002",
     "A2 设备分类码 GP 是泵，名称却是电机 → 期望建议改 A1A2 或留空说明"),
    ("00LBA10GG001", "00LBA10", "1号机组送风机", "1号机组", "暖通", "送风系统", "轴流风机", "00LBA10GG001",
     "A3 设备分类码 GG 未收录，名称是轴流风机 → 期望建议 GQ"),
    ("00LBA10QI001", "00LBA10", "1号机组给水截止阀", "1号机组", "暖通", "给水系统", "截止阀", "00LBA10QI001",
     "A4 设备分类码含禁用字母 I → 期望建议改掉 I（QM/其他合法阀码）"),
    ("00LBA10QA001", "00LBA10", "1号机组防火阀", "1号机组", "暖通", "消防系统", "防火阀", "00LBA10QA001",
     "A5 防火阀用了开关码 QA → 期望建议 FM（同类设备已有先例）"),
    ("00lba10gp003", "00LBA10", "1号机组排水泵", "1号机组", "暖通", "排水系统", "离心泵", "00LBA10GP003",
     "A6 编码含小写字母 → 期望建议转大写 00LBA10GP003"),
    # C 类：改了就是错，必须拦住
    ("01XAM20GQ001", "01XAM20", "1号机组主厂房各设备间方型轴流风机1", "2号机组", "暖通", "主厂房通风系统", "轴流风机", "01XAM20GQ001",
     "C1 前缀 01 但机组列写 2号机组，名称写 1号机组 → 必须拦住「改前缀」，应校正文本列"),
    ("01XAM20GQ002", "01XAM20", "1号机组主厂房各设备间方型轴流风机2", "2号机组", "暖通", "主厂房通风系统", "轴流风机", "01XAM20GQ002",
     "C2 同上（防误改机组前缀的第二例）"),
    # B 类：改码解决不了，应明确写「无需改码」
    ("00LBA10GQ002", "00LBA10", "1号机组轴流风机（壁式）", "1号机组", "暖通", "送风系统", "轴流风机", "00LBA10GQ002",
     "B1 名称含全角括号（KKS-30 文本卫生）→ 期望「无需改码」，按问题列整改"),
    ("00LBA10FM001", "00LBA10", "1号机组防火阀1", "1号机组", "暖通", "消防系统", "防火阀", "00LBA10FM001",
     "B2 同父级编号断号（KKS-F）→ 期望「无需改码」，须人工确认是否漏编"),
    # 干净行：不应被报
    ("00LBA10GP009", "00LBA10", "1号机组轴流风机正常", "1号机组", "暖通", "送风系统", "轴流风机", "00LBA10GP009",
     "对照：编码与名称都正常，不应产生任何问题"),
]

wb = Workbook()
ws = wb.active
ws.title = "暖通专业"
ws.append(HEADER)
for row in ROWS:
    ws.append(list(row[:8]))

# 说明 sheet：让使用者知道每行期望什么，而不是只看结论
ws2 = wb.create_sheet("验证用例说明")
ws2.append(["类别", "行号", "编码", "用例说明"])
title_font = Font(bold=True, size=12, color="1F4E79")
ws2["A1"].font = title_font
for idx, row in enumerate(ROWS, start=2):
    category = row[0][:2]
    label = {"00": "A/B 类", "01": "C 类"}[category]
    ws2.append([label, idx, row[0], row[8]])

fill = PatternFill("solid", fgColor="DDEBF7")
for cell in ws2[1]:
    cell.fill = fill
    cell.font = Font(bold=True)
ws2.column_dimensions["A"].width = 10
ws2.column_dimensions["B"].width = 7
ws2.column_dimensions["C"].width = 16
ws2.column_dimensions["D"].width = 78

for column, width in zip("ABCDEFGH", (16, 14, 34, 11, 8, 18, 12, 16)):
    ws.column_dimensions[column].width = width
for cell in ws[1]:
    cell.font = Font(bold=True)

wb.save(OUT)
print(f"已生成: {OUT}（{len(ROWS)} 行用例）")
