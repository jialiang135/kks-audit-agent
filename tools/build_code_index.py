# -*- coding: utf-8 -*-
"""生成 KKS 分类码索引（config/kks_code_index.json）。

## 数据来源（**单一事实来源**）

本索引**不再**从 `kks-audit/references/rules_gb50549.json`（另一条早期抽取链路）取数，
改为直接读**可复跑的字典产物**：

```
tools/extract_gb_dictionaries.py   → config/gb_system_codes.json / gb_device_codes.json
tools/extract_zhongmei_codes.py    → config/project_system_codes.json / project_device_codes.json
```

即：`project_*` = 国标附录 ∪ 中煤细则，与 `kks_gen/revise.py`、`ai_review.py` 判码时**读同一份数据**，
消除"同一份标准两套抽取、结果不一致"的隐患。

## 为什么改（踩过的两个坑）

1. **国标附录B 的 A/B/C 行在 OCR 里被拆到多行**，早期按**单行**跑 `<tr>` 正则 → 整行丢弃。
   上一轮用手工补丁 `T1_SUPPLEMENT` 补了 `A`/`B`/`C` 三条，但**同样被丢的 `AF`/`AG`/`BA` 没补**
   （`BA` 正是测点「电压/电位」的 T₂ 码）。现已把 `extract_gb_dictionaries.py` 改为**按区块解析**，
   一次性补回 5 条，**手工补丁随之删除**（不再需要，也不会再漂移）。
2. 早期按码长路由猜表类型，把中煤附录B 续表的 T₂ 设备码错标成系统码 —— 已在
   `extract_zhongmei_codes.py` 里改成"续表继承 + 长度硬校验"。

## 产出结构（`ai_review.py` 依赖，勿改字段名）

  standard / coverage_notes / source / generated_by / note
  system_s1_groups   S1 主组（1 位字母）→ 中文系统名
  system_codes       S1S2S3 系统分类码（2/3 位字母）→ 中文系统名
  device_t1_groups   T1 主组（1 位字母）→ 中文名称
  device_codes       T1T2 设备/产品分类码（2 位字母）→ 中文名称
"""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CFG = ROOT / "config"
OUT = CFG / "kks_code_index.json"


def load(name: str, fallback: str | None = None) -> dict:
    path = CFG / name
    if not path.is_file():
        if fallback:
            path = CFG / fallback
        if not path.is_file():
            raise SystemExit(f"缺少数据文件 {name}（先跑 extract_gb_dictionaries.py / extract_zhongmei_codes.py）")
    return json.loads(path.read_text(encoding="utf-8"))


def split_by_len(table: dict, lengths: tuple[int, ...]) -> dict[str, str]:
    """取指定码长的 {code: 释义}。

    ⚠️ 释义必须是 **`中文名` + ` / ` + 说明与实例**（如 `GP` → `…流动 / 提升泵、清洗泵、计量泵…`）。
    说明列里装着**应用实例（同义词）**，而 AI 建议编码的「按描述匹配分类码」正是靠这些实例命中的
    （`给水泵` → `GP`）；只留 `中文名` 会让匹配能力大幅退化（实测踩过：`给水泵` 只匹配到 `BM 水分`）。
    """
    out: dict[str, str] = {}
    for code, item in table.items():
        if len(code) not in lengths:
            continue
        it = item or {}
        parts = [str(it.get("cn") or "").strip(), str(it.get("note") or "").strip()]
        text = " / ".join(p for p in parts if p and p not in ("—", "-", "一"))
        if not text:
            continue
        out[code] = text[:200]
    return out


def main() -> None:
    # 项目可用码表 = 国标附录 ∪ 中煤细则（与引擎同源）
    sys_all = load("project_system_codes.json", "gb_system_codes.json")
    dev_all = load("project_device_codes.json", "gb_device_codes.json")

    system_s1 = split_by_len(sys_all, (1,))
    system_codes = split_by_len(sys_all, (2, 3))
    device_t1 = split_by_len(dev_all, (1,))
    device_codes = split_by_len(dev_all, (2,))

    payload = {
        "standard": "GB/T 50549-2020《电厂标识系统编码标准》（2021-03-01 实施；同时废止 GB/T 50549-2010）",
        "coverage_notes": (
            "S1 主组、S1S2S3 系统分类码来自 GB/T 50549-2020 附录A（表A.0.1/表A.0.2）；"
            "T1 主组、T1T2 设备及产品分类码来自附录B（表B.0.1/表B.0.2）；"
            "中煤集团《电厂标识系统编码应用细则（试行）》作为集团补充（含其附录与正文表格中的"
            "独创/赋义码，见 config/zhongmei_unique_codes.json）。"
            "本索引由 tools/build_code_index.py 从 config/project_*.json 生成，"
            "与 kks_gen/revise.py 判码读的是**同一份数据**。"
            "标准附录在 PDF 中分页，抽取后仍可能有极少数码未收录——"
            "因此「码未收录」只能作为待核线索，不能直接判定为错误。"
            "本项目不引用 GB/T 50549-2010 体系（含继承该体系的 2015 年火电导则）的分类码。"
        ),
        "source": [
            "GB/T 50549-2020 附录A 系统分类索引",
            "GB/T 50549-2020 附录B 设备及产品分类索引",
            "中煤集团《电厂标识系统编码应用细则（试行）》（附录 + 正文表格）",
        ],
        "generated_by": "tools/build_code_index.py（读 config/project_*.json，二者同源）",
        "note": "分类码取值是否合法以 GB/T 50549-2020 为准；本索引供自动校验与 AI 建议编码使用。",
        "system_s1_groups": dict(sorted(system_s1.items())),
        "system_codes": dict(sorted(system_codes.items())),
        "device_t1_groups": dict(sorted(device_t1.items())),
        "device_codes": dict(sorted(device_codes.items())),
    }
    OUT.write_text(json.dumps(payload, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    print(f"-> {OUT.relative_to(ROOT)} {OUT.stat().st_size} B")
    print(f"  S1 主组 {len(system_s1)} 个；系统分类码 {len(system_codes)} 个"
          f"（2 位 {sum(1 for c in system_codes if len(c) == 2)} / 3 位 {sum(1 for c in system_codes if len(c) == 3)}）")
    print(f"  T1 主组 {len(device_t1)} 个；设备分类码 {len(device_codes)} 个")
    print()
    for k in ("A", "B", "BA", "AF", "AG", "BP", "PA", "KB", "SGA", "AA", "BY"):
        where = []
        for label, d in (("S1", system_s1), ("系统码", system_codes),
                         ("T1", device_t1), ("设备码", device_codes)):
            if k in d:
                where.append(f"{label}:{d[k][:30]}")
        print(f"  {k:5s} -> " + ("; ".join(where) if where else "（未收录）"))


if __name__ == "__main__":
    main()
