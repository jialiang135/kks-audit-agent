# -*- coding: utf-8 -*-
"""第一步智能体 · AI 语义层。

只做规则层给不出答案的 4 件事，其余一律不调模型：
  ① 原名称为空 → 按编码语义 + 系统/父级上下文推出设备名称
  ② 原名称 → 设备名称 的语义重写（补限定词、调序、统一写法）
  ③ 部件名称细化（规则只给出类型词）
  ④ 无设计编码的设备级行：判断"该设备是否存在"及其分类码

设计要点（都是从本项目实测教训来的）：
- **few-shot 用竣工版真实变换**：模型学"规律"，不是自由发挥。
  实测规律：① 机组号统一（07=7号机组）② 分支号 7A/7B 提前 ③ 补父级/系统限定词
  ④ 名称里带空格分隔层级 ⑤ 全角转半角
- **本地复核**：生成的名称必须过结构校验（非空、长度合理、不含非法字符、
  机组号与编码前缀一致），不过的退回给人核，绝不写进成果。
- **可复现**：temperature=0，同批输入同批输出；失败的批次不阻塞、标记后继续。
- **批量**：一次请求带多行（默认 15 行），复用现有 AI 管道的并发与重试。
"""
from __future__ import annotations

import json
import logging
import re
import time
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Callable

from ai_review import (
    AIReviewError,
    OpenAICompatibleClient,
    RETRY_BACKOFF_SECONDS,
    _is_retriable_batch_error,
    _parse_int,
    environment_value,
    load_config,
)

LOGGER = logging.getLogger("kks-audit.gen.ai")

BATCH_SIZE = 40        # 每批行数：太小时请求数暴涨（15 行 → 978 个请求）
BATCH_RETRIES = 3
# 并发默认值。⚠️ **2026-10-10 实测推翻了 10-09 的旧结论**（旧结论说并发 8 是并发 4 的 6.5 倍）：
#   旧口径量的是"每秒请求数"（多半是小请求），而名称生成发的是**大请求**——两码事。
#   今天用**真实载荷**（40 行 / 请求体 39 KB，工具 `tools/probe_ai_endpoint.py`）量：
#     · 小请求（16 token）串行 0.4~0.9s；**8 并发也一样** → 端点不按请求数限流
#     · 真实批量**空闲时串行 28~35s**（远小于 90s 超时）
#     · 生成时并发 8 → 20 批里 **5 次撞 90s 超时**；而 AI 阶段总时长 525s ≈ 20×30s 串行
#   ⇒ **并发给不了加速，只换来超时**（大请求是端点吞吐瓶颈，不是连接数瓶颈）。
#   故默认降到 3：偶发卡顿不至于拖住全局，总时长不变、超时消失。
#   可用 `.env.local` 的 `AI_NAME_CONCURRENCY` 覆盖（不必改代码）。
MAX_CONCURRENCY = 3
# 超时不再是这里的常量，而是 `ai_review.DEFAULT_TIMEOUT_SECONDS`（`AI_TIMEOUT_SECONDS` 覆盖）。
#   **AI 名称生成是可选步骤**：`--no-ai` 走规则，秒级出结果。


def _tuning(env_name: str, explicit: int | None, default: int, low: int, high: int) -> int:
    """调优参数的取值优先级：**调用方显式传参 > 环境变量/.env.local > 代码默认值**。

    单独抽出来是为了**可测**（`test_ai_name_tuning.py`）：端点是快是慢会变，
    参数得能在 `.env.local` 里调，不能写死、也不该为调参去改代码。
    """
    if explicit is not None:
        return _parse_int(explicit, default, low, high)
    return _parse_int(environment_value(env_name, default), default, low, high)

# 机组号口径：本项目 07=7号机组、08=8号机组（不是 1/2 号）
UNIT_BY_PREFIX = {"07": "7", "08": "8", "01": "1", "02": "2", "03": "3",
                  "04": "4", "05": "5", "06": "6"}

NAME_PROMPT = """你是电厂 KKS 编码成果的编制工程师。任务：把图纸摘抄的「原名称」改写成规范的「设备名称」。

**改写规律**（全部来自中煤鄂州四期竣工版的真实成果，必须照这个风格）：
1. 机组号统一：本工程 07=7号机组、08=8号机组。凡写成"1号机/2号机"的都是错的，要改成 7号/8号；
   但 **"1号主变""1号高厂变"里的"1号"是设备编号，不是机组，必须保持原样**。
2. 分支号提前：`烟气再循环风机7A变频器` → `7A烟气再循环风机变频器`
   （把名称中间的 7A/7B/8A 这类分支号提到最前面）
3. 补父级/系统限定词：名称过于孤立时，用"上级设备名称"做前缀，中间**用空格分隔**：
   `母线PT` + 上级`10kV工作7A段` → `10kV工作7A段 母线PT`
   `试验电源屏` + 上级`7号机励磁系统` → `7号机试验电源屏`
4. 名称已在上级里体现的，可在上级名后断开：`10kV工作7A段工作电源` → `10kV工作7A段 工作电源进线开关`
   （这类还常需补功能词，如"进线开关""出线开关"——但**只在证据充分时才补**）
5. 全角括号/字母转半角：`DEH（ATC）控制柜` → `DEH(ATC)控制柜`
6. **分相设备必须带分相后缀**：若给出的"分相后缀"非空（如 `A相`），名称结尾要带上它：
   `1号主变 高压侧PT` + 分相`A相` → `1号主变 高压侧PT A相`
7. 若给了"同类设备已有名称/同父级兄弟名称"，**优先与它们保持同一写法**（同厂惯例最强）。
8. 不改变设备实质语义，不编造型号、编号、图号。

**证据不足时**：name 返回空字符串，并在 reason 里说明缺什么。宁可留空，不要猜。

只返回 JSON 对象：
{"items":[{"row":行号,"name":"规范的设备名称或空串","reason":"中文理由","confidence":0.0}]}
每个 row 最多返回一次，confidence 为 0~1。"""

FEWSHOT = [
    # (原名称, 上级设备名称, 编码, 正确的设备名称) —— 全部来自竣工版真实数据
    ("7号机组NCS测控屏", "1号机NCS测控屏系统", "07AAB02UH001", "7号机NCS测控屏"),
    ("烟气再循环风机7A变频器", "7号机高压变频装置系统", "07BBA00UC004", "7A烟气再循环风机变频器"),
    ("母线PT", "10kV工作7A段", "07BBA01UC003", "10kV工作7A段 母线PT"),
    ("7A循环水泵", "10kV工作7A段", "07BBA01UC007", "10kV工作7A段 7A循环水泵"),
    ("试验电源屏", "7号机励磁系统", "07AAC40UH001", "7号机试验电源屏"),
    ("凝结水泵变频器切换柜", "7号机高压变频装置系统", "07BBA00UC001", "7A凝结水泵变频器切换柜"),
    ("7号机DEH（ATC）控制柜", "7号机DEH系统", "07CCA01UH001", "7号机DEH(ATC)控制柜"),
    ("1号主变 高压侧PT", "500kV 1号主变进线间隔系统", "07ABG01BA001", "1号主变 高压侧PT"),
]

# 允许出现在名称里的字符（中文/字母/数字/常用标点）
ILLEGAL_NAME_CHARS = re.compile(r"[\x00-\x1f\x7f]")
UNIT_PREFIX_RE = re.compile(r"^([0-9]{1,2})")


def _build_payload(rows: list[dict[str, Any]]) -> dict[str, Any]:
    """组装请求体：few-shot 用竣工版真实样例，待处理行只给证据。

    ⚠️ OpenAICompatibleClient.review(model, system_prompt, payload) 会把 payload 序列化成
    user 消息，所以这里返回的是**结构化数据**，不是 messages 数组。
    """
    return {
        "task": "generate_device_name",
        "改写规律的真实样例": [
            {"原名称": old, "上级设备名称": parent, "KKS编码": code, "正确的设备名称": name}
            for old, parent, code, name in FEWSHOT
        ],
        "待改写的行": [
            {
                "row": r["row"],
                "原名称": r.get("original_name", ""),
                "上级设备名称": r.get("parent_name", ""),
                "上级KKS编码": r.get("parent_kks", ""),
                "KKS编码": r.get("kks", ""),
                "专业": r.get("profession", ""),
                "图纸名称或图号": str(r.get("drawing", ""))[:80],
                "系统分类码含义": r.get("system_meaning", ""),
                "设备分类码含义": r.get("device_meaning", ""),
                # ⚠️ 分相提示：13 位末尾 A/B/C 的设备名称必须带「A相/B相/C相」后缀
                #    （实测：不给这个提示，模型会把 `1号主变 高压侧PT A相` 写成不带分相的版本）
                "分相后缀": r.get("phase_suffix", ""),
                # ⚠️ 同类设备已有名称：最强证据。实测给了几条之后，模型能照同类写法一次写对。
                "同类设备已有名称": r.get("same_class_names", [])[:6],
                "同父级兄弟名称": r.get("sibling_names", [])[:6],
            }
            for r in rows
        ],
    }


def _validate_name(item: dict[str, Any], row: dict[str, Any]) -> tuple[bool, str]:
    """本地复核：生成的名称必须自洽，否则退回人工。"""
    name = str(item.get("name", "") or "").strip()
    if not name:
        return False, ""
    if ILLEGAL_NAME_CHARS.search(name):
        return False, "含控制字符"
    if len(name) < 3:
        return False, f"名称过短（{len(name)} 字）"
    if len(name) > 60:
        return False, f"名称过长（{len(name)} 字）"
    kks = str(row.get("kks", "") or "")
    want_unit = UNIT_BY_PREFIX.get(kks[:2], "")
    # 机组号必须与编码前缀一致（"1号机"这种错误写法不允许出现在成果里）
    if want_unit:
        for bad, label in (("1号机", "1号机"), ("2号机", "2号机")):
            if bad in name and want_unit not in (bad[0],):
                if want_unit + "号机" not in name:
                    return False, f"机组号与编码前缀不符（出现「{label}」，应为「{want_unit}号机」）"
    # 原名称的主体应保留（允许加前缀/后缀/调序，但不能整句换掉）。
    # ⚠️ 实测教训：规则太严会误杀——模型把 `7号机组NCS测控屏` 正确改成 `7号机NCS测控屏`
    # 时，被"语义偏离过大"拦下（因为"号机组"被缩成了"号机"）。所以这里只拦"明显跑掉"：
    # 长度差超过一半，或原名称里的中文词一个都没保留。
    old = str(row.get("original_name", "") or "").strip()
    if old and len(old) >= 4:
        words = [w for w in re.findall(r"[\u4e00-\u9fff]{2,}", old)]
        if words:
            kept = sum(1 for w in words if w in name)
            if kept == 0:
                return False, "与原名称语义偏离过大（原名称中的中文词一个都没保留）"
        # 短名称补父级前缀是**设计意图**（`母线PT` → `10kV工作7A段 母线PT`），
        # 所以不能用"长度差"拦。只拦"明显变短"（丢信息）与"暴涨但一个字都没保留"。
        if len(name) < len(old) * 0.6:
            return False, f"名称比原名称明显变短（原 {len(old)} 字 → 生成 {len(name)} 字），可能丢信息"
    return True, ""


def _notify_retry(cb: Callable[[dict[str, Any]], None] | None,
                  index: int, total: int, attempt: int, err: str) -> None:
    """把「正在重试」上报给界面（AI 阶段进度 30%~75%，按批号线性推进）。

    ⚠️ 为什么必须上报：端点偶发超时后要退避重试，最坏 3×90s + 退避 ≈ 4.5 分钟**没有任何
        进度更新**，界面上看起来就是"任务卡死"（用户实测反馈过）。多报一条消息，
        比让人以为程序死了要好。上报失败绝不影响生成。
    """
    if not cb:
        return
    try:
        percent = 30 + int((index / max(1, total)) * 45)
        cb({
            "stage": "ai_retry",
            "percent": percent,
            "message": f"AI 第 {index}/{total} 批请求超时，正在重试（第 {attempt + 1}/{BATCH_RETRIES} 次）",
            "hint": f"端点无响应时会自动重试，请稍候…（{str(err)[:40]}）",
        })
    except Exception:      # 进度上报绝不能影响生成
        pass


def resolve_names(
    rows: list[dict[str, Any]],
    *,
    progress_callback: Callable[[dict[str, Any]], None] | None = None,
    batch_size: int | None = None,
    concurrency: int | None = None,
) -> dict[str, Any]:
    """批量生成设备名称。返回 {row: {name, reason, confidence, valid, invalid_reason, model}}。

    rows 每项需含：row / kks / original_name / parent_name / parent_kks / profession / drawing。

    `batch_size` / `concurrency` 传 None（默认）时按
    **环境变量 `AI_NAME_BATCH_SIZE` / `AI_NAME_CONCURRENCY` > 代码默认值** 解析（见 `_tuning`）。
    """
    batch_size = _tuning("AI_NAME_BATCH_SIZE", batch_size, BATCH_SIZE, 1, 200)
    concurrency = _tuning("AI_NAME_CONCURRENCY", concurrency, MAX_CONCURRENCY, 1, 16)
    config = load_config()
    out: dict[str, Any] = {
        "status": "disabled", "model": config.model or None, "requested": len(rows),
        "returned": 0, "valid": 0, "rejected": 0, "errors": [], "items": {},
    }
    if not config.enabled or not config.api_key:
        out["reason"] = "AI 未启用（AI_ENABLED=false 或未配 AI_API_KEY）"
        return out
    if not rows:
        out["status"] = "completed"
        out["reason"] = "没有需要生成名称的行"
        return out

    client = OpenAICompatibleClient(config)
    try:
        model, _ = client.choose_model()
    except AIReviewError as exc:
        out["status"] = "error"
        out["errors"].append(str(exc))
        return out
    out["model"] = model

    batches = [rows[i:i + batch_size] for i in range(0, len(rows), batch_size)]
    total = len(batches)
    workers = max(1, min(concurrency, total))
    LOGGER.info("gen_name_start rows=%s batches=%s concurrency=%s model=%s",
                len(rows), total, workers, model)

    def run_batch(index: int, batch: list[dict[str, Any]]) -> dict[str, Any]:
        local: dict[str, Any] = {"index": index, "items": {}, "errors": [],
                                 "returned": 0, "valid": 0, "rejected": 0}
        by_row = {str(r["row"]): r for r in batch}
        payload = None
        for attempt in range(1, BATCH_RETRIES + 1):
            try:
                payload = client.review(model, NAME_PROMPT, _build_payload(batch))
                break
            except Exception as exc:
                if attempt < BATCH_RETRIES and _is_retriable_batch_error(str(exc)):
                    LOGGER.warning("gen_name_batch_retry batch=%s attempt=%s err=%s",
                                   index, attempt, str(exc)[:100])
                    # ⚠️ 必须让"重试"**可见**：端点偶发超时（实测撞到 90s 超时）后要退避重试，
                    #    最坏 3×90s + 退避 ≈ 4.5 分钟**没有任何进度更新** → 界面上就像"卡死了"
                    #    （用户实测反馈过）。宁可多报一条消息，也不要让人干等。
                    _notify_retry(progress_callback, index, total, attempt, str(exc))
                    time.sleep(RETRY_BACKOFF_SECONDS * attempt)
                    continue
                local["errors"].append(str(exc)[:200])
                LOGGER.warning("gen_name_batch_failed batch=%s err=%s", index, exc)
                return local
        if payload is None:
            local["errors"].append("模型未返回可解析内容")
            return local

        items = payload.get("items")
        if not isinstance(items, list):
            local["errors"].append("模型返回的 items 不是数组")
            return local
        for item in items:
            if not isinstance(item, dict):
                continue
            row_no = str(item.get("row", "")).strip()
            row = by_row.get(row_no)
            if row is None:
                continue
            local["returned"] += 1
            valid, reason = _validate_name(item, row)
            if not valid and str(item.get("name", "") or "").strip():
                local["rejected"] += 1
            elif valid:
                local["valid"] += 1
            try:
                conf = float(item.get("confidence", 0) or 0)
            except (TypeError, ValueError):
                conf = 0.0
            local["items"][row_no] = {
                "row": row_no,
                "name": str(item.get("name", "") or "").strip() if valid else "",
                "reason": str(item.get("reason", ""))[:300],
                "confidence": max(0.0, min(conf, 1.0)) if valid else min(conf, 0.3),
                "valid": valid,
                "invalid_reason": "" if valid else (reason or "模型未给出名称"),
                "model": model,
            }
        return local

    done = 0
    with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="kks-gen-ai") as pool:
        for local in pool.map(run_batch, range(1, total + 1), batches):
            out["errors"].extend(local["errors"])
            out["returned"] += local["returned"]
            out["valid"] += local["valid"]
            out["rejected"] += local["rejected"]
            out["items"].update(local["items"])
            done += 1
            if progress_callback:
                try:
                    # AI 阶段占总进度 30%~75%，按批数线性推进
                    percent = 30 + int((done / max(1, total)) * 45)
                    progress_callback({
                        "stage": "ai_batch", "percent": percent,
                        "message": f"AI 正在生成设备名称（{done}/{total} 批）",
                        "hint": f"可用 {out['valid']:,} 条 / 退回 {out['rejected']:,} 条",
                        "batch": done, "total_batches": total,
                    })
                except Exception:  # 进度上报不能影响生成
                    pass

    out["status"] = "completed" if out["items"] or not out["errors"] else "error"
    LOGGER.info("gen_name_done rows=%s returned=%s valid=%s rejected=%s errors=%s",
                out["requested"], out["returned"], out["valid"], out["rejected"], len(out["errors"]))
    return out
