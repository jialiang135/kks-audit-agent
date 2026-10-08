"""审计台账：每次审核自动追加单次记录，并维护跨文件的累计汇总报告。

- ledger（audit_ledger.json）：追加式 JSON 数组，每次 write_outputs=True 的审核记一条；
- 趋势总结（audit_trend.json）：跨文件 AI 趋势总结的缓存，随台账更新而刷新；
- 汇总报告（KKS审核台账汇总.html）：随每次审核自动重渲染，含累计 KPI、
  逐文件明细表、AI 跨文件趋势总结、以及“本次文件 vs 历史平均水平”的对比小节。
- 台账落在 output_dir 的上一级（服务器为 runs/ 根，随卷持久化）。
"""
from __future__ import annotations

import hashlib
import json
import logging
import os
import re
import shutil
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from ai_review import summarize_ledger
from report_business import CSS, _esc

LOGGER = logging.getLogger("kks-audit.history")

LEDGER_NAME = "audit_ledger.json"
TREND_NAME = "audit_trend.json"
HISTORY_NAME = "KKS审核台账汇总.html"

# 产物索引：台账只记 run_id，报告与问题清单本身留在 runs/<run_id>/ 下，
# 由保留策略按天清理；清理后把台账对应记录标记为已归档。
RUN_ID_RE = re.compile(r"^[0-9]{8}T[0-9]{6}_[0-9a-f]{8}$")
ARTIFACT_AVAILABLE = "available"
ARTIFACT_ARCHIVED = "archived"
RETENTION_ENV = "KKS_ARTIFACT_RETENTION_DAYS"
DEFAULT_RETENTION_DAYS = 90

_HIST_CSS = """
.hrow td,.hrow th{padding:6px 10px;border-top:1px solid var(--line);font-size:13px}
.hrow th{background:var(--blue2);text-align:left}
.hrow td.num{text-align:right;font-variant-numeric:tabular-nums}
.hrow tr.hl td{background:#fffbe6}
.hbad{color:var(--p0);font-weight:500}.hgood{color:var(--ok);font-weight:500}
.dup{display:inline-block;margin-left:6px;padding:0 6px;border-radius:9px;background:#e8f0fe;color:#1a56b8;font-size:11px;font-weight:500;cursor:help}
.trend{margin-top:12px;padding:12px 14px;border:1px solid var(--line);border-radius:8px;background:#f7faff}
.trend-h{font-weight:600;color:var(--ink);margin-bottom:6px;display:flex;align-items:center;gap:8px}
.trend-m{font-weight:400;font-size:12px;color:var(--muted)}
.trend p{margin:0 0 6px;line-height:1.7}
.trend ul{margin:0;padding-left:20px;line-height:1.8}
"""


def ledger_path(output_dir: Path) -> Path:
    """台账与汇总报告位于 output_dir 的上一级（runs/ 根）。"""
    return output_dir.parent / LEDGER_NAME


def trend_path(ledger_dir: Path) -> Path:
    """趋势缓存与台账同目录（服务器为 runs 根）；参数是台账所在目录。"""
    return ledger_dir / TREND_NAME


def load_trend(ledger_dir: Path) -> dict[str, Any]:
    """读取最近一次成功的跨文件 AI 趋势总结（无则返回 {}）。"""
    path = trend_path(ledger_dir)
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return data if isinstance(data, dict) else {}


def save_trend(ledger_dir: Path, trend: dict[str, Any]) -> None:
    path = trend_path(ledger_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(trend, ensure_ascii=False, indent=1), encoding="utf-8")


def history_path(output_dir: Path) -> Path:
    return output_dir.parent / HISTORY_NAME


def load_ledger(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    return [entry for entry in data if isinstance(entry, dict)]


def _file_digest(path: Path) -> str:
    """文件内容 SHA256 前 16 位；不可读时返回空串（此时退化为不去重）。"""
    try:
        digest = hashlib.sha256()
        with path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
        return digest.hexdigest()[:16]
    except OSError:
        return ""


def _run_id_of(output_dir: Path) -> str:
    """产物目录名即 run_id（服务器流程写作 runs/<run_id>/）；CLI 自定义目录返回空串。"""
    name = output_dir.name
    return name if RUN_ID_RE.fullmatch(name) else ""


def _entry(input_path: Path, result: dict[str, Any], run_id: str) -> dict[str, Any]:
    priority = result.get("priority_counts", {}) or {}
    ai = result.get("ai_review", {}) or {}
    final_counts = result.get("final_decision_counts", {}) or {}
    return {
        "ts": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "file": input_path.name,
        # 产物索引：run_id 指向 runs/<run_id>/，artifacts 记录产物是否还在。
        # run_id 为空表示本次不是服务器运行（CLI 自定义输出目录），无法在线导出。
        "run_id": run_id,
        "artifacts": ARTIFACT_AVAILABLE if run_id else "",
        "hash": _file_digest(input_path),
        "repeats": 1,
        "sheet": str(result.get("sheet", "") or ""),
        "rows": int(result.get("data_rows", 0) or 0),
        "issues": int(result.get("issue_count", 0) or 0),
        "priority": {k: int(priority.get(k, 0) or 0) for k in ("P0", "P1", "P2")},
        "ai_status": str(ai.get("status", "") or ""),
        "ai_candidate": int(ai.get("candidate_count", 0) or 0),
        "ai_reviewed": int(ai.get("reviewed_count", 0) or 0),
        "final_counts": {k: int(final_counts.get(k, 0) or 0) for k in ("confirmed_issue", "likely_false_positive", "needs_human")},
        "conclusion": str(result.get("conclusion", "") or "")[:200],
        "incremental": bool((result.get("scope_validation", {}) or {}).get("incremental_batch")),
    }


def append_ledger(input_path: Path, output_dir: Path, result: dict[str, Any]) -> list[dict[str, Any]]:
    """记录本次审核并返回完整台账（供随后的汇总渲染使用）。

    去重口径：以**文件内容哈希**为准。
    - 同一文件重复上传（内容未变）：不新增记录，把原记录移到末尾（保证 ledger[-1]
      恒为"本次"）、用最新结果覆盖并刷新时间戳，同时累加 repeats；
    - 文件内容发生变化（如更新为增量版）：视为新的一次审核，正常追加。
    哈希不可得（文件不可读）时退化为每次都追加，不会误合并。
    """
    path = ledger_path(output_dir)
    ledger = load_ledger(path)
    entry = _entry(input_path, result, _run_id_of(output_dir))
    digest = entry.get("hash", "")
    index = next((i for i, e in enumerate(ledger) if digest and e.get("hash") == digest), None)
    if index is None:
        entry["first_ts"] = entry["ts"]
    else:
        previous = ledger.pop(index)  # 移到末尾，保证 ledger[-1] 恒为本次
        entry["first_ts"] = str(previous.get("first_ts") or previous.get("ts") or entry["ts"])
        entry["repeats"] = int(previous.get("repeats", 1) or 1) + 1
    ledger.append(entry)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(ledger, ensure_ascii=False, indent=1), encoding="utf-8")
    return ledger


def _parse_run_start(run_id: str) -> datetime | None:
    """从 run_id 的 YYYYmmddTHHMMSS 前缀解析运行开始时间。"""
    try:
        return datetime.strptime(run_id.split("_", 1)[0], "%Y%m%dT%H%M%S")
    except ValueError:
        return None


def backfill_run_ids(ledger_dir: Path) -> int:
    """给缺 run_id 的历史台账记录补上产物索引（按输入文件内容哈希精确匹配）。

    台账的 hash 是"上传文件内容的 SHA256 前 16 位"，而 runs/<run_id>/input/ 下正是
    那次上传的原始副本，因此用哈希匹配是精确的、不依赖时间戳。同一文件多次运行时取
    最新一次（目录名按时间递增，后写覆盖）。幂等：已带 run_id 的记录不动。
    """
    path = ledger_dir / LEDGER_NAME
    ledger = load_ledger(path)
    pending = [e for e in ledger if not e.get("run_id") and e.get("hash")]
    if not pending or not ledger_dir.is_dir():
        return 0

    latest: dict[str, str] = {}
    for child in sorted(ledger_dir.iterdir()):
        if not child.is_dir() or not RUN_ID_RE.fullmatch(child.name):
            continue
        input_dir = child / "input"
        if not input_dir.is_dir():
            continue
        for source in sorted(input_dir.glob("*.xlsx")):
            digest = _file_digest(source)
            if digest:
                latest[digest] = child.name

    filled = 0
    for entry in ledger:
        if entry.get("run_id"):
            continue
        run_id = latest.get(str(entry.get("hash", "")))
        if run_id:
            entry["run_id"] = run_id
            entry["artifacts"] = ARTIFACT_AVAILABLE
            filled += 1
    if filled:
        path.write_text(json.dumps(ledger, ensure_ascii=False, indent=1), encoding="utf-8")
        LOGGER.info("ledger_backfill filled=%s candidates=%s", filled, len(latest))
    else:
        LOGGER.info("ledger_backfill filled=0 pending=%s candidates=%s", len(pending), len(latest))
    return filled


def retention_days() -> int:
    """产物保留天数；未配置用默认 90，配置 0（或非法值）表示不自动清理。"""
    raw = (os.environ.get(RETENTION_ENV) or "").strip()
    if not raw:
        return DEFAULT_RETENTION_DAYS
    try:
        return max(0, int(raw))
    except ValueError:
        return DEFAULT_RETENTION_DAYS


def purge_expired_artifacts(
    ledger_dir: Path,
    *,
    days: int | None = None,
    now: datetime | None = None,
) -> list[str]:
    """按保留策略清理过期产物目录，并把台账对应记录标记为已归档。

    安全约束：只处理名字严格符合 run_id 格式的目录，绝不触碰 runs/ 下的其他目录
    （回归样本、手工目录等）。days<=0 时直接跳过，不做任何删除。
    """
    if days is None:
        days = retention_days()
    if days <= 0 or not ledger_dir.is_dir():
        return []
    cutoff = (now or datetime.now()) - timedelta(days=days)

    purged: list[str] = []
    for child in sorted(ledger_dir.iterdir()):
        if not child.is_dir() or not RUN_ID_RE.fullmatch(child.name):
            continue
        started = _parse_run_start(child.name)
        if started is None or started >= cutoff:
            continue
        shutil.rmtree(child, ignore_errors=True)
        purged.append(child.name)
    if not purged:
        return []

    path = ledger_dir / LEDGER_NAME
    ledger = load_ledger(path)
    changed = False
    for entry in ledger:
        if entry.get("run_id") in purged:
            entry["artifacts"] = ARTIFACT_ARCHIVED
            changed = True
    if changed:
        path.write_text(json.dumps(ledger, ensure_ascii=False, indent=1), encoding="utf-8")
    LOGGER.info("artifacts_purged days=%s count=%s", days, len(purged))
    return purged


def _rate(issues: int, rows: int) -> str:
    if rows <= 0:
        return "—"
    return f"{issues / rows * 100:.1f}%"


def _dup_badge(entry: dict[str, Any]) -> str:
    """同一文件被重复审核时，在文件名后标注 ×N（悬停显示首次时间）。"""
    repeats = int(entry.get("repeats", 1) or 1)
    if repeats <= 1:
        return ""
    first = _esc(entry.get("first_ts", "") or "")
    return f' <span class="dup" title="重复审核 {repeats} 次，首次 {first}">×{repeats}</span>'


def render_history_html(path: Path, ledger: list[dict[str, Any]], ai_trend: dict[str, Any] | None = None) -> None:
    """渲染跨文件汇总报告：累计 KPI + 本次对比 + AI 趋势总结 + 逐文件明细。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    if not ledger:
        path.write_text("<!DOCTYPE html><html lang=zh-CN><meta charset=utf-8><body>暂无审核记录</body></html>", encoding="utf-8")
        return

    total_rows = sum(int(e.get("rows", 0) or 0) for e in ledger)
    total_issues = sum(int(e.get("issues", 0) or 0) for e in ledger)
    total_ai = sum(int(e.get("ai_reviewed", 0) or 0) for e in ledger)
    total_p0 = sum(int(e.get("priority", {}).get("P0", 0) or 0) for e in ledger)
    total_p1 = sum(int(e.get("priority", {}).get("P1", 0) or 0) for e in ledger)
    avg_rate = (total_issues / total_rows * 100) if total_rows else 0

    rows_html = ""
    for idx, entry in enumerate(reversed(ledger)):  # 最新在前
        hl = " class='hl'" if idx == 0 else ""
        ai_txt = f"{entry.get('ai_reviewed', 0)}/{entry.get('ai_candidate', 0)}" if entry.get("ai_reviewed", 0) else ("未启用" if entry.get("ai_status") in ("", "disabled") else str(entry.get("ai_reviewed", 0)))
        rows_html += (
            f"<tr{hl}><td>{_esc(entry.get('ts', ''))}</td>"
            f"<td>{_esc(entry.get('file', ''))}{_dup_badge(entry)}</td>"
            f"<td class='num'>{int(entry.get('rows', 0) or 0):,}</td>"
            f"<td class='num'>{int(entry.get('issues', 0) or 0):,}</td>"
            f"<td class='num'>{_rate(int(entry.get('issues', 0) or 0), int(entry.get('rows', 0) or 0))}</td>"
            f"<td class='num'>{int(entry.get('priority', {}).get('P0', 0) or 0):,}</td>"
            f"<td class='num'>{int(entry.get('priority', {}).get('P1', 0) or 0):,}</td>"
            f"<td class='num'>{int(entry.get('priority', {}).get('P2', 0) or 0):,}</td>"
            f"<td class='num'>{_esc(ai_txt)}</td>"
            f"<td>{'增量' if entry.get('incremental') else '全量'}</td></tr>"
        )

    latest = ledger[-1]
    latest_rate = (int(latest.get("issues", 0) or 0) / int(latest.get("rows", 0) or 0) * 100) if int(latest.get("rows", 0) or 0) else 0
    others = ledger[:-1]
    cmp_html = "（首个文件，暂无历史可对比）"
    if others:
        o_rows = sum(int(e.get("rows", 0) or 0) for e in others)
        o_issues = sum(int(e.get("issues", 0) or 0) for e in others)
        o_rate = (o_issues / o_rows * 100) if o_rows else 0
        diff = latest_rate - o_rate
        cls = "hbad" if diff > 0 else "hgood"
        cmp_html = (
            f"本次问题率 <b>{latest_rate:.1f}%</b>，历史文件（{len(others)} 个）平均 <b>{o_rate:.1f}%</b>，"
            f"<span class='{cls}'>{'高' if diff > 0 else '低'} {abs(diff):.1f} 个百分点</span>。"
        )

    trend_html = ""
    if isinstance(ai_trend, dict) and str(ai_trend.get("overall", "")).strip():
        pts = "".join(f"<li>{_esc(p)}</li>" for p in (ai_trend.get("points") or []) if str(p).strip())
        meta = (
            f"模型 {_esc(ai_trend.get('model', ''))} ｜ 基于 {int(ai_trend.get('file_count', len(ledger)) or len(ledger))} 次审核"
            f" ｜ {_esc(ai_trend.get('ts', ''))}"
        )
        trend_html = (
            '<div class="sec"><h2>AI 跨文件趋势总结</h2>'
            f'<div class="trend"><div class="trend-h">趋势与改进建议<span class="trend-m">{meta}</span></div>'
            f"<p>{_esc(ai_trend['overall'])}</p>"
            + (f"<ul>{pts}</ul>" if pts else "")
            + "</div></div>"
        )

    html = f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>KKS 审核台账汇总</title>
<style>{CSS}{_HIST_CSS}</style></head>
<body><div class="wrap">
<header><h1>KKS 审核台账汇总</h1><div class="meta">
累计 {len(ledger)} 次审核 ｜ 更新于 {datetime.now().strftime('%Y-%m-%d %H:%M')}</div></header>

<div class="kpis">
<div class="kpi"><div class="n">{len(ledger)}</div><div class="l">累计审核文件</div></div>
<div class="kpi"><div class="n">{total_rows:,}</div><div class="l">累计数据行</div></div>
<div class="kpi {'bad' if total_p0 else 'ok'}"><div class="n">{total_issues:,}</div><div class="l">累计问题（P0 {total_p0:,} / P1 {total_p1:,}）</div></div>
<div class="kpi"><div class="n">{avg_rate:.1f}%</div><div class="l">平均问题率</div></div>
<div class="kpi"><div class="n">{total_ai:,}</div><div class="l">累计 AI 复核</div></div>
</div>

<div class="sec"><h2>本次审核 vs 历史平均</h2>
<div class="sub">最新一次：<b>{_esc(latest.get('file', ''))}</b>（{_esc(latest.get('ts', ''))}）｜{_esc(latest.get('conclusion', ''))}</div>
<p>{cmp_html}</p></div>
{trend_html}
<div class="sec"><h2>审核历史明细（新→旧）</h2>
<table class="hrow" style="width:100%;border-collapse:collapse">
<tr><th>时间</th><th>文件</th><th>数据行</th><th>问题</th><th>问题率</th><th>P0</th><th>P1</th><th>P2</th><th>AI复核</th><th>口径</th></tr>
{rows_html}
</table>
<div class="sub" style="margin-top:10px">同一文件重复上传按内容哈希去重：仅保留最新结果并标注 ×N 次数，文件内容变化则记为一次新审核。单次明细见各次审核报告（HTML/Excel）；台账数据 audit_ledger.json 可回溯。</div>
</div>

<div class="foot">审计台账自动维护，每次审核后更新 ｜ {datetime.now().strftime('%Y-%m-%d')}</div>
</div></body></html>"""
    path.write_text(html, encoding="utf-8")


def record_and_render(
    input_path: Path,
    output_dir: Path,
    result: dict[str, Any],
    *,
    progress_callback: Any = None,
) -> Path:
    """审核完成后的统一入口：追加台账 + 清理过期产物 + 刷新 AI 趋势总结 + 重渲染汇总报告。"""
    ledger = append_ledger(input_path, output_dir, result)
    summary = history_path(output_dir)
    ledger_dir = output_dir.parent
    # 先按保留策略清理，再渲染：汇总里的"已归档"状态才是本轮的实际结果
    if purge_expired_artifacts(ledger_dir):
        ledger = load_ledger(ledger_dir / LEDGER_NAME)
    trend: dict[str, Any] = {}
    if len(ledger) >= 2:  # 仅 1 条记录时不存在"跨文件"趋势
        try:
            trend = summarize_ledger(ledger, progress_callback=progress_callback) or {}
        except Exception as exc:  # 趋势失败不阻断台账
            LOGGER.warning("ai_trend_failed_unexpected error=%s", exc)
            trend = {}
        if str(trend.get("overall", "")).strip():
            save_trend(ledger_dir, trend)
        else:
            trend = load_trend(ledger_dir)  # 本次未产出则沿用上次成功结果
    render_history_html(summary, ledger, ai_trend=trend)
    return summary
