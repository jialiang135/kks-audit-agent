"""审计台账：每次审核自动追加单次记录，并维护跨文件的累计汇总报告。

- ledger（audit_ledger.json）：追加式 JSON 数组，每次 write_outputs=True 的审核记一条；
- 汇总报告（KKS审核台账汇总.html）：随每次审核自动重渲染，含累计 KPI、
  逐文件明细表、以及“本次文件 vs 历史平均水平”的对比小节。
- 台账落在 output_dir 的上一级（服务器为 runs/ 根，随卷持久化）。
"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

from report_business import CSS, _esc

LEDGER_NAME = "audit_ledger.json"
HISTORY_NAME = "KKS审核台账汇总.html"

_HIST_CSS = """
.hrow td,.hrow th{padding:6px 10px;border-top:1px solid var(--line);font-size:13px}
.hrow th{background:var(--blue2);text-align:left}
.hrow td.num{text-align:right;font-variant-numeric:tabular-nums}
.hrow tr.hl td{background:#fffbe6}
.hbad{color:var(--p0);font-weight:500}.hgood{color:var(--ok);font-weight:500}
"""


def ledger_path(output_dir: Path) -> Path:
    """台账与汇总报告位于 output_dir 的上一级（runs/ 根）。"""
    return output_dir.parent / LEDGER_NAME


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


def _entry(input_path: Path, result: dict[str, Any]) -> dict[str, Any]:
    priority = result.get("priority_counts", {}) or {}
    ai = result.get("ai_review", {}) or {}
    final_counts = result.get("final_decision_counts", {}) or {}
    return {
        "ts": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "file": input_path.name,
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
    """追加本次审核记录并返回完整台账（供随后的汇总渲染使用）。"""
    path = ledger_path(output_dir)
    ledger = load_ledger(path)
    ledger.append(_entry(input_path, result))
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(ledger, ensure_ascii=False, indent=1), encoding="utf-8")
    return ledger


def _rate(issues: int, rows: int) -> str:
    if rows <= 0:
        return "—"
    return f"{issues / rows * 100:.1f}%"


def render_history_html(path: Path, ledger: list[dict[str, Any]]) -> None:
    """渲染跨文件汇总报告：累计 KPI + 逐文件明细 + 本次 vs 历史对比。"""
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
            f"<td>{_esc(entry.get('file', ''))}</td>"
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

<div class="sec"><h2>审核历史明细（新→旧）</h2>
<table class="hrow" style="width:100%;border-collapse:collapse">
<tr><th>时间</th><th>文件</th><th>数据行</th><th>问题</th><th>问题率</th><th>P0</th><th>P1</th><th>P2</th><th>AI复核</th><th>口径</th></tr>
{rows_html}
</table>
<div class="sub" style="margin-top:10px">单次明细见各次审核报告（HTML/Excel）；台账数据 audit_ledger.json 可回溯。</div>
</div>

<div class="foot">审计台账自动维护，每次审核后更新 ｜ {datetime.now().strftime('%Y-%m-%d')}</div>
</div></body></html>"""
    path.write_text(html, encoding="utf-8")


def record_and_render(input_path: Path, output_dir: Path, result: dict[str, Any]) -> Path:
    """审核完成后的统一入口：追加台账 + 重渲染汇总报告，返回汇总报告路径。"""
    ledger = append_ledger(input_path, output_dir, result)
    summary = history_path(output_dir)
    render_history_html(summary, ledger)
    return summary
