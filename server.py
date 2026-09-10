# -*- coding: utf-8 -*-
"""Standalone KKS audit HTTP service.

No WorkBuddy/FastAPI dependency is required. The service accepts multipart
Excel uploads, runs the read-only audit engine, and exposes report downloads.
"""
from __future__ import annotations

import argparse
import json
import logging
import re
import threading
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, quote, unquote, urlparse
from uuid import uuid4

from ai_review import load_config, normalize_base_url, test_ai_connection
from agent_runtime import KksAuditAgent
from app_runtime import (
    APP_ROOT,
    configure_logging,
    install_skill_zip,
    load_app_config,
    load_skill_context,
    mask_secret,
    save_app_config,
    save_skill_document,
    skill_documents,
    tail_log,
)
from run_audit import issue_rule_tree, output_artifact_names, quality_metrics, rule_source
from report_business import build_overview as business_overview


MAX_UPLOAD_BYTES = 50 * 1024 * 1024
RUN_ID_RE = re.compile(r"^[0-9]{8}T[0-9]{6}_[0-9a-f]{8}$")
LOGGER = logging.getLogger("kks-audit")


def _light_tree(result: dict[str, Any]) -> list[dict[str, Any]]:
    """把问题按「规则 → KKS → 行号明细」组织为 Web 前端预览树。

    直接复用 run_audit.issue_rule_tree 产出的轻量可序列化节点。
    """
    return issue_rule_tree(result)


def _update_job_state(server: ThreadingHTTPServer, run_id: str, **updates: object) -> dict[str, object]:
    with server.jobs_lock:  # type: ignore[attr-defined]
        state = dict(server.jobs.get(run_id, {}))  # type: ignore[attr-defined]
        state.update(updates)
        state["updated_at"] = __import__("datetime").datetime.now().isoformat(timespec="seconds")
        server.jobs[run_id] = state  # type: ignore[attr-defined]
        return dict(state)


def _get_job_state(server: ThreadingHTTPServer, run_id: str) -> dict[str, object] | None:
    with server.jobs_lock:  # type: ignore[attr-defined]
        state = server.jobs.get(run_id)  # type: ignore[attr-defined]
        return dict(state) if isinstance(state, dict) else None


def _apply_progress_event(server: ThreadingHTTPServer, run_id: str, event: dict[str, object]) -> None:
    state = _get_job_state(server, run_id) or {}
    ai_review = dict(state.get("ai_review", {})) if isinstance(state.get("ai_review"), dict) else {}
    agent = dict(state.get("agent", {})) if isinstance(state.get("agent"), dict) else {}
    for key in ("agent_status", "agent_stage", "agent_sequence", "tool", "plan_step"):
        if key in event:
            agent[key] = event[key]
    if event.get("phase") == "ai":
        for key in ("status", "model", "candidate_count", "reviewed_count", "parallelism", "completed_groups", "group_count"):
            if key in event:
                ai_review[key] = event[key]
    _update_job_state(
        server,
        run_id,
        status="running",
        phase=event.get("stage", "running"),
        percent=int(event.get("percent", state.get("percent", 0))),
        message=str(event.get("message", state.get("message", "正在审核"))),
        hint=str(event.get("hint", state.get("hint", ""))),
        ai_review=ai_review,
        agent=agent,
    )


def _run_audit_job(
    server: ThreadingHTTPServer,
    run_id: str,
    input_path: Path,
    run_dir: Path,
    comparison_path: Path | None,
) -> None:
    try:
        _update_job_state(
            server,
            run_id,
            status="running",
            phase="starting",
            percent=5,
            message="审核任务已启动",
            hint="准备创建 KKS 审核 Agent",
            agent={"name": "kks-coding-audit", "status": "created", "current_step": "created", "sequence": 0},
        )
        agent = KksAuditAgent()
        result = agent.run(
            input_path,
            run_dir,
            run_id=run_id,
            comparison_path=comparison_path,
            progress_callback=lambda event: _apply_progress_event(server, run_id, event),
        )
        html_name, xlsx_name = output_artifact_names(input_path)
        files = {label: f"/api/audits/{run_id}/{artifact}" for label, artifact in (("HTML报告", html_name), ("Excel问题清单", xlsx_name))}
        result["run_id"] = run_id
        result["files"] = files
        _update_job_state(
            server,
            run_id,
            status="completed",
            phase="completed",
            percent=100,
            message="审核完成",
            hint="报告已生成，可以下载结果",
            conclusion=result["conclusion"],
            data_rows=result["data_rows"],
            issue_count=result.get("issue_count", sum(1 for item in result.get("issues", []) if item.get("status") != "resolved")),
            resolved_count=result.get("resolved_count", len(result.get("resolved_reviews", []))),
            priority_counts=result["priority_counts"],
            quality_metrics=quality_metrics(result),
            ai_review=result.get("ai_review", {}),
            scope_comparison=result.get("scope_comparison"),
            agent=result.get("agent_runtime", {}),
            issue_preview=_light_tree(result),
            business_overview=business_overview(result),
            files=files,
        )
        LOGGER.info("audit_http_done run_id=%s ai_status=%s", run_id, result.get("ai_review", {}).get("status", "disabled"))
    except Exception as exc:
        LOGGER.exception("audit_http_failed run_id=%s", run_id)
        _update_job_state(server, run_id, status="error", phase="error", percent=0, message="审核失败", hint=str(exc), error=str(exc))


RESULT_UI_STYLE = """<style>
.result-panel{margin-top:22px;padding:30px 32px 32px;scroll-margin-top:24px}
.result-panel-head{align-items:center;margin-bottom:22px}
.result-panel-head h3{font-size:22px}
.section-kicker{margin-bottom:6px;color:#6366f1;font-size:11px;font-weight:800;letter-spacing:1.6px}
.result-state{display:inline-flex;align-items:center;gap:7px;padding:8px 12px;border:1px solid #e2e8f0;border-radius:999px;background:#f8fafc;color:#64748b;font-size:12px;font-weight:750;white-space:nowrap}
.result-state:before{content:"";width:7px;height:7px;border-radius:50%;background:#94a3b8}
.result-state.running{background:#eff6ff;border-color:#bfdbfe;color:#1d4ed8}.result-state.running:before{background:#3b82f6;box-shadow:0 0 0 4px #dbeafe}
.result-state.done{background:#ecfdf5;border-color:#bbf7d0;color:#047857}.result-state.done:before{background:#10b981}
.result-state.failed{background:#fff1f2;border-color:#fecdd3;color:#be123c}.result-state.failed:before{background:#f43f5e}
.progress-panel{margin:0 0 16px;padding:16px 18px;border:1px solid #c7d2fe;border-radius:14px;background:linear-gradient(180deg,#fbfbff,#f5f6ff)}
.progress-topline{display:flex;justify-content:space-between;gap:12px;align-items:center;color:#334155;font-size:13px}.progress-topline b{color:#4f46e5;font-size:13px}
.progress-track{height:9px;margin-top:11px;border-radius:999px;background:#e0e7ff;overflow:hidden}.progress-track>div{height:100%;width:0;border-radius:inherit;background:linear-gradient(90deg,#6366f1,#4f46e5);transition:width .45s ease}
.progress-hint{margin-top:9px;color:#64748b;font-size:12px;line-height:1.5}
.result-box{min-height:96px;padding:20px 22px;border:1px solid #e8edf5;border-radius:14px;background:#f8fafc;white-space:normal;line-height:1.6}
.result-box-running{border-color:#bfdbfe;background:#f8fbff}.result-box-complete{border-color:#bbf7d0;background:linear-gradient(135deg,#f7fffb,#f8fafc)}.result-box-failed{border-color:#fecdd3;background:#fff8f9}
.result-message{display:flex;align-items:flex-start;gap:13px}.result-message-icon{display:grid;place-items:center;flex:0 0 30px;width:30px;height:30px;border-radius:10px;background:#e0e7ff;color:#4f46e5;font-size:21px;line-height:1}.result-box-complete .result-message-icon{background:#d1fae5;color:#059669}.result-box-failed .result-message-icon{background:#ffe4e6;color:#e11d48}
.result-message strong{display:block;color:#1e293b;font-size:16px;font-weight:750}.result-message span:not(.result-message-icon){display:block;margin-top:4px;color:#64748b;font-size:13px}
.result-stats{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin-top:14px}.result-stat{min-height:86px;padding:15px 16px;border:1px solid #e2e8f0;border-radius:12px;background:#fff}.result-stat small{display:block;color:#64748b;font-size:12px;font-weight:650}.result-stat>b{display:block;margin-top:6px;color:#111827;font-size:23px;font-weight:750}.result-stat-hint{display:block;margin-top:3px;color:#94a3b8;font-size:11px}.result-priority>div{display:flex;gap:6px;margin-top:10px;flex-wrap:wrap}.priority-chip{display:inline-flex;gap:4px;align-items:center;padding:4px 7px;border-radius:6px;font-size:11px;font-weight:700}.priority-chip b{font-size:12px}.priority-chip.p0{background:#fff1f2;color:#be123c}.priority-chip.p1{background:#fffbeb;color:#b45309}.priority-chip.p2{background:#eff6ff;color:#1d4ed8}
.result-downloads{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;margin-top:16px}.download-card{display:flex;align-items:center;gap:13px;min-height:76px;padding:14px 15px;border:1px solid #e2e8f0;border-radius:12px;background:#fff;color:inherit;text-decoration:none;transition:border-color .18s,box-shadow .18s,transform .18s}.download-card:hover{border-color:#a5b4fc;box-shadow:0 8px 20px rgba(79,70,229,.1);transform:translateY(-1px)}.download-icon{display:grid;place-items:center;flex:0 0 38px;width:38px;height:38px;border-radius:11px;background:#eef2ff;color:#4f46e5;font-size:14px;font-weight:800}.download-card:nth-child(2) .download-icon{background:#ecfdf5;color:#047857}.download-copy{min-width:0;flex:1}.download-copy strong{display:block;color:#1e293b;font-size:14px}.download-copy span{display:block;margin-top:3px;color:#94a3b8;font-size:11px}.download-action{color:#4f46e5;font-size:12px;font-weight:750;white-space:nowrap}
@media(max-width:900px){.result-panel{padding:26px 24px 28px}.result-stats{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media(max-width:560px){.result-panel{padding:22px 18px 22px}.result-panel-head{align-items:flex-start;flex-direction:column;gap:12px}.result-stats,.result-downloads{grid-template-columns:1fr}.result-stat{min-height:76px}}
</style>"""


AUDIT_REDESIGN_STYLE = """<style>
.audit-intro{display:flex;align-items:flex-end;justify-content:space-between;gap:22px;padding:4px 2px 4px}
.audit-intro h2{margin:0;color:#111827;font-size:24px;line-height:1.3;letter-spacing:-.45px}.audit-intro p{margin:8px 0 0;color:#64748b;font-size:14px;line-height:1.7}
.audit-source-badge{display:inline-flex;align-items:center;gap:8px;padding:9px 12px;border:1px solid #dbeafe;border-radius:9px;background:#eff6ff;color:#1d4ed8;font-size:12px;font-weight:750;white-space:nowrap}.audit-source-badge:before{content:"";width:7px;height:7px;border-radius:50%;background:#3b82f6}
.audit-input-grid{display:grid;grid-template-columns:minmax(0,1.55fr) minmax(280px,.75fr);gap:18px;margin-top:18px}.upload-panel,.audit-guide{margin-top:0}.upload-panel .panel-head{margin-bottom:15px}.upload-panel .panel-head h3,.audit-guide h3{font-size:18px}.upload-panel .upload-zone.compact{min-height:224px;padding:24px 20px}.upload-zone.compact .upload-icon{width:44px;height:44px;margin-bottom:10px;font-size:24px;border-radius:13px}.upload-zone.compact .upload-select{margin-top:13px}.upload-zone.compact .upload-actions{margin-top:17px}.upload-zone.compact .file-name{margin-top:9px}
.audit-guide{display:flex;flex-direction:column}.audit-guide-header{display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:15px}.guide-caption{color:#94a3b8;font-size:12px}.workflow-list{display:grid;gap:4px}.workflow-item{display:flex;gap:11px;align-items:flex-start;padding:11px 10px;border-radius:10px}.workflow-item:hover{background:#f8fafc}.workflow-number{display:grid;place-items:center;flex:0 0 27px;width:27px;height:27px;border-radius:8px;background:#eef2ff;color:#4f46e5;font-size:11px;font-weight:800}.workflow-item strong{display:block;color:#1e293b;font-size:13px}.workflow-item span:not(.workflow-number){display:block;margin-top:3px;color:#94a3b8;font-size:11px;line-height:1.45}.audit-guide-note{margin-top:auto;padding:11px 12px;border:1px solid #e2e8f0;border-radius:10px;background:#f8fafc;color:#64748b;font-size:11px;line-height:1.6}.audit-section-bar{display:flex;justify-content:space-between;align-items:center;gap:14px;margin-top:22px;margin-bottom:10px}.audit-section-bar h3{margin:0;color:#1e293b;font-size:16px}.audit-section-bar p{margin:3px 0 0;color:#94a3b8;font-size:12px}.overview-state{padding:5px 9px;border-radius:999px;background:#f8fafc;color:#64748b;font-size:11px;font-weight:750}.summary-grid{margin-top:0}.metric{position:relative;overflow:hidden}.metric:after{content:"";position:absolute;right:-20px;bottom:-22px;width:70px;height:70px;border-radius:50%;background:#f8fafc}.metric b,.metric small{position:relative;z-index:1}.result-panel{margin-top:18px}.result-panel-head{padding-bottom:2px}
@media(max-width:900px){.audit-input-grid{grid-template-columns:1fr}.audit-guide{min-height:0}.audit-guide-note{margin-top:15px}}
@media(max-width:560px){.audit-intro{display:block}.audit-source-badge{margin-top:12px}.audit-section-bar{align-items:flex-start;flex-direction:column;gap:8px}.upload-panel .panel-head{display:block}}
 .overview-state.running{background:#eff6ff;color:#1d4ed8}.overview-state.done{background:#ecfdf5;color:#047857}.overview-state.failed{background:#fff1f2;color:#be123c}
 *[hidden]{display:none!important}body:has(#auditPage.active){overflow:hidden}body:has(#auditPage.active) .content{height:100vh;overflow:hidden;padding:14px 28px 16px}body:has(#auditPage.active) .topbar{min-height:48px;margin-bottom:12px}body:has(#auditPage.active) .topbar h1{font-size:26px;margin-top:2px}body:has(#auditPage.active) .topbar-subtitle{display:none}body:has(#auditPage.active) .topbar .pill{min-height:32px;padding:7px 12px}
 #auditPage.active .upload-panel{margin-top:0;padding:16px 18px}#auditPage.active .upload-panel .panel-head{margin-bottom:10px}#auditPage.active .upload-panel .panel-head>div{display:flex;align-items:baseline;gap:12px}#auditPage.active .upload-panel .panel-head h3{font-size:18px}#auditPage.active .upload-panel .muted{font-size:12px;line-height:1.4}#auditPage.active .upload-zone.compact{min-height:112px;padding:14px 16px;display:grid;grid-template-columns:44px minmax(180px,1fr) auto minmax(118px,auto);grid-template-areas:"icon title select action" "icon desc file action" "icon hint hint action";align-items:center;column-gap:14px;row-gap:2px;text-align:left}#auditPage.active .upload-zone.compact .upload-icon{grid-area:icon;width:44px;height:44px;margin:0}#auditPage.active .upload-zone.compact .upload-title{grid-area:title}#auditPage.active .upload-zone.compact .upload-desc{grid-area:desc;margin:0}#auditPage.active .upload-zone.compact .upload-select{grid-area:select;margin:0;white-space:nowrap}#auditPage.active .upload-zone.compact .file-name{grid-area:file;min-height:0;margin:0}#auditPage.active .upload-zone.compact .upload-drop-hint{grid-area:hint;margin:0}#auditPage.active .upload-zone.compact .upload-actions{grid-area:action;margin:0;align-self:center}#auditPage.active .audit-section-bar{margin-top:10px;margin-bottom:7px}#auditPage.active .audit-section-bar p{display:none}#auditPage.active .summary-grid{gap:10px;margin-top:0}#auditPage.active .metric{min-height:64px;padding:11px 14px}#auditPage.active .metric b{font-size:24px;margin-top:3px}#auditPage.active .result-panel{margin-top:10px;padding:14px 16px 16px}#auditPage.active .result-panel-head{margin-bottom:10px}#auditPage.active .result-panel-head .section-kicker,#auditPage.active .result-panel-head .muted{display:none}#auditPage.active .result-panel-head h3{font-size:18px}#auditPage.active .progress-panel{margin-bottom:8px;padding:10px 12px}#auditPage.active .progress-track{height:7px;margin-top:7px}#auditPage.active .progress-hint{margin-top:5px}#auditPage.active .result-box{min-height:54px;padding:12px 14px}#auditPage.active .result-message{gap:10px}#auditPage.active .result-message-icon{flex-basis:26px;width:26px;height:26px;font-size:18px}#auditPage.active .result-message strong{font-size:14px}#auditPage.active .result-message span:not(.result-message-icon){margin-top:2px;font-size:12px}#auditPage.active .result-stats{gap:8px;margin-top:8px}#auditPage.active .result-stat{min-height:62px;padding:10px 12px}#auditPage.active .result-stat>b{margin-top:3px;font-size:20px}#auditPage.active .result-priority>div{margin-top:6px}#auditPage.active .result-downloads{gap:8px;margin-top:8px}#auditPage.active .download-card{min-height:52px;padding:10px 12px}#auditPage.active .download-icon{flex-basis:30px;width:30px;height:30px;border-radius:9px;font-size:11px}#auditPage.active .download-copy span{display:none}
 @media(max-width:900px){body:has(#auditPage.active){overflow:auto}body:has(#auditPage.active) .content{height:auto;overflow:visible;padding:20px 18px 32px}#auditPage.active .upload-panel .panel-head>div{display:block}#auditPage.active .upload-zone.compact{min-height:220px;padding:20px;display:flex;flex-direction:column;align-items:center;text-align:center;gap:0}#auditPage.active .upload-zone.compact .upload-icon,#auditPage.active .upload-zone.compact .upload-title,#auditPage.active .upload-zone.compact .upload-desc,#auditPage.active .upload-zone.compact .upload-select,#auditPage.active .upload-zone.compact .file-name,#auditPage.active .upload-zone.compact .upload-drop-hint,#auditPage.active .upload-zone.compact .upload-actions{display:block}#auditPage.active .upload-zone.compact .upload-icon{margin-bottom:9px}#auditPage.active .upload-zone.compact .upload-desc{margin-top:3px}#auditPage.active .upload-zone.compact .upload-select{margin-top:12px}#auditPage.active .upload-zone.compact .file-name{margin-top:9px}#auditPage.active .upload-zone.compact .upload-drop-hint{margin-top:4px}#auditPage.active .upload-zone.compact .upload-actions{margin-top:15px}}
 </style>"""


DASHBOARD_UI_STYLE = """<style>
body:has(#auditPage.active) .content{height:100vh;overflow:hidden;padding:18px 28px 18px}body:has(#auditPage.active) .topbar{min-height:52px;margin-bottom:12px}body:has(#auditPage.active) .topbar h1{font-size:28px;margin-top:2px}body:has(#auditPage.active) .topbar-subtitle{display:block;margin-top:4px;color:#64748b;font-size:14px}.top-actions{display:flex;align-items:center;gap:10px}.btn-header{min-height:38px;padding:8px 13px;border:1px solid #e2e8f0;background:#fff;color:#1e3a8a;box-shadow:0 3px 10px rgba(15,23,42,.04);font-size:13px}.btn-header:hover{background:#f8faff;border-color:#a5b4fc}.nav-link .nav-icon{display:grid;place-items:center;width:24px;height:24px;color:#64748b;font-size:18px;font-weight:700;line-height:1}.nav-link.active .nav-icon{color:#2563eb}.side-ai-status{margin-top:auto;padding:13px 14px;border:1px solid #dbeafe;border-radius:12px;background:#f8fbff}.side-ai-title{display:block;color:#16a34a;font-size:14px;font-weight:750}.side-ai-title:not(.ready){color:#64748b}.side-ai-status small{display:block;margin-top:5px;color:#94a3b8;font-size:11px}.side-footer{margin-top:18px;padding:0 10px;color:#94a3b8;font-size:11px;line-height:2}.side-footer span{color:#cbd5e1}.dashboard-upload{margin-top:0;padding:14px 16px}.upload-strip{display:grid;grid-template-columns:minmax(285px,1.05fr) minmax(190px,1.25fr) 220px;gap:20px;align-items:center}.dashboard-drop{min-height:84px;padding:13px 16px;display:flex;flex-direction:row;align-items:center;justify-content:flex-start;gap:13px;text-align:left;border:1px dashed #8faeff;border-radius:12px;background:#fbfdff}.dashboard-drop.dragging{background:#eef4ff;border-color:#2563eb}.upload-file-icon{display:grid;place-items:center;flex:0 0 45px;width:45px;height:49px;border:2px solid #94a3b8;border-radius:5px;background:linear-gradient(135deg,#fff 70%,#dbeafe 70%);color:#16a34a;font-size:25px;font-weight:800}.upload-copy{min-width:0}.dashboard-drop .upload-title{font-size:15px}.dashboard-drop .upload-desc{margin-top:4px;font-size:12px}.upload-link{color:#2563eb;font-weight:700;cursor:pointer}.file-details{min-width:0}.dashboard-upload .file-name{margin:0;min-height:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#1e293b;font-size:14px;font-weight:750}.file-meta{display:flex;gap:8px;align-items:center;margin-top:8px;color:#94a3b8;font-size:12px}.file-meta span:first-child{color:#64748b}.upload-state{margin-top:7px;color:#16a34a;font-size:11px;font-weight:700}.upload-actions{display:flex;flex-direction:column;align-items:stretch;gap:6px;margin:0}.upload-submit{min-height:54px;font-size:16px}.upload-actions>span{color:#64748b;text-align:center;font-size:11px}.dashboard-metrics{grid-template-columns:repeat(7,minmax(0,1fr));gap:12px;margin-top:12px}.dashboard-metric{min-height:76px;padding:13px 15px;display:flex;flex-direction:row;align-items:center;gap:12px}.dashboard-metric:after{display:none}.metric-icon{display:grid;place-items:center;flex:0 0 42px;width:42px;height:42px;border-radius:50%;background:#eff6ff;color:#2563eb;font-size:22px;font-weight:800}.dashboard-metric.quality .metric-icon{background:#eef2ff;color:#4f46e5}.dashboard-metric.p0 .metric-icon{background:#fee2e2;color:#dc2626}.dashboard-metric.p1 .metric-icon{background:#fef3c7;color:#d97706}.dashboard-metric.p2 .metric-icon{background:#dbeafe;color:#2563eb}.dashboard-metric small{font-size:12px}.dashboard-metric b{margin-top:3px;font-size:24px}.metric-note{display:block;margin-top:2px;color:#94a3b8;font-size:11px}.dashboard-results{margin-top:12px;padding:13px 16px 14px}.dashboard-results .result-panel-head{margin-bottom:8px;align-items:center}.dashboard-results .result-panel-head h3{font-size:18px}.dashboard-results .result-panel-head .muted{margin-top:2px;font-size:11px}.dashboard-results .progress-panel{margin-bottom:8px;padding:9px 12px}.dashboard-results .progress-track{height:7px;margin-top:6px}.dashboard-results .progress-hint{margin-top:4px;font-size:11px}.dashboard-results .result-box{min-height:48px;padding:10px 12px}.dashboard-results .result-message{gap:9px}.dashboard-results .result-message-icon{width:25px;height:25px;flex-basis:25px;font-size:17px}.dashboard-results .result-message strong{font-size:13px}.dashboard-results .result-message span:not(.result-message-icon){margin-top:1px;font-size:11px}.result-dashboard{margin-top:8px}.severity-card{padding:10px 12px;border:1px solid #edf1f6;border-radius:13px;background:#fcfdff}.section-row{display:flex;align-items:center;justify-content:space-between;gap:10px}.section-row h3,.download-panel h3{margin:0;color:#1e293b;font-size:14px}.info-dot{display:grid;place-items:center;width:15px;height:15px;border:1px solid #94a3b8;border-radius:50%;color:#64748b;font-size:10px;font-weight:800}.severity-layout{display:grid;grid-template-columns:minmax(250px,1fr) 112px 150px 230px;gap:14px;align-items:center;margin-top:8px}.severity-bar{display:flex;height:18px;overflow:hidden;border-radius:8px;background:#eef2f7}.severity-bar span{display:block;height:100%;min-width:0;transition:width .3s ease}.severity-bar span:nth-child(1){background:#ef4444}.severity-bar span:nth-child(2){background:#f59e0b}.severity-bar span:nth-child(3){background:#2563eb}.severity-legend{display:flex;gap:15px;margin-top:7px;flex-wrap:wrap;color:#64748b;font-size:11px}.severity-legend span{display:flex;align-items:center;gap:5px;white-space:nowrap}.legend-dot{display:inline-block;width:8px;height:8px;border-radius:2px;background:#2563eb}.legend-dot.p0{background:#ef4444}.legend-dot.p1{background:#f59e0b}.legend-dot.p2{background:#2563eb}.severity-legend b{color:#475569}.issue-donut{display:grid;place-items:center;width:92px;height:92px;margin:auto;border-radius:50%;background:conic-gradient(#2563eb 0 100%);position:relative}.issue-donut:after{content:"";position:absolute;inset:9px;border-radius:50%;background:#fff}.issue-donut>div{position:relative;z-index:1;text-align:center}.issue-donut strong{display:block;color:#172033;font-size:22px;line-height:1.1}.issue-donut small{display:block;margin-top:3px;color:#64748b;font-size:10px}.valid-card{display:flex;align-items:center;gap:10px;padding:10px 12px;border:1px solid #e2e8f0;border-radius:11px}.valid-icon{display:grid;place-items:center;width:38px;height:38px;border-radius:50%;background:#ecfdf5;color:#16a34a;font-size:22px;font-weight:800}.valid-card small{display:block;color:#64748b;font-size:11px}.valid-card b{display:block;margin-top:2px;color:#172033;font-size:20px}.valid-card span:last-child{display:block;margin-top:2px;color:#94a3b8;font-size:10px}.priority-card{padding:8px 11px;border:1px solid #e2e8f0;border-radius:11px}.priority-row{display:flex;align-items:center;justify-content:space-between;min-height:20px;color:#64748b;font-size:11px}.priority-row span{display:flex;align-items:center;gap:7px}.priority-row b{color:#334155}.priority-row.total{margin-top:3px;padding-top:4px;border-top:1px solid #eef2f7;color:#64748b}.result-lower{display:grid;grid-template-columns:minmax(0,1.3fr) minmax(270px,.7fr);gap:16px;margin-top:10px;padding-top:10px;border-top:1px solid #eef2f7}.issue-preview-card{min-width:0}.text-action{border:0;background:transparent;color:#2563eb;font:inherit;font-size:11px;font-weight:700;cursor:pointer}.issue-table-wrap{margin-top:7px;overflow:hidden;border:1px solid #e2e8f0;border-radius:10px}.issue-table-wrap table{width:100%;border-collapse:collapse;table-layout:fixed}.issue-table-wrap th,.issue-table-wrap td{padding:5px 7px;border-bottom:1px solid #eef2f7;text-align:left;font-size:10px;line-height:1.35;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.issue-table-wrap th{color:#94a3b8;font-weight:600;background:#fbfcfe}.issue-table-wrap tr:last-child td{border-bottom:0}.issue-table-wrap th:nth-child(1),.issue-table-wrap td:nth-child(1){width:48px}.issue-table-wrap th:nth-child(3),.issue-table-wrap td:nth-child(3){width:100px}.issue-table-wrap th:nth-child(4),.issue-table-wrap td:nth-child(4){width:66px}.issue-table-wrap th:nth-child(5),.issue-table-wrap td:nth-child(5){width:72px}.priority-cell{font-weight:800}.priority-cell.p0{color:#dc2626}.priority-cell.p1{color:#d97706}.priority-cell.p2{color:#2563eb}.preview-status{color:#f59e0b;font-weight:700}.src-cell{color:#0e5a44;font-size:12px;font-weight:600}.empty-preview{text-align:center!important;color:#94a3b8!important}.issue-tree-box{max-height:340px;overflow:auto;padding:6px 8px}.it-node{border:1px solid #e2e8f0;border-radius:9px;margin:4px 0;background:#fbfdff;overflow:hidden}.it-node>summary{list-style:none;display:flex;align-items:center;gap:8px;flex-wrap:wrap;padding:6px 10px;cursor:pointer;font-weight:650;color:#1e3a8a;background:linear-gradient(180deg,#f2f7fd,#eaf2fa);font-size:12px}.it-node>summary::-webkit-details-marker{display:none}.it-node>summary:before{content:"▸";color:#8ba3bf;font-size:11px;width:12px}.it-node[open]>summary:before{content:"▾"}.it-node[open]>summary{background:#eaf2fa;border-bottom:1px solid #e2e8f0}.it-name{font-family:Consolas,Menlo,monospace;font-size:11.5px;color:#0b5394;font-weight:800}.it-title{color:#5b7590;font-size:11.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:300px}.it-count{color:#8aa0b8;font-size:11px;margin-left:auto}.tg-rule>summary .it-count{margin-left:0;flex:0 0 108px;text-align:right}.tb{min-width:28px;padding:1px 6px;border-radius:999px;font-size:10.5px;font-weight:800;text-align:center}.tb-p0{background:#fef2f2;color:#b91c1c;border:1px solid #fecaca}.tb-p1{background:#fffbeb;color:#b45309;border:1px solid #fde68a}.tb-p2{background:#eff6ff;color:#1d4ed8;border:1px solid #bfdbfe}.it-kids{padding:2px 6px 6px 16px}.it-issues{padding:1px 10px 4px;background:#fff}.it-issue{padding:5px 2px 3px;border-top:1px dashed #e2e8f0;font-size:11px;line-height:1.5;color:#334155}.it-issue .priority{min-width:auto;padding:1px 6px;font-size:10px}.it-issue code{margin:0 4px;font-size:10px}.it-issue .it-def{color:#5b7590;font-size:10.5px}.it-issue b{margin-right:3px;font-size:10.5px}.it-suggest{color:#94a3b8;font-size:10.5px;margin-top:1px}.it-leaf{display:flex;align-items:center;gap:8px;flex-wrap:wrap;padding:4px 10px;color:#627d98;font-size:11px}.tg-rule{border-left:3px solid #2563eb;background:#f6f9fe}.it-node.tg-rule>summary{display:grid;grid-template-columns:12px auto max-content minmax(0,1fr) 58px 108px;align-items:center;column-gap:8px;background:linear-gradient(180deg,#eef4fc,#e2ecf8);font-size:12.5px;padding:7px 10px}.tg-rule[open]>summary{background:#e2ecf8;border-bottom:1px solid #dbe7f5}.it-node.tg-rule>summary:before{color:#5b86c6}.tg-rid{font-size:12.5px}.tg-def{flex:0 1 auto;min-width:0;color:#42536b;font-size:11px;max-width:280px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.tg-act{display:inline-flex;align-items:center;justify-content:center;line-height:1;padding:1px 8px;border-radius:999px;font-size:10px;font-weight:800;white-space:nowrap}.tg-act.act-fix{background:#fee2e2;color:#b91c1c;border:1px solid #fecaca}.tg-act.act-hm{background:#fef3c7;color:#a16207;border:1px solid #fde68a}.tg-act.act-ck{background:#dbeafe;color:#1d4ed8;border:1px solid #bfdbfe}.tg-src{padding:2px 12px 7px;color:#7a93ac;font-size:10.5px;background:#fff}.it-code{margin-left:2px;border-color:#e7eef6;background:#fff}.it-code>summary{background:#fff;border-bottom:1px dashed #e2e8f0;padding:4px 10px;font-size:11.5px;font-weight:700}.it-code[open]>summary{background:#f7fafd}.it-row{display:inline-flex;flex:0 0 auto;align-items:center;padding:0 7px;border-radius:999px;background:#eef2f7;color:#40556b;font-size:10px;font-weight:800;white-space:nowrap}.it-issue .it-msg{flex:1 1 300px;color:#334155}.it-issue{display:flex;align-items:flex-start;gap:6px;flex-wrap:wrap}.download-panel{min-width:0;border-left:1px solid #eef2f7;padding-left:16px}.download-panel>h3{margin-bottom:7px}.download-panel .result-downloads{display:grid;grid-template-columns:1fr;gap:8px;margin:0}.download-panel .download-card{min-height:54px;padding:9px 10px;gap:9px}.download-panel .download-icon{width:30px;height:30px;flex-basis:30px;border-radius:8px;font-size:10px}.download-panel .download-copy strong{font-size:12px}.download-panel .download-copy span{font-size:10px}.download-panel .download-action{font-size:11px}
 @media(max-width:1000px){.upload-strip{grid-template-columns:minmax(250px,1fr) minmax(180px,1fr)}.upload-actions{grid-column:1/-1;display:grid;grid-template-columns:1fr 1fr;align-items:center}.upload-actions>span{text-align:left}.dashboard-metrics{grid-template-columns:repeat(4,minmax(0,1fr))}.severity-layout{grid-template-columns:minmax(220px,1fr) 100px 150px}.priority-card{grid-column:1/-1}.result-lower{grid-template-columns:minmax(0,1fr) 260px}}
 @media(max-width:700px){body:has(#auditPage.active){overflow:auto}body:has(#auditPage.active) .content{height:auto;overflow:visible;padding:18px}.top-actions{margin-top:10px}.topbar{display:block}.upload-strip,.result-lower{grid-template-columns:1fr}.upload-actions{grid-column:auto;display:flex}.upload-actions>span{text-align:center}.severity-layout{grid-template-columns:1fr 100px}.severity-main{grid-column:1/-1}.valid-card{grid-column:1/-1}.priority-card{grid-column:1/-1}.download-panel{border-left:0;border-top:1px solid #eef2f7;padding:10px 0 0}.dashboard-metrics{grid-template-columns:repeat(2,1fr)}.dashboard-metric{min-height:68px;padding:10px}.metric-icon{width:34px;height:34px;flex-basis:34px;font-size:17px}.dashboard-metric b{font-size:21px}}
 @media(max-width:460px){.dashboard-metrics{grid-template-columns:1fr}.severity-layout{grid-template-columns:1fr}.issue-donut{grid-column:1}.top-actions{flex-wrap:wrap}.btn-header{flex:1}}
</style>"""


UPLOAD_PAGE = """<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>KKS 审核工作台</title>
<style>
:root{--ink:#172b3a;--muted:#6b7c88;--line:#dce8ed;--teal:#0b6570;--teal-dark:#084953;--mint:#e7f5f3;--orange:#e88942;--bg:#f5f8fa}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font-family:Segoe UI,Microsoft YaHei,sans-serif}.shell{min-height:100vh;display:grid;grid-template-columns:236px 1fr}.sidebar{background:#083f4a;color:#d9eef0;padding:26px 16px;display:flex;flex-direction:column}.brand{display:flex;align-items:center;gap:11px;padding:4px 10px 35px}.brand-mark{width:36px;height:36px;border-radius:11px;background:#e9a15a;color:#083f4a;display:grid;place-items:center;font-weight:800;font-size:20px}.brand strong{display:block;color:#fff;font-size:16px}.brand small{display:block;color:#9bc4c7;margin-top:3px}.nav-label{font-size:11px;letter-spacing:1.5px;color:#82b3b7;padding:0 12px 9px}.nav-link{width:100%;border:0;background:transparent;color:#b8d8da;text-align:left;padding:12px;border-radius:10px;margin:3px 0;cursor:pointer;font:inherit;display:flex;align-items:center;gap:10px}.nav-link span{font-size:11px;color:#77aeb3}.nav-link:hover,.nav-link.active{background:#125b66;color:#fff}.side-note{margin-top:auto;border-top:1px solid #2b6d75;padding:18px 10px 0;color:#8ebabe;font-size:12px;line-height:1.8}.content{padding:28px 42px 48px;max-width:1400px;width:100%}.topbar{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:28px}.eyebrow{font-size:12px;letter-spacing:1.5px;color:var(--teal);font-weight:700}.topbar h1{font-size:29px;margin:7px 0 0;letter-spacing:-.5px}.pill{padding:8px 13px;border-radius:999px;font-size:13px;font-weight:700;background:#fff1e6;color:#a95118;border:1px solid #f2c59e}.pill.ready{background:var(--mint);color:#18736d;border-color:#b9e1db}.page{display:none}.page.active{display:block}.hero{background:linear-gradient(135deg,#0b6570,#0b4d59);color:#fff;border-radius:18px;padding:30px 32px;display:flex;justify-content:space-between;gap:28px;align-items:center;box-shadow:0 12px 30px #0b657020}.hero h2{font-size:25px;margin:0 0 10px}.hero p{color:#c9e7e8;line-height:1.7;margin:0;max-width:620px}.hero-tag{background:#ffffff18;border:1px solid #ffffff35;border-radius:12px;padding:15px 18px;min-width:170px}.hero-tag b{font-size:22px;display:block}.hero-tag span{font-size:12px;color:#c9e7e8}.panel{background:#fff;border:1px solid var(--line);border-radius:16px;padding:25px;margin-top:18px;box-shadow:0 7px 24px #19445209}.panel-head{display:flex;justify-content:space-between;gap:15px;align-items:flex-start;margin-bottom:18px}.panel h3{margin:0;font-size:18px}.muted{color:var(--muted);font-size:13px;line-height:1.7}.upload-zone{border:1px dashed #9bc2c6;border-radius:14px;background:#f8fcfc;padding:30px;text-align:center}.upload-zone input{display:block;margin:0 auto 13px;max-width:100%}.file-name{color:var(--muted);font-size:13px;min-height:20px}.btn{border:0;border-radius:9px;padding:11px 20px;font:inherit;font-weight:700;cursor:pointer}.btn-primary{background:var(--orange);color:#fff;box-shadow:0 5px 12px #e8894230}.btn-secondary{background:#edf4f5;color:var(--teal)}.btn:disabled{opacity:.55;cursor:not-allowed}.summary-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-top:18px}.metric{border:1px solid var(--line);border-radius:12px;padding:15px}.metric small{color:var(--muted);display:block}.metric b{font-size:25px;display:block;margin-top:5px}.metric.p0{border-color:#f3c0b8;background:#fff8f7}.metric.p1{border-color:#f0d29d;background:#fffaf1}.metric.p2{border-color:#bdd7ed;background:#f7fbff}.result-box{background:#f6f9fa;border-radius:11px;padding:15px;white-space:pre-wrap;min-height:75px;color:#344d59;line-height:1.7}.links a{display:inline-block;margin:10px 14px 0 0;color:var(--teal);font-weight:600;text-decoration:none}.links a:hover{text-decoration:underline}.skill-editor{width:100%;min-height:440px;border:1px solid var(--line);border-radius:11px;padding:16px;font:14px/1.8 Consolas,Microsoft YaHei,sans-serif;color:var(--ink);resize:vertical}.log-box{background:#102d35;color:#b9e1df;border-radius:12px;padding:18px;min-height:390px;white-space:pre-wrap;overflow:auto;font:12px/1.7 Consolas,monospace}.status{padding:11px 13px;border-radius:9px;background:#f0f7f7;color:#477078;font-size:13px;margin-top:12px}.status.error{background:#fff1ef;color:#a2463e}
@media(max-width:900px){.shell{grid-template-columns:1fr}.sidebar{padding:13px 14px;display:block}.brand{padding:4px 8px 12px}.nav-label,.side-note{display:none}.nav-link{width:auto;display:inline-flex;padding:8px 11px}.content{padding:22px 17px 40px}.hero{display:block}.hero-tag{margin-top:20px}.summary-grid{grid-template-columns:repeat(2,1fr)}}@media(max-width:520px){.topbar{display:block}.pill{display:inline-block;margin-top:14px}.summary-grid{grid-template-columns:1fr}}</style>
<style>
:root{--ink:#172033;--muted:#64748b;--line:#e2e8f0;--brand:#4f46e5;--brand-dark:#3730a3;--brand-soft:#eef2ff;--canvas:#f6f7fb;--surface:#fff;--success:#0f766e;--warning:#b45309;--danger:#b42318;--shadow:0 10px 30px rgba(15,23,42,.06)}
body{background:var(--canvas);color:var(--ink);font-family:Inter,Segoe UI,Microsoft YaHei,sans-serif}.shell{grid-template-columns:248px 1fr}.sidebar{background:#fff;color:var(--ink);border-right:1px solid var(--line);padding:24px 15px}.brand{padding:4px 10px 34px}.brand-mark{display:grid!important;place-items:center!important;width:38px;height:38px;border-radius:10px;background:#4f46e5!important;color:#fff!important;box-shadow:none}.brand strong{color:var(--ink);font-size:15px}.brand small{color:var(--muted)}.nav-label{color:#94a3b8;font-weight:700;letter-spacing:1.2px}.nav-link{color:#64748b;padding:12px 13px;transition:background .18s,color .18s}.nav-link span{color:#94a3b8;font-weight:700}.nav-link:hover{background:#f8fafc;color:var(--ink)}.nav-link:focus{outline:none}.nav-link:focus-visible{outline:2px solid #c7d2fe;outline-offset:2px}.nav-link.active{background:var(--brand-soft);color:var(--brand-dark)}.nav-link.active span{color:var(--brand)}.side-note{border-top:1px solid var(--line);color:#94a3b8}.content{padding:34px 48px 56px;max-width:1480px}.topbar{margin-bottom:24px}.eyebrow{color:var(--brand);letter-spacing:1.4px}.topbar h1{font-size:28px;letter-spacing:-.6px}.pill{background:#fff7ed;color:var(--warning);border:1px solid #fed7aa}.pill.ready{background:#ecfdf5;color:var(--success);border-color:#bbf7d0}.hero{background:#fff;color:var(--ink);border:1px solid var(--line);box-shadow:var(--shadow);padding:31px 34px}.hero h2{font-size:24px;letter-spacing:-.35px}.hero p{color:var(--muted)}.hero-tag{background:var(--brand-soft);border:1px solid #c7d2fe;color:var(--brand-dark)}.hero-tag span{color:#6366f1}.panel{border-color:var(--line);box-shadow:var(--shadow);padding:24px}.panel h3{letter-spacing:-.15px}.muted{color:var(--muted)}.upload-zone{border-color:#c7d2fe;background:#fafaff;padding:34px}.btn{transition:transform .15s,box-shadow .15s,background .15s}.btn:hover{transform:translateY(-1px)}.btn:disabled{opacity:.7;background:#c7d2fe;color:#4338ca;box-shadow:none}.btn-primary{background:var(--brand);box-shadow:0 6px 14px rgba(79,70,229,.22)}.btn-primary:hover{background:var(--brand-dark)}.btn-secondary{background:#f1f5f9;color:#334155}.summary-grid{gap:16px;margin-top:20px}.metric{background:#fff;border-color:var(--line);box-shadow:0 5px 16px rgba(15,23,42,.035)}.metric.p0{border-color:#fecaca;border-left:4px solid #dc2626;background:#fffafa}.metric.p1{border-color:#fed7aa;border-left:4px solid #d97706;background:#fffdf7}.metric.p2{border-color:#bfdbfe;border-left:4px solid #2563eb;background:#fbfdff}.metric small{color:var(--muted)}.metric b{color:var(--ink)}.result-box{background:#f8fafc;border:1px solid #eef2f7;color:#334155}.links a{color:var(--brand)}.skill-editor{border-color:var(--line);background:#fbfcfe}.log-box{background:#172033;color:#dbeafe;box-shadow:inset 0 0 0 1px rgba(255,255,255,.05)}.status{background:#f0fdf4;color:#166534;border:1px solid #dcfce7}.status.error{background:#fff1f2;color:#9f1239;border-color:#fecdd3}.admin-tag{background:#fef3c7;color:#92400e}
#configPage .panel>div[style*="display:grid"]{grid-template-columns:repeat(2,minmax(0,1fr))!important;gap:16px!important}#configPage input,#configPage select{border-color:var(--line)!important;background:#fbfcfe!important}#configPage input:focus,#configPage select:focus{outline:3px solid #e0e7ff!important;border-color:var(--brand)!important}#configPage button{margin-right:8px}
@media(max-width:900px){.shell{grid-template-columns:1fr}.sidebar{border-right:0;border-bottom:1px solid var(--line)}.content{padding:24px 18px 42px}.hero{padding:26px}.summary-grid{grid-template-columns:repeat(2,1fr)}#configPage .panel>div[style*="display:grid"]{grid-template-columns:1fr!important}}
@media(max-width:520px){.summary-grid{grid-template-columns:1fr}.hero{padding:22px}.panel{padding:18px}}
<style>
.topbar-copy{min-width:0}.topbar-subtitle{margin-top:6px;color:var(--muted);font-size:13px;line-height:1.6}.upload-zone{position:relative;transition:border-color .18s,background .18s,box-shadow .18s,transform .18s}.upload-zone.dragging{border-color:var(--brand);background:#f3f4ff;box-shadow:0 0 0 4px #e0e7ff;transform:translateY(-1px)}.upload-drop-hint{margin-top:8px;color:#94a3b8;font-size:12px}.skill-meta{display:flex;align-items:center;gap:10px;min-height:24px;margin:13px 0 7px;color:var(--muted);font-size:12px}.editor-toolbar{display:flex;justify-content:space-between;align-items:center;margin:7px 2px 8px;color:#94a3b8;font-size:12px}.skill-badge{display:inline-flex;align-items:center;padding:4px 8px;border-radius:999px;background:#f1f5f9;color:#64748b}.skill-badge.dirty{background:#fff7ed;color:#b45309}.skill-editor{min-height:500px;background:#fbfcfe;line-height:1.75;tab-size:4;transition:border-color .18s,box-shadow .18s}.skill-editor:focus{outline:none;border-color:#a5b4fc;box-shadow:0 0 0 4px #e0e7ff}.skill-upload-name{max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.status{transition:background .18s,color .18s,border-color .18s}.status.pending{background:#eff6ff;color:#1d4ed8;border-color:#dbeafe}.status.success{background:#f0fdf4;color:#166534;border-color:#dcfce7}.btn.is-loading{position:relative;pointer-events:none}.btn.is-loading:before{content:"";display:inline-block;width:12px;height:12px;margin-right:7px;border:2px solid currentColor;border-right-color:transparent;border-radius:50%;vertical-align:-2px;animation:spin .7s linear infinite}@keyframes spin{to{transform:rotate(360deg)}}
@media(max-width:700px){.topbar-subtitle{max-width:100%}.skill-editor{min-height:420px}.skill-upload-name{max-width:180px}}
</style>
<style>
.config-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px}.config-field{display:flex;flex-direction:column;gap:7px}.config-field label{font-size:13px;font-weight:700}.config-field input,.config-field select{width:100%;border:1px solid var(--line);border-radius:9px;padding:11px 12px;font:inherit;color:var(--ink);background:#fbfcfe;transition:border-color .18s,box-shadow .18s}.config-field input:focus,.config-field select:focus{outline:none;border-color:var(--brand);box-shadow:0 0 0 4px #e0e7ff}.field-help{color:#94a3b8;font-size:12px;line-height:1.5}.config-actions{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-top:16px}.config-key-option{display:flex;align-items:center;gap:8px;margin:16px 0 0;color:var(--muted);font-size:13px}.config-key-option input{accent-color:var(--brand)}.log-toolbar{display:flex;align-items:center;gap:12px}.log-meta{color:#94a3b8;font-size:12px}.log-box{min-height:430px;max-height:58vh}.empty-state{display:grid;place-items:center;min-height:220px;border:1px dashed #cbd5e1;border-radius:12px;background:#fbfcfe;color:#94a3b8;font-size:13px}.status:empty{display:none}@media(max-width:700px){.config-grid{grid-template-columns:1fr}.config-actions .btn{flex:1}.log-toolbar{align-items:flex-start;flex-direction:column;gap:7px}.log-box{max-height:none}}
<style>
html{font-size:16px}body{font-family:Inter,"Segoe UI","Microsoft YaHei",Arial,sans-serif;font-size:14px;line-height:1.55;letter-spacing:0}.content{padding:40px 56px 64px}.topbar{min-height:72px;margin-bottom:30px;align-items:center}.topbar h1{font-size:32px;font-weight:750;line-height:1.2;color:#111827;letter-spacing:-.8px}.topbar-subtitle{font-size:14px}.topbar .pill{min-height:40px;display:inline-flex;align-items:center;justify-content:center;padding:9px 16px}.brand strong{font-size:16px;font-weight:750}.nav-link{min-height:44px;font-size:14px;font-weight:600;letter-spacing:.1px}.nav-link span{font-size:11px;font-weight:800}.hero{min-height:164px;padding:34px 40px}.hero h2{font-size:26px;font-weight:750;line-height:1.3}.hero p{font-size:15px;line-height:1.8}.hero-tag{min-width:190px;padding:18px 20px}.hero-tag b{font-size:25px;font-weight:750}.panel{padding:28px;border-radius:18px}.panel h3{font-size:20px;font-weight:750;line-height:1.35}.panel-head{margin-bottom:20px}.muted{font-size:14px;line-height:1.7}.btn{min-height:42px;padding:10px 19px;border:1px solid transparent;border-radius:10px;display:inline-flex;align-items:center;justify-content:center;gap:7px;font-size:14px;font-weight:700;line-height:1;letter-spacing:.1px;transition:transform .16s,box-shadow .16s,background .16s,border-color .16s}.btn:focus-visible{outline:3px solid #c7d2fe;outline-offset:2px}.btn-secondary{border-color:#dbe2ee;background:#fff;color:#3730a3;box-shadow:0 3px 8px rgba(15,23,42,.04)}.btn-secondary:hover{background:#f8faff;border-color:#a5b4fc;color:#312e81}.btn-primary{background:#4f46e5;color:#fff;box-shadow:0 8px 18px rgba(79,70,229,.22)}.btn-primary:hover{background:#4338ca;box-shadow:0 10px 22px rgba(79,70,229,.28)}.btn:disabled{transform:none}.upload-zone{min-height:310px;padding:30px 24px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:0;border:1px dashed #b8c5f5;background:linear-gradient(180deg,#fbfbff,#f7f8ff);transition:border-color .18s,background .18s,box-shadow .18s,transform .18s}.upload-zone.dragging{border-color:#4f46e5;background:#eef2ff;box-shadow:0 0 0 5px #e0e7ff;transform:translateY(-1px)}.upload-icon{width:52px;height:52px;margin-bottom:13px;border-radius:16px;background:#e0e7ff;color:#4f46e5;display:grid;place-items:center;font-size:29px;font-weight:800;line-height:1}.upload-title{font-size:17px;font-weight:750;color:#1e293b}.upload-desc{margin-top:4px;color:#94a3b8;font-size:13px}.upload-select{margin-top:17px;cursor:pointer}.file-input{position:absolute!important;width:1px!important;height:1px!important;padding:0!important;margin:-1px!important;overflow:hidden!important;clip:rect(0,0,0,0)!important;white-space:nowrap!important;border:0!important;opacity:0!important}.file-name{min-height:22px;margin-top:12px;color:#475569;font-size:13px;font-weight:600}.upload-drop-hint{margin-top:5px;color:#94a3b8;font-size:12px}.upload-actions{margin-top:22px}.summary-grid{gap:18px;margin-top:22px}.metric{min-height:112px;padding:18px 19px;border-radius:14px;display:flex;flex-direction:column;justify-content:center}.metric small{font-size:13px;font-weight:600}.metric b{font-size:29px;font-weight:750;margin-top:8px}.skill-panel{max-width:1180px}.skill-editor{min-height:540px;padding:19px;font-size:13.5px;line-height:1.8}.skill-meta{padding:8px 11px;border:1px solid #e2e8f0;border-radius:9px;background:#f8fafc}.editor-toolbar{margin:9px 2px}.config-grid{max-width:1120px;align-items:start}.config-field input,.config-field select{min-height:44px;font-size:14px}.config-actions{padding-top:17px;border-top:1px solid #eef2f7}.log-box{font-size:12px;line-height:1.75}.log-toolbar .btn{min-height:38px;padding:9px 15px}
@media(max-width:900px){.content{padding:28px 22px 48px}.topbar h1{font-size:28px}.hero{padding:28px 26px}.upload-zone{min-height:280px}.config-grid{max-width:none}}
@media(max-width:520px){.content{padding:22px 15px 38px}.topbar h1{font-size:25px}.topbar .pill{margin-top:12px}.hero{min-height:0;padding:24px 20px}.hero h2{font-size:22px}.hero-tag{min-width:0;margin-top:20px}.panel{padding:20px}.upload-zone{min-height:270px;padding:24px 15px}.summary-grid{gap:12px}.metric{min-height:96px;padding:15px}.metric b{font-size:25px}.config-actions .btn{flex:1}}
</style></head><body><div class="shell">
 <aside class="sidebar"><div class="brand"><div class="brand-mark">K</div><div><strong>KKS 审核工作台</strong><small>编码质量控制中心</small></div></div><div class="nav-label">工作台</div><button class="nav-link active" data-panel="auditPage"><span class="nav-icon">⌂</span>开始审核</button><button class="nav-link" data-panel="skillPage"><span class="nav-icon">▧</span>Skill 规则</button><button class="nav-link" data-panel="configPage"><span class="nav-icon">⚙</span>AI 配置</button><button class="nav-link" data-panel="logsPage"><span class="nav-icon">◷</span>运行日志</button><div class="side-ai-status"><span id="aiBadge" class="side-ai-title">正在读取 AI 状态</span><small id="sideAiHint">已连接 AI 审核引擎</small></div><div class="side-footer">当前版本　v1.0.0<br><span>© 2024 KKS QUALITY CONTROL</span></div></aside>
 <main class="content"><header class="topbar"><div class="topbar-copy"><div id="pageEyebrow" class="eyebrow">KKS QUALITY CONTROL</div><h1 id="pageTitle">编码审核工作台</h1><div id="pageSubtitle" class="topbar-subtitle">上传 Excel 文件，AI 智能识别并生成审核结果</div></div><div class="top-actions"><button class="btn btn-header page-switch" data-panel="logsPage">◷ 运行日志</button><button class="btn btn-header page-switch" data-panel="configPage">⚙ 系统配置</button></div></header>
<section id="auditPage" class="page active">
  <div class="panel upload-panel dashboard-upload"><form id="form" enctype="multipart/form-data"><div class="upload-strip"><div id="uploadZone" class="upload-zone dashboard-drop"><div class="upload-file-icon">▤</div><div class="upload-copy"><div class="upload-title">拖拽 Excel 文件到此处</div><div class="upload-desc">或 <label for="file" class="upload-link">点击选择文件</label></div></div><input id="file" class="file-input" name="file" type="file" accept=".xlsx" required></div><div class="file-details"><div id="fileName" class="file-name">尚未选择文件</div><div class="file-meta"><span id="fileSize">—</span><span>·</span><span id="fileFormat">支持 .xlsx</span><span>·</span><span>最大文件：50MB</span></div><div id="uploadState" class="upload-state">等待上传</div></div><div class="upload-actions"><button id="submit" class="btn btn-primary upload-submit" type="submit">▶ 开始审核</button><span>AI 智能识别并生成结果</span></div></div></form></div>
 <div class="summary-grid dashboard-metrics"><div class="metric dashboard-metric quality"><span class="metric-icon">◉</span><div><small>有效 KKS 数量</small><b id="effectiveKks">—</b><span class="metric-note">本次有效编码</span></div></div><div class="metric dashboard-metric quality"><span class="metric-icon">▦</span><div><small>设备级编码</small><b id="deviceLevelCodes">—</b><span class="metric-note">严格 12 位</span></div></div><div class="metric dashboard-metric quality"><span class="metric-icon">≡</span><div><small>重复编码</small><b id="duplicateCodes">—</b><span class="metric-note">完整 KKS 码组</span></div></div><div class="metric dashboard-metric quality"><span class="metric-icon">↳</span><div><small>父级错误</small><b id="parentErrors">—</b><span class="metric-note">孤儿或层级关系</span></div></div><div class="metric dashboard-metric p0"><span class="metric-icon">!</span><div><small>高风险问题（P0）</small><b id="qualityP0">—</b><span class="metric-note">必须整改</span></div></div><div class="metric dashboard-metric p1"><span class="metric-icon">△</span><div><small>待整改问题（P1）</small><b id="qualityP1">—</b><span class="metric-note">整改或人工确认</span></div></div><div class="metric dashboard-metric p2"><span class="metric-icon">i</span><div><small>提示问题（P2）</small><b id="qualityP2">—</b><span class="metric-note">需人工确认或治理</span></div></div></div>
 <div id="resultPanel" class="panel result-panel dashboard-results"><div class="panel-head result-panel-head"><div><h3>审核结果</h3><div class="muted">问题分布、问题预览和下载结果</div></div><span id="resultState" class="result-state idle">等待审核</span></div><div id="progressPanel" class="progress-panel" hidden><div class="progress-topline"><span id="progressText">准备审核</span><b id="progressPercent">0%</b></div><div class="progress-track"><div id="progressBar"></div></div><div id="progressHint" class="progress-hint">正在准备文件……</div></div><div id="result" class="result-box result-box-empty"><div class="result-message"><span class="result-message-icon" aria-hidden="true">◎</span><div><strong id="resultTitle">等待上传文件</strong><span id="resultSubtitle">审核完成后，问题分布和下载入口会显示在这里。</span></div></div></div><div id="resultDashboard" hidden><div class="severity-card"><div class="section-row"><h3>问题级别分布</h3><span class="info-dot">i</span></div><div class="severity-layout"><div class="severity-main"><div class="severity-bar"><span id="barP0"></span><span id="barP1"></span><span id="barP2"></span></div><div class="severity-legend"><span><i class="legend-dot p0"></i>P0 严重 <b id="legendP0">0</b></span><span><i class="legend-dot p1"></i>P1 重要 <b id="legendP1">0</b></span><span><i class="legend-dot p2"></i>P2 一般 <b id="legendP2">0</b></span></div></div><div id="issueDonut" class="issue-donut"><div><strong id="donutTotal">—</strong><small>问题总数</small></div></div><div class="valid-card"><span class="valid-icon">✓</span><div><small>有效编码</small><b id="resultDataRows">—</b><span id="validRate">—</span></div></div><div class="priority-card"><div class="priority-row"><span><i class="legend-dot p0"></i>P0 严重</span><b id="priorityP0Table">0</b></div><div class="priority-row"><span><i class="legend-dot p1"></i>P1 重要</span><b id="priorityP1Table">0</b></div><div class="priority-row"><span><i class="legend-dot p2"></i>P2 一般</span><b id="priorityP2Table">0</b></div><div class="priority-row total"><span>总计</span><b id="priorityTotal">0</b></div></div></div></div><div class="result-lower"><div class="issue-preview-card"><div class="section-row"><h3>问题预览（按规则分组）</h3><span class="text-action">点击节点展开/收起</span></div><div class="issue-table-wrap"><div id="issuePreview" class="issue-tree-box"><div class="empty-preview">审核完成后显示问题预览</div></div></div></div><div class="download-panel"><h3>下载结果</h3><div id="links" class="links result-downloads" hidden></div></div></div></div></div></section>
<section id="skillPage" class="page"><div class="panel"><div class="panel-head"><div><h3>Skill 规则管理</h3><div class="muted">管理员可以查看、修改或上传 Skill 文件。Markdown 会影响 AI 参考上下文；audit_template.py 是可执行规则代码，修改后必须重启服务。</div></div><span style="padding:6px 10px;border-radius:999px;background:#fff1e6;color:#a95118;font-size:12px;font-weight:700;white-space:nowrap">管理员功能</span></div><div style="display:grid;grid-template-columns:minmax(230px,1fr) 2fr;gap:12px;align-items:center;margin:15px 0 10px"><label for="skillFile" style="font-size:13px;font-weight:700">当前文件</label><select id="skillFile" style="width:100%;border:1px solid var(--line);border-radius:9px;padding:11px 12px;font:inherit;color:var(--ink);background:#fff"></select></div><div id="skillMeta" class="skill-meta">选择文件后显示说明</div><div class="editor-toolbar"><span id="skillDirty" class="skill-badge">未修改</span><span id="skillCharCount">0 字符</span></div><textarea id="skill" class="skill-editor" spellcheck="false"></textarea><div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-top:12px"><button id="saveSkill" class="btn btn-primary" type="button">保存当前文件</button><input id="skillUpload" type="file" accept=".md,.py,.zip" style="max-width:260px"><span id="skillUploadName" class="skill-upload-name muted">未选择文件</span><button id="uploadSkill" class="btn btn-secondary" type="button">上传 Skill 文件 / 包</button></div><div class="muted" style="margin-top:10px">支持单个 `.md` / `.py` 文件，也支持完整 `.zip` 包；系统只接收 SKILL.md、references/*.md 和 scripts/audit_template.py，并自动保留旧版本备份。</div><div id="skillStatus" class="status">正在加载 Skill……</div></div></section>
<section id="configPage" class="page"><div class="panel"><div class="panel-head"><div><h3>AI 管理员配置</h3><div class="muted">修改接口地址、模型和 AI 复核开关。API 密钥只显示掩码，不会写入日志、报告或审核结果。</div></div><span style="padding:6px 10px;border-radius:999px;background:#fff1e6;color:#a95118;font-size:12px;font-weight:700;white-space:nowrap">管理员</span></div><div class="config-grid"><div class="config-field"><label for="baseUrl">AI 接口地址</label><input id="baseUrl" type="url" placeholder="https://example.com/v1"><span class="field-help">兼容 OpenAI 格式的 /v1 接口地址。</span></div><div class="config-field"><label for="model">模型名称</label><input id="model" type="text" placeholder="例如 deepseek-v4-flash"><span class="field-help">填写服务商实际提供的模型标识。</span></div><div class="config-field"><label for="apiKey">API 密钥</label><input id="apiKey" type="password" autocomplete="new-password" placeholder="留空表示保持现有密钥"><span class="field-help">保存后只显示掩码，留空不会覆盖原密钥。</span></div><div class="config-field"><label for="enabled">AI 复核开关</label><select id="enabled"><option value="true">启用</option><option value="false">停用</option></select><span class="field-help">停用时仍执行规则审核，但不调用模型。</span></div></div><label class="config-key-option"><input id="clearKey" type="checkbox"> 清空当前已保存的 API 密钥</label><div class="config-actions"><button id="saveConfig" class="btn btn-primary" type="button">保存配置</button><button id="testAi" class="btn btn-secondary" type="button">测试 AI 连接</button><span id="configDirty" class="skill-badge">未修改</span></div><div id="configStatus" class="status">正在读取配置……</div></div></section>
<section id="logsPage" class="page"><div class="panel"><div class="panel-head"><div><h3>运行日志</h3><div class="muted">记录服务启动、文件审核、AI 状态和报告生成过程；不会记录 API 密钥。</div></div><div class="log-toolbar"><span id="logMeta" class="log-meta">尚未刷新</span><button id="refreshLogs" class="btn btn-secondary" type="button">刷新日志</button></div></div><pre id="logs" class="log-box">正在读取日志……</pre></div></section>
</main></div><script>
const form=document.getElementById('form'),submit=document.getElementById('submit'),result=document.getElementById('result'),links=document.getElementById('links');
async function jsonFetch(url,options={}){const r=await fetch(url,options);const d=await r.json();if(!r.ok)throw new Error(d.error||'请求失败');return d}
 const pageMeta={auditPage:['KKS QUALITY CONTROL','编码审核工作台','上传 Excel 文件，AI 智能识别并生成审核结果'],skillPage:['RULE MANAGEMENT','Skill 规则管理','维护 AI 参考文档和可执行审核规则'],configPage:['AI ADMINISTRATION','AI 配置','管理模型接口、密钥和审核参数'],logsPage:['SYSTEM ACTIVITY','运行日志','查看审核任务、配置变更和服务运行记录']};
function setPageMeta(page){const meta=pageMeta[page]||pageMeta.auditPage;document.getElementById('pageEyebrow').textContent=meta[0];document.getElementById('pageTitle').textContent=meta[1];document.getElementById('pageSubtitle').textContent=meta[2]}
 function activatePage(page){document.querySelectorAll('.nav-link').forEach(x=>x.classList.toggle('active',x.dataset.panel===page));document.querySelectorAll('.page').forEach(x=>x.classList.toggle('active',x.id===page));setPageMeta(page);if(page==='logsPage')loadLogs();if(page==='configPage')loadConfig()}
 document.querySelectorAll('.nav-link,.page-switch').forEach(btn=>btn.addEventListener('click',()=>activatePage(btn.dataset.panel)));setPageMeta('auditPage');
const fileInput=document.getElementById('file'),uploadZone=document.getElementById('uploadZone');
 function formatBytes(value){if(!value)return '—';if(value<1024*1024)return (value/1024).toFixed(1)+' KB';return (value/1024/1024).toFixed(1)+' MB'}
 function showSelectedFile(file){setText('fileName',file?file.name:'尚未选择文件');setText('fileSize',file?formatBytes(file.size):'—');setText('fileFormat',file?'.'+file.name.split('.').pop().toLowerCase():'支持 .xlsx');setText('uploadState',file?'已选择文件':'等待上传')}
fileInput.addEventListener('change',e=>showSelectedFile(e.target.files[0]));
['dragenter','dragover'].forEach(type=>uploadZone.addEventListener(type,e=>{e.preventDefault();uploadZone.classList.add('dragging')}));
['dragleave','drop'].forEach(type=>uploadZone.addEventListener(type,e=>{e.preventDefault();uploadZone.classList.remove('dragging')}));
 uploadZone.addEventListener('drop',e=>{const file=e.dataTransfer.files[0];if(!file)return;if(!/\\.xlsx$/i.test(file.name)){showSelectedFile(null);setResultMessage('文件格式不支持','请选择 .xlsx 格式的 Excel 文件。','failed');setResultState('文件错误','failed');return}const transfer=new DataTransfer();transfer.items.add(file);fileInput.files=transfer.files;showSelectedFile(file)});
  function setAiBadge(ready){const badge=document.getElementById('aiBadge');badge.textContent=ready?'AI 已启用':'AI 尚未配置';badge.className='side-ai-title'+(ready?' ready':'');setText('sideAiHint',ready?'已连接 AI 审核引擎':'请到 AI 配置页完成设置')}
 
 function setText(id,value){const el=document.getElementById(id);if(el)el.textContent=String(value)}
 function setResultState(label,kind){const state=document.getElementById('resultState');state.textContent=label;state.className='result-state '+kind;const overview=document.getElementById('overviewState');if(overview){overview.textContent=label;overview.className='overview-state '+kind}}
 function setResultMessage(title,subtitle,kind){result.className='result-box result-box-'+kind;setText('resultTitle',title);setText('resultSubtitle',subtitle)}
 function setProgress(value,text,hint){document.getElementById('progressBar').style.width=value+'%';setText('progressPercent',value+'%');setText('progressText',text);setText('progressHint',hint);if(value>0&&value<100)setResultState('审核中','running')}
 function startProgress(){document.getElementById('progressPanel').hidden=false;document.getElementById('resultDashboard').hidden=true;result.hidden=false;links.hidden=true;setText('uploadState','正在审核');setResultMessage('正在创建审核任务','文件已接收，正在执行规则检查和 AI 复核。','running');setResultState('排队中','running');setProgress(5,'审核任务已启动','已接收文件，正在等待后台任务开始')}
 function finishProgress(ok){setProgress(ok?100:0,ok?'审核完成':'审核未完成',ok?'报告已生成，可以下载结果':'请根据错误信息修正后重试');setResultState(ok?'已完成':'未完成',ok?'done':'failed')}
 function renderQualityMetrics(metrics){const q=metrics||{};setText('effectiveKks',q.effective_kks===undefined?'—':q.effective_kks);setText('deviceLevelCodes',q.device_level_codes===undefined?'—':q.device_level_codes);setText('duplicateCodes',q.duplicate_codes===undefined?'—':q.duplicate_codes);setText('parentErrors',q.parent_errors===undefined?'—':q.parent_errors);setText('qualityP0',q.p0===undefined?'—':q.p0);setText('qualityP1',q.p1===undefined?'—':q.p1);setText('qualityP2',q.p2===undefined?'—':q.p2)}
 function renderAuditSummary(d){if(d.quality_metrics)renderQualityMetrics(d.quality_metrics)}
 function renderDistribution(counts,total){const p0=Number(counts.P0||0),p1=Number(counts.P1||0),p2=Number(counts.P2||0),sum=p0+p1+p2,totalValue=Number(total)||sum||1,a=p0/totalValue*100,b=p1/totalValue*100,c=p2/totalValue*100;setText('legendP0',p0);setText('legendP1',p1);setText('legendP2',p2);setText('priorityP0Table',p0);setText('priorityP1Table',p1);setText('priorityP2Table',p2);setText('priorityTotal',sum);setText('donutTotal',total||sum);const bars=[['barP0',a],['barP1',b],['barP2',c]];bars.forEach(([id,value])=>{const el=document.getElementById(id);if(el)el.style.width=value.toFixed(2)+'%'});const donut=document.getElementById('issueDonut');if(donut)donut.style.background='conic-gradient(#ef4444 0 '+a+'%,#f59e0b '+a+'% '+(a+b)+'%,#2563eb '+(a+b)+'% 100%)'}
function escHtml(v){return String(v==null?'\u2014':v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;')}
function _nodeBadge(n){let b='';if(Number(n.p0))b+='<span class="tb tb-p0">'+n.p0+'</span>';if(Number(n.p1))b+='<span class="tb tb-p1">'+n.p1+'</span>';if(Number(n.p2))b+='<span class="tb tb-p2">'+n.p2+'</span>';return b}
function _issueRowsHtml(rows){if(!Array.isArray(rows)||!rows.length)return '';return '<div class="it-issues">'+rows.map(function(r){return '<div class="it-issue"><span class="it-row">行号 '+escHtml(r.excel_row||'—')+'</span><span class="it-msg">'+escHtml(r.message||'待核问题')+'</span><div class="it-suggest">建议：'+escHtml(r.suggestion||'')+'</div></div>'}).join('')+'</div>'}
function _treeNodeHtml(n,open){const isCode=n.kind==='code',act=String(n.action||'待确认'),actCls=act==='需整改'?'act-fix':(act==='需确认'?'act-hm':'act-ck');let summary=isCode?'<span class="it-name">'+escHtml(n.title||'未挂接编码的问题')+'</span>'+(n.sub?'<span class="it-title">'+escHtml(n.sub)+'</span>':''):'<span class="priority p'+String(n.priority||'P2').slice(-1)+'">'+escHtml(n.priority||'P2')+'</span><span class="it-name tg-rid">'+escHtml(n.title||'')+'</span><span class="tg-def">'+escHtml(n.sub||'')+'</span><span class="tg-act '+actCls+'">'+escHtml(act)+'</span>';const kids=(n.children||[]).map(function(c){return _treeNodeHtml(c,'')}).join('');const src=n.source?'<div class="tg-src">规则依据：'+escHtml(n.source)+'</div>':'';let inner=src;if(isCode){inner+=_issueRowsHtml(n.issues||[])}else if(kids){inner+='<div class="it-kids">'+kids+'</div>'}return '<details class="it-node '+(isCode?'it-code':'tg-rule')+'"'+(open?' open':'')+'><summary>'+summary+'<span class="it-count">'+escHtml(n.note||'')+'</span>'+(isCode?_nodeBadge(n):'')+'</summary>'+inner+'</details>'}
function renderIssueTree(container,nodes,openDepth){if(!container)return;if(!Array.isArray(nodes)||!nodes.length){container.innerHTML='<div class="empty-preview">本次没有可展示的问题</div>';return}container.innerHTML=nodes.map(function(n){return _treeNodeHtml(n,Number(n.level||0)<Number(openDepth||1))}).join('')}
function renderIssuePreview(items){const body=document.getElementById('issuePreview');if(!body)return;renderIssueTree(body,items,1)}

function renderCompletedResult(d){const counts=d.priority_counts||{};document.getElementById('result').hidden=true;document.getElementById('progressPanel').hidden=true;document.getElementById('resultDashboard').hidden=false;setText('uploadState','审核完成');renderQualityMetrics(d.quality_metrics||{});renderDistribution(counts,d.issue_count);renderIssuePreview(d.issue_preview||[]);links.innerHTML='';for(const [label,url] of Object.entries(d.files||{})){const a=document.createElement('a');a.className='download-card';a.href=url;a.download='';const icon=document.createElement('span');icon.className='download-icon';icon.textContent=label.includes('HTML')?'</>':'XLSX';const copy=document.createElement('span');copy.className='download-copy';const title=document.createElement('strong');title.textContent=label;const desc=document.createElement('span');desc.textContent=label.includes('HTML')?'正式质量审核报告，包含结论和整改建议':'精简问题清单，支持筛选和排序';copy.append(title,desc);const action=document.createElement('span');action.className='download-action';action.textContent='下载';a.append(icon,copy,action);links.appendChild(a)}links.hidden=!Object.keys(d.files||{}).length}
 async function watchAudit(runId){while(true){const d=await jsonFetch('/api/audits/'+encodeURIComponent(runId)+'/status'),ai=d.ai_review||{};if(d.percent!==undefined)setProgress(d.percent,d.message||'正在审核',d.hint||'');renderAuditSummary(d);if(d.status==='completed'){renderCompletedResult(d);finishProgress(true);return}if(d.status==='error')throw new Error(d.hint||d.error||'审核失败');setResultMessage(d.message||'正在审核，请稍候……',d.hint||'后台正在处理，请保持页面打开。','running');await new Promise(resolve=>setTimeout(resolve,600))}}
async function loadState(){try{const d=await jsonFetch('/healthz');setAiBadge(Boolean(d.ai_configured))}catch(err){setAiBadge(false)}}
  form.addEventListener('submit',async e=>{e.preventDefault();submit.disabled=true;submit.classList.add('is-loading');submit.textContent='审核准备中…';startProgress();try{const created=await jsonFetch('/api/audits',{method:'POST',body:new FormData(form)});await watchAudit(created.run_id)}catch(err){setResultMessage('审核失败',err.message||'请根据日志排查问题后重试。','failed');finishProgress(false)}finally{submit.disabled=false;submit.classList.remove('is-loading');submit.textContent='开始审核'}});
let configSnapshot='';
function configValues(){return JSON.stringify({base_url:document.getElementById('baseUrl').value.trim(),model:document.getElementById('model').value.trim(),enabled:document.getElementById('enabled').value,api_key:document.getElementById('apiKey').value.trim()?'__set__':'',clear_key:document.getElementById('clearKey').checked})}
function updateConfigDirty(){const dirty=configValues()!==configSnapshot,badge=document.getElementById('configDirty'),box=document.getElementById('configStatus');badge.textContent=dirty?'未保存修改':'未修改';badge.className='skill-badge'+(dirty?' dirty':'');if(dirty&&box.className.indexOf('pending')<0){box.textContent='配置已修改，请保存后生效';box.className='status pending'}if(!dirty&&box.className.indexOf('pending')>=0){box.textContent='已恢复为当前配置';box.className='status success'}return dirty}
function setButtonBusy(button,busy,label,restore){button.disabled=busy;button.classList.toggle('is-loading',busy);button.textContent=busy?label:restore}
async function loadConfig(){try{const d=await jsonFetch('/api/config'),a=d.ai||{};document.getElementById('baseUrl').value=a.base_url||'';document.getElementById('model').value=a.model||'';document.getElementById('enabled').value=String(Boolean(a.enabled));const key=document.getElementById('apiKey');key.value='';key.placeholder=a.api_key_configured?'已配置 '+(a.api_key_hint||'掩码')+'，留空保持不变':'尚未配置，请输入 API 密钥';document.getElementById('clearKey').checked=false;configSnapshot=configValues();const box=document.getElementById('configStatus');box.textContent='已读取配置；当前密钥状态：'+(a.api_key_configured?'已配置（'+(a.api_key_hint||'已隐藏')+'）':'未配置');box.className='status success';updateConfigDirty()}catch(err){document.getElementById('configStatus').textContent='读取配置失败：'+err.message;document.getElementById('configStatus').className='status error'}}
document.querySelectorAll('#configPage input,#configPage select').forEach(field=>{field.addEventListener('input',updateConfigDirty);field.addEventListener('change',updateConfigDirty)});
document.getElementById('saveConfig').onclick=async()=>{const box=document.getElementById('configStatus'),button=document.getElementById('saveConfig');const ai={base_url:document.getElementById('baseUrl').value.trim(),model:document.getElementById('model').value.trim(),enabled:document.getElementById('enabled').value==='true'};const key=document.getElementById('apiKey').value.trim();if(key)ai.api_key=key;if(document.getElementById('clearKey').checked)ai.clear_api_key=true;setButtonBusy(button,true,'保存中…','保存配置');box.textContent='正在保存配置，请稍候……';box.className='status pending';try{const d=await jsonFetch('/api/config',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({ai})});document.getElementById('apiKey').value='';document.getElementById('clearKey').checked=false;await loadConfig();box.textContent='配置已保存，下一次审核生效；当前密钥状态：'+(d.config.ai.api_key_configured?'已配置':'未配置');box.className='status success';await loadState()}catch(err){box.textContent='保存失败：'+err.message;box.className='status error'}finally{setButtonBusy(button,false,'保存中…','保存配置')}};
document.getElementById('testAi').onclick=async()=>{const box=document.getElementById('configStatus'),button=document.getElementById('testAi');setButtonBusy(button,true,'测试中…','测试 AI 连接');box.textContent='正在测试 AI 连接，请稍候……';box.className='status pending';try{const d=await jsonFetch('/api/ai/test',{method:'POST'});box.textContent='AI 连接成功：'+(d.model||'当前模型')+'；已收到模型响应';box.className='status success';setAiBadge(true)}catch(err){box.textContent='AI 连接失败：'+err.message;box.className='status error';setAiBadge(false)}finally{setButtonBusy(button,false,'测试中…','测试 AI 连接')}};
let skillDocuments=[],skillOriginalContent='',skillCurrentPath='',skillDirty=false;
function setSkillStatus(text,type=''){const box=document.getElementById('skillStatus');box.textContent=text;box.className='status'+(type?' '+type:'')}
function updateSkillEditorState(){const editor=document.getElementById('skill'),dirty=editor.value!==skillOriginalContent;skillDirty=dirty;document.getElementById('skillDirty').textContent=dirty?'未保存修改':'未修改';document.getElementById('skillDirty').className='skill-badge'+(dirty?' dirty':'');document.getElementById('skillCharCount').textContent=editor.value.length.toLocaleString()+' 字符';const doc=skillDocuments.find(x=>x.path===skillCurrentPath);if(doc){document.getElementById('skillMeta').textContent=doc.kind==='executable_rule'?'可执行规则代码 · 修改后需重启服务':'AI 参考文档 · 修改后保存即可生效'}}
function renderSkillDocument(){const selected=document.getElementById('skillFile').value;const doc=skillDocuments.find(x=>x.path===selected);skillCurrentPath=selected;document.getElementById('skill').value=doc?doc.content:'';skillOriginalContent=doc?doc.content:'';updateSkillEditorState();setSkillStatus(doc?(doc.kind==='executable_rule'?'已加载可执行规则代码':'已加载 Skill 文档，可直接修改保存'):'请选择 Skill 文件',doc?'success':'')}
async function loadSkill(){try{const d=await jsonFetch('/api/skill');skillDocuments=d.documents||[];const select=document.getElementById('skillFile'),previous=skillCurrentPath;select.innerHTML='';for(const doc of skillDocuments){const option=document.createElement('option');option.value=doc.path;option.textContent=doc.path;select.appendChild(option)}if(previous&&skillDocuments.some(x=>x.path===previous))select.value=previous;renderSkillDocument()}catch(err){setSkillStatus('读取失败：'+err.message,'error')}}
document.getElementById('skill').addEventListener('input',()=>{updateSkillEditorState();if(skillDirty)setSkillStatus('当前文件有未保存修改，请保存后再切换或离开','pending')});
document.getElementById('skillFile').onchange=()=>{if(skillDirty&&!window.confirm('当前文件有未保存修改，切换文件将丢失这些修改。是否继续？')){document.getElementById('skillFile').value=skillCurrentPath;return}renderSkillDocument()};
document.getElementById('skillUpload').addEventListener('change',e=>{const file=e.target.files[0];document.getElementById('skillUploadName').textContent=file?file.name:'未选择文件'});
document.getElementById('saveSkill').onclick=async()=>{const path=document.getElementById('skillFile').value,content=document.getElementById('skill').value,box=document.getElementById('skillStatus'),button=document.getElementById('saveSkill');if(!path){setSkillStatus('请先选择 Skill 文件','error');return}button.disabled=true;button.classList.add('is-loading');button.textContent='保存中…';setSkillStatus('正在保存 Skill，请稍候……','pending');try{const d=await jsonFetch('/api/skill',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({path,content})});const doc=skillDocuments.find(x=>x.path===path);if(doc)doc.content=content;skillOriginalContent=content;updateSkillEditorState();setSkillStatus(d.message||'Skill 文件已保存','success')}catch(err){setSkillStatus('保存失败：'+err.message,'error')}finally{button.disabled=false;button.classList.remove('is-loading');button.textContent='保存当前文件'}};
document.getElementById('uploadSkill').onclick=async()=>{const input=document.getElementById('skillUpload'),box=document.getElementById('skillStatus'),button=document.getElementById('uploadSkill'),file=input.files[0];if(!file){setSkillStatus('请先选择 .md、.py 或 .zip 文件','error');return}button.disabled=true;button.classList.add('is-loading');button.textContent='上传中…';setSkillStatus('正在上传 Skill，请稍候……','pending');try{const formData=new FormData();formData.append('file',file);formData.append('target',document.getElementById('skillFile').value||'');const d=await jsonFetch('/api/skill/upload',{method:'POST',body:formData});input.value='';document.getElementById('skillUploadName').textContent='未选择文件';await loadSkill();setSkillStatus(d.message||'Skill 文件已上传','success')}catch(err){setSkillStatus('上传失败：'+err.message,'error')}finally{button.disabled=false;button.classList.remove('is-loading');button.textContent='上传 Skill 文件 / 包'}};
async function loadLogs(){const box=document.getElementById('logs'),button=document.getElementById('refreshLogs'),meta=document.getElementById('logMeta');setButtonBusy(button,true,'刷新中…','刷新日志');box.className='log-box';box.textContent='正在读取日志……';meta.textContent='正在刷新';try{const d=await jsonFetch('/api/logs?limit=300'),lines=Array.isArray(d.lines)?d.lines:[];box.textContent=lines.join('\\n')||'暂无运行日志';box.className=lines.length?'log-box':'log-box empty-state';meta.textContent='最近刷新：'+new Date().toLocaleTimeString()}catch(err){box.textContent='读取日志失败：'+err.message;box.className='log-box empty-state';meta.textContent='刷新失败'}finally{setButtonBusy(button,false,'刷新中…','刷新日志')}}document.getElementById('refreshLogs').onclick=loadLogs;loadState();loadConfig();loadSkill();loadLogs();
</script><style>
.biz-kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:4px 0 12px}
.biz-kpi{border:1px solid #e2e8f0;border-radius:10px;padding:10px 12px;background:#fff}
.biz-kpi b{display:block;font-size:20px;color:#2563eb}.biz-kpi.warn b{color:#d97706}.biz-kpi.bad b{color:#dc2626}.biz-kpi.ok b{color:#059669}
.biz-kpi span{display:block;margin-top:2px;color:#6b7a8d;font-size:11px}
.biz-chips{display:flex;flex-wrap:wrap;gap:7px;margin:2px 0 10px}
.biz-chip{border:1px solid #dbeafe;background:#eff6ff;color:#1d4ed8;border-radius:999px;padding:4px 11px;font-size:12px;cursor:pointer;font-weight:600}
.biz-chip b{margin-left:4px}
.biz-chip.p0{border-color:#fecdd3;background:#fff1f2;color:#be123c}.biz-chip.p1{border-color:#fde0b2;background:#fffbeb;color:#b45309}
.biz-chip.active{outline:2px solid #2563eb;background:#dbeafe}
.biz-panel{margin-top:2px}
.biz-table-wrap{overflow:auto;max-height:360px;border:1px solid #eef2f7;border-radius:10px}
.biz-table{border-collapse:collapse;width:100%;font-size:12px;background:#fff}
.biz-table th{position:sticky;top:0;background:#eef2f7;text-align:left;padding:6px 8px;border-bottom:1px solid #e2e8f0;color:#334155;white-space:nowrap;z-index:1}
.biz-table td{padding:5px 8px;border-bottom:1px solid #eef2f7;vertical-align:top;word-break:break-all}
.biz-more{color:#94a3b8;text-align:right;font-size:12px}
@media(max-width:700px){.biz-kpis{grid-template-columns:repeat(2,1fr)}}
</style><script>
function _bizEsc(s){s=(s===null||s===undefined)?'':String(s);var d=document.createElement('div');d.textContent=s;return d.innerHTML}
function renderBusinessOverview(ov){
  var host=document.getElementById('issuePreview');if(!host)return;
  var buckets=Array.isArray(ov&&ov.buckets)?ov.buckets.filter(function(b){return b.items&&b.items.length}):[];
  var head=document.querySelector('.issue-preview-card .section-row h3');if(head)head.textContent='问题预览（业务分组）';
  var act=document.querySelector('.issue-preview-card .section-row .text-action');if(act)act.textContent='点击分组查看明细，完整清单见 Excel';
  if(!ov||!buckets.length){host.innerHTML='<div class="empty-preview">本次没有需要列入问题清单的问题</div>';return;}
  var kpi='<div class="biz-kpis">'+(ov.kpis||[]).map(function(k){return '<div class="biz-kpi '+k.tone+'"><b>'+_bizEsc(k.value)+'</b><span>'+_bizEsc(k.label)+'</span></div>'}).join('')+'</div>';
  var chips=buckets.map(function(b,i){return '<button type="button" class="biz-chip '+b.level.toLowerCase()+' biz-chip-'+i+'">'+_bizEsc(b.level)+'·'+_bizEsc(b.title)+' <b>'+b.count+'</b></button>'}).join('');
  host.innerHTML=kpi+'<div class="biz-chips">'+chips+'</div><div class="biz-panel"></div>';
  function renderTable(i){
    var bkt=buckets[i],cols=bkt.columns||[],items=bkt.items||[];
    var show=items.slice(0,80);
    var th=cols.map(function(c){return '<th>'+_bizEsc(c.t)+'</th>'}).join('');
    var tr=show.map(function(r){return '<tr>'+cols.map(function(c){return '<td>'+_bizEsc(r[c.k])+'</td>'}).join('')+'</tr>'}).join('');
    if(items.length>show.length)tr+='<tr><td colspan="'+cols.length+'" class="biz-more">…共 '+items.length+' 条，完整清单见 Excel 下载</td></tr>';
    var panel=host.querySelector('.biz-panel');if(!panel)return;
    panel.innerHTML='<div class="biz-table-wrap"><table class="biz-table"><thead><tr>'+th+'</tr></thead><tbody>'+tr+'</tbody></table></div>';
    var chipEls=host.querySelectorAll('.biz-chip');for(var j=0;j<chipEls.length;j++){chipEls[j].classList.toggle('active',j===i)}
  }
  var els=host.querySelectorAll('.biz-chip');for(var j=0;j<els.length;j++){(function(i){els[i].addEventListener('click',function(){renderTable(i)})})(j)}
  renderTable(0);
}
function renderCompletedResult(d){
  var counts=d.priority_counts||{};
  document.getElementById('result').hidden=true;document.getElementById('progressPanel').hidden=true;document.getElementById('resultDashboard').hidden=false;
  setText('uploadState','审核完成');
  renderQualityMetrics(d.quality_metrics||{});
  renderDistribution(counts,d.issue_count);
  renderBusinessOverview(d.business_overview||null);
  var links=document.getElementById('links');links.innerHTML='';
  for(var key in (d.files||{})){var url=d.files[key];var a=document.createElement('a');a.className='download-card';a.href=url;a.download='';var icon=document.createElement('span');icon.className='download-icon';icon.textContent=key.indexOf('HTML')>=0?'</>':'XLSX';var copy=document.createElement('span');copy.className='download-copy';var title=document.createElement('strong');title.textContent=key;var desc=document.createElement('span');desc.textContent=key.indexOf('HTML')>=0?'八维业务化审核报告（总体结论/八维总表/分组明细）':'业务分组问题清单（概览+各类明细）';copy.append(title,desc);var action=document.createElement('span');action.className='download-action';action.textContent='下载';a.append(icon,copy,action);links.appendChild(a);}
  links.hidden=!Object.keys(d.files||{}).length;
}
</script><script>
if(!window.__kksHistoryInjected){
window.__kksHistoryInjected=true;
(function(){
function esc(s){s=(s===null||s===undefined)?'':String(s);var d=document.createElement('div');d.textContent=s;return d.innerHTML}
function rate(i,r){return r>0?(i/r*100).toFixed(1)+'%':'—'}
fetch('/api/history').then(function(r){return r.json()}).then(function(d){
  var entries=(d&&d.entries)||[];
  var host=document.createElement('div');
  host.style.cssText='max-width:1100px;margin:26px auto 40px;padding:0 22px';
  var det=document.createElement('details');
  det.style.cssText='background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:14px 20px';
  var sum=document.createElement('summary');
  sum.style.cssText='cursor:pointer;font-weight:600;color:#1d4ed8;font-size:15px';
  var tr=0,ti=0,ta=0,tp0=0;
  entries.forEach(function(e){tr+=+e.rows||0;ti+=+e.issues||0;ta+=+e.ai_reviewed||0;tp0+=+(e.priority&&e.priority.P0)||0});
  sum.textContent='审计台账（累计 '+entries.length+' 次审核 · '+(ti).toLocaleString()+' 个问题）';
  det.appendChild(sum);
  if(!entries.length){
    var p=document.createElement('p');p.style.cssText='color:#6b7a8d;font-size:13px';p.textContent='暂无审核记录，完成一次上传审核后自动生成。';det.appendChild(p);
  }else{
    var k=document.createElement('div');k.className='biz-kpis';
    k.innerHTML='<div class="biz-kpi"><b>'+entries.length+'</b><span>累计审核文件</span></div>'
      +'<div class="biz-kpi"><b>'+tr.toLocaleString()+'</b><span>累计数据行</span></div>'
      +'<div class="biz-kpi '+(tp0?'bad':'ok')+'"><b>'+ti.toLocaleString()+'</b><span>累计问题（P0 '+tp0+'）</span></div>'
      +'<div class="biz-kpi"><b>'+(tr?(ti/tr*100).toFixed(1)+'%':'—')+'</b><span>平均问题率</span></div>'
      +'<div class="biz-kpi"><b>'+ta.toLocaleString()+'</b><span>累计 AI 复核</span></div>';
    det.appendChild(k);
    var tw=document.createElement('div');tw.className='biz-table-wrap';
    var h='<table class="biz-table"><thead><tr><th>时间</th><th>文件</th><th>数据行</th><th>问题</th><th>问题率</th><th>P0</th><th>P1</th><th>P2</th><th>AI复核</th><th>口径</th></tr></thead><tbody>';
    entries.slice().reverse().forEach(function(e,idx){
      var ai=e.ai_reviewed?(e.ai_reviewed+'/'+e.ai_candidate):(e.ai_status==='completed'?e.ai_reviewed:'未启用');
      h+='<tr'+(idx===0?' style="background:#fffbe6"':'')+'><td>'+esc(e.ts)+'</td><td>'+esc(e.file)+'</td><td>'+(+e.rows||0).toLocaleString()+'</td><td>'+(+e.issues||0).toLocaleString()+'</td><td>'+rate(+e.issues||0,+e.rows||0)+'</td><td>'+((e.priority&&e.priority.P0)||0)+'</td><td>'+((e.priority&&e.priority.P1)||0)+'</td><td>'+((e.priority&&e.priority.P2)||0)+'</td><td>'+esc(ai)+'</td><td>'+(e.incremental?'增量':'全量')+'</td></tr>';
    });
    h+='</tbody></table>';
    tw.innerHTML=h;det.appendChild(tw);
    var note=document.createElement('div');note.style.cssText='color:#94a3b8;font-size:12px;margin-top:6px';note.textContent='跨文件汇总报告见 runs 目录「KKS审核台账汇总.html」；点击标题可折叠。';det.appendChild(note);
  }
  host.appendChild(det);document.body.appendChild(host);
}).catch(function(){});
})();
}
</script></body></html>"""
UPLOAD_PAGE = UPLOAD_PAGE.replace("</head>", RESULT_UI_STYLE + AUDIT_REDESIGN_STYLE + DASHBOARD_UI_STYLE + "</head>")


def json_bytes(payload: object) -> bytes:
    return json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")


def safe_filename(name: str) -> str:
    name = Path(name or "upload.xlsx").name
    name = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff._-]+", "_", name)
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return name[:160]


def parse_upload(content_type: str, body: bytes, fallback_name: str) -> tuple[str, bytes, tuple[str, bytes] | None]:
    if content_type.lower().startswith("multipart/form-data"):
        envelope = BytesParser(policy=default).parsebytes(
            f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8") + body
        )
        main_file: tuple[str, bytes] | None = None
        comparison_file: tuple[str, bytes] | None = None
        for part in envelope.iter_parts():
            field = part.get_param("name", header="content-disposition")
            if field not in {"file", "compare_file"}:
                continue
            data = part.get_payload(decode=True) or b""
            # Browsers include an empty optional file input in FormData. It is
            # not a comparison workbook and must not be passed to openpyxl.
            if field == "compare_file" and not data:
                continue
            item = (safe_filename(part.get_filename() or fallback_name), data)
            if field == "file":
                main_file = item
            else:
                comparison_file = item
        if main_file is None:
            raise ValueError("multipart 请求中没有名为 file 的文件字段")
        return main_file[0], main_file[1], comparison_file
    return safe_filename(fallback_name), body, None


class AuditHandler(BaseHTTPRequestHandler):
    server_version = "KKS-Audit/0.1"

    @property
    def runs_dir(self) -> Path:
        return self.server.runs_dir  # type: ignore[attr-defined]

    @property
    def max_upload_bytes(self) -> int:
        return self.server.max_upload_bytes  # type: ignore[attr-defined]

    def log_message(self, fmt: str, *args) -> None:
        LOGGER.info("http %s", fmt % args)

    def send_bytes(self, data: bytes, content_type: str, status: int = 200, download_name: str | None = None) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        if download_name:
            ascii_name = re.sub(r"[^A-Za-z0-9._-]+", "_", download_name) or "audit.xlsx"
            encoded_name = quote(download_name, safe="")
            self.send_header("Content-Disposition", f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{encoded_name}")
        self.end_headers()
        self.wfile.write(data)

    def send_json(self, payload: object, status: int = 200) -> None:
        self.send_bytes(json_bytes(payload), "application/json; charset=utf-8", status)

    def read_body(self, max_bytes: int = 2 * 1024 * 1024) -> bytes:
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            length = 0
        if length <= 0:
            raise ValueError("请求体为空")
        if length > max_bytes:
            raise ValueError(f"请求体超过 {max_bytes // 1024 // 1024} MB 限制")
        return self.rfile.read(length)

    def read_json(self) -> dict[str, object]:
        try:
            value = json.loads(self.read_body().decode("utf-8"))
        except json.JSONDecodeError as exc:
            raise ValueError("请求体不是有效 JSON") from exc
        if not isinstance(value, dict):
            raise ValueError("JSON 顶层必须是对象")
        return value

    def public_config(self) -> dict[str, object]:
        app_config = load_app_config()
        ai = load_config()
        audit = app_config.get("audit", {})
        return {
            "ai": {
                "enabled": ai.enabled,
                "base_url": ai.base_url,
                "model": ai.model,
                "timeout_seconds": ai.timeout_seconds,
                "api_key_configured": bool(ai.api_key),
                "api_key_hint": mask_secret(ai.api_key),
            },
            "audit": {
                "max_upload_mb": int(audit.get("max_upload_mb", 50)),
                "runs_dir": str(audit.get("runs_dir", "runs")),
            },
        }

    def handle_config(self) -> None:
        payload = self.read_json()
        incoming_ai = payload.get("ai", {})
        if not isinstance(incoming_ai, dict):
            raise ValueError("ai 配置必须是对象")
        patch: dict[str, object] = {"ai": {}}
        ai_patch = patch["ai"]  # type: ignore[index]
        assert isinstance(ai_patch, dict)
        if "base_url" in incoming_ai:
            ai_patch["base_url"] = normalize_base_url(str(incoming_ai["base_url"]))
        if "api_key" in incoming_ai and str(incoming_ai["api_key"]).strip():
            ai_patch["api_key"] = str(incoming_ai["api_key"]).strip()
        elif incoming_ai.get("clear_api_key"):
            ai_patch["api_key"] = ""
        if "model" in incoming_ai and str(incoming_ai["model"]).strip():
            ai_patch["model"] = str(incoming_ai["model"]).strip()
        if "enabled" in incoming_ai:
            ai_patch["enabled"] = bool(incoming_ai["enabled"])
        for key in ("timeout_seconds",):
            if key in incoming_ai:
                ai_patch[key] = int(incoming_ai[key])
        incoming_audit = payload.get("audit", {})
        if isinstance(incoming_audit, dict):
            patch["audit"] = {}
            audit_patch = patch["audit"]  # type: ignore[index]
            assert isinstance(audit_patch, dict)
            if "max_upload_mb" in incoming_audit:
                audit_patch["max_upload_mb"] = max(1, min(int(incoming_audit["max_upload_mb"]), 500))
            if "runs_dir" in incoming_audit:
                audit_patch["runs_dir"] = str(incoming_audit["runs_dir"])
        save_app_config(patch)
        LOGGER.info("config_updated ai_model=%s ai_enabled=%s", load_config().model, load_config().enabled)
        self.send_json({"ok": True, "config": self.public_config()})

    def handle_skill_get(self) -> None:
        documents = skill_documents()
        files = [item["path"] for item in documents]
        self.send_json({"ok": True, "files": files, "documents": documents, "rule_engine": "kks-audit/scripts/audit_template.py", "editable": True, "content": load_skill_context(), "backup_dir": "skill_backups"})

    def handle_skill_save(self) -> None:
        payload = self.read_json()
        path = str(payload.get("path", ""))
        content = payload.get("content")
        if not isinstance(content, str):
            raise ValueError("Skill 内容必须是文本")
        result = save_skill_document(path, content)
        self.send_json({"ok": True, **result, "message": "Skill 文件已保存；如果修改了 audit_template.py，请重启服务后再审核。"})

    def handle_skill_upload(self) -> None:
        body = self.read_body(max_bytes=10 * 1024 * 1024)
        content_type = self.headers.get("Content-Type", "")
        if not content_type.lower().startswith("multipart/form-data"):
            raise ValueError("Skill 上传必须使用 multipart/form-data")
        envelope = BytesParser(policy=default).parsebytes(
            f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8") + body
        )
        upload_name = ""
        upload_data = b""
        target = ""
        for part in envelope.iter_parts():
            field = part.get_param("name", header="content-disposition")
            if field == "target":
                target = (part.get_payload(decode=True) or b"").decode("utf-8", errors="strict").strip()
            elif field == "file":
                upload_name = Path(part.get_filename() or "").name
                upload_data = part.get_payload(decode=True) or b""
        if not upload_name or not upload_data:
            raise ValueError("没有收到 Skill 文件")
        if upload_name.lower().endswith(".zip"):
            result = install_skill_zip(upload_data, upload_name)
            message = "Skill 压缩包已安装；如果包含 audit_template.py，请重启服务后再审核。"
        else:
            if not upload_name.lower().endswith((".md", ".py")):
                raise ValueError("单文件 Skill 只支持 .md 或 .py；完整包请上传 .zip")
            if not target:
                if upload_name == "SKILL.md":
                    target = "kks-audit/SKILL.md"
                elif upload_name == "audit_template.py":
                    target = "kks-audit/scripts/audit_template.py"
                else:
                    target = f"kks-audit/references/{upload_name}"
            if upload_name.lower().endswith(".py") and not target.endswith("audit_template.py"):
                raise ValueError(".py 文件只能上传为 kks-audit/scripts/audit_template.py")
            if upload_name.lower().endswith(".md") and not target.endswith(".md"):
                raise ValueError(".md 文件只能上传为 Skill 文档或 references/*.md")
            result = save_skill_document(target, upload_data.decode("utf-8", errors="strict"))
            message = "Skill 文件已上传；如果修改了 audit_template.py，请重启服务后再审核。"
        self.send_json({"ok": True, **result, "message": message})

    def handle_ai_test(self) -> None:
        LOGGER.info("ai_connection_test_start")
        result = test_ai_connection()
        LOGGER.info("ai_connection_test_done status=%s model=%s", result.get("status"), result.get("model", ""))
        self.send_json(result, 200 if result.get("ok") else 400)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = unquote(parsed.path.rstrip("/") or "/")
        if path == "/":
            self.send_bytes(UPLOAD_PAGE.encode("utf-8"), "text/html; charset=utf-8")
            return
        if path == "/healthz":
            ai = load_config()
            self.send_json({"ok": True, "service": "kks-audit", "max_upload_bytes": self.max_upload_bytes, "ai_configured": bool(ai.api_key), "log_file": "logs/kks-audit.log"})
            return
        if path == "/api/config":
            self.send_json(self.public_config())
            return
        if path == "/api/skill":
            try:
                self.handle_skill_get()
            except Exception as exc:
                self.send_json({"error": str(exc)}, 400)
            return
        if path == "/api/history":
            from audit_history import LEDGER_NAME, load_ledger
            self.send_json({"entries": load_ledger(self.runs_dir / LEDGER_NAME)})
            return
        if path == "/api/logs":
            query = parse_qs(parsed.query)
            try:
                limit = int(query.get("limit", ["300"])[0])
            except ValueError:
                limit = 300
            self.send_json({"ok": True, "lines": tail_log(limit)})
            return

        parts = path.split("/")
        if len(parts) == 4 and parts[1:3] == ["api", "audits"]:
            run_id, artifact = parts[3], None
        elif len(parts) == 5 and parts[1:3] == ["api", "audits"]:
            run_id, artifact = parts[3], parts[4]
        else:
            self.send_json({"error": "Not Found"}, 404)
            return

        if not RUN_ID_RE.fullmatch(run_id):
            self.send_json({"error": "无效的审核编号"}, 400)
            return
        if artifact == "status":
            state = _get_job_state(self.server, run_id)  # type: ignore[arg-type]
            if state is None:
                self.send_json({"error": "审核任务不存在或服务已重启"}, 404)
            else:
                self.send_json(state)
            return
        run_root = (self.runs_dir / run_id).resolve()
        input_candidates = sorted(
            path for path in (run_root / "input").glob("*.xlsx")
            if not path.name.startswith("compare_")
        )
        html_name, xlsx_name = output_artifact_names(input_candidates[0]) if input_candidates else ("audit.html", "audit.xlsx")
        if artifact is None:
            artifact = html_name
        allowed = {
            html_name: ("text/html; charset=utf-8", None),
            xlsx_name: ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", xlsx_name),
        }
        if artifact not in allowed:
            self.send_json({"error": "不支持的报告类型"}, 404)
            return
        target = (self.runs_dir / run_id / artifact).resolve()
        if run_root not in target.parents or not target.is_file():
            self.send_json({"error": "审核结果不存在"}, 404)
            return
        content_type, download_name = allowed[artifact]
        self.send_bytes(target.read_bytes(), content_type, download_name=download_name)

    def do_POST(self) -> None:
        path = urlparse(self.path).path.rstrip("/")
        if path == "/api/config":
            try:
                self.handle_config()
            except Exception as exc:
                LOGGER.exception("config_update_failed")
                self.send_json({"error": str(exc)}, 400)
            return
        if path == "/api/skill":
            try:
                self.handle_skill_save()
            except Exception as exc:
                LOGGER.exception("skill_update_failed")
                self.send_json({"error": str(exc)}, 400)
            return
        if path == "/api/skill/upload":
            try:
                self.handle_skill_upload()
            except Exception as exc:
                LOGGER.exception("skill_upload_failed")
                self.send_json({"error": str(exc)}, 400)
            return
        if path == "/api/ai/test":
            try:
                self.handle_ai_test()
            except Exception as exc:
                LOGGER.exception("ai_connection_test_failed")
                self.send_json({"error": str(exc)}, 400)
            return
        if path != "/api/audits":
            self.send_json({"error": "Not Found"}, 404)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            length = 0
        if length <= 0:
            self.send_json({"error": "请求体为空"}, 400)
            return
        if length > self.max_upload_bytes:
            self.send_json({"error": f"文件超过 {self.max_upload_bytes // 1024 // 1024} MB 限制"}, 413)
            self.close_connection = True
            return
        body = self.rfile.read(length)
        try:
            filename, data, comparison_upload = parse_upload(self.headers.get("Content-Type", ""), body, self.headers.get("X-Filename", "upload.xlsx"))
            if not data:
                raise ValueError("上传文件为空")
            if not filename.lower().endswith(".xlsx"):
                raise ValueError("目前只支持 .xlsx 文件")
            run_id = f"{__import__('datetime').datetime.now():%Y%m%dT%H%M%S}_{uuid4().hex[:8]}"
            run_dir = self.runs_dir / run_id
            input_dir = run_dir / "input"
            input_dir.mkdir(parents=True, exist_ok=False)
            input_path = input_dir / filename
            input_path.write_bytes(data)
            LOGGER.info("audit_upload run_id=%s file=%s bytes=%s", run_id, filename, len(data))
            comparison_path = None
            if comparison_upload:
                comparison_name, comparison_data = comparison_upload
                if not comparison_name.lower().endswith(".xlsx"):
                    raise ValueError("对比文件目前只支持 .xlsx 文件")
                comparison_path = input_dir / f"compare_{comparison_name}"
                comparison_path.write_bytes(comparison_data)
                LOGGER.info("audit_compare_upload run_id=%s file=%s bytes=%s", run_id, comparison_name, len(comparison_data))
            _update_job_state(
                self.server,  # type: ignore[arg-type]
                run_id,
                status="queued",
                phase="queued",
                percent=2,
                message="审核任务已创建",
                hint="正在排队启动后台审核",
                ai_review={"status": "pending", "candidate_count": 0, "reviewed_count": 0},
            )
            worker = threading.Thread(
                target=_run_audit_job,
                args=(self.server, run_id, input_path, run_dir, comparison_path),  # type: ignore[arg-type]
                name=f"kks-audit-{run_id}",
                daemon=True,
            )
            worker.start()
            self.send_json({"run_id": run_id, "status": "queued", "status_url": f"/api/audits/{run_id}/status"}, 202)
        except Exception as exc:  # keep client errors JSON-readable
            LOGGER.exception("audit_http_failed")
            self.send_json({"error": str(exc)}, 400)


def main() -> int:
    configure_logging(clear=True)
    parser = argparse.ArgumentParser(description="启动独立 KKS 编码审核服务")
    parser.add_argument("--host", default="127.0.0.1", help="监听地址；局域网共享可用 0.0.0.0")
    parser.add_argument("--port", type=int, default=8080)
    parser.add_argument("--runs-dir", type=Path, default=None)
    parser.add_argument("--max-upload-mb", type=int, default=None)
    args = parser.parse_args()
    app_config = load_app_config()
    audit_config = app_config.get("audit", {}) if isinstance(app_config.get("audit"), dict) else {}
    runs_dir = args.runs_dir or Path(str(audit_config.get("runs_dir", "runs")))
    if not runs_dir.is_absolute():
        runs_dir = APP_ROOT / runs_dir
    max_upload_mb = args.max_upload_mb or int(audit_config.get("max_upload_mb", 50))
    runs_dir.mkdir(parents=True, exist_ok=True)
    server = ThreadingHTTPServer((args.host, args.port), AuditHandler)
    server.runs_dir = runs_dir.resolve()  # type: ignore[attr-defined]
    server.max_upload_bytes = max_upload_mb * 1024 * 1024  # type: ignore[attr-defined]
    server.jobs = {}  # type: ignore[attr-defined]
    server.jobs_lock = threading.Lock()  # type: ignore[attr-defined]
    print(f"KKS audit service: http://{args.host}:{args.port}/")
    print(f"runs directory: {server.runs_dir}")
    LOGGER.info("service_started host=%s port=%s runs_dir=%s", args.host, args.port, server.runs_dir)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务已停止")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
