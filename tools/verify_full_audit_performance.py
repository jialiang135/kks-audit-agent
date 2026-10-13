"""Verify a complete workbook with the measured, conservative fast settings."""
from pathlib import Path
from dataclasses import replace
import hashlib
import json
import sys
import time

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import ai_review
import run_audit
from agent_runtime import KksAuditAgent
from audit_metrics import summary_metrics
from report_business import build_overview


def main():
    source, output = Path(sys.argv[1]), Path(sys.argv[2])
    output.mkdir(parents=True, exist_ok=True)
    before = hashlib.sha256(source.read_bytes()).hexdigest()
    config = replace(ai_review.load_config(), thinking_mode="disabled", review_concurrency=4,
                     review_batch_size=12, context_rows=12, brief_output=False, max_candidates=0)
    ai_review.load_config = lambda: config
    started = time.perf_counter()
    def progress(event):
        if event.get("stage") in {"model_ready", "review_running", "verification_running", "completed", "report"}:
            print(json.dumps({"seconds": round(time.perf_counter() - started, 1),
                "stage": event.get("stage"), "hint": event.get("hint"),
                "candidate_count": event.get("candidate_count"), "reviewed_count": event.get("reviewed_count"),
                "verified_count": event.get("verified_count")}, ensure_ascii=False), flush=True)
    result = KksAuditAgent().run(source, output, progress_callback=progress)
    elapsed = time.perf_counter() - started
    ai = result["ai_review"]
    metrics = run_audit.quality_metrics(result)
    overview = build_overview(result)
    checks = {
        "source_unchanged": before == hashlib.sha256(source.read_bytes()).hexdigest(),
        "all_candidates_reviewed": ai.get("candidate_count") == ai.get("reviewed_count"),
        "all_false_positives_verified": ai.get("verification_count") == ai.get("verified_count"),
        "hard_findings_protected": all(i.get("final_decision") == "confirmed_issue" for i in result["issues"] if i.get("compliance_locked")),
        "missing_evidence_pending": all(i.get("final_decision") == "needs_human" for i in result["issues"] if i.get("evidence_missing")),
        "target_device_not_flagged_as_signal": not any(i["rule_id"] == "KKS-38" and i.get("excel_row") == 76 for i in result["issues"]),
        "api_html_metrics_equal": all(int(k["value"].replace(",", "")) == metrics[k["key"]] for k in overview["kpis"] if "key" in k),
        "no_errors": not ai.get("errors"),
    }
    measurements = {"seconds": round(elapsed, 3), "reviewed_rows": result["reviewed_rows"],
        "coded_rows": result["data_rows"], "candidate_count": ai.get("candidate_count"),
        "reviewed_count": ai.get("reviewed_count"), "verification_count": ai.get("verification_count"),
        "verified_count": ai.get("verified_count"), "checks": checks, "metrics": metrics,
        "settings": {"thinking": "disabled", "concurrency": 4, "batch": 12, "context": 12, "brief": False}}
    (output / "整文件耗时与质量核查.json").write_text(json.dumps(measurements, ensure_ascii=False, indent=2), encoding="utf-8")
    (output / "审查结果.json").write_text(json.dumps({k:v for k,v in result.items() if not k.startswith("_")}, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(measurements, ensure_ascii=False), flush=True)
    return 0 if all(checks.values()) else 1


if __name__ == "__main__":
    raise SystemExit(main())
