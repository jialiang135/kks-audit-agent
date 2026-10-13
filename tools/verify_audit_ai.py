"""Run a reproducible, stratified AI audit evaluation using real workbook rows."""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import random
import sys
from collections import Counter, defaultdict
from dataclasses import replace
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import ai_review
import run_audit
from report_business import build_overview, render_html, render_xlsx


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--per-rule", type=int, default=4)
    parser.add_argument("--previous", type=Path, help="Compare the preceding evaluation against current rule findings")
    parser.add_argument("--include-previous", action="store_true", help="Retest surviving findings from the preceding evaluation")
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    before = hashlib.sha256(args.input.read_bytes()).hexdigest()
    full = run_audit.audit_file(args.input, args.output, run_ai_review=False, write_outputs=False)
    (args.output / "完整规则审查.json").write_text(json.dumps(
        {k: v for k, v in full.items() if not k.startswith("_")}, ensure_ascii=False, indent=2), encoding="utf-8")
    full_overview = build_overview(full)
    (args.output / "完整规则审查报告.html").write_text(render_html(full, full_overview), encoding="utf-8")
    render_xlsx(args.output / "完整规则审查问题清单.xlsx", full, full_overview)
    groups = defaultdict(list)
    for item in full["issues"]:
        groups[(item.get("sheet"), item["rule_id"], item.get("category"),
                bool(item.get("compliance_locked")), bool(item.get("evidence_missing")))].append(item)
    rng = random.Random(20261012)
    selected = []
    for key in sorted(groups, key=str):
        selected.extend(rng.sample(groups[key], min(args.per_rule, len(groups[key]))))
    sample = copy.copy(full)
    sample["issues"] = copy.deepcopy(selected)
    sample["evaluation_scope"] = {"type": "按工作表、规则、类别分层抽样", "seed": 20261012,
                                  "full_issues": len(full["issues"]), "sample_issues": len(selected)}
    if args.previous:
        previous = json.loads(args.previous.read_text(encoding="utf-8"))
        if args.include_previous:
            previous_keys = {(i.get("sheet"), str(i.get("excel_row")), i["rule_id"]) for i in previous["issues"]}
            selected_ids = {id(i) for i in selected}
            selected.extend(i for i in full["issues"] if id(i) not in selected_ids
                            and (i.get("sheet"), str(i.get("excel_row")), i["rule_id"]) in previous_keys)
            sample["issues"] = copy.deepcopy(selected)
            sample["evaluation_scope"]["sample_issues"] = len(selected)
            sample["evaluation_scope"]["includes_previous_findings"] = True
        keys = {(i.get("sheet"), str(i.get("excel_row")), i["rule_id"]) for i in full["issues"]}
        comparison = [{"sheet": i.get("sheet"), "excel_row": i.get("excel_row"), "rule_id": i["rule_id"],
                       "previous_final_decision": i.get("final_decision"),
                       "still_triggered": (i.get("sheet"), str(i.get("excel_row")), i["rule_id"]) in keys}
                      for i in previous["issues"]]
        (args.output / "前后规则对照.json").write_text(json.dumps(comparison, ensure_ascii=False, indent=2), encoding="utf-8")
    sample["priority_counts"] = dict(Counter(i["priority"] for i in selected))
    sample["issue_count"] = len(selected)
    config = ai_review.load_config()
    ai_review.load_config = lambda: replace(config, max_candidates=0)
    print(json.dumps(sample["evaluation_scope"], ensure_ascii=False), flush=True)
    def progress(event):
        if event.get("phase") == "ai":
            print(json.dumps({k: event.get(k) for k in ("stage", "hint", "candidate_count", "reviewed_count")},
                             ensure_ascii=False), flush=True)
    run_audit.apply_ai_review_stage(sample, progress_callback=progress)
    sample["_ai_context"] = {"records": full["_ai_context"]["records"]}
    ov = build_overview(sample)
    (args.output / "AI抽样复核报告.html").write_text(render_html(sample, ov), encoding="utf-8")
    render_xlsx(args.output / "AI抽样复核问题清单.xlsx", sample, ov)
    serializable = {k: v for k, v in sample.items() if not k.startswith("_")}
    (args.output / "AI复核结果.json").write_text(json.dumps(serializable, ensure_ascii=False, indent=2), encoding="utf-8")
    ai = sample.get("ai_review", {})
    checks = {"source_unchanged": before == hashlib.sha256(args.input.read_bytes()).hexdigest(),
              "hard_findings_protected": all(i.get("final_decision") == "confirmed_issue" for i in sample["issues"] if i.get("compliance_locked")),
              "missing_evidence_pending": all(i.get("final_decision") == "needs_human" for i in sample["issues"] if i.get("evidence_missing")),
              "all_candidates_reviewed": ai.get("candidate_count") == ai.get("reviewed_count"),
              "false_positives_verified": ai.get("verification_count", 0) == ai.get("verified_count", 0)}
    evaluation = {"checks": checks, "ai_status": ai.get("status"), "candidate_count": ai.get("candidate_count"),
                  "reviewed_count": ai.get("reviewed_count"), "final_decisions": sample.get("final_decision_counts"),
                  "errors": ai.get("errors", []), "scope": sample["evaluation_scope"]}
    (args.output / "效果检查.json").write_text(json.dumps(evaluation, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(evaluation, ensure_ascii=False), flush=True)
    return 0 if all(checks.values()) and not ai.get("errors") else 1


if __name__ == "__main__":
    raise SystemExit(main())
