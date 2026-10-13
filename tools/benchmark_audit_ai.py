"""Paired, one-factor-at-a-time measurements on a fixed real workbook sample.

Credentials remain in memory; only timing/token metadata and audit evidence
are saved. Experimental monkeypatches apply only inside this CLI process.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import random
import sys
import threading
import time
from collections import Counter
from dataclasses import replace
from pathlib import Path
from statistics import mean

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import ai_review
import run_audit

ORIGINAL_CLIENT = ai_review.OpenAICompatibleClient
ORIGINAL_PROMPT = ai_review.SYSTEM_PROMPT
CONFIG = ai_review.load_config()


def dump(path, value):
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def prepare(path, output):
    started = time.perf_counter()
    full = run_audit.audit_file(path, output, run_ai_review=False, write_outputs=False)
    elapsed = time.perf_counter() - started
    candidates = [i for i in full["issues"] if ai_review._is_candidate(i)]
    semantic = [i for i in candidates if i["rule_id"] == "KKS-42"]
    other = [i for i in candidates if i["rule_id"] == "KKS-A"]
    anchors = [i for i in semantic if i.get("excel_row") in {11, 41, 43, 45, 47, 66, 76}]
    rng = random.Random(20261013)
    remaining = [i for i in semantic if i not in anchors]
    selected = anchors + rng.sample(remaining, min(36 - len(anchors), len(remaining)))
    selected += rng.sample(other, min(12, len(other)))
    # Exactly 48 issues where possible, with 36 semantic and 12 type-lineage
    # issues, yielding four equal 12-item groups in the baseline.
    if len(selected) < 48:
        pool = [i for i in candidates if i not in selected]
        selected += rng.sample(pool, min(48 - len(selected), len(pool)))
    dump(output / "fixed_sample.json", {"seed": 20261013, "issues": selected,
        "rule_seconds": elapsed, "workbook_rows": full["reviewed_rows"],
        "all_ai_candidates": len(candidates), "sample_size": len(selected)})
    return full, selected


def trial(full, selected, condition, output):
    config = replace(CONFIG, max_candidates=0, summary_enabled=False,
                     review_concurrency=condition["concurrency"], review_batch_size=condition["batch"],
                     context_rows=condition["context"], thinking_mode="" if condition["thinking"] == "default" else condition["thinking"],
                     brief_output=condition.get("brief", False))
    traces, lock, activity = [], threading.Lock(), {"active": 0, "peak": 0}

    class TracedClient(ORIGINAL_CLIENT):
        def _request_json(self, path, payload=None):
            if payload is not None:
                payload = dict(payload)
                if condition["thinking"] != "default":
                    payload["thinking"] = {"type": condition["thinking"]}
            trace = {"request_chars": len(json.dumps(payload, ensure_ascii=False)) if payload else 0}
            started = time.perf_counter()
            with lock:
                activity["active"] += 1
                activity["peak"] = max(activity["peak"], activity["active"])
            try:
                response = super()._request_json(path, payload)
                usage = response.get("usage", {})
                message = ((response.get("choices") or [{}])[0].get("message") or {})
                trace.update({"ok": True, "usage": usage,
                    "reasoning_chars": len(str(message.get("reasoning_content") or "")),
                    "answer_chars": len(str(message.get("content") or "")),
                    "finish_reason": (response.get("choices") or [{}])[0].get("finish_reason")})
                return response
            except Exception as exc:
                trace.update(ok=False, error_type=type(exc).__name__)
                raise
            finally:
                trace["seconds"] = round(time.perf_counter() - started, 3)
                with lock:
                    activity["active"] -= 1
                    traces.append(trace)

    ai_review.OpenAICompatibleClient = TracedClient
    ai_review.load_config = lambda: config
    ai_review.MODEL_REQUEST_SLOTS = threading.BoundedSemaphore(condition["concurrency"])
    ai_review.INTERNAL_GROUP_SIZE = condition["batch"]
    ai_review.MAX_CONTEXT_ROWS = condition["context"]
    ai_review.SYSTEM_PROMPT = ORIGINAL_PROMPT
    sample = copy.copy(full)
    sample.pop("_ai_evidence_index", None)
    sample["issues"] = copy.deepcopy(selected)
    sample["evaluation_scope"] = {"type": "固定实际样本控制变量实验", "sample_issues": len(selected),
                                  "full_issues": len(full["issues"])}
    started = time.perf_counter()
    print(json.dumps({"event": "start", "condition": condition, "sample_size": len(selected)}, ensure_ascii=False), flush=True)
    def progress(event):
        if event.get("stage") in {"review_running", "verification_running", "completed"}:
            print(json.dumps({"event": "progress", "condition": condition["name"],
                "stage": event.get("stage"), "reviewed": event.get("reviewed_count"),
                "verified": event.get("verified_count"), "hint": event.get("hint")}, ensure_ascii=False), flush=True)
    ai = ai_review.review_issue_candidates(sample, progress_callback=progress)
    elapsed = time.perf_counter() - started
    target = next((i for i in sample["issues"] if i.get("excel_row") == 76 and i["rule_id"] == "KKS-42"), {})
    citations = [i for i in sample["issues"] if i.get("ai_evidence")]
    checks = {"complete": ai.get("status") == "completed" and ai.get("reviewed_count") == len(selected),
              "verification_complete": ai.get("verification_count") == ai.get("verified_count"),
              "no_errors": not ai.get("errors"),
              "device_anchor_safe": not target or target.get("final_decision") in {"likely_false_positive", "needs_human"},
              "no_missing_evidence": len(citations) == len(selected) and all(str(i.get("ai_evidence", "")).strip() for i in citations)}
    result = {"condition": condition, "seconds": round(elapsed, 3), "sample_size": len(selected),
        "throughput_per_minute": round(len(selected) / max(elapsed, .001) * 60, 3),
        "request_count": len(traces), "peak_requests": activity["peak"],
        "request_mean_seconds": round(mean(t["seconds"] for t in traces), 3) if traces else 0,
        "request_chars": sum(t["request_chars"] for t in traces),
        "reasoning_chars": sum(t.get("reasoning_chars", 0) for t in traces),
        "answer_chars": sum(t.get("answer_chars", 0) for t in traces),
        "checks": checks, "eligible": all(checks.values()), "ai_status": ai.get("status"),
        "reviewed": ai.get("reviewed_count"), "verified": ai.get("verified_count"),
        "errors": ai.get("errors", []), "final_decisions": ai.get("final_decision_counts"),
        "anchor_decision": target.get("final_decision"), "requests": traces,
        "reviews": [{k: i.get(k) for k in ("sheet", "excel_row", "rule_id", "kks_code", "name",
            "final_decision", "ai_evidence", "ai_reason", "ai_suggestion")} for i in sample["issues"]]}
    dump(output / (condition["name"] + ".json"), result)
    print(json.dumps({k: v for k, v in result.items() if k not in {"requests", "reviews"}}, ensure_ascii=False), flush=True)
    return result


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--phase", choices=["pilot", "matrix"], default="pilot")
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    before = hashlib.sha256(args.input.read_bytes()).hexdigest()
    full, selected = prepare(args.input, args.output)
    base = {"concurrency": 2, "batch": 12, "context": 12, "thinking": "disabled", "brief": False}
    if args.phase == "pilot":
        # Paired pilot: same 12 semantic issues, all evidence and settings equal.
        selected = [i for i in selected if i["rule_id"] == "KKS-42"][:12]
        conditions = [{**base, "name": "pilot_default", "thinking": "default"},
                      {**base, "name": "pilot_disabled"}]
    else:
        conditions = [{**base, "name": "baseline"}, {**base, "name": "concurrency_1", "concurrency": 1},
            {**base, "name": "concurrency_4", "concurrency": 4}, {**base, "name": "batch_6", "batch": 6},
            {**base, "name": "batch_24", "batch": 24}, {**base, "name": "context_6", "context": 6},
            {**base, "name": "brief_output", "brief": True}, {**base, "name": "baseline_repeat"}]
    results = []
    for condition in conditions:
        results.append(trial(full, selected, condition, args.output))
        dump(args.output / "summary.json", {"phase": args.phase, "source_hash": before,
             "source_unchanged": before == hashlib.sha256(args.input.read_bytes()).hexdigest(),
             "results": [{k: v for k, v in r.items() if k not in {"requests", "reviews"}} for r in results]})
    if args.phase == "matrix":
        passing = [r for r in results if r["eligible"]]
        if passing:
            best = min(passing, key=lambda r: r["seconds"])
            # Re-test the measured best rather than assume factor effects add.
            winner = dict(best["condition"], name="winner_repeat")
            repeat = trial(full, selected, winner, args.output)
            dump(args.output / "winner.json", {"first": {k:v for k,v in best.items() if k not in {"requests","reviews"}},
                 "repeat": {k:v for k,v in repeat.items() if k not in {"requests","reviews"}}})
            # Compare the combined configuration explicitly, rather than
            # assuming independent factor improvements add up.
            trial(full, selected, {**winner, "name": "combined_brief", "brief": True}, args.output)
            trial(full, selected, {**base, "name": "default_thinking_48", "thinking": "default"}, args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
