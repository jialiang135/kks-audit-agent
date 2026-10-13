"""One metric definition for the API, HTML and Excel reports."""
import re
from collections import Counter

DEVICE = re.compile(r"[0-9]{2}[A-HJ-NP-Z]{3}[0-9]{2}[A-HJ-NP-Z]{2}[0-9]{3}")


def summary_metrics(result):
    records = result.get("_ai_context", {}).get("records", [])
    codes = [str(r.get("kks_code") or "").strip() for r in records if isinstance(r, dict)]
    issues = [i for i in result.get("issues", []) if isinstance(i, dict) and i.get("status") != "resolved"]
    # Severity measures confirmed findings. Pending reviews and likely false
    # positives are reported separately so review workload is not presented as
    # confirmed noncompliance.
    decisions = Counter(str(i.get("final_decision") or "confirmed_issue") for i in issues)
    counts = Counter(
        str(i.get("priority", "P2")) for i in issues
        if str(i.get("final_decision") or "confirmed_issue") == "confirmed_issue"
    )
    metrics = result.get("metrics", {})
    cached = result.get("quality_metrics", {})
    return {
        "effective_kks": int(result.get("data_rows", len([c for c in codes if c])) or 0),
        "device_level_codes": sum(bool(DEVICE.fullmatch(c)) for c in codes) if records else int(cached.get("device_level_codes", 0)),
        "duplicate_codes": int(metrics.get("duplicate_code_groups", 0) or 0),
        "parent_errors": int(metrics.get("orphan_rows", 0) or 0) + int(metrics.get("prefix_mismatch_rows", 0) or 0),
        "p0": counts.get("P0", 0), "p1": counts.get("P1", 0), "p2": counts.get("P2", 0),
        "confirmed_issues": decisions.get("confirmed_issue", 0),
        "pending_review": decisions.get("needs_human", 0),
        "suspected_false_positives": decisions.get("likely_false_positive", 0),
        "rule_hits": len(issues),
    }


def confirmed_priority_counts(result):
    """Severity counts for confirmed findings only, shared by API and reports."""
    counts = Counter()
    for item in result.get("issues", []):
        if not isinstance(item, dict) or item.get("status") == "resolved":
            continue
        if str(item.get("final_decision") or "confirmed_issue") == "confirmed_issue":
            counts[str(item.get("priority", "P2"))] += 1
    return {key: counts.get(key, 0) for key in ("P0", "P1", "P2")}
