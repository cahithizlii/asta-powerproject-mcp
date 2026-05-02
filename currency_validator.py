"""Phase 6.1 — currency mode cross-validation pure module.

Extends Phase 5a `_evm_detect_currency_mode` from binary cost/hours
into 4-mode RULE 3 detection with cross-source validation.

4 modes returned by all detect_* functions:
    'cost'      — cost loaded (real cost data present)
    'hours'     — hours-only (no cost or RULE 3 pattern: cost == qty)
    'mixed'     — partial cost loading (some entries have cost, some don't)
    'uncertain' — insufficient data to decide

Functions:
    detect_mode_from_xer_assignments(assignments) -> str
    detect_mode_from_tasks_resources(tasks, resources) -> str
    extract_currency_code(xer_header_fields) -> str | None
    cross_validate_modes(sources) -> dict

Pure Python, zero dependencies. Yaklaşım C: pure module + I/O adapters
in core + dispatcher.
"""
from collections import Counter

# RULE 3: target_cost == target_qty within this tolerance -> not cost loaded
_RULE3_TOLERANCE = 0.01


def _safe_float(v):
    """Return float(v) or None if non-numeric / None."""
    if v is None:
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def detect_mode_from_xer_assignments(assignments):
    """RULE 3 cost-loading detection from XER TASKRSRC rows.

    Logic per row (target_qty, target_cost both numeric):
        target_qty > 0:
            target_cost == target_qty (within tolerance) -> hours_signal
            target_cost differs                          -> cost_signal
        target_qty == 0:
            target_cost > 0 -> cost_signal
            target_cost == 0 -> no signal (skipped)

    Aggregate:
        all hours_signal, no cost_signal -> 'hours'
        all cost_signal, no hours_signal -> 'cost'
        both signals present             -> 'mixed'
        no signals (empty/all zero)      -> 'uncertain'
    """
    if not assignments:
        return "uncertain"
    hours_signal = 0
    cost_signal = 0
    for row in assignments:
        qty = _safe_float(row.get("target_qty"))
        cost = _safe_float(row.get("target_cost"))
        if qty is None or cost is None:
            continue
        if qty > 0:
            if abs(cost - qty) <= _RULE3_TOLERANCE:
                hours_signal += 1
            else:
                cost_signal += 1
        else:
            if cost > 0:
                cost_signal += 1
    if hours_signal == 0 and cost_signal == 0:
        return "uncertain"
    if cost_signal == 0:
        return "hours"
    if hours_signal == 0:
        return "cost"
    return "mixed"


def detect_mode_from_tasks_resources(tasks, resources):
    """Aggregate-based detection from MSP-shape task/resource cost fields.

    Logic per entry (cost field numeric):
        cost > 0  -> cost_signal
        cost == 0 -> hours_signal
        cost None / non-numeric -> skipped

    Aggregate:
        only hours_signal -> 'hours'
        only cost_signal  -> 'cost'
        both              -> 'mixed'
        no signals        -> 'uncertain'
    """
    hours_signal = 0
    cost_signal = 0
    for entry in (tasks or []):
        c = _safe_float(entry.get("cost"))
        if c is None:
            continue
        if c > 0:
            cost_signal += 1
        else:
            hours_signal += 1
    for entry in (resources or []):
        c = _safe_float(entry.get("cost"))
        if c is None:
            continue
        if c > 0:
            cost_signal += 1
        else:
            hours_signal += 1
    if hours_signal == 0 and cost_signal == 0:
        return "uncertain"
    if cost_signal == 0:
        return "hours"
    if hours_signal == 0:
        return "cost"
    return "mixed"


def extract_currency_code(xer_header_fields):
    """ERMHDR.currency field — 3-letter code or None.

    Returns None if header dict is None/empty or 'currency' key absent
    or value is empty/blank.
    """
    if not xer_header_fields:
        return None
    code = xer_header_fields.get("currency")
    if not code:
        return None
    code = str(code).strip()
    return code if code else None


def cross_validate_modes(sources):
    """Cross-validate currency mode across multiple sources.

    Args:
        sources: list of (source_name, mode) tuples. mode in
                 {'cost','hours','mixed','uncertain'}.

    Returns:
        {
            consensus_mode: dominant non-uncertain mode (or 'uncertain'),
            confidence: 'high' | 'medium' | 'low',
                'high'   = all non-uncertain sources agree (single mode)
                'medium' = >=66% sources share dominant mode
                'low'    = no clear majority
            conflicts: list of (source_a, mode_a, source_b, mode_b) pairs
                where modes differ (uncertain pairs excluded),
            warnings: list of human-readable strings,
            source_counts: dict {mode: count} for ALL inputs (incl. uncertain),
        }
    """
    sources = sources or []
    # Full counter (including uncertain) for transparency.
    source_counts = dict(Counter(m for _, m in sources))
    warnings = []

    # Note any 'mixed' source.
    mixed_sources = [s for s, m in sources if m == "mixed"]
    if mixed_sources:
        warnings.append(
            "Source(s) report 'mixed' cost loading: "
            + ", ".join(mixed_sources)
            + ". Partial cost data detected — RULE 3 violation likely."
        )

    # Filter out uncertain for consensus.
    real = [(s, m) for s, m in sources if m != "uncertain"]
    if not real:
        return {
            "consensus_mode": "uncertain",
            "confidence": "low",
            "conflicts": [],
            "warnings": warnings,
            "source_counts": source_counts,
        }

    # Compute consensus.
    mode_counts = Counter(m for _, m in real)
    most_common_mode, top_count = mode_counts.most_common(1)[0]
    total = len(real)
    distinct_modes = len(mode_counts)

    if distinct_modes == 1:
        confidence = "high"
    elif top_count / total >= 2.0 / 3.0:
        confidence = "medium"
    else:
        confidence = "low"

    # Conflicts: pairs of sources with different modes.
    conflicts = []
    for i, (sa, ma) in enumerate(real):
        for (sb, mb) in real[i + 1:]:
            if ma != mb:
                conflicts.append((sa, ma, sb, mb))

    return {
        "consensus_mode": most_common_mode,
        "confidence": confidence,
        "conflicts": conflicts,
        "warnings": warnings,
        "source_counts": source_counts,
    }
