from __future__ import annotations

"""Validate the recruiter-facing sanitized Streamlit portfolio snapshot.

The validator is intentionally strict for public deployment: it fails when the snapshot is
missing, contains no portfolio visuals, lacks core metrics, or exposes forbidden position-level
fields. This prevents the public showcase from silently falling back to demo data.
"""

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SNAPSHOT = ROOT / "showcase" / "data" / "portfolio_snapshot.json"

REQUIRED_METRICS = {
    "annualized_return",
    "annualized_volatility",
    "sharpe",
    "tracking_error",
    "information_ratio",
    "max_drawdown",
}

FORBIDDEN_KEYS = {
    "ticker",
    "symbol",
    "company",
    "companyname",
    "shares",
    "averagecost",
    "costbasis",
    "marketvalue",
    "unrealizedpnl",
    "unrealizedpnlpct",
    "transaction",
    "transactions",
}


def _walk_keys(value):
    if isinstance(value, dict):
        for key, child in value.items():
            yield str(key)
            yield from _walk_keys(child)
    elif isinstance(value, list):
        for child in value:
            yield from _walk_keys(child)


def main():
    if not SNAPSHOT.exists():
        raise SystemExit(
            "Recruiter showcase validation failed: sanitized portfolio snapshot is missing. "
            "Run institutional_research/run_research.py and automation/build_showcase_snapshot.py first."
        )

    data = json.loads(SNAPSHOT.read_text(encoding="utf-8"))

    if data.get("snapshot_type") != "sanitized_real_portfolio_analytics":
        raise SystemExit("Recruiter showcase validation failed: snapshot is not marked as real sanitized analytics.")

    metrics = data.get("metrics") or {}
    missing_metrics = sorted(REQUIRED_METRICS - set(metrics))
    if missing_metrics:
        raise SystemExit(
            "Recruiter showcase validation failed: missing core metrics: "
            + ", ".join(missing_metrics)
        )

    holdings = data.get("holdings") or []
    if len(holdings) < 2:
        raise SystemExit(
            "Recruiter showcase validation failed: fewer than two anonymous holdings were exported."
        )

    for idx, holding in enumerate(holdings, start=1):
        if not holding.get("holding") or holding.get("weight") is None:
            raise SystemExit(
                f"Recruiter showcase validation failed: anonymous holding {idx} is incomplete."
            )

    timeseries = data.get("timeseries") or []
    if len(timeseries) < 2:
        raise SystemExit(
            "Recruiter showcase validation failed: portfolio growth history is missing or too short."
        )

    seen_forbidden = sorted(
        {
            key
            for key in _walk_keys(data)
            if key.replace("_", "").replace(" ", "").lower() in FORBIDDEN_KEYS
        }
    )
    if seen_forbidden:
        raise SystemExit(
            "Recruiter showcase validation failed: forbidden public fields detected: "
            + ", ".join(seen_forbidden)
        )

    total_weight = sum(float(row.get("weight") or 0.0) for row in holdings)
    if not 0.97 <= total_weight <= 1.03:
        raise SystemExit(
            f"Recruiter showcase validation failed: anonymous portfolio weights sum to {total_weight:.4f}."
        )

    print("Recruiter showcase validation passed.")
    print(f"Anonymous holdings: {len(holdings)}")
    print(f"Growth observations: {len(timeseries)}")
    print(f"Core metrics: {len(REQUIRED_METRICS)} / {len(REQUIRED_METRICS)}")


if __name__ == "__main__":
    main()
