from __future__ import annotations

"""Build a resume-safe public snapshot from the latest private portfolio outputs.

The snapshot preserves real aggregate portfolio analytics while removing identifiers and
sensitive position economics. It never exports tickers, company names, shares, average cost,
market value, unrealized P&L, transaction history, or private source files.
"""

import csv
import json
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "institutional_research" / "outputs" / "latest"
DEST = ROOT / "showcase" / "data" / "portfolio_snapshot.json"


def _read_csv(name: str) -> list[dict]:
    path = OUT / f"{name}.csv"
    if not path.exists() or path.stat().st_size == 0:
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def _float(value):
    try:
        if value in (None, "", "nan", "NaN"):
            return None
        return float(value)
    except Exception:
        return None


def _safe_metric_map(portfolio: dict) -> dict:
    allowed = [
        "annualized_return",
        "annualized_volatility",
        "sharpe",
        "sortino",
        "tracking_error",
        "information_ratio",
        "max_drawdown",
        "beta",
        "daily_var_95",
        "daily_expected_shortfall_95",
        "active_annualized_return",
        "daily_active_hit_rate",
        "top_1_weight",
        "top_5_weight",
        "effective_number_of_holdings",
        "largest_risk_contribution",
        "annualized_alpha",
        "alpha_t_stat",
        "alpha_p_value",
        "alpha_r_squared",
    ]
    return {k: portfolio.get(k) for k in allowed if k in portfolio}


def _anonymous_holdings() -> list[dict]:
    rows = _read_csv("holdings_analysis")
    clean = []
    for row in rows:
        weight = _float(row.get("Weight"))
        risk = _float(row.get("RiskContributionPct"))
        if weight is None:
            continue
        clean.append({"weight": weight, "risk_contribution": risk})
    clean.sort(key=lambda x: x["weight"], reverse=True)
    out = []
    for idx, row in enumerate(clean):
        label = f"Holding {chr(65 + idx)}" if idx < 26 else f"Holding {idx + 1}"
        out.append({"holding": label, **row})
    return out


def _alpha() -> list[dict]:
    rows = _read_csv("alpha_summary")
    out = []
    for row in rows:
        model = str(row.get("Model") or "").strip()
        if not model:
            continue
        display = {
            "CAPM - Benchmark": "CAPM / Jensen",
            "Fama-French 3": "Fama-French 3",
            "Carhart 4": "Carhart 4",
            "Fama-French 5": "Fama-French 5",
            "Public Style Proxy": "Style Proxy",
        }.get(model, model)
        out.append(
            {
                "model": display,
                "annualized_alpha": _float(row.get("AnnualizedAlpha")),
                "t_stat": _float(row.get("AlphaTStat")),
                "p_value": _float(row.get("AlphaPValue")),
                "r2": _float(row.get("R2")),
                "significant_5pct": str(row.get("Significant5Pct") or "").lower() in {"true", "1", "yes"},
                "interpretation": row.get("Interpretation") or None,
            }
        )
    return out


def _factor_exposures() -> list[dict]:
    rows = _read_csv("factor_proxy_sensitivity")
    out = []
    for row in rows:
        name = str(row.get("Proxy") or "").strip()
        beta = _float(row.get("BetaToProxy"))
        if name and beta is not None:
            out.append({"factor": name, "exposure": beta})
    if out:
        return out
    rows = _read_csv("factor_exposure")
    for row in rows:
        name = str(row.get("Factor") or "").strip()
        val = _float(row.get("PortfolioExposure"))
        if name and val is not None:
            out.append({"factor": name, "exposure": val})
    return out


def _stress() -> list[dict]:
    rows = _read_csv("stress_tests")
    out = []
    for row in rows:
        if str(row.get("Ticker") or "").strip().upper() != "PORTFOLIO":
            continue
        scenario = str(row.get("Scenario") or "").strip()
        value = _float(row.get("EstimatedHoldingReturn"))
        if scenario and value is not None:
            out.append({"scenario": scenario, "estimated_return": value})
    return out


def _timeseries() -> list[dict]:
    rows = _read_csv("portfolio_timeseries")
    out = []
    for row in rows:
        date = str(row.get("Date") or "").strip()
        p = _float(row.get("PortfolioGrowth"))
        b = _float(row.get("BenchmarkGrowth"))
        if date and p is not None:
            out.append({"date": date, "portfolio_growth": p, "benchmark_growth": b})
    # Weekly/downsampled public snapshot keeps the repo small while preserving the real path.
    if len(out) > 800:
        step = max(1, len(out) // 500)
        sampled = out[::step]
        if sampled[-1] != out[-1]:
            sampled.append(out[-1])
        out = sampled
    return out


def main():
    summary_path = OUT / "summary.json"
    if not summary_path.exists():
        raise SystemExit(
            "No portfolio outputs found. Run institutional_research/run_research.py first."
        )
    summary = json.loads(summary_path.read_text(encoding="utf-8"))
    portfolio = summary.get("portfolio", {})
    snapshot = {
        "snapshot_type": "sanitized_real_portfolio_analytics",
        "generated_utc": datetime.now(timezone.utc).isoformat(),
        "privacy_note": (
            "Aggregate analytics are derived from the real portfolio. Holdings are anonymized; "
            "tickers, company names, shares, cost basis, market value and transaction data are excluded."
        ),
        "metrics": _safe_metric_map(portfolio),
        "holdings": _anonymous_holdings(),
        "alpha": _alpha(),
        "factors": _factor_exposures(),
        "stress": _stress(),
        "timeseries": _timeseries(),
    }
    DEST.parent.mkdir(parents=True, exist_ok=True)
    DEST.write_text(json.dumps(snapshot, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Sanitized real portfolio snapshot written to: {DEST}")
    print(f"Anonymous holdings exported: {len(snapshot['holdings'])}")
    print(f"Alpha models exported: {len(snapshot['alpha'])}")


if __name__ == "__main__":
    main()
