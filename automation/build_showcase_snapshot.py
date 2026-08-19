from __future__ import annotations

"""Build the recruiter-safe public portfolio snapshot from the latest private outputs.

The public snapshot intentionally exposes company names, portfolio weights and recruiter-safe
investment-thesis content while continuing to exclude tickers, share counts, average cost,
market value, unrealized P&L, transactions, private notes, credentials and private workbooks.
"""

import csv
import json
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "institutional_research" / "outputs" / "latest"
THESIS_PATH = ROOT / "institutional_research" / "portfolio_thesis_public.json"
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


def _holding_rows() -> list[dict]:
    rows = _read_csv("holdings_analysis")
    clean: list[dict] = []
    for row in rows:
        ticker = str(row.get("Ticker") or "").strip().upper()
        company = str(row.get("Company") or "").strip()
        sector = str(row.get("Sector") or "").strip() or None
        weight = _float(row.get("Weight"))
        risk = _float(row.get("RiskContributionPct"))
        if weight is None:
            continue
        clean.append(
            {
                "ticker": ticker or None,
                "company": company or ticker or "Portfolio holding",
                "sector": sector,
                "weight": weight,
                "risk_contribution": risk,
            }
        )
    clean.sort(key=lambda x: x["weight"], reverse=True)
    return clean


def _public_holdings(holding_rows: list[dict]) -> list[dict]:
    return [
        {
            "company": row["company"],
            "sector": row.get("sector"),
            "weight": row["weight"],
            "risk_contribution": row.get("risk_contribution"),
        }
        for row in holding_rows
    ]


def _load_thesis_payload() -> dict:
    if not THESIS_PATH.exists() or THESIS_PATH.stat().st_size == 0:
        return {"portfolio_philosophy": {}, "company_theses": []}
    try:
        data = json.loads(THESIS_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {"portfolio_philosophy": {}, "company_theses": []}


def _portfolio_philosophy(payload: dict) -> dict:
    philosophy = payload.get("portfolio_philosophy") or {}
    if not isinstance(philosophy, dict):
        return {}
    return {str(k): v for k, v in philosophy.items() if v not in (None, "")}


def _public_theses(payload: dict, holding_rows: list[dict]) -> list[dict]:
    holding_lookup = {
        str(row.get("ticker") or "").upper(): row
        for row in holding_rows
        if row.get("ticker")
    }
    out: list[dict] = []
    for thesis in payload.get("company_theses") or []:
        if not isinstance(thesis, dict):
            continue
        ticker = str(thesis.get("ticker") or "").strip().upper()
        holding = holding_lookup.get(ticker)
        if not holding:
            continue

        public = {
            "company": holding["company"],
            "sector": holding.get("sector"),
            "weight": holding.get("weight"),
            "status": thesis.get("status"),
            "time_horizon": thesis.get("time_horizon"),
            "conviction": _float(thesis.get("conviction")),
            "composite_score": _float(thesis.get("composite_score")),
            "expected_annual_return": _float(thesis.get("expected_annual_return")),
            "investment_thesis": thesis.get("investment_thesis"),
            "why_owned": thesis.get("why_owned"),
            "competitive_advantage": thesis.get("competitive_advantage"),
            "growth_drivers": thesis.get("growth_drivers"),
            "valuation_rationale": thesis.get("valuation_rationale"),
            "catalysts": thesis.get("catalysts"),
            "key_risks": thesis.get("key_risks"),
            "sell_condition": thesis.get("sell_condition"),
            "monitoring_kpi": thesis.get("monitoring_kpi"),
            "review_date": thesis.get("review_date"),
            "public_notes": thesis.get("public_notes"),
            "scores": thesis.get("scores") if isinstance(thesis.get("scores"), dict) else {},
        }
        out.append({k: v for k, v in public.items() if v not in (None, "", {})})
    out.sort(key=lambda x: x.get("weight", 0), reverse=True)
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
                "significant_5pct": str(row.get("Significant5Pct") or "").lower()
                in {"true", "1", "yes"},
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
    holding_rows = _holding_rows()
    thesis_payload = _load_thesis_payload()

    snapshot = {
        "snapshot_type": "sanitized_real_portfolio_analytics",
        "generated_utc": datetime.now(timezone.utc).isoformat(),
        "privacy_note": (
            "Company names, portfolio weights and recruiter-safe investment theses are public. "
            "Tickers, shares, cost basis, market value, unrealized P&L, transactions, private "
            "notes, credentials and private workbooks are excluded."
        ),
        "data_note": (
            "Company names, weights and analytics are generated from the latest production "
            "portfolio run. Investment-thesis text is user-authored and synced from the private "
            "portfolio thesis workbook."
        ),
        "metrics": _safe_metric_map(portfolio),
        "holdings": _public_holdings(holding_rows),
        "portfolio_philosophy": _portfolio_philosophy(thesis_payload),
        "theses": _public_theses(thesis_payload, holding_rows),
        "alpha": _alpha(),
        "factors": _factor_exposures(),
        "stress": _stress(),
        "timeseries": _timeseries(),
    }
    DEST.parent.mkdir(parents=True, exist_ok=True)
    DEST.write_text(json.dumps(snapshot, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Recruiter-safe real portfolio snapshot written to: {DEST}")
    print(f"Named holdings exported: {len(snapshot['holdings'])}")
    print(f"Company theses exported: {len(snapshot['theses'])}")
    print(f"Alpha models exported: {len(snapshot['alpha'])}")


if __name__ == "__main__":
    main()
