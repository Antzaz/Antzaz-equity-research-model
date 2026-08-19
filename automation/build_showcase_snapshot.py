from __future__ import annotations

"""Build the recruiter-safe public portfolio snapshot from the latest private outputs.

The public snapshot exposes named holdings, aggregate portfolio analytics, public market
characteristics and recruiter-safe investment theses. Sensitive position economics remain
private: tickers, share counts, average cost, market value, unrealized P&L, transactions,
private notes, credentials and private workbooks are never serialized.
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


def _text(value):
    value = str(value or "").strip()
    return value or None


def _safe_metric_map(portfolio: dict) -> dict:
    allowed = [
        "annualized_return",
        "benchmark_annualized_return",
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
        "benchmark_correlation",
        "up_capture",
        "down_capture",
        "top_1_weight",
        "top_3_weight",
        "top_5_weight",
        "effective_number_of_holdings",
        "effective_number_of_sectors",
        "largest_risk_contribution",
        "annualized_alpha",
        "alpha_t_stat",
        "alpha_p_value",
        "alpha_r_squared",
        "active_share",
    ]
    return {k: portfolio.get(k) for k in allowed if portfolio.get(k) is not None}


def _holding_rows() -> list[dict]:
    rows = _read_csv("holdings_analysis")
    clean: list[dict] = []
    for row in rows:
        ticker = str(row.get("Ticker") or "").strip().upper()
        company = _text(row.get("Company"))
        weight = _float(row.get("Weight"))
        if weight is None:
            continue
        clean.append(
            {
                "ticker": ticker or None,
                "company": company or ticker or "Portfolio holding",
                "sector": _text(row.get("Sector")),
                "industry": _text(row.get("Industry")),
                "country": _text(row.get("Country")),
                "currency": _text(row.get("Currency")),
                "weight": weight,
                "risk_contribution": _float(row.get("RiskContributionPct")),
                "forward_pe": _float(row.get("ForwardPE")),
                "revenue_growth": _float(row.get("RevenueGrowth")),
                "operating_margin": _float(row.get("OperatingMargin")),
                "roe": _float(row.get("ROE")),
            }
        )
    clean.sort(key=lambda x: x["weight"], reverse=True)
    return clean


def _public_holdings(holding_rows: list[dict]) -> list[dict]:
    allowed = [
        "company",
        "sector",
        "industry",
        "country",
        "currency",
        "weight",
        "risk_contribution",
        "forward_pe",
        "revenue_growth",
        "operating_margin",
        "roe",
    ]
    return [
        {k: row.get(k) for k in allowed if row.get(k) not in (None, "")}
        for row in holding_rows
    ]


def _weighted_characteristics(holding_rows: list[dict]) -> dict:
    fields = {
        "forward_pe": "weighted_forward_pe",
        "revenue_growth": "weighted_revenue_growth",
        "operating_margin": "weighted_operating_margin",
        "roe": "weighted_roe",
    }
    out: dict[str, float] = {}
    for field, target in fields.items():
        usable = [r for r in holding_rows if r.get(field) is not None and r.get("weight") is not None]
        denom = sum(float(r["weight"]) for r in usable)
        if denom > 0:
            out[target] = sum(float(r["weight"]) * float(r[field]) for r in usable) / denom
    return out


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
            "industry": holding.get("industry"),
            "country": holding.get("country"),
            "weight": holding.get("weight"),
            "forward_pe": holding.get("forward_pe"),
            "revenue_growth": holding.get("revenue_growth"),
            "operating_margin": holding.get("operating_margin"),
            "roe": holding.get("roe"),
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


def _lookup_by_ticker(holding_rows: list[dict]) -> dict[str, dict]:
    return {
        str(row.get("ticker") or "").upper(): row
        for row in holding_rows
        if row.get("ticker")
    }


def _attribution(holding_rows: list[dict]) -> list[dict]:
    lookup = _lookup_by_ticker(holding_rows)
    out = []
    for row in _read_csv("return_attribution"):
        holding = lookup.get(str(row.get("Ticker") or "").strip().upper())
        if not holding:
            continue
        contribution = _float(row.get("StaticArithmeticContribution"))
        if contribution is None:
            continue
        out.append(
            {
                "company": holding["company"],
                "sector": holding.get("sector"),
                "weight": _float(row.get("Weight")) or holding.get("weight"),
                "asset_total_return": _float(row.get("AssetTotalReturn")),
                "contribution": contribution,
                "contribution_share_abs": _float(row.get("ContributionShareAbs")),
            }
        )
    out.sort(key=lambda x: x.get("contribution", 0), reverse=True)
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


def _historical_stress() -> list[dict]:
    out = []
    for row in _read_csv("historical_stress_windows"):
        scenario = _text(row.get("Scenario"))
        if not scenario:
            continue
        out.append(
            {
                "scenario": scenario,
                "start": _text(row.get("Start")),
                "end": _text(row.get("End")),
                "portfolio_return": _float(row.get("PortfolioReturn")),
                "benchmark_return": _float(row.get("BenchmarkReturn")),
                "active_return": _float(row.get("ActiveReturn")),
            }
        )
    return out


def _rolling_risk() -> list[dict]:
    out = []
    for row in _read_csv("rolling_risk"):
        window = _text(row.get("Window"))
        if not window:
            continue
        out.append(
            {
                "window": window,
                "latest_return": _float(row.get("LatestReturn")),
                "latest_volatility": _float(row.get("LatestVolatility")),
                "latest_tracking_error": _float(row.get("LatestTrackingError")),
                "worst_rolling_return": _float(row.get("WorstRollingReturn")),
                "best_rolling_return": _float(row.get("BestRollingReturn")),
            }
        )
    return out


def _reverse_dcf(holding_rows: list[dict]) -> list[dict]:
    lookup = _lookup_by_ticker(holding_rows)
    out = []
    for row in _read_csv("reverse_dcf"):
        holding = lookup.get(str(row.get("Ticker") or "").strip().upper())
        if not holding:
            continue
        out.append(
            {
                "company": holding["company"],
                "implied_annual_fcf_growth": _float(row.get("ImpliedAnnualFCFGrowth")),
                "wacc": _float(row.get("WACC")),
                "terminal_growth": _float(row.get("TerminalGrowth")),
                "forecast_years": _float(row.get("ForecastYears")),
                "status": _text(row.get("Status")),
            }
        )
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
    philosophy = _portfolio_philosophy(thesis_payload)
    timeseries = _timeseries()

    analysis_start = timeseries[0]["date"] if timeseries else None
    analysis_end = timeseries[-1]["date"] if timeseries else None
    benchmark_code = portfolio.get("benchmark") or "SPY"
    benchmark_display = philosophy.get("benchmark") or benchmark_code

    snapshot = {
        "snapshot_type": "sanitized_real_portfolio_analytics",
        "generated_utc": datetime.now(timezone.utc).isoformat(),
        "privacy_note": (
            "Company names, weights, public market characteristics, aggregate analytics and "
            "recruiter-safe investment theses are public. Tickers, shares, cost basis, market "
            "value, unrealized P&L, transactions, private notes, credentials and private "
            "workbooks are excluded."
        ),
        "data_note": (
            "Portfolio analytics are generated from the latest production research run. "
            "The historical series is a current-weight research model rather than a "
            "transaction-weighted realized client track record."
        ),
        "metadata": {
            "benchmark": benchmark_display,
            "benchmark_proxy": benchmark_code,
            "analysis_start": analysis_start,
            "analysis_end": analysis_end,
            "track_record_type": "Current-weight historical research model",
            "performance_basis": (
                "Current portfolio weights are applied across available adjusted price history; "
                "figures are research diagnostics, not realized client performance."
            ),
            "currency_basis": (
                "Listing-currency adjusted price returns; foreign-exchange effects are not "
                "normalized to a single reporting currency in the current model."
            ),
        },
        "metrics": _safe_metric_map(portfolio),
        "portfolio_characteristics": _weighted_characteristics(holding_rows),
        "holdings": _public_holdings(holding_rows),
        "portfolio_philosophy": philosophy,
        "theses": _public_theses(thesis_payload, holding_rows),
        "attribution": _attribution(holding_rows),
        "alpha": _alpha(),
        "factors": _factor_exposures(),
        "stress": _stress(),
        "historical_stress": _historical_stress(),
        "rolling_risk": _rolling_risk(),
        "reverse_dcf": _reverse_dcf(holding_rows),
        "timeseries": timeseries,
    }
    DEST.parent.mkdir(parents=True, exist_ok=True)
    DEST.write_text(json.dumps(snapshot, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Recruiter-safe real portfolio snapshot written to: {DEST}")
    print(f"Named holdings exported: {len(snapshot['holdings'])}")
    print(f"Company theses exported: {len(snapshot['theses'])}")
    print(f"Attribution rows exported: {len(snapshot['attribution'])}")
    print(f"Alpha models exported: {len(snapshot['alpha'])}")
    print(f"Reverse-DCF rows exported: {len(snapshot['reverse_dcf'])}")


if __name__ == "__main__":
    main()
