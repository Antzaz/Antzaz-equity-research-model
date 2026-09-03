from __future__ import annotations

from pathlib import Path
import sys

import numpy as np
import pandas as pd

# institutional_research can be launched from its own directory. Keep the project-level business
# model registry importable without duplicating sector logic in the portfolio engine.
ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from business_model_registry import get_business_model_policy, reverse_dcf_applicability_message


def equity_value_from_fcf(
    fcf0: float,
    growth: float,
    years: int,
    wacc: float,
    terminal_growth: float,
    cash: float,
    debt: float,
) -> float:
    if wacc <= terminal_growth:
        return np.nan

    fcf = fcf0
    pv = 0.0
    for year in range(1, years + 1):
        fcf *= 1 + growth
        pv += fcf / ((1 + wacc) ** year)

    terminal_value = fcf * (1 + terminal_growth) / (wacc - terminal_growth)
    pv_terminal = terminal_value / ((1 + wacc) ** years)
    return pv + pv_terminal + cash - debt


def solve_implied_growth(
    target_market_cap: float,
    fcf0: float,
    years: int,
    wacc: float,
    terminal_growth: float,
    cash: float,
    debt: float,
    min_growth: float = -0.20,
    max_growth: float = 0.50,
):
    if not all(np.isfinite(x) for x in [target_market_cap, fcf0, cash, debt]):
        return np.nan
    if target_market_cap <= 0 or fcf0 <= 0 or wacc <= terminal_growth:
        return np.nan

    lo, hi = min_growth, max_growth
    vlo = equity_value_from_fcf(fcf0, lo, years, wacc, terminal_growth, cash, debt)
    vhi = equity_value_from_fcf(fcf0, hi, years, wacc, terminal_growth, cash, debt)

    if not np.isfinite(vlo) or not np.isfinite(vhi):
        return np.nan
    if not (vlo <= target_market_cap <= vhi):
        return np.nan

    for _ in range(100):
        mid = (lo + hi) / 2
        value = equity_value_from_fcf(fcf0, mid, years, wacc, terminal_growth, cash, debt)
        if value < target_market_cap:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


def reverse_dcf_table(
    tickers: list[str],
    info: dict[str, dict],
    assumptions: dict,
) -> pd.DataFrame:
    rows = []
    for ticker in tickers:
        d = info.get(ticker) or {}
        policy = get_business_model_policy(
            ticker,
            d.get("sector"),
            d.get("industry"),
            d.get("longName") or d.get("shortName"),
        )
        market_cap = d.get("marketCap")
        fcf = d.get("freeCashflow")
        cash = d.get("totalCash") or 0.0
        debt = d.get("totalDebt") or 0.0

        if not policy.reverse_dcf_allowed:
            rows.append({
                "Ticker": ticker,
                "MarketCap": market_cap,
                "CurrentFCF": fcf,
                "Cash": cash,
                "Debt": debt,
                "WACC": np.nan,
                "TerminalGrowth": np.nan,
                "ForecastYears": assumptions.get("years"),
                "ImpliedAnnualFCFGrowth": np.nan,
                "BusinessModel": policy.key,
                "PrimaryValuation": policy.primary_valuation,
                "Status": reverse_dcf_applicability_message(policy),
            })
            continue

        implied = solve_implied_growth(
            target_market_cap=float(market_cap) if market_cap else np.nan,
            fcf0=float(fcf) if fcf else np.nan,
            years=int(assumptions["years"]),
            wacc=float(assumptions["wacc"]),
            terminal_growth=float(assumptions["terminal_growth"]),
            cash=float(cash),
            debt=float(debt),
            min_growth=float(assumptions.get("min_growth", -0.20)),
            max_growth=float(assumptions.get("max_growth", 0.50)),
        )

        rows.append({
            "Ticker": ticker,
            "MarketCap": market_cap,
            "CurrentFCF": fcf,
            "Cash": cash,
            "Debt": debt,
            "WACC": assumptions["wacc"],
            "TerminalGrowth": assumptions["terminal_growth"],
            "ForecastYears": assumptions["years"],
            "ImpliedAnnualFCFGrowth": implied,
            "BusinessModel": policy.key,
            "PrimaryValuation": policy.primary_valuation,
            "Status": "Solved" if np.isfinite(implied) else "Insufficient data / outside search range",
        })

    return pd.DataFrame(rows)
