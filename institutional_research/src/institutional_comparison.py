from __future__ import annotations

from pathlib import Path
from typing import Dict, Iterable

import numpy as np
import pandas as pd


USER_THESIS_COLUMNS = [
    "Ticker", "Metric", "UserValue", "Unit", "AsOfDate", "Notes",
]

INSTITUTIONAL_VIEW_COLUMNS = [
    "Ticker", "Institution", "Metric", "InstitutionValue", "Unit", "AsOfDate",
    "SourceType", "Source", "DisclosureStatus", "Notes",
]

THESIS_RISK_COLUMNS = [
    "Ticker", "Thesis", "KeyRisk", "FalsificationCondition", "MonitoringMetric",
    "Threshold", "ReviewDate", "Status", "Notes",
]


def _empty(columns: Iterable[str]) -> pd.DataFrame:
    return pd.DataFrame(columns=list(columns))


def _safe_num(value):
    try:
        if value is None:
            return np.nan
        out = float(value)
        return out if np.isfinite(out) else np.nan
    except (TypeError, ValueError):
        return np.nan


def _ratio(value):
    """Normalize a ratio that may arrive as a percent-like number."""
    value = _safe_num(value)
    if pd.isna(value):
        return np.nan
    if abs(value) > 10:
        return value / 100.0
    return value


def _score_high(value, strong, acceptable, weak):
    value = _safe_num(value)
    if pd.isna(value):
        return np.nan
    if value >= strong:
        return 1.0
    if value >= acceptable:
        return 0.75
    if value >= weak:
        return 0.50
    return 0.0


def _score_low(value, strong, acceptable, weak):
    value = _safe_num(value)
    if pd.isna(value):
        return np.nan
    if value <= strong:
        return 1.0
    if value <= acceptable:
        return 0.75
    if value <= weak:
        return 0.50
    return 0.0


def build_buffett_style_scorecard(info: Dict[str, dict]) -> pd.DataFrame:
    """
    Public-data quality/value screen inspired by themes Berkshire/Buffett has
    discussed publicly. It is NOT Berkshire Hathaway's private valuation model,
    target price, or a claim about Buffett's actual investment decision.
    """
    rows = []

    for ticker, raw in info.items():
        d = raw or {}

        revenue = _safe_num(d.get("totalRevenue"))
        fcf = _safe_num(d.get("freeCashflow"))
        market_cap = _safe_num(d.get("marketCap"))
        net_income = _safe_num(d.get("netIncomeToCommon", d.get("netIncome")))
        cash = _safe_num(d.get("totalCash"))
        debt = _safe_num(d.get("totalDebt"))
        ebitda = _safe_num(d.get("ebitda"))

        roe = _ratio(d.get("returnOnEquity"))
        operating_margin = _ratio(d.get("operatingMargins"))
        revenue_growth = _ratio(d.get("revenueGrowth"))
        earnings_growth = _ratio(d.get("earningsGrowth"))
        debt_to_equity = _ratio(d.get("debtToEquity"))

        fcf_margin = (
            fcf / revenue
            if pd.notna(fcf) and pd.notna(revenue) and revenue != 0
            else np.nan
        )
        fcf_yield = (
            fcf / market_cap
            if pd.notna(fcf) and pd.notna(market_cap) and market_cap > 0
            else np.nan
        )
        cash_conversion = (
            fcf / net_income
            if pd.notna(fcf) and pd.notna(net_income) and net_income > 0
            else np.nan
        )
        net_debt = debt - cash if pd.notna(debt) and pd.notna(cash) else np.nan
        net_debt_ebitda = (
            net_debt / ebitda
            if pd.notna(net_debt) and pd.notna(ebitda) and ebitda > 0
            else np.nan
        )
        ev_ebitda = _safe_num(d.get("enterpriseToEbitda"))

        metric_scores = {
            "ROE": _score_high(roe, 0.20, 0.15, 0.10),
            "OperatingMargin": _score_high(operating_margin, 0.20, 0.12, 0.07),
            "FCFMargin": _score_high(fcf_margin, 0.15, 0.10, 0.05),
            "CashConversion": _score_high(cash_conversion, 1.00, 0.80, 0.60),
            "DebtToEquity": _score_low(debt_to_equity, 0.50, 1.00, 2.00),
            "NetDebtEBITDA": _score_low(net_debt_ebitda, 1.00, 2.00, 3.00),
            "RevenueGrowth": _score_high(revenue_growth, 0.10, 0.05, 0.00),
            "EarningsGrowth": _score_high(earnings_growth, 0.10, 0.05, 0.00),
            "FCFYield": _score_high(fcf_yield, 0.06, 0.04, 0.02),
            "EVEBITDA": _score_low(ev_ebitda, 12.0, 18.0, 25.0),
        }

        available = [v for v in metric_scores.values() if pd.notna(v)]
        coverage = len(available) / len(metric_scores)
        total_score = float(np.mean(available) * 10.0) if available else np.nan

        rows.append({
            "Ticker": str(ticker).upper(),
            "Company": d.get("shortName") or d.get("longName"),
            "ROE": roe,
            "OperatingMargin": operating_margin,
            "FCFMargin": fcf_margin,
            "CashConversionFCFToNetIncome": cash_conversion,
            "DebtToEquity": debt_to_equity,
            "NetDebtEBITDA": net_debt_ebitda,
            "RevenueGrowth": revenue_growth,
            "EarningsGrowth": earnings_growth,
            "FCFYield": fcf_yield,
            "EVEBITDA": ev_ebitda,
            "BuffettStyleScore10": total_score,
            "DataCoveragePct": coverage,
            "FrameworkNote": (
                "Public-data Buffett-style screen; not Berkshire Hathaway's "
                "private valuation model or a disclosed Buffett score."
            ),
        })

    return pd.DataFrame(rows)


def load_user_theses(path: str | Path) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        return _empty(USER_THESIS_COLUMNS)

    df = pd.read_csv(path)
    missing = [c for c in USER_THESIS_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"{path.name} missing required columns: {', '.join(missing)}")

    df = df[USER_THESIS_COLUMNS].copy()
    df["Ticker"] = df["Ticker"].astype(str).str.upper().str.strip()
    df["Metric"] = df["Metric"].astype(str).str.strip()
    df["UserValue"] = pd.to_numeric(df["UserValue"], errors="coerce")
    return df[df["Ticker"].ne("") & df["Metric"].ne("")]


def load_institutional_views(path: str | Path) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        return _empty(INSTITUTIONAL_VIEW_COLUMNS)

    df = pd.read_csv(path)
    missing = [c for c in INSTITUTIONAL_VIEW_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"{path.name} missing required columns: {', '.join(missing)}")

    df = df[INSTITUTIONAL_VIEW_COLUMNS].copy()
    df["Ticker"] = df["Ticker"].astype(str).str.upper().str.strip()
    df["Institution"] = df["Institution"].astype(str).str.strip()
    df["Metric"] = df["Metric"].astype(str).str.strip()
    df["InstitutionValue"] = pd.to_numeric(df["InstitutionValue"], errors="coerce")
    return df[
        df["Ticker"].ne("") & df["Institution"].ne("") & df["Metric"].ne("")
    ]


def compare_institutional_views(
    user_theses: pd.DataFrame,
    institutional_views: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    comparison_columns = [
        "Ticker", "Metric", "UserValue", "Institution", "InstitutionValue",
        "Unit", "Delta", "DeltaPctVsInstitution", "AsOfDateInstitution",
        "SourceType", "Source", "DisclosureStatus", "InstitutionNotes", "UserNotes",
    ]
    consensus_columns = [
        "Ticker", "Metric", "Unit", "UserValue", "InstitutionCount",
        "InstitutionMean", "InstitutionMedian", "InstitutionMin", "InstitutionMax",
        "DeltaVsMean", "DeltaPctVsMean",
    ]

    if user_theses.empty or institutional_views.empty:
        return _empty(comparison_columns), _empty(consensus_columns)

    user = user_theses.rename(
        columns={"AsOfDate": "AsOfDateUser", "Notes": "UserNotes"}
    ).copy()
    inst = institutional_views.rename(
        columns={"AsOfDate": "AsOfDateInstitution", "Notes": "InstitutionNotes"}
    ).copy()

    merged = user.merge(
        inst,
        on=["Ticker", "Metric"],
        how="inner",
        suffixes=("_User", "_Institution"),
    )
    if merged.empty:
        return _empty(comparison_columns), _empty(consensus_columns)

    merged["Unit"] = merged["Unit_User"].where(
        merged["Unit_User"].notna() & merged["Unit_User"].astype(str).ne(""),
        merged["Unit_Institution"],
    )
    merged["Delta"] = merged["UserValue"] - merged["InstitutionValue"]
    merged["DeltaPctVsInstitution"] = np.where(
        merged["InstitutionValue"].abs() > 1e-12,
        merged["Delta"] / merged["InstitutionValue"].abs(),
        np.nan,
    )
    comparison = merged[comparison_columns].copy()

    grouped = (
        institutional_views.groupby(["Ticker", "Metric", "Unit"], dropna=False)["InstitutionValue"]
        .agg(
            InstitutionCount="count",
            InstitutionMean="mean",
            InstitutionMedian="median",
            InstitutionMin="min",
            InstitutionMax="max",
        )
        .reset_index()
    )
    consensus = user_theses.merge(grouped, on=["Ticker", "Metric", "Unit"], how="inner")
    if consensus.empty:
        return comparison, _empty(consensus_columns)

    consensus["DeltaVsMean"] = consensus["UserValue"] - consensus["InstitutionMean"]
    consensus["DeltaPctVsMean"] = np.where(
        consensus["InstitutionMean"].abs() > 1e-12,
        consensus["DeltaVsMean"] / consensus["InstitutionMean"].abs(),
        np.nan,
    )
    return comparison, consensus[consensus_columns].copy()


def load_thesis_risks(path: str | Path) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        return _empty(THESIS_RISK_COLUMNS)

    df = pd.read_csv(path)
    missing = [c for c in THESIS_RISK_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"{path.name} missing required columns: {', '.join(missing)}")

    df = df[THESIS_RISK_COLUMNS].copy()
    df["Ticker"] = df["Ticker"].astype(str).str.upper().str.strip()
    return df[df["Ticker"].ne("")]
