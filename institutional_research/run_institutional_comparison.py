from __future__ import annotations

import json
from pathlib import Path
import shutil

import pandas as pd

from src.data import fetch_info
from src.institutional_comparison import (
    build_buffett_style_scorecard,
    compare_institutional_views,
    load_institutional_views,
    load_thesis_risks,
    load_user_theses,
)


BASE = Path(__file__).resolve().parent
OUT = BASE / "outputs" / "institutional_comparison"


def _ensure_private_copy(template_name: str, private_name: str) -> Path:
    private_path = BASE / private_name
    template_path = BASE / template_name
    if not private_path.exists() and template_path.exists():
        shutil.copyfile(template_path, private_path)
    return private_path


def _expected_return_inputs() -> pd.DataFrame:
    path = BASE / "expected_returns.csv"
    columns = ["Ticker", "Metric", "UserValue", "Unit", "AsOfDate", "Notes"]
    if not path.exists():
        return pd.DataFrame(columns=columns)

    df = pd.read_csv(path)
    if "Ticker" not in df.columns or "ExpectedReturn" not in df.columns:
        return pd.DataFrame(columns=columns)

    out = pd.DataFrame({
        "Ticker": df["Ticker"].astype(str).str.upper().str.strip(),
        "Metric": "ExpectedReturn",
        "UserValue": pd.to_numeric(df["ExpectedReturn"], errors="coerce"),
        "Unit": "decimal",
        "AsOfDate": "",
        "Notes": "Imported automatically from expected_returns.csv",
    })
    return out[out["Ticker"].ne("")]


def main():
    thesis_path = _ensure_private_copy(
        "company_thesis_template.csv",
        "company_thesis.csv",
    )
    institutional_path = _ensure_private_copy(
        "institutional_views_template.csv",
        "institutional_views.csv",
    )
    risks_path = _ensure_private_copy(
        "thesis_risks_template.csv",
        "thesis_risks.csv",
    )

    user_theses = load_user_theses(thesis_path)
    expected = _expected_return_inputs()
    if not expected.empty:
        user_theses = pd.concat([user_theses, expected], ignore_index=True)
        user_theses = user_theses.drop_duplicates(
            subset=["Ticker", "Metric"], keep="first"
        )

    institutional_views = load_institutional_views(institutional_path)
    thesis_risks = load_thesis_risks(risks_path)

    tickers = sorted(set(
        user_theses.get("Ticker", pd.Series(dtype=str)).dropna().tolist()
        + institutional_views.get("Ticker", pd.Series(dtype=str)).dropna().tolist()
        + thesis_risks.get("Ticker", pd.Series(dtype=str)).dropna().tolist()
    ))

    if not tickers:
        portfolio_path = BASE / "portfolio.csv"
        if portfolio_path.exists():
            portfolio = pd.read_csv(portfolio_path)
            if "Ticker" in portfolio.columns:
                tickers = sorted(
                    set(
                        portfolio["Ticker"]
                        .dropna()
                        .astype(str)
                        .str.upper()
                        .str.strip()
                        .tolist()
                    )
                )

    info = fetch_info(tickers) if tickers else {}
    buffett_scorecard = build_buffett_style_scorecard(info)
    comparison, consensus = compare_institutional_views(
        user_theses,
        institutional_views,
    )

    OUT.mkdir(parents=True, exist_ok=True)

    tables = {
        "buffett_style_scorecard.csv": buffett_scorecard,
        "user_thesis_inputs.csv": user_theses,
        "institutional_views.csv": institutional_views,
        "institutional_comparison.csv": comparison,
        "institutional_consensus.csv": consensus,
        "thesis_risks.csv": thesis_risks,
    }
    for filename, df in tables.items():
        df.to_csv(OUT / filename, index=False)

    summary = {
        "tickers_analyzed": tickers,
        "institutional_rows": int(len(institutional_views)),
        "comparison_rows": int(len(comparison)),
        "consensus_rows": int(len(consensus)),
        "thesis_risk_rows": int(len(thesis_risks)),
        "important_notes": [
            "The Buffett-style score is a public-data screening framework, not Berkshire Hathaway's private valuation model.",
            "Institutional comparisons are only as reliable as the cited public disclosures entered in institutional_views.csv.",
            "A 13F shows reportable US equity holdings with a reporting lag; it does not reveal a manager's intrinsic value estimate.",
            "Fund letters, annual reports and investor presentations may disclose thesis elements but usually not complete internal models.",
        ],
    }
    with open(OUT / "summary.json", "w", encoding="utf-8") as f:
        json.dump(summary, f, indent=2)

    print("Institutional comparison complete.")
    print(f"Outputs: {OUT}")
    if institutional_views.empty:
        print(
            "Next: add public institutional disclosures to institutional_views.csv "
            "and rerun this script."
        )
    if user_theses.empty:
        print(
            "Next: add your company assumptions to company_thesis.csv "
            "(or expected_returns.csv) and rerun this script."
        )


if __name__ == "__main__":
    main()
