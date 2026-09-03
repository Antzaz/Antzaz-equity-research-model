from __future__ import annotations

"""Run the AI-growth forecasting layer for one equity-research workbook.

Examples:
    python ai_growth_research.py GOOGL
    python ai_growth_research.py GOOGL --llm
    python ai_growth_research.py GOOGL --llm --model gpt-5.6-terra

The engine never changes the authoritative DCF assumptions. It writes a separate
"AI Growth Forecast" research sheet and a JSON audit artifact.
"""

import argparse
from datetime import datetime
import json
from pathlib import Path
import re

from machine_learning.ai_growth import (
    LightGBMGrowthForecaster,
    ai_adjustments,
    apply_ai_overlay,
    deterministic_ai_signals,
    expectations_gap,
    growth_training_frame_from_history,
    llm_ai_signals,
    load_kpi_evidence,
    reverse_dcf_from_workbook,
    workbook_current_growth_features,
    write_ai_growth_sheet,
)
from machine_learning.data import latest_workbook

BASE = Path(__file__).resolve().parent
DEFAULT_DB = BASE / "ml_data" / "ml_history.sqlite"
RUNS = BASE / "ml_runs"


def ticker_type(raw: str) -> str:
    ticker = str(raw or "").upper().strip()
    if not re.fullmatch(r"[A-Z0-9.\-]{1,10}", ticker):
        raise argparse.ArgumentTypeError("Enter a ticker such as GOOGL, MSFT, NVDA, TSM, or SIE.DE")
    return ticker


def parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="LLM evidence extraction + LightGBM fundamental growth + reverse-DCF expectations gap"
    )
    p.add_argument("ticker", type=ticker_type)
    p.add_argument("--workbook", help="Workbook path; defaults to the newest generated model")
    p.add_argument("--history-db", default=str(DEFAULT_DB), help="Persistent point-in-time ML SQLite database")
    p.add_argument("--llm", action="store_true", help="Use OpenAI structured extraction when OPENAI_API_KEY is set")
    p.add_argument("--model", default="gpt-5.6-luna", help="OpenAI extraction model override")
    p.add_argument("--no-workbook-write", action="store_true")
    return p


def main() -> int:
    args = parser().parse_args()
    ticker = args.ticker
    workbook = Path(args.workbook).resolve() if args.workbook else latest_workbook(BASE, ticker)
    history_db = Path(args.history_db).resolve()

    rows, corpus = load_kpi_evidence(BASE, ticker)
    signals = (
        llm_ai_signals(rows, corpus, model=args.model)
        if args.llm
        else deterministic_ai_signals(rows, corpus)
    )
    print(
        f"[ai-growth] evidence: {signals.evidence_count} KPI row(s); "
        f"extraction={signals.extraction_mode}; confidence={signals.confidence:.0%}"
    )

    frame = growth_training_frame_from_history(history_db)
    current = workbook_current_growth_features(workbook, ticker)
    model = LightGBMGrowthForecaster()
    revenue = model.fit_predict(
        frame,
        current,
        target_col="target_next_revenue_growth",
        target_label="Next FY revenue growth",
    )
    fcf = model.fit_predict(
        frame,
        current,
        target_col="target_next_fcf_growth",
        target_label="Next FY FCF growth",
    )

    adjustments = ai_adjustments(signals)
    ai_revenue = apply_ai_overlay(revenue, adjustments["revenue_growth_adjustment"])
    ai_fcf = apply_ai_overlay(fcf, adjustments["fcf_growth_adjustment"])
    reverse_dcf = reverse_dcf_from_workbook(workbook)
    gap = expectations_gap(ai_fcf, reverse_dcf)

    payload = {
        "ticker": ticker,
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "workbook": str(workbook),
        "history_db": str(history_db),
        "architecture": {
            "evidence_layer": signals.extraction_mode,
            "forecast_model": "LightGBM",
            "baseline_model": "ElasticNet",
            "explainability": "SHAP with LightGBM feature-importance fallback",
            "valuation_bridge": "reverse DCF implied FCF growth",
            "governance": (
                "AI evidence is a bounded overlay until sufficient dated AI KPI history exists "
                "to train AI features directly. This layer does not overwrite DCF assumptions."
            ),
        },
        "ai_signals": signals.to_dict(),
        "revenue_forecast": revenue.to_dict(),
        "fcf_forecast": fcf.to_dict(),
        "ai_adjustments": adjustments,
        "ai_adjusted_revenue_growth": ai_revenue,
        "ai_adjusted_fcf_growth": ai_fcf,
        "reverse_dcf": reverse_dcf,
        "expectations_gap": gap,
    }

    stamp = datetime.now().astimezone().strftime("%Y%m%d_%H%M%S")
    run_dir = RUNS / ticker / stamp
    run_dir.mkdir(parents=True, exist_ok=True)
    output = run_dir / "ai_growth_results.json"
    output.write_text(json.dumps(payload, indent=2, ensure_ascii=False, default=str) + "\n", encoding="utf-8")

    if not args.no_workbook_write:
        write_ai_growth_sheet(workbook, ticker, payload)

    def show(label: str, value):
        print(f"[ai-growth] {label}: {'N/M' if value is None else f'{value:.1%}'}")

    show("LightGBM revenue growth", revenue.prediction)
    show("AI-adjusted revenue growth", ai_revenue)
    show("LightGBM FCF growth", fcf.prediction)
    show("AI-adjusted FCF growth", ai_fcf)
    show("reverse-DCF implied FCF growth", reverse_dcf.get("implied_annual_fcf_growth"))
    show("AI expectations gap", gap.get("fcf_growth_gap"))
    print(f"[ai-growth] output: {output}")
    if not args.no_workbook_write:
        print(f"[ai-growth] workbook sheet updated: {workbook} -> AI Growth Forecast")

    # Forecast insufficiency is not treated as a hard pipeline failure: the evidence
    # extraction and reverse-DCF diagnostics remain useful while history accumulates.
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
