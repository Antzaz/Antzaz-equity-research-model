from __future__ import annotations

"""Workbook output for the AI Growth Forecast layer.

Kept separate from the forecasting logic so workbook-format changes cannot affect model
training or inference.
"""

from datetime import datetime, timezone
from pathlib import Path
from typing import Any
import math

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill


def _finite(value: Any) -> float | None:
    try:
        x = float(value)
        return x if math.isfinite(x) else None
    except Exception:
        return None


def _fmt_pct(value: Any) -> str:
    x = _finite(value)
    return "N/M" if x is None else f"{x:.1%}"


def write_ai_growth_sheet(
    workbook_path: Path,
    ticker: str,
    payload: dict[str, Any],
    *,
    sheet_name: str = "AI Growth Forecast",
) -> None:
    path = Path(workbook_path)
    wb = load_workbook(path)
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A5"
    for col, width in {"A": 31, "B": 18, "C": 20, "D": 18, "E": 50, "F": 24}.items():
        ws.column_dimensions[col].width = width

    navy = "17365D"; blue = "2F75B5"; white = "FFFFFF"; gold = "FFF2CC"
    green = "008000"; red = "C00000"; grey = "666666"

    ws.merge_cells("A1:F2")
    ws["A1"] = f"{ticker} — AI Growth Forecast"
    ws["A1"].fill = PatternFill("solid", fgColor=navy)
    ws["A1"].font = Font(bold=True, color=white, size=18)
    ws["A1"].alignment = Alignment(vertical="center")
    ws.merge_cells("A3:F3")
    ws["A3"] = (
        "LLM/deterministic AI evidence extraction + LightGBM fundamental growth + "
        "reverse-DCF expectations gap. AI is a bounded evidence overlay until enough "
        "dated AI observations exist for supervised training."
    )
    ws["A3"].font = Font(italic=True, color=grey)
    ws["A3"].alignment = Alignment(wrap_text=True)
    ws.row_dimensions[3].height = 34
    ws["F4"] = datetime.now(timezone.utc).isoformat(timespec="seconds")
    ws["F4"].font = Font(color=grey, italic=True, size=9)

    def section(row: int, title: str) -> None:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        c = ws.cell(row, 1, title)
        c.fill = PatternFill("solid", fgColor=navy)
        c.font = Font(bold=True, color=white)

    def header(row: int, labels: list[str]) -> None:
        for col, label in enumerate(labels, 1):
            c = ws.cell(row, col, label)
            c.fill = PatternFill("solid", fgColor=blue)
            c.font = Font(bold=True, color=white)
            c.alignment = Alignment(horizontal="center", wrap_text=True)

    signals = payload.get("ai_signals") or {}
    section(5, "AI Evidence Signals — normalized 0 to 1")
    header(6, ["Signal", "Score", "Interpretation", "Extraction", "Evidence", "Notes"])
    signal_rows = [
        ("AI demand", "demand_score", "Higher = stronger demand/backlog/workload evidence"),
        ("AI monetization", "monetization_score", "Higher = stronger revenue/pricing/paid-usage evidence"),
        ("AI adoption", "adoption_score", "Higher = stronger users/customers/deployment evidence"),
        ("AI efficiency", "efficiency_score", "Higher = stronger margins/unit-economics/productivity evidence"),
        ("AI capex burden", "capex_burden_score", "Higher = heavier cash/capital burden"),
        ("AI risk", "risk_score", "Higher = greater competition/capacity/regulatory/execution risk"),
    ]
    for r, (label, field, note) in enumerate(signal_rows, 7):
        ws.cell(r, 1, label)
        ws.cell(r, 2, _finite(signals.get(field)))
        ws.cell(r, 2).number_format = "0%"
        ws.cell(r, 3, note)
        ws.cell(r, 4, signals.get("extraction_mode"))
        ws.cell(r, 5, signals.get("evidence_count"))
        ws.cell(r, 6, signals.get("summary") if r == 7 else "")
        for c in range(1, 7):
            ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[7].height = 44

    section(14, "Growth Forecast & Market Expectations")
    header(15, ["Metric", "Fundamental ML", "AI adjustment", "AI-adjusted", "Market implied", "Gap / validation"])
    revenue = payload.get("revenue_forecast") or {}
    fcf = payload.get("fcf_forecast") or {}
    adj = payload.get("ai_adjustments") or {}
    reverse = payload.get("reverse_dcf") or {}
    gap = payload.get("expectations_gap") or {}
    rows = [
        (
            "Next FY revenue growth", revenue.get("prediction"), adj.get("revenue_growth_adjustment"),
            payload.get("ai_adjusted_revenue_growth"), None,
            f"LightGBM holdout MAE {_fmt_pct((revenue.get('metrics') or {}).get('time_purged_holdout_mae'))}; "
            f"Elastic Net {_fmt_pct((revenue.get('metrics') or {}).get('elastic_net_holdout_mae'))}",
        ),
        (
            "Next FY FCF growth", fcf.get("prediction"), adj.get("fcf_growth_adjustment"),
            payload.get("ai_adjusted_fcf_growth"), reverse.get("implied_annual_fcf_growth"),
            gap.get("fcf_growth_gap"),
        ),
    ]
    for r, item in enumerate(rows, 16):
        for c, value in enumerate(item, 1):
            ws.cell(r, c, value)
        for c in (2, 3, 4, 5):
            ws.cell(r, c).number_format = "0.0%"
        if r == 17 and _finite(ws.cell(r, 6).value) is not None:
            ws.cell(r, 6).number_format = "0.0%"
        ws.cell(r, 4).fill = PatternFill("solid", fgColor=gold)
        ws.cell(r, 4).font = Font(bold=True)

    ws.merge_cells("A19:F19")
    ws["A19"] = gap.get("interpretation") or "Expectations gap unavailable."
    gap_value = _finite(gap.get("fcf_growth_gap"))
    ws["A19"].font = Font(bold=True, color=green if gap_value is not None and gap_value >= 0 else red)
    ws["A19"].alignment = Alignment(wrap_text=True)

    section(21, "LightGBM Explainability — current forecast drivers")
    header(22, ["Target", "Feature", "Direction", "SHAP / importance", "Current value", "Model confidence"])
    rr = 23
    for label, forecast in (("Revenue", revenue), ("FCF", fcf)):
        for driver in (forecast.get("drivers") or [])[:6]:
            ws.cell(rr, 1, label)
            ws.cell(rr, 2, driver.get("feature"))
            ws.cell(rr, 3, driver.get("direction"))
            ws.cell(rr, 4, driver.get("shap_value", driver.get("importance")))
            ws.cell(rr, 5, driver.get("current_value"))
            ws.cell(rr, 6, forecast.get("confidence"))
            rr += 1
    if rr == 23:
        ws.cell(rr, 1, "No explainability output; growth history may still be insufficient.")
        rr += 1

    section(rr + 1, "Evidence & Governance")
    header(rr + 2, ["Evidence", "", "", "", "", ""])
    cursor = rr + 3
    for item in (signals.get("evidence") or [])[:8]:
        ws.merge_cells(start_row=cursor, start_column=1, end_row=cursor, end_column=6)
        ws.cell(cursor, 1, str(item))
        ws.cell(cursor, 1).alignment = Alignment(wrap_text=True)
        cursor += 1
    for note in [
        "LightGBM is benchmarked against Elastic Net on a chronological holdout; lower holdout MAE is better.",
        "If LightGBM fails to match the Elastic Net benchmark, model confidence is downgraded.",
        "The AI overlay is bounded and confidence-scaled because current AI KPI history is sparse relative to financial history.",
        "The LLM extracts and scores evidence only; it does not directly set the target price or DCF assumptions.",
        "Reverse DCF is a simplified FCF-growth hurdle and is not meaningful for every business model.",
        "No trades are executed and no private portfolio economics are exposed by this sheet.",
    ]:
        ws.merge_cells(start_row=cursor, start_column=1, end_row=cursor, end_column=6)
        ws.cell(cursor, 1, "• " + note)
        ws.cell(cursor, 1).font = Font(color=grey)
        ws.cell(cursor, 1).alignment = Alignment(wrap_text=True)
        cursor += 1

    wb.save(path)
