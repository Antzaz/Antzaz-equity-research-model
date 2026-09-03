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
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList


def _finite(value: Any) -> float | None:
    try:
        x = float(value)
        return x if math.isfinite(x) else None
    except Exception:
        return None


def _fmt_pct(value: Any) -> str:
    x = _finite(value)
    return "N/M" if x is None else f"{x:.1%}"


def _add_ai_growth_charts(ws, signals: dict[str, Any], revenue: dict[str, Any], fcf: dict[str, Any],
                          payload: dict[str, Any], reverse: dict[str, Any]) -> None:
    """Create decision-oriented charts using hidden helper tables to keep the visible sheet clean."""
    ws.column_dimensions["P"].hidden=True
    ws.column_dimensions["Q"].hidden=True

    signal_rows=[
        ("Demand",_finite(signals.get("demand_score"))),
        ("Monetization",_finite(signals.get("monetization_score"))),
        ("Adoption",_finite(signals.get("adoption_score"))),
        ("Efficiency",_finite(signals.get("efficiency_score"))),
        ("Capex burden",_finite(signals.get("capex_burden_score"))),
        ("Risk",_finite(signals.get("risk_score"))),
    ]
    ws["P2"]="Signal"; ws["Q2"]="Score"
    for i,(name,val) in enumerate(signal_rows,3):
        ws.cell(i,16,name); ws.cell(i,17,val); ws.cell(i,17).number_format="0%"
    ch=BarChart(); ch.type="bar"; ch.style=10
    ch.title="AI evidence scores (burden/risk are adverse)"
    ch.y_axis.title="Signal"; ch.x_axis.title="Score"
    ch.x_axis.scaling.min=0; ch.x_axis.scaling.max=1; ch.x_axis.numFmt="0%"
    ch.height=6.7; ch.width=12.5; ch.legend=None
    ch.add_data(Reference(ws,min_col=17,min_row=2,max_row=8),titles_from_data=True)
    ch.set_categories(Reference(ws,min_col=16,min_row=3,max_row=8))
    ch.dLbls=DataLabelList(); ch.dLbls.showVal=True; ch.dLbls.numFmt="0%"
    ws.add_chart(ch,"H5")

    ws["P11"]="FCF growth"; ws["Q11"]="Value"
    fcf_rows=[
        ("Fundamental ML",_finite(fcf.get("prediction"))),
        ("AI-adjusted",_finite(payload.get("ai_adjusted_fcf_growth"))),
        ("Market implied",_finite(reverse.get("implied_annual_fcf_growth"))),
    ]
    end=11
    for i,(name,val) in enumerate(fcf_rows,12):
        if val is None: continue
        ws.cell(i,16,name); ws.cell(i,17,val); ws.cell(i,17).number_format="0.0%"; end=max(end,i)
    if end>=12:
        ch=BarChart(); ch.type="col"; ch.style=11
        ch.title="FCF growth forecast vs market hurdle"
        ch.y_axis.title="Annual growth"; ch.y_axis.numFmt="0%"
        vals=[v for _,v in fcf_rows if v is not None]
        if vals and min(vals)>=0: ch.y_axis.scaling.min=0
        ch.height=6.7; ch.width=12.5; ch.legend=None
        ch.add_data(Reference(ws,min_col=17,min_row=11,max_row=end),titles_from_data=True)
        ch.set_categories(Reference(ws,min_col=16,min_row=12,max_row=end))
        ch.dLbls=DataLabelList(); ch.dLbls.showVal=True; ch.dLbls.numFmt="0.0%"
        ws.add_chart(ch,"H20")

    driver_rows=[]
    for label,forecast in (("Revenue",revenue),("FCF",fcf)):
        for d in (forecast.get("drivers") or [])[:6]:
            val=_finite(d.get("shap_value",d.get("importance")))
            if val is None: continue
            driver_rows.append((f"{label}: {d.get('feature')}",val))
    driver_rows=sorted(driver_rows,key=lambda x:abs(x[1]),reverse=True)[:8]
    if driver_rows:
        ws["P17"]="Driver"; ws["Q17"]="Contribution"
        for i,(name,val) in enumerate(driver_rows,18):
            ws.cell(i,16,name); ws.cell(i,17,val); ws.cell(i,17).number_format="0.0%"
        ch=BarChart(); ch.type="bar"; ch.style=12
        ch.title="Largest LightGBM forecast drivers"
        ch.y_axis.title="Driver"; ch.x_axis.title="Signed forecast contribution"; ch.x_axis.numFmt="0.0%"
        ch.height=7.5; ch.width=12.5; ch.legend=None
        ch.add_data(Reference(ws,min_col=17,min_row=17,max_row=17+len(driver_rows)),titles_from_data=True)
        ch.set_categories(Reference(ws,min_col=16,min_row=18,max_row=17+len(driver_rows)))
        ch.dLbls=DataLabelList(); ch.dLbls.showVal=True; ch.dLbls.numFmt="0.0%"
        ws.add_chart(ch,"H35")


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
    for col, width in {
        "A":31,"B":18,"C":20,"D":18,"E":50,"F":24,"G":3,
        "H":15,"I":15,"J":15,"K":15,"L":15,"M":15,"N":15
    }.items():
        ws.column_dimensions[col].width = width

    navy = "17365D"; blue = "2F75B5"; white = "FFFFFF"; gold = "FFF2CC"
    green = "008000"; red = "C00000"; grey = "666666"; pale_red="FCE4D6"; pale_green="E2F0D9"

    ws.merge_cells("A1:N2")
    ws["A1"] = f"{ticker} — AI Growth Forecast"
    ws["A1"].fill = PatternFill("solid", fgColor=navy)
    ws["A1"].font = Font(bold=True, color=white, size=18)
    ws["A1"].alignment = Alignment(vertical="center")
    ws.merge_cells("A3:N3")
    ws["A3"] = (
        "LLM/deterministic AI evidence extraction + LightGBM fundamental growth + "
        "reverse-DCF expectations gap. AI is a bounded evidence overlay until enough "
        "dated AI observations exist for supervised training."
    )
    ws["A3"].font = Font(italic=True, color=grey)
    ws["A3"].alignment = Alignment(wrap_text=True)
    ws.row_dimensions[3].height = 34
    ws["N4"] = datetime.now(timezone.utc).isoformat(timespec="seconds")
    ws["N4"].font = Font(color=grey, italic=True, size=9)

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
            f"LightGBM MAE {_fmt_pct((revenue.get('metrics') or {}).get('time_purged_holdout_mae'))}; "
            f"Elastic Net {_fmt_pct((revenue.get('metrics') or {}).get('elastic_net_holdout_mae'))}; "
            f"confidence {revenue.get('confidence') or 'N/M'}",
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
            ws.cell(rr,4).number_format='0.0%;[Red](0.0%);-'
            ws.cell(rr,5).number_format='0.0%;[Red](0.0%);-'
            if str(forecast.get("confidence") or "").lower()=="low":
                ws.cell(rr,6).fill=PatternFill("solid",fgColor=pale_red)
            elif str(forecast.get("confidence") or "").lower()=="high":
                ws.cell(rr,6).fill=PatternFill("solid",fgColor=pale_green)
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

    _add_ai_growth_charts(ws,signals,revenue,fcf,payload,reverse)
    wb.save(path)
