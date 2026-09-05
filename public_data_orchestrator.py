from __future__ import annotations

"""Orchestrate conservative public-data recovery and canonical-history synchronization."""

from typing import Any

from business_model_registry import get_business_model_policy
from public_data_backfill import backfill_public_data


def _num(v):
    try:
        if isinstance(v, bool) or v in (None, ""):
            return None
        return float(v)
    except Exception:
        return None


def _find(ws, label, start=1, end=None):
    needle = str(label or "").strip().lower(); end = end or ws.max_row
    for r in range(start, min(end, ws.max_row) + 1):
        if str(ws.cell(r, 1).value or "").strip().lower() == needle:
            return r
    return None


def _year_cols(ws, header):
    return {
        int(ws.cell(header, c).value): c
        for c in range(2, min(ws.max_column, 12) + 1)
        if isinstance(ws.cell(header, c).value, (int, float)) and 1900 <= int(ws.cell(header, c).value) <= 2100
    }


def _sync_history(wb, ticker: str) -> int:
    if "Financial Statements" not in wb.sheetnames or "Historical Financials" not in wb.sheetnames:
        return 0
    fs = wb["Financial Statements"]; hs = wb["Historical Financials"]
    i0 = _find(fs, "Income Statement"); b0 = _find(fs, "Balance Sheet"); c0 = _find(fs, "Cash Flow Statement")
    if not all((i0, b0, c0)):
        return 0
    ih = next((r for r in range(i0 + 1, min(b0, i0 + 6)) if str(fs.cell(r, 1).value or "").strip().lower() == "metric"), None)
    ch = next((r for r in range(c0 + 1, min(fs.max_row + 1, c0 + 6)) if str(fs.cell(r, 1).value or "").strip().lower() == "metric"), None)
    if not ih or not ch:
        return 0
    iy = _year_cols(fs, ih); cy = _year_cols(fs, ch)
    hy = {
        int(hs.cell(3, c).value): c
        for c in range(2, min(hs.max_column, 8) + 1)
        if isinstance(hs.cell(3, c).value, (int, float)) and 1900 <= int(hs.cell(3, c).value) <= 2100
    }
    try:
        cd = wb["Company Data"]
        policy = get_business_model_policy(ticker, cd["B6"].value, cd["B7"].value, cd["B5"].value)
    except Exception:
        policy = get_business_model_policy(ticker)

    revenue_label = "Total Net Revenue" if policy.key in {"bank", "capital_markets"} else "Total Revenues" if policy.key in {"insurance", "insurance_conglomerate"} else "Revenue"
    net_label = "Net Earnings Attributable to Berkshire" if policy.key in {"insurance", "insurance_conglomerate"} and _find(fs, "Net Earnings Attributable to Berkshire", i0, b0 - 1) else "Net Income"
    mapping = [
        (4, revenue_label, iy, i0, b0 - 1, False),
        (11, net_label, iy, i0, b0 - 1, False),
        (14, "Operating Cash Flow", cy, c0, fs.max_row, False),
        (15, "Capital Expenditures", cy, c0, fs.max_row, True),
        (18, "Depreciation, Amortization & Accretion", cy, c0, fs.max_row, False),
        (21, "Stock-Based Compensation", cy, c0, fs.max_row, False),
    ]
    if policy.industrial_fcf_primary:
        mapping.insert(1, (9, "Operating Income", iy, i0, b0 - 1, False))
    else:
        for _year, hc in hy.items():
            hs.cell(9, hc).value = None

    written = 0
    for hrow, label, cols, start, end, absolute in mapping:
        r = _find(fs, label, start, end)
        if not r:
            continue
        for year, hc in hy.items():
            fc = cols.get(year); value = _num(fs.cell(r, fc).value) if fc else None
            if value is None:
                continue
            hs.cell(hrow, hc).value = abs(value) if absolute else value
            written += 1
    return written


def run_public_data_recovery(wb, ticker: str, info: dict | None = None) -> dict[str, Any]:
    result = backfill_public_data(wb, ticker, info or {})
    result["history_sync"] = _sync_history(wb, ticker)
    try:
        wb.calculation.calcMode = "auto"; wb.calculation.fullCalcOnLoad = True; wb.calculation.forceFullCalc = True
    except Exception:
        pass
    return result
