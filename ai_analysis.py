"""Compatibility entrypoint for the cross-company AI Impact Analysis.

analysis_charts historically imported ``ensure_ai_analysis`` from this module.  Keep that
contract, remove retired Alphabet-only tabs, and always build the newer company-agnostic
AI Impact Analysis so every generated research workbook receives the same AI diligence
framework.
"""

from ai_effect_analysis import ensure_ai_impact_analysis


def ensure_ai_analysis(wb, ticker):
    ticker = str(ticker or "").upper()
    for name in ("AI Analysis", "AI Valuation"):
        if name in wb.sheetnames:
            wb.remove(wb[name])
    return ensure_ai_impact_analysis(wb, ticker)
