"""Compatibility shim for the retired Alphabet-only AI Analysis / AI Valuation tabs.

analysis_charts historically imported ``ensure_ai_analysis`` from this module. The current
research model uses the newer, cross-company ``AI Impact Analysis`` layer instead, so the
legacy pair is intentionally not recreated. Keeping this small shim preserves the import
contract while preventing duplicate AI presentation tabs from returning.
"""


def ensure_ai_analysis(wb, ticker):
    ticker = str(ticker or "").upper()
    # Clean up any stale legacy tabs that may still exist in the base template or an older
    # workbook. The newer AI Impact Analysis is created by ai_effect_analysis later in the
    # final reliability pipeline.
    for name in ("AI Analysis", "AI Valuation"):
        if name in wb.sheetnames:
            wb.remove(wb[name])
    return None
