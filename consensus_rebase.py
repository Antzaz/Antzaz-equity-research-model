"""Optional near-term Street-consensus anchoring.

Consensus is an external benchmark by default, not an input to the independent Base case.
Historically this module overwrote the first two Base revenue-growth assumptions with Yahoo
0y/+1y estimates.  That made the workbook's "Your Model" revenue mechanically identical to
consensus and destroyed the variant-perception comparison.

The legacy anchor is retained only as an explicit opt-in sensitivity.  Set
``EQUITY_CONSENSUS_ANCHOR=1`` when a deliberately Street-anchored case is wanted.  Normal
research runs leave Bear/Base/Bull scenario assumptions untouched and show consensus only on
the expectations sheet.
"""

import os

try:
    import yfinance as yf
except Exception:
    yf=None

from openpyxl.styles import Font
from currency_normalization import convert_financial_amount_to_quote

GREY="666666"; FMT_PCT='0.0%;[Red](0.0%);-'


def _num(v,default=None):
    try:
        if isinstance(v,bool): return default
        return float(v)
    except Exception: return default


def _enabled():
    return str(os.getenv("EQUITY_CONSENSUS_ANCHOR") or "").strip().lower() in {"1","true","yes","on"}


def _revenue_estimates(ticker):
    if yf is None: return None,None,None
    try:
        t=yf.Ticker(ticker); df=t.get_revenue_estimate()
        if df is None or getattr(df,"empty",True): return None,None,None
        try: info=t.info or {}
        except Exception: info={}
        def avg(idx):
            if idx not in df.index: return None
            v=_num(df.loc[idx].get("avg"))
            if v is None: return None
            converted=convert_financial_amount_to_quote(v,info)
            if converted is None: return None
            return converted/1e9 if abs(converted)>1e6 else converted
        return avg("0y"),avg("+1y"),info
    except Exception:
        return None,None,None


def rebase_near_term_revenue(wb,ticker):
    """Apply the old consensus anchor only when explicitly enabled.

    Returning ``False`` means no scenario input was changed.  This is the expected result for
    production research runs.  Callers do not need to change: the safety policy lives here so a
    future pipeline reorder cannot accidentally make consensus and the Base case identical again.
    """
    if not {"Historical Financials","Three-Case Scenarios"}.issubset(wb.sheetnames): return False
    s=wb["Three-Case Scenarios"]

    if not _enabled():
        s["AK8"]="Consensus policy"
        s["AL8"]=(
            "Independent-model mode: public analyst consensus is NOT written into Bear/Base/Bull "
            "scenario assumptions. Consensus is an external benchmark on Expectations & Consensus. "
            "Set EQUITY_CONSENSUS_ANCHOR=1 only for an explicit Street-anchored sensitivity run."
        )
        s["AL8"].font=Font(italic=True,color=GREY)
        return False

    h=wb["Historical Financials"]; latest=_num(h["G4"].value)
    if not latest or latest<=0: return False
    y0,y1,info=_revenue_estimates(ticker)
    if y0 is None or y0<=0: return False

    # Reject obviously stale/mismatched estimates.  Currency normalization above should handle
    # cross-border issuers, but this guard prevents a malformed provider row driving scenarios.
    g0=y0/latest-1
    if g0<=-.7 or g0>=1.0:
        print(f"Warning: consensus revenue anchor rejected for {ticker}: implied growth {g0:.1%}")
        return False
    g1=(y1/y0-1) if y1 and y1>0 else None

    s["N12"]=g0; s["N12"].number_format=FMT_PCT
    s["B12"]=max(-.7,g0-.03); s["B12"].number_format=FMT_PCT
    s["Z12"]=min(1.0,g0+.03); s["Z12"].number_format=FMT_PCT
    if g1 is not None and -.7<g1<1.0:
        s["O12"]=g1; s["O12"].number_format=FMT_PCT
        s["C12"]=max(-.7,g1-.03); s["C12"].number_format=FMT_PCT
        s["AA12"]=min(1.0,g1+.03); s["AA12"].number_format=FMT_PCT

    quote=str((info or {}).get("currency") or "model currency")
    financial=str((info or {}).get("financialCurrency") or quote)
    s["AK8"]="Consensus policy"
    s["AL8"]=(
        f"EXPLICIT STREET-ANCHORED SENSITIVITY: yfinance revenue_estimate (0y/+1y), normalized "
        f"{financial} → {quote} when required. This mode is not the independent Base forecast."
    )
    s["AL8"].font=Font(italic=True,color=GREY)
    return True
