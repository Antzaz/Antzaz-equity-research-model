"""Best-effort near-term revenue anchoring from current public consensus.

The long-term scenario engine should not blindly extrapolate historical CAGR into the
next fiscal year when a current analyst revenue estimate is available. This module uses
yfinance's 0y/+1y revenue estimates to anchor the first two Base forecast years, then
leaves the model's long-run fade intact. Bear/Bull retain explicit +/-3ppt spreads around
the consensus-anchored Base growth for those two years.

Foreign issuers are normalized to the traded security's quote currency before comparing
consensus revenue with the Historical Financials sheet. This prevents ADR models from
mixing local-currency revenue estimates with USD-normalized history.
"""

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
    except Exception: return None,None,None

def rebase_near_term_revenue(wb,ticker):
    if not {"Historical Financials","Three-Case Scenarios"}.issubset(wb.sheetnames): return False
    h=wb["Historical Financials"]; s=wb["Three-Case Scenarios"]; latest=_num(h["G4"].value)
    if not latest or latest<=0: return False
    y0,y1,info=_revenue_estimates(ticker)
    if y0 is None or y0<=0: return False
    # Reject obviously stale/mismatched estimates. A currency mismatch should now be caught
    # by normalization above, but retain the guard so bad provider rows never drive scenarios.
    g0=y0/latest-1
    if g0<=-.7 or g0>=1.0:
        print(f"Warning: consensus revenue anchor rejected for {ticker}: implied growth {g0:.1%}")
        return False
    g1=(y1/y0-1) if y1 and y1>0 else None
    s["N12"]=g0; s["N12"].number_format=FMT_PCT; s["B12"]=max(-.7,g0-.03); s["B12"].number_format=FMT_PCT; s["Z12"]=min(1.0,g0+.03); s["Z12"].number_format=FMT_PCT
    if g1 is not None and -.7<g1<1.0:
        s["O12"]=g1; s["O12"].number_format=FMT_PCT; s["C12"]=max(-.7,g1-.03); s["C12"].number_format=FMT_PCT; s["AA12"]=min(1.0,g1+.03); s["AA12"].number_format=FMT_PCT
    quote=str((info or {}).get("currency") or "model currency")
    financial=str((info or {}).get("financialCurrency") or quote)
    s["AK8"]="Near-term revenue anchor"
    s["AL8"]=f"Public analyst consensus via yfinance revenue_estimate (0y/+1y), normalized {financial} → {quote} when required; long-run scenario fade remains model-driven."
    s["AL8"].font=Font(italic=True,color=GREY)
    return True
