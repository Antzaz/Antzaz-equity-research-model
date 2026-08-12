from __future__ import annotations

"""Runtime guards for provider-specific ticker symbols and stale template metadata.

The workbook template is intentionally reused across tickers.  Raw company cells must therefore
be cleared before a new issuer is written.  Provider aliases are kept at the data-source boundary
so the analyst-facing ticker can remain BRK.B while Yahoo/SEC receive BRK-B.
"""

import re


PROVIDER_SYMBOLS={
    "BRK.B":"BRK-B",
    "BRK.A":"BRK-A",
}

# Fallbacks are used only when the live provider omitted classification.  They prevent the
# original GOOGL template classification from surviving into another issuer's workbook.
CLASSIFICATION_FALLBACKS={
    "SIE.DE":("Industrials","Specialty Industrial Machinery"),
    "BRK.B":("Financial Services","Insurance - Diversified"),
    "BRK-B":("Financial Services","Insurance - Diversified"),
    "JPM":("Financial Services","Banks - Diversified"),
}

_INSTALLED=False


def provider_symbol(symbol):
    s=str(symbol or "").strip()
    return PROVIDER_SYMBOLS.get(s.upper(),s)


def _normalize_ticker_arg(value):
    if isinstance(value,str):
        # yfinance accepts a single symbol or a whitespace/comma-delimited symbol string.
        parts=re.split(r"([,\s]+)",value)
        return "".join(provider_symbol(p) if p and not re.fullmatch(r"[,\s]+",p) else p for p in parts)
    if isinstance(value,tuple): return tuple(provider_symbol(x) for x in value)
    if isinstance(value,list): return [provider_symbol(x) for x in value]
    if isinstance(value,set): return {provider_symbol(x) for x in value}
    return value


def install_runtime_data_guards():
    global _INSTALLED
    if _INSTALLED: return

    import yfinance as yf
    import update_model

    original_ticker=yf.Ticker
    original_download=yf.download
    original_tickers=getattr(yf,"Tickers",None)

    def guarded_ticker(symbol,*args,**kwargs):
        return original_ticker(provider_symbol(symbol),*args,**kwargs)

    def guarded_download(tickers,*args,**kwargs):
        return original_download(_normalize_ticker_arg(tickers),*args,**kwargs)

    yf.Ticker=guarded_ticker
    yf.download=guarded_download
    if original_tickers is not None:
        def guarded_tickers(tickers,*args,**kwargs):
            return original_tickers(_normalize_ticker_arg(tickers),*args,**kwargs)
        yf.Tickers=guarded_tickers

    original_cik_for=update_model.cik_for
    def guarded_cik_for(ticker):
        return original_cik_for(provider_symbol(ticker))
    update_model.cik_for=guarded_cik_for

    original_put_company=update_model.put_company
    def guarded_put_company(wb,ticker,info):
        ws=wb["Company Data"]
        # B5:B15 are issuer-specific raw metadata/market fields in the clean template.  Clearing
        # them first is safer than preserving a prior GOOGL value when a provider returns None.
        for row in range(5,16):
            ws.cell(row,2).value=None
        original_put_company(wb,ticker,info or {})
        fallback=CLASSIFICATION_FALLBACKS.get(str(ticker).upper().strip())
        if fallback:
            if not str(ws["B6"].value or "").strip(): ws["B6"]=fallback[0]
            if not str(ws["B7"].value or "").strip(): ws["B7"]=fallback[1]
        # A known stale-template pair is never valid for Siemens.  This extra assertion-style
        # repair protects old local templates that may already contain GOOGL metadata.
        if str(ticker).upper().strip()=="SIE.DE":
            pair=(str(ws["B6"].value or "").strip(),str(ws["B7"].value or "").strip())
            if pair==("Communication Services","Internet Content & Information"):
                ws["B6"],ws["B7"]="Industrials","Specialty Industrial Machinery"
    update_model.put_company=guarded_put_company

    _INSTALLED=True
