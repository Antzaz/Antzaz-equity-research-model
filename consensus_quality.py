from __future__ import annotations

"""Post-generation quality control for public analyst consensus.

Yahoo can report analyst revenue estimates in an issuer's reporting currency even when the
traded security and workbook are valued in another currency. This module rewrites only the
Revenue consensus rows in Expectations & Consensus so they use the workbook/quote currency.
EPS is left untouched because ADR EPS conventions can differ from local-share EPS and should
not be silently transformed without an explicit depositary ratio source.
"""

import math

import yfinance as yf
from openpyxl.styles import Alignment, Font, PatternFill

from currency_normalization import convert_financial_amount_to_quote

GOLD="FFF2CC"; PALE_GREEN="E2F0D9"; PALE_RED="FCE4D6"; GREY="666666"
FMT_BN='#,##0.0;[Red](#,##0.0);-'; FMT_PCT='0.0%;[Red](0.0%);-'


def _num(v,default=None):
    try:
        if isinstance(v,bool) or v in (None,""): return default
        x=float(v)
        return x if math.isfinite(x) else default
    except Exception:
        return default


def _latest_year(wb):
    if "Historical Financials" not in wb.sheetnames: return None
    ws=wb["Historical Financials"]
    years=[int(ws.cell(3,c).value) for c in range(2,8) if isinstance(ws.cell(3,c).value,(int,float))]
    return max(years) if years else None


def _quality_row(wb,status,observed):
    if "Data Quality" not in wb.sheetnames: return
    ws=wb["Data Quality"]
    row=None
    for r in range(1,ws.max_row+1):
        if str(ws.cell(r,1).value or "").strip()=="Consensus currency alignment":
            row=r; break
    row=row or ws.max_row+1
    ws.cell(row,1,"Consensus currency alignment")
    ws.cell(row,2,status)
    ws.cell(row,3,observed)
    ws.cell(row,4,"Revenue consensus and model revenue must use the same currency before gaps or revision signals are interpreted.")
    ws.cell(row,2).fill=PatternFill("solid",fgColor=PALE_GREEN if status=="PASS" else (PALE_RED if status=="FAIL" else GOLD))
    ws.cell(row,2).font=Font(bold=True)
    for c in range(1,5): ws.cell(row,c).alignment=Alignment(wrap_text=True,vertical="top")


def normalize_expectations_consensus(wb,ticker:str,info:dict|None=None)->bool:
    if "Expectations & Consensus" not in wb.sheetnames:
        return False
    info=info or {}
    financial=str(info.get("financialCurrency") or "").upper()
    quote=str(info.get("currency") or "").upper()
    if not financial or not quote:
        try:
            live=yf.Ticker(ticker).info or {}
            info={**live,**info}
            financial=str(info.get("financialCurrency") or "").upper()
            quote=str(info.get("currency") or "").upper()
        except Exception:
            pass

    try:
        df=yf.Ticker(ticker).revenue_estimate
    except Exception:
        df=None
    if df is None or getattr(df,"empty",True):
        _quality_row(wb,"REVIEW","Public revenue-estimate table unavailable; existing consensus cells were not changed.")
        return False

    latest=_latest_year(wb)
    if latest is None:
        _quality_row(wb,"REVIEW","Could not determine latest fiscal year for consensus alignment.")
        return False

    ws=wb["Expectations & Consensus"]
    changed=0
    rows=[]
    for r in range(7,ws.max_row+1):
        metric=str(ws.cell(r,1).value or "").strip()
        year=ws.cell(r,2).value
        if metric=="Revenue" and isinstance(year,(int,float)):
            rows.append((r,int(year)))

    for r,year in rows:
        idx="0y" if year==latest+1 else ("+1y" if year==latest+2 else None)
        if idx is None or idx not in df.index: continue
        row=df.loc[idx]
        avg=_num(row.get("avg")); low=_num(row.get("low")); high=_num(row.get("high"))
        if avg is None: continue
        converted=convert_financial_amount_to_quote(avg,info,year)
        converted_low=convert_financial_amount_to_quote(low,info,year) if low is not None else None
        converted_high=convert_financial_amount_to_quote(high,info,year) if high is not None else None
        if converted is None:
            continue
        cons=converted/1e9
        ws.cell(r,3,cons); ws.cell(r,3).number_format=FMT_BN
        if cons and converted_low is not None and converted_high is not None:
            ws.cell(r,9,(converted_high-converted_low)/abs(converted)); ws.cell(r,9).number_format=FMT_PCT
        note="Yahoo Finance public consensus snapshot"
        if financial and quote and financial!=quote:
            note+=f"; currency-normalized {financial}→{quote}"
        ws.cell(r,11,note)
        ws.cell(r,11).font=Font(color=GREY)
        changed+=1

    if changed:
        _quality_row(wb,"PASS",f"Normalized {changed} revenue consensus row(s); reporting={financial or 'unknown'}, quote={quote or 'unknown'}.")
        return True
    _quality_row(wb,"REVIEW",f"No revenue consensus rows could be normalized; reporting={financial or 'unknown'}, quote={quote or 'unknown'}.")
    return False
