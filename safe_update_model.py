from __future__ import annotations

"""Guarded entry point for the deterministic equity model.

This wrapper installs source-integrity, scoring and valuation controls before running
update_model.py:
- date-like cells cannot become financial values or fake fiscal years;
- cross-border annual history is independently sanity-checked;
- sparse SEC statement sheets receive structured fallback data rather than silent blanks;
- WACC is calculated per company from public market inputs instead of inheriting 9% from the template;
- one sector-aware Score Engine feeds Advanced Analytics, Investment Summary and institutional lenses;
- every reusable score receives an audit trail with formula, inputs and sources;
- verified issuer/regulator segment data can replace failed generic parsing;
- the final Decision View reconciles quality, valuation, downside and margin of safety;
- alternate exchange listings are deduplicated in peer discovery;
- a disk-space preflight avoids building a workbook that cannot be saved.
"""

from datetime import date, datetime
from pathlib import Path
import re
import shutil

import pandas as pd

import data_integrity
import issuer_source_engine
import model_reliability
import dynamic_peer_engine
import segment_chart_fix
import update_model
import advanced_analytics_v2
import institutional_lenses
import research_extensions
from financial_statement_repair_final import repair_financial_statements
from score_engine_v2 import advanced_scorecard
from score_integration_v2 import institutional_dimensions, leadership_proxy, finalize_score_transparency
from wacc_engine import apply_dynamic_wacc
from costco_segment_analysis import ensure_costco_segment_analysis
from decision_view_v2 import ensure_decision_view

BASE=Path(__file__).resolve().parent


def _safe_year(value):
    if isinstance(value,(datetime,date,pd.Timestamp)): return int(pd.Timestamp(value).year)
    if isinstance(value,bool) or value is None: return None
    if isinstance(value,(int,float)):
        n=int(value); return n if 1900<=n<=2100 and float(value)==n else None
    text=str(value).strip()
    if re.fullmatch(r"(?:19|20)\d{2}",text): return int(text)
    m=re.search(r"(?:19|20)\d{2}",text); return int(m.group(0)) if m else None


def _safe_number(value):
    if isinstance(value,(datetime,date,pd.Timestamp)): return None
    if isinstance(value,str):
        text=value.strip()
        if re.fullmatch(r"\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}(?:[ T].*)?",text): return None
        if re.fullmatch(r"(?:19|20)\d{2}-\d{2}-\d{2}.*",text): return None
    return _ORIGINAL_NUMBER(value)


_ORIGINAL_NUMBER=issuer_source_engine._number
_ORIGINAL_BUILD=issuer_source_engine.build_crossborder_history
_ORIGINAL_UPDATE_FILINGS=update_model.update_filings
_ORIGINAL_SELECT_PEERS=dynamic_peer_engine.select_dynamic_peers
_ORIGINAL_NORMALIZE=update_model.normalize_workbook_currency
_ORIGINAL_UPDATE_SCENARIOS=update_model.update_scenarios
_ORIGINAL_SEGMENT_V2=update_model.ensure_segment_analysis_v2
_ORIGINAL_SEGMENT_ENRICH=segment_chart_fix.enrich_segment_analysis
_ORIGINAL_FINANCIAL_STATEMENTS=update_model.ensure_financial_statements
_ORIGINAL_RESEARCH_EXTENSIONS=update_model.ensure_research_extensions

issuer_source_engine._year=_safe_year
issuer_source_engine._number=_safe_number
# Operating income must be preferred to pretax income in Yahoo statement recovery.
issuer_source_engine.YF_INCOME_ROWS["op"]=["Operating Income","Operating Income Loss","Pretax Income"]

issuer_source_engine.ISSUER_IR_PAGES.setdefault("SIE.DE",[
    "https://www.siemens.com/es-es/company/investor-relations/financial-results/",
    "https://www.siemens.com/en-us/company/investor-relations/financial-results-archive/",
])

dynamic_peer_engine.STRATEGIC_PEERS.setdefault("SIE.DE",[
    "SU.PA","ABBN.SW","HON","ETN","ROK","EMR","PH","GE","AME","IR"
])


def _safe_build_crossborder_history(ticker,info,facts):
    hist,meta=_ORIGINAL_BUILD(ticker,info,facts)
    return data_integrity.sanitize_crossborder_history(ticker,info,hist,meta)

issuer_source_engine.build_crossborder_history=_safe_build_crossborder_history
model_reliability.build_crossborder_history=_safe_build_crossborder_history


def _safe_update_filings(wb,ticker):
    if "Filings" in wb.sheetnames:
        ws=wb["Filings"]
        for r in range(4,20):
            for c in range(1,6): ws.cell(r,c).value=None
    return _ORIGINAL_UPDATE_FILINGS(wb,ticker)

update_model.update_filings=_safe_update_filings


def _entity_key(symbol,info):
    name=str(info.get("longName") or info.get("shortName") or "").lower()
    name=re.sub(r"\b(ag|aktiengesellschaft|se|sa|plc|inc|corp|corporation|ltd|limited|nv)\b","",name)
    name=re.sub(r"[^a-z0-9]+"," ",name).strip()
    return name or str(symbol).upper().split(".")[0]


def _safe_select_dynamic_peers(wb,ticker,count=9):
    target,sector,industry,raw_peers=_ORIGINAL_SELECT_PEERS(wb,ticker,max(count*3,18))
    seen={_entity_key(ticker,target)}; peers=[]
    for symbol,info in raw_peers:
        key=_entity_key(symbol,info)
        if key in seen: continue
        seen.add(key); peers.append((symbol,info))
        if len(peers)>=count: break
    return target,sector,industry,peers

dynamic_peer_engine.select_dynamic_peers=_safe_select_dynamic_peers


def _safe_normalize_workbook_currency(wb,ticker,info):
    result=_ORIGINAL_NORMALIZE(wb,ticker,info)
    data_integrity.apply_workbook_integrity_controls(wb,ticker,info,getattr(wb,"_issuer_source_meta",None))
    return result
update_model.normalize_workbook_currency=_safe_normalize_workbook_currency


def _ticker_from_wb(wb):
    try: return str(wb["Company Data"]["B4"].value or "").upper().strip()
    except Exception: return ""


def _safe_update_scenarios(wb,hist,info):
    result=_ORIGINAL_UPDATE_SCENARIOS(wb,hist,info)
    ticker=_ticker_from_wb(wb)
    # Preserve the exact market-data snapshot used for beta/market-cap inputs so the
    # post-statements and final WACC refresh cannot switch from provider beta to a
    # regression beta after DCF/Monte Carlo have already been calculated.
    setattr(wb,"_wacc_info",info or {})
    if ticker:
        try: apply_dynamic_wacc(wb,ticker,getattr(wb,"_wacc_info",{}))
        except Exception as exc: print(f"Warning: initial dynamic WACC failed: {exc}")
    return result
update_model.update_scenarios=_safe_update_scenarios


def _safe_segment_v2(wb,ticker,headers):
    if str(ticker).upper()=="COST":
        return ensure_costco_segment_analysis(wb,ticker)
    return _ORIGINAL_SEGMENT_V2(wb,ticker,headers)
update_model.ensure_segment_analysis_v2=_safe_segment_v2


def _safe_segment_enrichment(wb,ticker,headers):
    # The verified Costco primary-source sheet must not be polluted by narrative-parser fragments.
    if str(ticker).upper()=="COST":
        return ensure_costco_segment_analysis(wb,ticker)
    return _ORIGINAL_SEGMENT_ENRICH(wb,ticker,headers)
segment_chart_fix.enrich_segment_analysis=_safe_segment_enrichment


def _safe_financial_statements(wb,ticker,facts):
    ws=_ORIGINAL_FINANCIAL_STATEMENTS(wb,ticker,facts)
    try:
        result=repair_financial_statements(wb,ticker); setattr(wb,"_financial_statement_repair",result)
        print(f"Financial Statements fallback repair: filled={result.get('filled',0)}, core coverage={result.get('coverage',0):.0%}")
    except Exception as exc: print(f"Warning: Financial Statements fallback repair failed: {exc}")
    # Recalculate WACC after the reported tax rate / statement repair is available, before DCF analytics,
    # but keep the same beta/market-cap snapshot captured at scenario construction.
    try: apply_dynamic_wacc(wb,ticker,getattr(wb,"_wacc_info",{}))
    except Exception as exc: print(f"Warning: post-statements dynamic WACC failed: {exc}")
    return ws
update_model.ensure_financial_statements=_safe_financial_statements

# Single scoring source of truth. Functions created earlier resolve these module globals at runtime.
advanced_analytics_v2._scorecard=advanced_scorecard
institutional_lenses._scorecard_dimensions=institutional_dimensions
research_extensions._leadership_proxy=leadership_proxy


def _safe_research_extensions(wb,ticker,info=None):
    fin={}
    try:
        fin=repair_financial_statements(wb,ticker); setattr(wb,"_financial_statement_repair",fin)
    except Exception as exc: print(f"Warning: final Financial Statements repair failed: {exc}")
    # Refresh source dates at the end without changing the market-data snapshot that fed valuation.
    wacc_info=info or getattr(wb,"_wacc_info",{})
    try: apply_dynamic_wacc(wb,ticker,wacc_info)
    except Exception as exc: print(f"Warning: final dynamic WACC refresh failed: {exc}")
    result=_ORIGINAL_RESEARCH_EXTENSIONS(wb,ticker,info)
    try: finalize_score_transparency(wb,ticker,fin.get("coverage") if isinstance(fin,dict) else None)
    except Exception as exc: print(f"Warning: score transparency finalization failed: {exc}")
    try: ensure_decision_view(wb,ticker)
    except Exception as exc: print(f"Warning: Decision View finalization failed: {exc}")
    return result
update_model.ensure_research_extensions=_safe_research_extensions


def _disk_preflight(min_free_gb=1.5):
    free=shutil.disk_usage(BASE).free/(1024**3)
    if free<min_free_gb:
        raise OSError(f"Only {free:.2f} GB free on the project drive. Free at least {min_free_gb:.1f} GB before running the equity model so Git/Excel temporary files and the final workbook can be written safely.")
    print(f"Disk-space preflight: {free:.2f} GB free")


def main():
    _disk_preflight(); update_model.main()

if __name__=="__main__": main()
