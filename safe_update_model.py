from __future__ import annotations

"""Guarded entry point for the deterministic equity model.

This wrapper installs source-integrity, scoring, statement and valuation controls before
running update_model.py. The final workbook uses canonical annual financial definitions,
company-specific WACC, reliability-gated scoring, verified segment adapters where available,
profile-aware full three-statement reporting, a main-products/company profile, and a final
quality/pruning pass before save.
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
import score_integration_v2
import decision_view_v2
from financial_statement_integrity_v3 import repair_financial_statements_v3 as repair_financial_statements
from full_financial_statements_v3 import expand_financial_statements
from verified_full_statement_adapters import apply_verified_full_statement_adapter
from company_profile_v2 import enrich_company_data
from score_engine_v3 import advanced_scorecard, compute_score_bundle
from score_integration_v2 import institutional_dimensions, leadership_proxy, finalize_score_transparency
from wacc_engine import apply_dynamic_wacc
from costco_segment_analysis import ensure_costco_segment_analysis
from constellation_segment_analysis import ensure_constellation_segment_analysis
from decision_view_v2 import ensure_decision_view
from output_quality_v3 import prune_low_value_tabs, reconcile_score_displays, ensure_quality_checks

BASE=Path(__file__).resolve().parent

# Verified consolidated annual figures from Constellation's investor financial table.
# These are a guardrail against SEC tag-selection choosing only contract-with-customer revenue.
# Network/SEC/issuer data remains the normal source; this adapter is ticker-specific and must be
# extended when a new audited fiscal year is published.
CEG_CANONICAL_HISTORY={
    2022:{"revenue":24.440e9,"op":0.495e9,"ni":-0.160e9,"capex":1.689e9,"ocf":-2.353e9},
    2023:{"revenue":24.918e9,"op":1.610e9,"ni":1.623e9,"capex":2.422e9,"ocf":-5.301e9},
    2024:{"revenue":23.568e9,"op":4.352e9,"ni":3.749e9,"capex":2.565e9,"ocf":-2.464e9},
    2025:{"revenue":25.533e9,"op":3.086e9,"ni":2.319e9,"capex":2.949e9,"ocf":4.237e9},
}
for _d in CEG_CANONICAL_HISTORY.values():
    _d["fcf"]=_d["ocf"]-_d["capex"]


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


def _pending_crossborder_normalization(wb):
    info=getattr(wb,"_wacc_info",{}) or {}
    financial=str(info.get("financialCurrency") or "").upper().strip()
    quote=str(info.get("currency") or "").upper().strip()
    return bool(financial and quote and financial!=quote),financial,quote


def _safe_update_scenarios(wb,hist,info):
    ticker=_ticker_from_wb(wb)
    guarded={y:dict(v) for y,v in (hist or {}).items()}
    if ticker=="CEG":
        for y,canonical in CEG_CANONICAL_HISTORY.items():
            base=dict(guarded.get(y) or {})
            base.update(canonical)
            guarded[y]=base
    result=_ORIGINAL_UPDATE_SCENARIOS(wb,guarded,info)
    setattr(wb,"_wacc_info",info or {})
    if ticker:
        try: apply_dynamic_wacc(wb,ticker,getattr(wb,"_wacc_info",{}))
        except Exception as exc: print(f"Warning: initial dynamic WACC failed: {exc}")
    return result
update_model.update_scenarios=_safe_update_scenarios


def _verified_segment(wb,ticker):
    t=str(ticker).upper()
    if t=="COST": return ensure_costco_segment_analysis(wb,ticker)
    if t=="CEG": return ensure_constellation_segment_analysis(wb,ticker)
    return None

def _safe_segment_v2(wb,ticker,headers):
    special=_verified_segment(wb,ticker)
    return special if special is not None else _ORIGINAL_SEGMENT_V2(wb,ticker,headers)
update_model.ensure_segment_analysis_v2=_safe_segment_v2

def _safe_segment_enrichment(wb,ticker,headers):
    special=_verified_segment(wb,ticker)
    return special if special is not None else _ORIGINAL_SEGMENT_ENRICH(wb,ticker,headers)
segment_chart_fix.enrich_segment_analysis=_safe_segment_enrichment


def _safe_financial_statements(wb,ticker,facts):
    _ORIGINAL_FINANCIAL_STATEMENTS(wb,ticker,facts)
    crossborder,financial,quote=_pending_crossborder_normalization(wb)
    initial={}
    if crossborder:
        print(f"Financial Statements canonical pre-repair deferred: reporting {financial} -> valuation {quote}")
    else:
        try:
            initial=repair_financial_statements(wb,ticker) or {}
            print(f"Financial Statements canonical pre-repair: filled={initial.get('filled',0)}, history_sync={initial.get('history_sync',0)}")
        except Exception as exc:
            print(f"Warning: Financial Statements canonical pre-repair failed: {exc}")
    expansion={}
    try:
        expansion=expand_financial_statements(wb,ticker,facts) or {}
        setattr(wb,"_full_statement_expansion",expansion)
        print(
            "Full Financial Statements: "
            f"profile={expansion.get('profile','default')}, "
            f"income={expansion.get('income_rows',0)}/{expansion.get('income_total',0)}, "
            f"balance={expansion.get('balance_rows',0)}/{expansion.get('balance_total',0)}, "
            f"cash={expansion.get('cash_rows',0)}/{expansion.get('cash_total',0)}"
        )
    except Exception as exc:
        print(f"Warning: full Financial Statements expansion failed: {exc}")
    try:
        verified=apply_verified_full_statement_adapter(wb,ticker) or {}
        if verified.get("written"): print(f"Verified full-statement adapter: {ticker} wrote {verified['written']} annual cells")
    except Exception as exc:
        print(f"Warning: verified full-statement adapter failed: {exc}")
    if crossborder:
        final={"deferred_for_currency_normalization":True,"reporting_currency":financial,"valuation_currency":quote,"full_statement_expansion":expansion}
        setattr(wb,"_financial_statement_repair",final)
        print("Financial Statements canonical final repair deferred until post-FX normalization")
    else:
        try:
            final=repair_financial_statements(wb,ticker) or {}
            final["full_statement_expansion"]=expansion
            setattr(wb,"_financial_statement_repair",final)
            print(f"Financial Statements canonical final repair: filled={final.get('filled',0)}, history_sync={final.get('history_sync',0)}")
        except Exception as exc:
            print(f"Warning: Financial Statements canonical final repair failed: {exc}")
    try: apply_dynamic_wacc(wb,ticker,getattr(wb,"_wacc_info",{}))
    except Exception as exc: print(f"Warning: post-statements dynamic WACC failed: {exc}")
    return wb["Financial Statements"] if "Financial Statements" in wb.sheetnames else None
update_model.ensure_financial_statements=_safe_financial_statements


def _advanced_sheet_scorecard(wb,current_price,forward_pe,base_value,severe_value):
    """Advanced Analytics' legacy renderer expects numeric rows while building.

    Reliability-excluded dimensions get a neutral temporary placeholder solely so chart/table
    construction cannot crash. The final reconciliation pass replaces those cells with the v3
    authoritative value (including blank/excluded) and status, so the neutral placeholder is
    never presented as investment evidence in the saved workbook.
    """
    rows=advanced_scorecard(wb,current_price,forward_pe,base_value,severe_value)
    return [(name,50.0 if score is None else score,note+(" Temporary neutral render placeholder; final score is excluded by reliability gate." if score is None else "")) for name,score,note in rows]

advanced_analytics_v2._scorecard=_advanced_sheet_scorecard
score_integration_v2.compute_score_bundle=compute_score_bundle
decision_view_v2.compute_score_bundle=compute_score_bundle
institutional_lenses._scorecard_dimensions=institutional_dimensions
research_extensions._leadership_proxy=leadership_proxy


def _safe_research_extensions(wb,ticker,info=None):
    fin={}
    try:
        # update_model.main() has already run its final currency normalization before this hook.
        # Only now is it safe to synchronize foreign reporting-currency statements into the
        # canonical valuation-currency Historical Financials sheet.
        apply_verified_full_statement_adapter(wb,ticker)
        fin=repair_financial_statements(wb,ticker) or {}
        fin["full_statement_expansion"]=getattr(wb,"_full_statement_expansion",{})
        setattr(wb,"_financial_statement_repair",fin)
    except Exception as exc: print(f"Warning: final Financial Statements repair failed: {exc}")
    result=_ORIGINAL_RESEARCH_EXTENSIONS(wb,ticker,info)
    wacc_info=info or getattr(wb,"_wacc_info",{})
    try: apply_dynamic_wacc(wb,ticker,wacc_info)
    except Exception as exc: print(f"Warning: final dynamic WACC refresh failed: {exc}")
    try: _verified_segment(wb,ticker)
    except Exception as exc: print(f"Warning: final verified Segment Analysis refresh failed: {exc}")
    try:
        profile=enrich_company_data(wb,ticker,info or {}) or {}
        setattr(wb,"_company_profile",profile)
        print(f"Company Data products/services: {profile.get('products',0)} rows ({profile.get('source','unknown')})")
    except Exception as exc: print(f"Warning: Company Data product/profile enrichment failed: {exc}")

    removed=[]
    try: removed=prune_low_value_tabs(wb)
    except Exception as exc: print(f"Warning: workbook pruning failed: {exc}")
    try: bundle=reconcile_score_displays(wb,ticker)
    except Exception as exc:
        print(f"Warning: final score reconciliation failed: {exc}"); bundle=None
    try: finalize_score_transparency(wb,ticker,fin.get("coverage") if isinstance(fin,dict) else None)
    except Exception as exc: print(f"Warning: score transparency finalization failed: {exc}")
    try: ensure_quality_checks(wb,ticker,bundle,removed)
    except Exception as exc: print(f"Warning: final Data Quality controls failed: {exc}")
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
