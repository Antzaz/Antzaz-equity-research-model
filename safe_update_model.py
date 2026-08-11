from __future__ import annotations

"""Guarded entry point for the deterministic equity model.

This wrapper installs source-integrity controls before importing/running the normal model:
- date-like spreadsheet cells cannot become financial amounts or fake fiscal years;
- cross-border annual history is sanitized against an independent annual-statement source;
- unfinished fiscal years are excluded from annual actuals;
- stale template filings/sources are cleared;
- alternate listings of the same company are deduplicated in peer discovery;
- core currency labels follow the traded security currency;
- a disk-space preflight avoids doing a full build that cannot be saved.

All downstream calculations remain in update_model.py.
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

BASE=Path(__file__).resolve().parent


def _safe_year(value):
    if isinstance(value,(datetime,date,pd.Timestamp)):
        return int(pd.Timestamp(value).year)
    if isinstance(value,bool) or value is None:
        return None
    if isinstance(value,(int,float)):
        n=int(value)
        return n if 1900<=n<=2100 and float(value)==n else None
    text=str(value).strip()
    if re.fullmatch(r"(?:19|20)\d{2}",text):
        return int(text)
    m=re.search(r"(?:19|20)\d{2}",text)
    return int(m.group(0)) if m else None


def _safe_number(value):
    if isinstance(value,(datetime,date,pd.Timestamp)):
        return None
    if isinstance(value,str):
        text=value.strip()
        # Explicit dates/timestamps are metadata, never financial values.
        if re.fullmatch(r"\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}(?:[ T].*)?",text):
            return None
        if re.fullmatch(r"(?:19|20)\d{2}-\d{2}-\d{2}.*",text):
            return None
    return _ORIGINAL_NUMBER(value)


_ORIGINAL_NUMBER=issuer_source_engine._number
_ORIGINAL_BUILD=issuer_source_engine.build_crossborder_history
_ORIGINAL_UPDATE_FILINGS=update_model.update_filings
_ORIGINAL_SELECT_PEERS=dynamic_peer_engine.select_dynamic_peers
_ORIGINAL_NORMALIZE=update_model.normalize_workbook_currency

# Fix the two permissive primitives used by issuer downloadable-table extraction.
issuer_source_engine._year=_safe_year
issuer_source_engine._number=_safe_number

# Explicit official issuer pages for companies where the generic website suffixes are weak.
issuer_source_engine.ISSUER_IR_PAGES.setdefault("SIE.DE",[
    "https://www.siemens.com/es-es/company/investor-relations/financial-results/",
    "https://www.siemens.com/en-us/company/investor-relations/financial-results-archive/",
])

# Prefer primary-company peers rather than alternate Siemens/Frankfurt listings.
dynamic_peer_engine.STRATEGIC_PEERS.setdefault("SIE.DE",[
    "SU.PA","ABBN.SW","HON","ETN","ROK","EMR","PH","GE","AME","IR"
])


def _safe_build_crossborder_history(ticker,info,facts):
    hist,meta=_ORIGINAL_BUILD(ticker,info,facts)
    hist,meta=data_integrity.sanitize_crossborder_history(ticker,info,hist,meta)
    return hist,meta


# model_reliability imported the function directly, so patch both module namespaces.
issuer_source_engine.build_crossborder_history=_safe_build_crossborder_history
model_reliability.build_crossborder_history=_safe_build_crossborder_history


def _safe_update_filings(wb,ticker):
    # The GOOGL template must never survive as a filing source when a foreign issuer has no SEC CIK.
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
    # A useful fallback for regional aliases when company names are unavailable.
    root=str(symbol).upper().split(".")[0]
    return name or root


def _safe_select_dynamic_peers(wb,ticker,count=9):
    target,sector,industry,raw_peers=_ORIGINAL_SELECT_PEERS(wb,ticker,max(count*3,18))
    target_key=_entity_key(ticker,target); seen={target_key}; peers=[]
    for symbol,info in raw_peers:
        key=_entity_key(symbol,info)
        if key in seen: continue
        seen.add(key); peers.append((symbol,info))
        if len(peers)>=count: break
    return target,sector,industry,peers


dynamic_peer_engine.select_dynamic_peers=_safe_select_dynamic_peers


def _safe_normalize_workbook_currency(wb,ticker,info):
    result=_ORIGINAL_NORMALIZE(wb,ticker,info)
    data_integrity.apply_workbook_integrity_controls(
        wb,ticker,info,getattr(wb,"_issuer_source_meta",None)
    )
    return result


# update_model imported this function directly; subsequent calls now include integrity controls.
update_model.normalize_workbook_currency=_safe_normalize_workbook_currency


def _disk_preflight(min_free_gb=1.5):
    free=shutil.disk_usage(BASE).free/(1024**3)
    if free<min_free_gb:
        raise OSError(
            f"Only {free:.2f} GB free on the project drive. Free at least {min_free_gb:.1f} GB "
            "before running the equity model so Git/Excel temporary files and the final workbook can be written safely."
        )
    print(f"Disk-space preflight: {free:.2f} GB free")


def main():
    _disk_preflight()
    update_model.main()


if __name__=="__main__":
    main()
