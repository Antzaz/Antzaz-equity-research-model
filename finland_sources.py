from __future__ import annotations

"""Free Finland/Nordics public-data helpers for Helsinki-listed equity research.

The core project is intentionally provider-agnostic.  This module adds a Finland-specific
source layer without introducing paid dependencies:

- Finnish OAM (Nasdaq Helsinki): official regulated disclosures / ESEF annual reports.
- XBRL International filings.xbrl.org: free machine-readable ESEF mirror sourced from OAMs.
- Nasdaq Nordic company news: official exchange announcements.
- PRH open data: Finnish company/register data and available digital IXBRL statements.
- Statistics Finland PxWeb: official Finnish macro/industry statistics.
- Bank of Finland / FIN-FSA open data: official financial-system and rates data.
- Fingrid open data: electricity-system data; free registration/API key required.

For .HE issuers, the ESEF helper is used as a conservative structured-financial fallback.
It only consumes consolidated, non-dimensional facts and never invents missing values.
"""

from datetime import date, datetime, timedelta
from functools import lru_cache
from io import StringIO
import json
import re
from urllib.parse import urljoin

import pandas as pd
import requests
try:
    from bs4 import BeautifulSoup
except Exception:  # optional; requests/ESEF core still works without it
    BeautifulSoup = None


FILINGS_XBRL_API = "https://filings.xbrl.org/api/filings"
FILINGS_XBRL_ROOT = "https://filings.xbrl.org"
_ESEF_PAYLOAD_CACHE: dict[str, dict] = {}

FINLAND_PUBLIC_SOURCES = {
    "finland_oam": {
        "provider": "Finnish OAM / Nasdaq Helsinki",
        "url": "https://oam.fi/",
        "purpose": "Official archive of regulated disclosures and ESEF annual financial reports for Finnish issuers",
        "source_type": "Official regulated-information archive",
        "env_key": None,
    },
    "xbrl_international_esef": {
        "provider": "XBRL International filings.xbrl.org",
        "url": "https://filings.xbrl.org/api/filings",
        "purpose": "Free machine-readable ESEF filing index and xBRL-JSON facts sourced from European OAMs",
        "source_type": "Open structured-filing repository",
        "env_key": None,
    },
    "nasdaq_nordic_company_news": {
        "provider": "Nasdaq Nordic Company News",
        "url": "https://www.nasdaq.com/european-market-activity/news/company-news",
        "purpose": "Official exchange announcements, earnings releases, guidance changes and other regulated news",
        "source_type": "Exchange / regulated company news",
        "env_key": None,
    },
    "prh_open_data": {
        "provider": "Finnish Patent and Registration Office (PRH) Open Data",
        "url": "https://avoindata.prh.fi/en",
        "purpose": "Finnish company/register information, registered notifications and available digital IXBRL financial statements",
        "source_type": "Official company-register open data",
        "env_key": None,
    },
    "statistics_finland": {
        "provider": "Statistics Finland PxWeb",
        "url": "https://pxdata.stat.fi/PxWeb/pxweb/en/",
        "purpose": "Official Finnish industry, employment, construction, trade, prices and other macro statistics",
        "source_type": "Official statistics / open API",
        "env_key": None,
    },
    "bank_of_finland": {
        "provider": "Bank of Finland / FIN-FSA Open Data",
        "url": "https://www.suomenpankki.fi/en/statistics/open-data",
        "purpose": "Finnish rates, banking, credit, deposits and financial-system time series",
        "source_type": "Central bank / supervisor open data",
        "env_key": None,
    },
    "fingrid": {
        "provider": "Fingrid Open Data",
        "url": "https://data.fingrid.fi/en",
        "purpose": "Finnish electricity-system, generation, consumption and market data",
        "source_type": "Official electricity-system open API",
        "env_key": "FINGRID_API_KEY",
        "free_registration": True,
    },
}


# Standard IFRS/ESEF concepts commonly present in consolidated annual reports.  Company-extension
# concepts are deliberately not guessed here; unsupported fields remain blank and can still be
# recovered from issuer IR/Yahoo fallbacks elsewhere in the project.
ESEF_CONCEPT_ALIASES = {
    "revenue": (
        "Revenue",
        "RevenueFromContractsWithCustomers",
        "RevenueFromContractsWithCustomersIncludingAssessedTax",
    ),
    "cost": (
        "CostOfSales",
    ),
    "gross": (
        "GrossProfit",
    ),
    "op": (
        "ProfitLossFromOperatingActivities",
        "OperatingProfitLoss",
        "OperatingProfit",
        "OperatingResult",
    ),
    "ni": (
        "ProfitLoss",
        "ProfitLossAttributableToOwnersOfParent",
        "ResultForPeriod",
    ),
    "eps": (
        "DilutedEarningsLossPerShare",
        "BasicAndDilutedEarningsLossPerShare",
    ),
    "ocf": (
        "CashFlowsFromUsedInOperatingActivities",
    ),
    "capex": (
        "PurchaseOfPropertyPlantAndEquipment",
    ),
    "depr": (
        "DepreciationAndAmortisationExpense",
        "DepreciationAmortisationAndImpairmentExpense",
    ),
    "rd": (
        "ResearchAndDevelopmentExpense",
    ),
    "sbc": (
        "ShareBasedPaymentExpense",
    ),
    "cash": (
        "CashAndCashEquivalents",
    ),
    "assets": (
        "Assets",
    ),
    "liabilities": (
        "Liabilities",
    ),
    "equity": (
        "Equity",
        "EquityAttributableToOwnersOfParent",
    ),
    "debt": (
        "Borrowings",
    ),
}


def is_finnish_ticker(ticker: str) -> bool:
    """True for Yahoo-style Nasdaq Helsinki symbols such as KNEBV.HE or NOKIA.HE."""
    return str(ticker or "").upper().strip().endswith(".HE")


def finland_public_sources() -> dict[str, dict]:
    return {key: dict(value) for key, value in FINLAND_PUBLIC_SOURCES.items()}


def _norm_name(value: str) -> str:
    text = str(value or "").lower()
    text = re.sub(r"\b(oyj|plc|oy|abp|corporation|corp|group|class [ab])\b", " ", text)
    return re.sub(r"[^a-z0-9åäö]+", " ", text).strip()


def _name_candidates(ticker: str, info: dict | None) -> list[str]:
    values = [
        (info or {}).get("longName"),
        (info or {}).get("shortName"),
    ]
    out = []
    for value in values:
        name = _norm_name(value)
        if len(name) >= 3 and name not in out:
            out.append(name)
    # Ticker root is only a final weak fallback for names such as NOKIA.HE.
    root = re.sub(r"\.HE$", "", str(ticker or "").upper().strip())
    root = _norm_name(root)
    if len(root) >= 4 and root not in out:
        out.append(root)
    return out


def _request_json(url: str, *, params: dict | None = None, timeout: int = 20) -> dict:
    response = requests.get(
        url,
        params=params,
        timeout=timeout,
        headers={
            "Accept": "application/vnd.api+json, application/json",
            "User-Agent": "AntzazEquityResearch/1.0 (public ESEF research)",
        },
    )
    response.raise_for_status()
    return response.json()


def _filing_json_urls(ticker: str, info: dict | None = None, limit: int = 8) -> tuple[list[str], dict]:
    """Find recent Finnish ESEF xBRL-JSON URLs using company-name filters.

    filings.xbrl.org uses JSON:API.  The related entity name is filtered server-side and then
    validated locally so a fuzzy result cannot silently become another company's history.
    """
    if not is_finnish_ticker(ticker):
        return [], {"status": "NOT_APPLICABLE", "ticker": ticker}

    candidates = _name_candidates(ticker, info)
    attempts = []
    for candidate in candidates:
        filt = [
            {"name": "country", "op": "eq", "val": "FI"},
            {"name": "entity.name", "op": "ilike", "val": f"%{candidate}%"},
        ]
        params = {
            "filter": json.dumps(filt, separators=(",", ":")),
            "include": "entity",
            "page[size]": max(1, min(int(limit), 20)),
            "sort": "-period_end",
        }
        try:
            payload = _request_json(FILINGS_XBRL_API, params=params)
        except Exception as exc:
            attempts.append({"candidate": candidate, "error": str(exc)[:240]})
            continue

        included = {
            str(x.get("id")): (x.get("attributes") or {})
            for x in payload.get("included", [])
            if x.get("type") == "entity"
        }
        urls = []
        filings = []
        for item in payload.get("data", []):
            attrs = item.get("attributes") or {}
            rel = (((item.get("relationships") or {}).get("entity") or {}).get("data") or {})
            entity = included.get(str(rel.get("id")), {})
            entity_name = _norm_name(entity.get("name"))
            # Require a meaningful local name match; server-side ilike alone is not enough.
            if entity_name and entity_name != candidate:
                continue
            raw_url = attrs.get("json_url")
            if raw_url:
                url = urljoin(FILINGS_XBRL_ROOT + "/", str(raw_url))
                if url not in urls:
                    urls.append(url)
            filings.append(
                {
                    "entity": entity.get("name"),
                    "lei": entity.get("identifier"),
                    "period_end": attrs.get("period_end"),
                    "published": attrs.get("published"),
                    "json_url": raw_url,
                    "viewer_url": attrs.get("viewer_url"),
                }
            )
        attempts.append({"candidate": candidate, "matches": len(urls)})
        if urls:
            return urls[:limit], {
                "status": "PASS",
                "ticker": ticker,
                "matched_name": candidate,
                "filings": filings[:limit],
                "attempts": attempts,
            }

    return [], {
        "status": "NO_MATCH",
        "ticker": ticker,
        "candidates": candidates,
        "attempts": attempts,
    }


def _concept_local_name(qname: str) -> str:
    return str(qname or "").split(":")[-1]


def _period_year(period: str) -> tuple[int | None, int | None]:
    """Return (year, duration_days). xBRL periods end at the start of the following day."""
    try:
        parts = str(period or "").split("/")
        if len(parts) == 2:
            start = pd.Timestamp(parts[0])
            end = pd.Timestamp(parts[1])
            return int((end - timedelta(days=1)).year), int((end - start).days)
        if len(parts) == 1 and parts[0]:
            end = pd.Timestamp(parts[0])
            return int((end - timedelta(days=1)).year), None
    except Exception:
        pass
    return None, None


def _numeric(value):
    try:
        if value in (None, ""):
            return None
        return float(value)
    except Exception:
        return None


def _consolidated_fact(fact: dict) -> bool:
    dims = fact.get("dimensions") or {}
    allowed = {"concept", "entity", "period", "unit", "language"}
    return not any(key not in allowed for key in dims)


def _extract_esef_report(payload: dict) -> tuple[dict[int, dict], dict]:
    facts = payload.get("facts") or {}
    out: dict[int, dict] = {}
    selected: dict[int, dict] = {}
    alias_rank = {
        field: {name: rank for rank, name in enumerate(names)}
        for field, names in ESEF_CONCEPT_ALIASES.items()
    }

    for fact in facts.values():
        if not isinstance(fact, dict) or not _consolidated_fact(fact):
            continue
        dims = fact.get("dimensions") or {}
        concept = _concept_local_name(dims.get("concept"))
        unit = str(dims.get("unit") or "")
        year, duration_days = _period_year(dims.get("period"))
        if year is None:
            continue
        value = _numeric(fact.get("value"))
        if value is None:
            continue

        for field, ranks in alias_rank.items():
            if concept not in ranks:
                continue
            # Income/cash-flow fields should represent approximately a full year. Balance-sheet
            # instants do not have a duration. EPS may have a shares unit; other values should be
            # monetary when a unit is supplied.
            if field in {"revenue", "cost", "gross", "op", "ni", "eps", "ocf", "capex", "depr", "rd", "sbc"}:
                if duration_days is None or duration_days < 250 or duration_days > 450:
                    continue
            if field != "eps" and unit and "iso4217:" not in unit:
                continue

            rank = ranks[concept]
            key = (year, field)
            previous = selected.get(key)
            if previous is not None and previous["rank"] <= rank:
                continue
            if field == "capex":
                value = abs(value)
            selected[key] = {"rank": rank, "concept": concept}
            out.setdefault(year, {})[field] = value

    # Borrowings are commonly split current/non-current. If both appeared in a year, the exact
    # alias selection above may only retain one. This remains conservative: downstream logic can
    # still recover total debt from Yahoo rather than fabricate the missing component.
    for year, values in out.items():
        if values.get("ni") is not None and values.get("eps") not in (None, 0):
            values.setdefault("shares", values["ni"] / values["eps"])
        if values.get("ocf") is not None and values.get("capex") is not None:
            values.setdefault("fcf", values["ocf"] - abs(values["capex"]))

    return out, {
        "years": sorted(out),
        "selected_concepts": {
            f"{year}:{field}": detail["concept"]
            for (year, field), detail in selected.items()
        },
    }


def finland_esef_filings(ticker: str, info: dict | None = None, limit: int = 12) -> tuple[list[dict], dict]:
    """Return recent Finnish ESEF filing metadata for workbook filing registers."""
    urls, discovery = _filing_json_urls(ticker, info, limit=limit)
    rows = []
    for item in discovery.get("filings", []) if isinstance(discovery, dict) else []:
        raw_json = item.get("json_url")
        raw_viewer = item.get("viewer_url")
        rows.append(
            {
                "form": "ESEF AFR",
                "period_end": item.get("period_end"),
                "published": item.get("published"),
                "url": (
                    urljoin(FILINGS_XBRL_ROOT + "/", str(raw_viewer))
                    if raw_viewer
                    else (
                        urljoin(FILINGS_XBRL_ROOT + "/", str(raw_json))
                        if raw_json
                        else "https://oam.fi/"
                    )
                ),
                "note": "Finnish annual financial report — ESEF / official OAM source mirrored by XBRL International",
                "entity": item.get("entity"),
                "lei": item.get("lei"),
            }
        )
    # Discovery can occasionally expose fewer metadata rows than JSON URLs. Preserve a usable
    # filing link rather than silently losing a machine-readable report.
    known={row["url"] for row in rows}
    for url in urls:
        if url not in known:
            rows.append(
                {
                    "form": "ESEF AFR",
                    "period_end": None,
                    "published": None,
                    "url": url,
                    "note": "Finnish annual financial report — machine-readable ESEF xBRL-JSON",
                    "entity": None,
                    "lei": None,
                }
            )
    return rows[:limit], discovery


def finland_esef_history(ticker: str, info: dict | None = None, limit: int = 8) -> tuple[dict, dict]:
    """Return annual financial history from free Finnish ESEF filings when available.

    The source is XBRL International's machine-readable mirror of official OAM filings.  Missing
    concepts are left blank so the existing issuer-IR/Yahoo recovery path can fill them later.
    """
    if not is_finnish_ticker(ticker):
        return {}, {
            "source": "XBRL International ESEF / Finnish OAM",
            "status": "NOT_APPLICABLE",
            "years": [],
        }

    urls, discovery = _filing_json_urls(ticker, info, limit=limit)
    merged: dict[int, dict] = {}
    provenance: dict[int, dict] = {}
    errors = []

    for url in urls:
        try:
            payload = _request_json(url, timeout=30)
            hist, meta = _extract_esef_report(payload)
        except Exception as exc:
            errors.append(f"{url}: {exc}"[:300])
            continue
        for year, values in hist.items():
            for field, value in values.items():
                if value is None:
                    continue
                if field not in merged.setdefault(year, {}):
                    merged[year][field] = value
                    provenance.setdefault(year, {})[field] = url

    years = sorted(
        year for year, values in merged.items()
        if values.get("revenue") is not None or values.get("ni") is not None
    )
    if len(years) > 6:
        keep = set(years[-6:])
        merged = {y: v for y, v in merged.items() if y in keep}
        provenance = {y: v for y, v in provenance.items() if y in keep}
        years = sorted(keep)

    return merged, {
        "source": "XBRL International ESEF xBRL-JSON sourced from Finnish OAM filings",
        "status": "PASS" if merged else discovery.get("status", "NO_DATA"),
        "years": years,
        "discovery": discovery,
        "json_urls": urls,
        "provenance": provenance,
        "errors": errors[:10],
    }


def _esef_payload(url: str, timeout: int = 30) -> dict:
    """Cache xBRL-JSON within one research run so history/statements/segments reuse downloads."""
    key=str(url)
    if key not in _ESEF_PAYLOAD_CACHE:
        _ESEF_PAYLOAD_CACHE[key]=_request_json(key, timeout=timeout)
    return _ESEF_PAYLOAD_CACHE[key]


def _xbrl_unit_to_companyfacts(unit: str) -> str:
    raw=str(unit or "")
    low=raw.lower()
    currency=None
    m=re.search(r"iso4217:([A-Za-z]{3})",raw)
    if m:
        currency=m.group(1).upper()
    if "share" in low and currency:
        return f"{currency}/shares"
    if "share" in low:
        return "shares"
    return currency or raw or "pure"


def _xbrl_period_to_companyfacts(period: str) -> tuple[str | None, str | None, int | None]:
    """Translate xBRL-JSON half-open periods to SEC-CompanyFacts-like inclusive dates."""
    try:
        parts=str(period or "").split("/")
        if len(parts)==2:
            start=pd.Timestamp(parts[0]).date()
            end_exclusive=pd.Timestamp(parts[1]).date()
            end=end_exclusive-timedelta(days=1)
            return start.isoformat(),end.isoformat(),int(end.year)
        if len(parts)==1 and parts[0]:
            instant=pd.Timestamp(parts[0]).date()-timedelta(days=1)
            return None,instant.isoformat(),int(instant.year)
    except Exception:
        pass
    return None,None,None


def _esef_companyfacts_from_payload(payload: dict, source_url: str | None = None) -> dict:
    """Convert consolidated standard-IFRS xBRL-JSON facts to the internal CompanyFacts shape.

    Only ifrs-full facts without segment/other dimensions are included.  Issuer-extension facts
    are intentionally left to dedicated parsers because guessing their accounting meaning would
    create false precision.
    """
    root={"facts":{"ifrs-full":{}},"_source_type":"Finnish OAM / ESEF","_source_url":source_url}
    for fact in (payload.get("facts") or {}).values():
        if not isinstance(fact,dict) or not _consolidated_fact(fact):
            continue
        dims=fact.get("dimensions") or {}
        qname=str(dims.get("concept") or "")
        if not qname.startswith("ifrs-full:"):
            continue
        tag=_concept_local_name(qname)
        value=_numeric(fact.get("value"))
        if value is None:
            continue
        start,end,fy=_xbrl_period_to_companyfacts(dims.get("period"))
        if end is None or fy is None:
            continue
        unit=_xbrl_unit_to_companyfacts(dims.get("unit"))
        item={"val":value,"end":end,"fy":fy,"form":"ESEF","filed":end}
        if start:
            item["start"]=start
        node=root["facts"]["ifrs-full"].setdefault(
            tag,
            {"label":tag,"description":"IFRS/ESEF fact from Finnish OAM filing","units":{}},
        )
        bucket=node["units"].setdefault(unit,[])
        signature=(item.get("start"),item.get("end"),item.get("val"))
        if not any((x.get("start"),x.get("end"),x.get("val"))==signature for x in bucket):
            bucket.append(item)
    return root


def finland_esef_companyfacts(ticker: str, info: dict | None = None, limit: int = 8) -> tuple[dict, dict]:
    """Return Finnish ESEF facts in the shape expected by the full-statement renderer."""
    if not is_finnish_ticker(ticker):
        return {},{"status":"NOT_APPLICABLE","ticker":ticker}
    urls,discovery=_filing_json_urls(ticker,info,limit=limit)
    merged={"facts":{"ifrs-full":{}},"_source_type":"Finnish OAM / ESEF","_source_urls":urls}
    errors=[]
    for url in urls:
        try:
            converted=_esef_companyfacts_from_payload(_esef_payload(url),url)
        except Exception as exc:
            errors.append(f"{url}: {exc}"[:300])
            continue
        for tag,node in ((converted.get("facts") or {}).get("ifrs-full") or {}).items():
            dest=merged["facts"]["ifrs-full"].setdefault(
                tag,
                {"label":node.get("label") or tag,"description":node.get("description") or "","units":{}},
            )
            for unit,items in (node.get("units") or {}).items():
                bucket=dest["units"].setdefault(unit,[])
                seen={(x.get("start"),x.get("end"),x.get("val")) for x in bucket}
                for item in items:
                    sig=(item.get("start"),item.get("end"),item.get("val"))
                    if sig not in seen:
                        bucket.append(item); seen.add(sig)
    tags=sorted(merged["facts"]["ifrs-full"])
    return merged,{
        "status":"PASS" if tags else discovery.get("status","NO_DATA"),
        "source":"Finnish OAM / ESEF via XBRL International",
        "tags":tags,
        "tag_count":len(tags),
        "json_urls":urls,
        "discovery":discovery,
        "errors":errors[:10],
    }


_SEGMENT_AXIS_HINTS=("segment","reportable","operatingsegment","businesssegment")
_SEGMENT_EXCLUDE=("consolidated","group","elimination","unallocated","reconciliation","total")


def _humanize_member(qname: str) -> str:
    name=_concept_local_name(qname)
    name=re.sub(r"(Member|Segment)$","",name)
    name=re.sub(r"([a-zåäö0-9])([A-ZÅÄÖ])",r"\1 \2",name)
    name=re.sub(r"[_\-]+"," ",name)
    return re.sub(r"\s+"," ",name).strip()


def _segment_member(dims: dict) -> str | None:
    allowed={"concept","entity","period","unit","language"}
    for axis,member in dims.items():
        if axis in allowed:
            continue
        axis_local=_concept_local_name(axis).lower()
        if any(h in axis_local for h in _SEGMENT_AXIS_HINTS):
            label=_humanize_member(member)
            low=label.lower()
            if label and not any(x in low for x in _SEGMENT_EXCLUDE):
                return label
    return None


def _segment_metric(concept_qname: str) -> tuple[str | None,int]:
    """Map standard or issuer-extension segment concepts conservatively."""
    local=_concept_local_name(concept_qname)
    low=re.sub(r"[^a-z0-9]","",local.lower())
    if local in ESEF_CONCEPT_ALIASES["revenue"]:
        return "revenue",100
    if "revenue" in low and not any(x in low for x in ("growth","percentage","share")):
        return "revenue",75
    if "adjustedoperatingprofit" in low or "adjustedoperatingresult" in low:
        return "op",100
    if local in ESEF_CONCEPT_ALIASES["op"]:
        return "op",90
    if ("operatingprofit" in low or "segmentprofit" in low) and not any(x in low for x in ("margin","percentage")):
        return "op",70
    return None,0


def _extract_esef_segments(payload: dict) -> dict:
    """Extract dimensioned annual segment revenue/profit without treating segments as group facts."""
    best={}
    for fact in (payload.get("facts") or {}).values():
        if not isinstance(fact,dict):
            continue
        dims=fact.get("dimensions") or {}
        member=_segment_member(dims)
        if not member:
            continue
        field,score=_segment_metric(dims.get("concept"))
        if not field:
            continue
        unit=str(dims.get("unit") or "")
        if unit and "iso4217:" not in unit:
            continue
        year,duration=_period_year(dims.get("period"))
        if year is None or duration is None or not 250<=duration<=450:
            continue
        value=_numeric(fact.get("value"))
        if value is None:
            continue
        key=(member,year,field)
        old=best.get(key)
        if old is None or score>old[0]:
            best[key]=(score,value,_concept_local_name(dims.get("concept")))
    segments={}
    for (member,year,field),(score,value,concept) in best.items():
        rec=segments.setdefault(member,{"name":member,"revenue":{},"op":{},"concepts":{}})
        rec[field][year]=value
        rec["concepts"][f"{year}:{field}"]=concept
    return segments


def finland_esef_segments(ticker: str, info: dict | None = None, limit: int = 4) -> tuple[list[dict], dict]:
    """Return recent operating-segment economics from dimensioned Finnish ESEF facts."""
    if not is_finnish_ticker(ticker):
        return [],{"status":"NOT_APPLICABLE","ticker":ticker}
    urls,discovery=_filing_json_urls(ticker,info,limit=limit)
    merged={}
    errors=[]
    for url in urls:  # newest filing first; retain first observation of an overlapping year
        try:
            found=_extract_esef_segments(_esef_payload(url))
        except Exception as exc:
            errors.append(f"{url}: {exc}"[:300]); continue
        for name,rec in found.items():
            dest=merged.setdefault(name,{"name":name,"revenue":{},"op":{},"concepts":{},"source":url})
            for field in ("revenue","op"):
                for year,value in rec[field].items():
                    dest[field].setdefault(year,value)
            dest["concepts"].update({k:v for k,v in rec["concepts"].items() if k not in dest["concepts"]})
    # Keep entities with useful segment economics and a recent annual observation.
    rows=[x for x in merged.values() if x["revenue"] or x["op"]]
    rows.sort(key=lambda x:max(list(x["revenue"])+list(x["op"]) or [0]),reverse=True)
    years=sorted({y for x in rows for field in ("revenue","op") for y in x[field]})
    return rows,{
        "status":"PASS" if rows else discovery.get("status","NO_DATA"),
        "source":"Finnish OAM / ESEF dimensioned segment facts",
        "json_urls":urls,
        "years":years[-3:],
        "errors":errors[:10],
        "discovery":discovery,
    }


def _issuer_candidate_pages(info: dict | None, suffixes: tuple[str,...]) -> list[str]:
    website=str((info or {}).get("website") or "").strip().rstrip("/")
    if not website:
        return []
    pages=[website+s for s in suffixes]
    # Also inspect investor landing pages for canonical links.
    for landing in (website+"/investors",website+"/en/investors",website+"/investor-relations"):
        try:
            r=requests.get(landing,timeout=12,headers={"User-Agent":"Mozilla/5.0 AntzazEquityResearch"})
            if not r.ok or BeautifulSoup is None:
                continue
            soup=BeautifulSoup(r.text,"html.parser")
            for a in soup.find_all("a",href=True):
                href=urljoin(r.url,a.get("href"))
                text=(a.get_text(" ",strip=True)+" "+href).lower()
                if any(token in text for token in ("shareholder","ownership","share information","stock exchange release","company announcement","releases")):
                    pages.append(href)
        except Exception:
            continue
    return list(dict.fromkeys(pages))


def finland_major_shareholders(ticker: str, info: dict | None = None, limit: int = 20) -> tuple[list[dict], dict]:
    """Best-effort issuer/Euroclear shareholder table for Helsinki-listed companies."""
    if not is_finnish_ticker(ticker):
        return [],{"status":"NOT_APPLICABLE"}
    pages=_issuer_candidate_pages(info,(
        "/investors/share-and-shareholders","/en/investors/share-and-shareholders",
        "/investors/shareholders","/en/investors/shareholders",
        "/investors/share-information","/en/investors/share-information",
        "/investor-relations/shareholders","/en/investor-relations/shareholders",
    ))
    errors=[]
    for page in pages[:12]:
        try:
            r=requests.get(page,timeout=15,headers={"User-Agent":"Mozilla/5.0 AntzazEquityResearch"})
            r.raise_for_status()
            page_text=(BeautifulSoup(r.text,"html.parser").get_text(" ",strip=True) if BeautifulSoup is not None else re.sub(r"<[^>]+>"," ",r.text))
            shares_outstanding=None
            sm=re.search(r"(?:total\s+)?number\s+of\s+shares.{0,80}?([0-9][0-9\s,.]{5,})",page_text,re.I)
            if sm:
                try: shares_outstanding=float(re.sub(r"[^0-9.]","",sm.group(1).replace(",","")))
                except Exception: shares_outstanding=None
            tables=pd.read_html(StringIO(r.text))
        except Exception as exc:
            errors.append(f"{page}: {exc}"[:240]); continue
        for df in tables:
            if df.empty:
                continue
            cols=[str(x).lower() for x in df.columns]
            text=" ".join(cols)+" "+" ".join(str(x).lower() for x in df.head(3).astype(str).values.flatten())
            if not any(x in text for x in ("shareholder","owner","osakkeenomist")):
                continue
            records=[]
            for _,row in df.iterrows():
                vals=[x for x in row.tolist()]
                if not vals:
                    continue
                holder=str(vals[0]).strip()
                if not holder or holder.lower() in {"nan","total","yhteensä"}:
                    continue
                nums=[]
                for v in vals[1:]:
                    if isinstance(v,(int,float)) and not pd.isna(v):
                        nums.append(float(v)); continue
                    s=str(v).replace("\xa0"," ").replace(" ","").replace(",",".")
                    m=re.search(r"-?\d+(?:\.\d+)?",s)
                    if m:
                        try: nums.append(float(m.group(0)))
                        except Exception: pass
                shares=next((n for n in nums if n>1000),None)
                pct=next((n/100 if n>1 else n for n in reversed(nums) if 0<n<=100),None)
                records.append({"holder":holder,"shares":shares,"pct":pct,"source":page})
            if len(records)>=3:
                return records[:limit],{"status":"PASS","source":page,"provider":"Issuer / Euroclear Finland","pages_checked":pages[:12],"shares_outstanding":shares_outstanding}
    return [],{"status":"NO_TABLE","pages_checked":pages[:12],"errors":errors[:10]}


def finland_company_news(ticker: str, info: dict | None = None, horizon_days: int = 730, limit: int = 80) -> tuple[list[dict], dict]:
    """Collect issuer stock-exchange/company releases for Finnish names when Yahoo coverage is thin."""
    if not is_finnish_ticker(ticker):
        return [],{"status":"NOT_APPLICABLE"}
    pages=_issuer_candidate_pages(info,(
        "/news-repository/stock-exchange-releases","/en/news-repository/stock-exchange-releases",
        "/investors/releases","/en/investors/releases",
        "/investors/stock-exchange-releases","/en/investors/stock-exchange-releases",
        "/investor-relations/releases","/en/investor-relations/releases",
    ))
    cutoff=date.today()-timedelta(days=int(horizon_days))
    out=[]; seen=set(); errors=[]
    date_re=re.compile(r"(20\d{2})[-./](\d{1,2})[-./](\d{1,2})")
    for page in pages[:12]:
        try:
            r=requests.get(page,timeout=15,headers={"User-Agent":"Mozilla/5.0 AntzazEquityResearch"})
            r.raise_for_status()
            if BeautifulSoup is None:
                continue
            soup=BeautifulSoup(r.text,"html.parser")
        except Exception as exc:
            errors.append(f"{page}: {exc}"[:240]); continue
        for a in soup.find_all("a",href=True):
            title=re.sub(r"\s+"," ",a.get_text(" ",strip=True)).strip()
            href=urljoin(r.url,a.get("href"))
            if len(title)<8 or href in seen:
                continue
            blob=re.sub(r"\s+"," ",a.parent.get_text(" ",strip=True) if a.parent else title)
            m=date_re.search(blob)
            published=None
            if m:
                try: published=date(int(m.group(1)),int(m.group(2)),int(m.group(3)))
                except Exception: published=None
            if published and published<cutoff:
                continue
            low=(title+" "+href).lower()
            if not any(k in low for k in ("release","acqui","divest","agreement","contract","invest","transaction","strategy","order","partnership","joint venture","guidance","result")):
                continue
            seen.add(href)
            out.append({
                "title":title,
                "summary":blob[:700],
                "date":published.isoformat() if published else None,
                "publisher":(info or {}).get("shortName") or (info or {}).get("longName") or ticker,
                "url":href,
                "source_kind":"press_release",
            })
            if len(out)>=limit:
                break
        if len(out)>=limit:
            break
    out.sort(key=lambda x:x.get("date") or "",reverse=True)
    return out,{"status":"PASS" if out else "NO_ITEMS","pages_checked":pages[:12],"errors":errors[:10],"source":"Issuer stock-exchange/company releases"}


def _table_number(value):
    if value in (None,""):
        return None
    if isinstance(value,(int,float)) and not pd.isna(value):
        return float(value)
    text=str(value).strip().replace("\u2212","-").replace("\xa0"," ")
    if text.lower() in {"nan","-","—","n/a","na"}:
        return None
    neg=text.startswith("(") and text.endswith(")")
    text=text.strip("()").replace(" ","")
    # European tables often use comma thousands separators; decimal-comma values are usually
    # short and contain no thousands grouping. Preserve both forms conservatively.
    if re.fullmatch(r"-?\d{1,3}(?:,\d{3})+",text):
        text=text.replace(",","")
    elif re.fullmatch(r"-?\d+,\d+",text):
        text=text.replace(",",".")
    text=re.sub(r"[^0-9.\-]","",text)
    try:
        out=float(text)
        return -out if neg else out
    except Exception:
        return None


def _year_from_header(value) -> int | None:
    m=re.search(r"\b(20\d{2})\b",str(value or ""))
    return int(m.group(1)) if m else None


def finland_key_figures(ticker: str, info: dict | None = None, limit: int = 40) -> tuple[list[dict], dict]:
    """Extract issuer-owned annual KPI tables while preserving issuer-defined metric names."""
    if not is_finnish_ticker(ticker):
        return [],{"status":"NOT_APPLICABLE"}
    pages=_issuer_candidate_pages(info,(
        "/investors/financial-information","/en/investors/financial-information",
        "/investors/key-figures","/en/investors/key-figures",
        "/investor-relations/financial-information","/en/investor-relations/financial-information",
        "/investors/financials","/en/investors/financials",
    ))
    errors=[]
    for page in pages[:12]:
        try:
            r=requests.get(page,timeout=18,headers={"User-Agent":"Mozilla/5.0 AntzazEquityResearch"})
            r.raise_for_status()
            tables=pd.read_html(StringIO(r.text))
        except Exception as exc:
            errors.append(f"{page}: {exc}"[:240]); continue
        best=[]
        for df in tables:
            if df is None or df.empty or df.shape[1]<3:
                continue
            x=df.copy()
            if isinstance(x.columns,pd.MultiIndex):
                x.columns=[" ".join(str(z) for z in col if str(z)!="nan").strip() for col in x.columns]
            else:
                x.columns=[str(col) for col in x.columns]
            year_cols={col:_year_from_header(col) for col in x.columns}
            year_cols={col:y for col,y in year_cols.items() if y}
            if len(set(year_cols.values()))<2:
                continue
            label_col=next((col for col in x.columns if col not in year_cols),x.columns[0])
            rows=[]
            for _,row in x.iterrows():
                metric=re.sub(r"\s+"," ",str(row.get(label_col) or "")).strip()
                if not metric or metric.lower() in {"nan","total"}:
                    continue
                values={}
                for col,year in year_cols.items():
                    num=_table_number(row.get(col))
                    if num is not None:
                        values[year]=num
                if values:
                    rows.append({"metric":metric,"values":values,"source":page})
            if len(rows)>=3:
                best=rows
                break
        if best:
            return best[:limit],{"status":"PASS","source":page,"rows":len(best),"provider":"Issuer investor relations"}
    return [],{"status":"NO_TABLE","pages_checked":pages[:12],"errors":errors[:10]}


def finland_latest_results(ticker: str, info: dict | None = None, horizon_days: int = 500, limit: int = 50) -> tuple[dict, dict]:
    """Extract the latest issuer-hosted interim/results table for Helsinki companies."""
    if not is_finnish_ticker(ticker):
        return {},{"status":"NOT_APPLICABLE"}
    items,news_meta=finland_company_news(ticker,info,horizon_days,limit=100)
    candidates=[]
    for item in items:
        title=str(item.get("title") or "")
        low=title.lower()
        if any(k in low for k in ("interim report","half-year","half year","financial statements bulletin","results","business review")):
            candidates.append(item)
    errors=[]
    for item in candidates[:12]:
        url=item.get("url")
        if not url:
            continue
        try:
            r=requests.get(url,timeout=18,headers={"User-Agent":"Mozilla/5.0 AntzazEquityResearch"})
            r.raise_for_status()
            tables=pd.read_html(StringIO(r.text))
        except Exception as exc:
            errors.append(f"{url}: {exc}"[:240]); continue
        for df in tables:
            if df is None or df.empty or df.shape[1]<2:
                continue
            x=df.copy()
            if isinstance(x.columns,pd.MultiIndex):
                x.columns=[" | ".join(str(z) for z in col if str(z)!="nan").strip() for col in x.columns]
            else:
                x.columns=[str(col) for col in x.columns]
            first=x.columns[0]
            labels=[str(v).lower() for v in x[first].head(30).tolist()]
            score=sum(any(k in lab for k in ("revenue","operating profit","earnings per share","order book","net interest-bearing debt","cash flow")) for lab in labels)
            if score<2:
                continue
            rows=[]
            for _,row in x.iterrows():
                metric=re.sub(r"\s+"," ",str(row.get(first) or "")).strip()
                if not metric or metric.lower()=="nan":
                    continue
                vals=[]
                for col in x.columns[1:]:
                    raw=row.get(col)
                    vals.append({"column":str(col),"value":_table_number(raw),"raw":None if pd.isna(raw) else str(raw)})
                if any(v["value"] is not None for v in vals):
                    rows.append({"metric":metric,"values":vals})
            if rows:
                return {
                    "title":item.get("title"),
                    "date":item.get("date"),
                    "url":url,
                    "columns":[str(x) for x in x.columns[1:]],
                    "rows":rows[:limit],
                },{"status":"PASS","source":url,"provider":"Issuer stock-exchange release","news":news_meta}
    return {},{"status":"NO_TABLE","candidates":len(candidates),"errors":errors[:10],"news":news_meta}
