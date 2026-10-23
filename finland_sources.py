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

from datetime import timedelta
import json
import re
from urllib.parse import urljoin

import pandas as pd
import requests


FILINGS_XBRL_API = "https://filings.xbrl.org/api/filings"
FILINGS_XBRL_ROOT = "https://filings.xbrl.org"

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
    "op": (
        "ProfitLossFromOperatingActivities",
        "OperatingProfitLoss",
    ),
    "ni": (
        "ProfitLoss",
        "ProfitLossAttributableToOwnersOfParent",
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
        "PurchaseOfIntangibleAssets",
    ),
    "depr": (
        "DepreciationAndAmortisationExpense",
        "DepreciationAmortisationAndImpairmentExpense",
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
        "CurrentBorrowings",
        "NoncurrentBorrowings",
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
            if entity_name and not (candidate in entity_name or entity_name in candidate):
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
            if field in {"revenue", "op", "ni", "eps", "ocf", "capex", "depr", "sbc"}:
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
