"""Central source registry for equity research.

Keep issuer-owned and specialist research URLs in one place so workbook, agents and
source-health checks do not maintain duplicate hard-coded lists. Financial statements
still follow the model's primary-source hierarchy: issuer disclosures -> regulator/XBRL
-> transparent market-data fallback.

Specialist market sources are context sources only. A market-share percentage may enter
Peer Comps only when the same source and market definition are comparable across the
companies being compared.
"""

from __future__ import annotations

from typing import Iterable


ISSUER_SOURCES = {
    "GOOGL": {
        "investor": "https://abc.xyz/investor/",
        "earnings": "https://abc.xyz/investor/earnings/",
        "annual_reports": "https://abc.xyz/investor/annual-reports/",
        "governance": "https://abc.xyz/investor/board-and-governance/",
        "workplace_governance": "https://abc.xyz/investor/board-and-governance/ldicc/default.aspx",
        "additional_information": "https://abc.xyz/investor/additional-information/default.aspx",
    },
    "MSFT": {
        "investor": "https://www.microsoft.com/en-us/investor",
        "investor_information": "https://www.microsoft.com/en-us/investor/investor-information",
        "annual_reports": "https://www.microsoft.com/en-us/investor/annual-reports",
        "filings": "https://www.microsoft.com/en-us/investor/sec-filings",
    },
    "AMZN": {
        "investor": "https://ir.aboutamazon.com/",
        "quarterly_results": "https://ir.aboutamazon.com/quarterly-results/default.aspx",
        "filings": "https://ir.aboutamazon.com/sec-filings/default.aspx",
        "annual_reports": "https://ir.aboutamazon.com/annual-reports-proxies-and-shareholder-letters/default.aspx",
    },
    "META": {
        "investor": "https://investor.atmeta.com/home/default.aspx",
        "financials": "https://investor.atmeta.com/financials/default.aspx",
    },
    "NVDA": {
        "investor": "https://investor.nvidia.com/",
        "financial_reports": "https://investor.nvidia.com/financial-info/financial-reports/default.aspx",
        "annual_reports": "https://investor.nvidia.com/financial-info/annual-reports-and-proxies/default.aspx",
        "filings": "https://investor.nvidia.com/financial-info/sec-filings/default.aspx",
    },
    "AAPL": {
        "investor": "https://investor.apple.com/investor-relations/",
        "filings": "https://investor.apple.com/sec-filings/default.aspx",
        "governance": "https://investor.apple.com/leadership-and-governance/default.aspx",
    },
    "TSM": {
        "investor": "https://investor.tsmc.com/english",
        "financial_reports": "https://investor.tsmc.com/english/financial-reports",
        "quarterly_results": "https://investor.tsmc.com/english/quarterly-results",
        "annual_reports": "https://investor.tsmc.com/static/annualReports/2025/english/index.html",
        "filings": "https://investor.tsmc.com/english/sec-filings",
        "leadership": "https://www.tsmc.com/english/aboutTSMC/executives",
        "governance": "https://investor.tsmc.com/english/board-of-directors",
        "workplace": "https://esg.tsmc.com/en-US/articles/358",
        "sustainability": "https://esg.tsmc.com/en-US/ESG-data-hub/latest-sustainability-information?tab=overview",
    },
    "UBS": {
        "investor": "https://www.ubs.com/global/en/investor-relations.html",
        "annual_reports": "https://www.ubs.com/global/en/investor-relations/financial-information/annual-reporting.html",
        "quarterly_results": "https://www.ubs.com/global/en/investor-relations/financial-information/quarterly-reporting.html",
        "business_structure": "https://www.ubs.com/global/en/our-firm/governance/ubs-group-ag/organization-structure.html",
    },
}

TICKER_ALIASES = {"GOOG": "GOOGL", "2330.TW": "TSM"}


SPECIALIST_MARKET_SOURCES = {
    "foundry": {
        "provider": "TrendForce",
        "url": "https://www.trendforce.com/presscenter/news/20260612-13095.html",
        "purpose": "Comparable global foundry revenue share",
        "source_type": "Specialist industry research",
    },
    "search_engine": {
        "provider": "StatCounter",
        "url": "https://gs.statcounter.com/search-engine-market-share/",
        "purpose": "Search-engine usage share by geography/device",
        "source_type": "Specialist web-usage measurement",
    },
    "cloud_infrastructure": {
        "provider": "Synergy Research Group",
        "url": "https://www.srgresearch.com/articles/cloud-market-annual-revenue-run-rate-topped-half-a-trillion-dollars-in-q1-as-growth-surge-continues",
        "purpose": "Cloud infrastructure services market share",
        "source_type": "Specialist industry research",
    },
    "smartphone_shipments": {
        "provider": "IDC",
        "url": "https://www.idc.com/promo/smartphone-market-share/",
        "purpose": "Worldwide smartphone shipment share",
        "source_type": "Specialist industry research",
    },
}


def canonical_ticker(ticker: str) -> str:
    symbol = str(ticker or "").upper().strip()
    return TICKER_ALIASES.get(symbol, symbol)


def issuer_sources(ticker: str, website: str | None = None) -> dict[str, str]:
    """Return deduplicated known issuer-owned sources, plus the company website fallback."""
    key = canonical_ticker(ticker)
    out = dict(ISSUER_SOURCES.get(key, {}))
    if website:
        base = str(website).strip().rstrip("/")
        if base and base not in out.values():
            out.setdefault("company_website", base)
    return out


def _selected_pages(ticker: str, keys: Iterable[str], website: str | None = None) -> list[str]:
    sources = issuer_sources(ticker, website)
    pages = [sources[k] for k in keys if sources.get(k)]
    return list(dict.fromkeys(pages))


def investor_pages(ticker: str, website: str | None = None) -> list[str]:
    return _selected_pages(
        ticker,
        (
            "investor", "investor_information", "financial_reports", "quarterly_results",
            "earnings", "annual_reports", "filings", "company_website",
        ),
        website,
    )


def segment_pages(ticker: str, website: str | None = None) -> list[str]:
    return _selected_pages(
        ticker,
        (
            "quarterly_results", "financial_reports", "annual_reports", "business_structure",
            "investor", "company_website",
        ),
        website,
    )


def specialist_sources() -> dict[str, dict[str, str]]:
    return {k: dict(v) for k, v in SPECIALIST_MARKET_SOURCES.items()}
