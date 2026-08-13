from __future__ import annotations

"""Issuer-first source discovery for segment analysis.

Policy:
- issuer investor-relations / annual-report / earnings materials are always attempted first;
- regulatory filings are a fallback and corroboration layer, never the first choice when an
  issuer-owned source is available;
- missing segment economics stay blank rather than being inferred from third-party data.

The engine deliberately separates two jobs:
- discover the issuer's current business/reportable-segment names from official narrative sources;
- expose filing/report documents to segment_analysis_v2 / segment_source_enrichment for conservative numeric extraction.

Network work is intentionally bounded so one slow issuer page cannot stall the whole build.
"""

from io import BytesIO
from urllib.parse import urlparse
import re

import requests
import yfinance as yf

try:
    from lxml import html as lxml_html
except Exception:
    lxml_html = None
try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None

try:
    from issuer_source_engine import _discover_ir_pages, _download_links
except Exception:
    _discover_ir_pages = None
    _download_links = None


ANNUAL_FORMS = {"10-K", "10-K/A", "20-F", "20-F/A", "40-F", "40-F/A"}
CURRENT_FORMS = {"6-K"}

# Explicit issuer-owned roots are a high-confidence accelerator for companies where generic
# /investors discovery is known to be brittle. Generic discovery still runs for every ticker.
OFFICIAL_SEGMENT_PAGES = {
    "CVX": [
        "https://www.chevron.com/investors/reports-and-filings",
        "https://www.chevron.com/annual-report",
        "https://www.chevron.com/investors",
        "https://www.chevron.com/newsroom/media/publications",
        "https://www.chevron.com/what-we-do/energy/oil-and-natural-gas",
        "https://www.chevron.com/what-we-do/energy/refining",
    ],
    "UBS": [
        "https://www.ubs.com/global/en/our-firm/governance/ubs-group-ag/organization-structure.html",
        "https://www.ubs.com/global/en/our-firm/what-we-do.html",
        "https://www.ubs.com/global/en/investor-relations/financial-information.html",
    ],
    "TSM": [
        "https://investor.tsmc.com/static/annualReports/2025/english/index.html",
        "https://investor.tsmc.com/english/quarterly-results/2026/q2",
        "https://investor.tsmc.com/english/monthly-revenue/2026",
        "https://investor.tsmc.com/english/sec-filings",
    ],
    "SIE.DE": [
        "https://www.siemens.com/global/en/company/investor-relations/financial-results.html",
        "https://www.siemens.com/global/en/company/investor-relations/reports-publications-ad-hoc.html",
    ],
}

# Verified structure fallbacks are allowed only when the source is issuer-owned and the fallback
# contains names/structure, never fabricated financial values. Numeric rows still require a table.
VERIFIED_SEGMENT_FALLBACKS = {
    "CVX": {
        "segments": ["Upstream", "Downstream", "All Other"],
        "source": "https://www.chevron.com/investors/reports-and-filings",
        "note": (
            "Chevron's issuer reporting center publishes the 2025 Annual Report, Supplement to the "
            "Annual Report, supplement Excel template and quarterly data supplements. The model uses "
            "Upstream / Downstream / All Other as the disclosed accounting structure fallback only when "
            "live issuer parsing cannot recover the labels; exact financial values still require an issuer "
            "annual/results table and are never estimated."
        ),
    },
    "UBS": {
        "segments": [
            "Global Wealth Management",
            "Personal & Corporate Banking",
            "Asset Management",
            "Investment Bank",
            "Non-core and Legacy",
            "Group Functions / Group Items",
        ],
        "source": "https://www.ubs.com/global/en/our-firm/governance/ubs-group-ag/organization-structure.html",
        "note": "UBS states that its operational structure comprises five business divisions plus Group Functions; financial statements provide segment reporting by the business divisions and Group Functions / Group Items.",
    }
}


def _clean(value) -> str:
    return re.sub(r"\s+", " ", str(value or "").replace("\xa0", " ")).strip()


def _html_text(raw_html: str | None) -> str:
    if not raw_html:
        return ""
    if lxml_html is not None:
        try:
            root = lxml_html.fromstring(raw_html)
            for node in root.xpath("//script|//style|//noscript"):
                try:
                    node.drop_tree()
                except Exception:
                    pass
            return _clean(" ".join(root.itertext()))
        except Exception:
            pass
    return _clean(re.sub(r"<[^>]+>", " ", raw_html))


def _fetch_html(url: str, headers: dict | None = None, timeout: int = 10):
    try:
        r = requests.get(url, headers=headers or {"User-Agent": "Mozilla/5.0 EquityResearch/1.0"}, timeout=timeout)
        r.raise_for_status()
        ctype = str(r.headers.get("content-type", "")).lower()
        if "pdf" in ctype or url.lower().split("?")[0].endswith(".pdf"):
            return None, _pdf_text(r.content), r.url
        # Excel is intentionally not interpreted as HTML. The issuer landing page / PDF stays the
        # primary citation, while an XLS/XLSX supplement can be added later by a verified adapter.
        if any(x in ctype for x in ("spreadsheet", "excel")) or url.lower().split("?")[0].endswith((".xls", ".xlsx")):
            return None, "", r.url
        return r.text, _html_text(r.text), r.url
    except Exception:
        return None, "", url


def _pdf_text(content: bytes) -> str:
    if not content or PdfReader is None:
        return ""
    try:
        reader = PdfReader(BytesIO(content))
        texts = []
        for page in reader.pages[:90]:
            try:
                texts.append(page.extract_text() or "")
            except Exception:
                continue
        return _clean(" ".join(texts))
    except Exception:
        return ""


def _cik_for(ticker: str, headers: dict):
    try:
        r=requests.get("https://www.sec.gov/files/company_tickers.json", headers=headers, timeout=10)
        r.raise_for_status()
        data=r.json()
        # SEC uses hyphenated Berkshire symbols while the project accepts BRK.B.
        aliases={ticker.upper(), ticker.upper().replace(".", "-")}
        for item in data.values():
            if str(item.get("ticker", "")).upper() in aliases:
                return str(item["cik_str"]).zfill(10)
    except Exception:
        return None
    return None


def sec_segment_documents(ticker: str, headers: dict) -> list[dict]:
    """Regulatory fallback/corroboration documents.

    Priorities deliberately come after issuer-owned material.
    """
    cik = _cik_for(ticker, headers)
    if not cik:
        return []
    try:
        r=requests.get(f"https://data.sec.gov/submissions/CIK{cik}.json", headers=headers, timeout=10)
        r.raise_for_status()
        recent=r.json().get("filings", {}).get("recent", {})
    except Exception:
        return []
    docs = []
    annual_added = False
    current_added = 0
    for form, acc, doc, filed in zip(
        recent.get("form", []), recent.get("accessionNumber", []),
        recent.get("primaryDocument", []), recent.get("filingDate", []),
    ):
        if form in ANNUAL_FORMS and not annual_added:
            url = f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{acc.replace('-', '')}/{doc}"
            raw, text, resolved = _fetch_html(url, headers, 12)
            if raw or text:
                docs.append({"kind": f"Regulatory fallback — SEC {form}", "url": resolved, "html": raw, "text": text, "filed": filed, "priority": 50, "issuer_owned": False})
                annual_added = True
        elif form in CURRENT_FORMS and current_added < 2:
            url = f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{acc.replace('-', '')}/{doc}"
            raw, text, resolved = _fetch_html(url, headers, 10)
            low = text.lower()
            if text and any(k in low for k in ("segment", "business division", "business divisions", "group items", "wafer revenue", "platform")):
                docs.append({"kind": f"Regulatory corroboration — SEC {form}", "url": resolved, "html": raw, "text": text, "filed": filed, "priority": 60, "issuer_owned": False})
                current_added += 1
        if annual_added and current_added >= 2:
            break
    return docs


def _website_candidates(ticker: str) -> list[str]:
    pages = list(OFFICIAL_SEGMENT_PAGES.get(ticker.upper(), []))
    try:
        info = yf.Ticker(ticker.replace(".", "-") if ticker.upper().startswith("BRK.") else ticker).info or {}
    except Exception:
        info = {}
    website = info.get("website")
    if website:
        base = str(website).rstrip("/")
        # Investor/report pages first. Narrative corporate pages are useful for names but should
        # never outrank current financial reporting pages for economics.
        pages.extend([
            base + "/investors/reports-and-filings",
            base + "/investors/financial-information",
            base + "/investors/events-presentations",
            base + "/investors",
            base + "/investor-relations/financial-results",
            base + "/investor-relations/reports",
            base + "/investor-relations",
            base + "/annual-report",
            base + "/reports",
            base + "/our-businesses",
            base + "/what-we-do",
            base + "/about",
        ])
    if _discover_ir_pages is not None:
        try:
            pages.extend(_discover_ir_pages(ticker, info))
        except Exception:
            pass
    return list(dict.fromkeys(x for x in pages if x))


def issuer_segment_documents(ticker: str) -> list[dict]:
    """Collect company-owned segment/reporting material before any regulator fallback."""
    docs = []
    seen = set()
    pages = _website_candidates(ticker)
    report_count=0
    for page in pages[:12]:
        raw, text, resolved = _fetch_html(page, None, 9)
        if not text:
            continue
        key = resolved.split("#")[0]
        if key not in seen:
            low=(text+" "+resolved).lower()
            kind="Issuer investor / reporting page" if any(k in low for k in ("investor", "annual report", "financial", "earnings", "results")) else "Issuer website"
            priority=8 if kind.startswith("Issuer investor") else 12
            docs.append({"kind": kind, "url": resolved, "html": raw, "text": text, "priority": priority, "issuer_owned": True})
            seen.add(key)

        if report_count>=5 or _download_links is None or not any(k in page.lower() for k in ("investor", "financial", "report", "annual", "result", "presentation")):
            continue
        try:
            links = _download_links(page, timeout=8)
        except TypeError:
            try: links = _download_links(page)
            except Exception: links=[]
        except Exception:
            links = []
        for link in links[:12]:
            path = urlparse(link).path.lower()
            if not any(k in (link + " " + path).lower() for k in ("annual", "supplement", "financial", "result", "earnings", "report", "presentation", "data")):
                continue
            raw2, text2, resolved2 = _fetch_html(link, None, 12)
            if not text2:
                continue
            key2 = resolved2.split("#")[0]
            if key2 in seen:
                continue
            docs.append({"kind": "Issuer annual/results report", "url": resolved2, "html": raw2, "text": text2, "priority": 5, "issuer_owned": True})
            seen.add(key2)
            report_count+=1
            if report_count>=5:
                break
    return docs


def collect_segment_documents(ticker: str, headers: dict) -> list[dict]:
    """Return official documents with issuer-owned sources strictly ahead of regulators."""
    issuer_docs = issuer_segment_documents(ticker)
    sec_docs = sec_segment_documents(ticker, headers)
    docs = issuer_docs + sec_docs
    docs.sort(key=lambda d: (0 if d.get("issuer_owned") else 1, int(d.get("priority", 99)), str(d.get("filed", ""))))
    return docs


def verified_fallback(ticker: str) -> dict:
    return dict(VERIFIED_SEGMENT_FALLBACKS.get(ticker.upper(), {}))
