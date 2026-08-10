from __future__ import annotations

"""Official/public source discovery for segment analysis.

The engine deliberately separates two jobs:
- discover the issuer's current business/reportable-segment names from official narrative sources;
- expose filing/report documents to segment_analysis_v2 for conservative numeric extraction.

Source priority is issuer/regulator data only. Network work is intentionally bounded so
one slow issuer page cannot stall the whole equity-research build.
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

OFFICIAL_SEGMENT_PAGES = {
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
}

VERIFIED_SEGMENT_FALLBACKS = {
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
        return r.text, _html_text(r.text), r.url
    except Exception:
        return None, "", url


def _pdf_text(content: bytes) -> str:
    if not content or PdfReader is None:
        return ""
    try:
        reader = PdfReader(BytesIO(content))
        texts = []
        for page in reader.pages[:60]:
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
        for item in data.values():
            if str(item.get("ticker", "")).upper() == ticker.upper():
                return str(item["cik_str"]).zfill(10)
    except Exception:
        return None
    return None


def sec_segment_documents(ticker: str, headers: dict) -> list[dict]:
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
                docs.append({"kind": f"SEC {form}", "url": resolved, "html": raw, "text": text, "filed": filed, "priority": 10})
                annual_added = True
        elif form in CURRENT_FORMS and current_added < 2:
            url = f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{acc.replace('-', '')}/{doc}"
            raw, text, resolved = _fetch_html(url, headers, 10)
            low = text.lower()
            if text and any(k in low for k in ("segment", "business division", "business divisions", "group items", "wafer revenue", "platform")):
                docs.append({"kind": f"SEC {form}", "url": resolved, "html": raw, "text": text, "filed": filed, "priority": 20})
                current_added += 1
        if annual_added and current_added >= 2:
            break
    return docs


def _website_candidates(ticker: str) -> list[str]:
    pages = list(OFFICIAL_SEGMENT_PAGES.get(ticker.upper(), []))
    try:
        info = yf.Ticker(ticker).info or {}
    except Exception:
        info = {}
    website = info.get("website")
    if website:
        base = str(website).rstrip("/")
        pages.extend([
            base + "/about", base + "/our-businesses", base + "/what-we-do",
            base + "/investors", base + "/investor-relations",
        ])
    if _discover_ir_pages is not None:
        try:
            pages.extend(_discover_ir_pages(ticker, info))
        except Exception:
            pass
    return list(dict.fromkeys(pages))


def issuer_segment_documents(ticker: str) -> list[dict]:
    docs = []
    seen = set()
    pages = _website_candidates(ticker)
    report_count=0
    for page in pages[:6]:
        raw, text, resolved = _fetch_html(page, None, 8)
        if not text:
            continue
        key = resolved.split("#")[0]
        if key not in seen:
            docs.append({"kind": "Issuer website", "url": resolved, "html": raw, "text": text, "priority": 30})
            seen.add(key)

        if report_count>=2 or _download_links is None or not any(k in page.lower() for k in ("investor", "financial", "report")):
            continue
        try:
            links = _download_links(page, timeout=8)
        except TypeError:
            try: links = _download_links(page)
            except Exception: links=[]
        except Exception:
            links = []
        for link in links[:5]:
            path = urlparse(link).path.lower()
            if not any(k in (link + " " + path).lower() for k in ("annual", "financial", "result", "report", "presentation", "management")):
                continue
            raw2, text2, resolved2 = _fetch_html(link, None, 10)
            if not text2:
                continue
            key2 = resolved2.split("#")[0]
            if key2 in seen:
                continue
            docs.append({"kind": "Issuer annual/results report", "url": resolved2, "html": raw2, "text": text2, "priority": 15})
            seen.add(key2)
            report_count+=1
            if report_count>=2:
                break
    return docs


def collect_segment_documents(ticker: str, headers: dict) -> list[dict]:
    docs = sec_segment_documents(ticker, headers) + issuer_segment_documents(ticker)
    docs.sort(key=lambda d: (int(d.get("priority", 99)), str(d.get("filed", ""))))
    return docs


def verified_fallback(ticker: str) -> dict:
    return dict(VERIFIED_SEGMENT_FALLBACKS.get(ticker.upper(), {}))
