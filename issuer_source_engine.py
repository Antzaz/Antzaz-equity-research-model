from __future__ import annotations

"""Cross-border historical financial source engine.

Source priority:
1) issuer investor-relations downloadable financial tables when discoverable/parsible;
2) SEC annual XBRL facts, including foreign-private-issuer forms (20-F/40-F)
   and both US-GAAP and IFRS namespaces;
3) yfinance annual financial statements as a transparent fallback.

The module is intentionally conservative: it merges only fields it can identify with
reasonable confidence and returns provenance metadata for the workbook.
"""

from io import BytesIO, StringIO
from pathlib import Path
from urllib.parse import urljoin, urlparse
import re

import requests
import pandas as pd
import yfinance as yf

try:
    from bs4 import BeautifulSoup
except Exception:
    BeautifulSoup = None


ANNUAL_FORMS = {"10-K", "10-K/A", "20-F", "20-F/A", "40-F", "40-F/A"}

ISSUER_IR_PAGES = {
    "UBS": [
        "https://www.ubs.com/global/en/investor-relations/financial-information/annual-reporting.html",
        "https://www.ubs.com/global/en/investor-relations/financial-information/quarterly-reporting.html",
    ],
}

FIELD_ALIASES = {
    "revenue": [
        "total revenues", "total revenue", "operating income", "revenue",
    ],
    "op": [
        "operating profit before tax", "operating profit", "profit before tax", "pretax income",
    ],
    "ni": [
        "net profit attributable to shareholders", "net income attributable to shareholders",
        "net profit", "net income", "profit attributable to owners",
    ],
    "eps": ["diluted earnings per share", "diluted eps"],
    "ocf": ["net cash from operating activities", "operating cash flow"],
    "capex": ["capital expenditures", "capital expenditure", "purchase of property plant and equipment"],
    "depr": ["depreciation and amortization", "depreciation amortisation"],
    "sbc": ["share based compensation", "share-based compensation"],
    "cash": ["cash and cash equivalents"],
    "assets": ["total assets"],
    "liabilities": ["total liabilities"],
    "equity": ["total equity", "equity attributable to shareholders", "stockholders equity"],
    "debt": ["long term debt", "long-term debt", "total debt"],
}

YF_INCOME_ROWS = {
    "revenue": ["Total Revenue", "Operating Revenue"],
    "op": ["Pretax Income", "Operating Income"],
    "ni": ["Net Income Common Stockholders", "Net Income", "Net Income Including Noncontrolling Interests"],
    "eps": ["Diluted EPS"],
}
YF_CASH_ROWS = {
    "ocf": ["Operating Cash Flow", "Total Cash From Operating Activities"],
    "capex": ["Capital Expenditure", "Capital Expenditures"],
    "depr": ["Depreciation And Amortization", "Depreciation Amortization Depletion"],
    "sbc": ["Stock Based Compensation"],
}
YF_BALANCE_ROWS = {
    "cash": ["Cash Cash Equivalents And Short Term Investments", "Cash And Cash Equivalents"],
    "assets": ["Total Assets"],
    "liabilities": ["Total Liabilities Net Minority Interest", "Total Liabilities"],
    "equity": ["Stockholders Equity", "Total Equity Gross Minority Interest"],
    "debt": ["Total Debt"],
}


def _norm(text) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(text or "").lower()).strip()


def _year(value):
    try:
        ts = pd.Timestamp(value)
        return int(ts.year)
    except Exception:
        m = re.search(r"(?:19|20)\d{2}", str(value))
        return int(m.group(0)) if m else None


def _number(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace("\u2212", "-")
    neg = text.startswith("(") and text.endswith(")")
    text = text.strip("()").replace(",", "")
    text = re.sub(r"[^0-9.\-]", "", text)
    if text in {"", "-", "."}:
        return None
    try:
        out = float(text)
        return -out if neg else out
    except Exception:
        return None


def _scale_hint(text: str) -> float:
    t = _norm(text)
    if "million" in t or "millions" in t:
        return 1e6
    if "billion" in t or "billions" in t:
        return 1e9
    if "thousand" in t or "thousands" in t:
        return 1e3
    return 1.0


def _row_match(index, aliases):
    normalized = {_norm(x): x for x in index}
    for alias in aliases:
        a = _norm(alias)
        if a in normalized:
            return normalized[a]
    for alias in aliases:
        a = _norm(alias)
        for n, original in normalized.items():
            if a and (a in n or n in a):
                return original
    return None


def _yf_table(ticker: str, attr: str) -> pd.DataFrame:
    try:
        obj = yf.Ticker(ticker)
        df = getattr(obj, attr)
        if callable(df):
            df = df()
        return df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()
    except Exception:
        return pd.DataFrame()


def yfinance_statement_history(ticker: str) -> tuple[dict, dict]:
    income = _yf_table(ticker, "income_stmt")
    cashflow = _yf_table(ticker, "cashflow")
    balance = _yf_table(ticker, "balance_sheet")
    out: dict[int, dict] = {}

    def apply(df, mapping, negate_capex=False):
        if df.empty:
            return
        for field, aliases in mapping.items():
            row = _row_match(df.index, aliases)
            if row is None:
                continue
            series = df.loc[row]
            for col, val in series.items():
                y = _year(col)
                num = _number(val)
                if y is None or num is None:
                    continue
                if negate_capex and field == "capex":
                    num = abs(num)
                out.setdefault(y, {})[field] = num

    apply(income, YF_INCOME_ROWS)
    apply(cashflow, YF_CASH_ROWS, negate_capex=True)
    apply(balance, YF_BALANCE_ROWS)

    for y, d in out.items():
        if d.get("ni") is not None and d.get("eps") not in (None, 0):
            d["shares"] = d["ni"] / d["eps"]
        if d.get("ocf") is not None and d.get("capex") is not None:
            d["fcf"] = d["ocf"] - abs(d["capex"])

    return out, {
        "source": "Yahoo Finance annual statements fallback",
        "years": sorted(out),
        "income_rows": list(income.index.astype(str)) if not income.empty else [],
    }


def _fact_candidates(facts: dict, aliases: list[str]):
    results = []
    for namespace, namespace_facts in (facts.get("facts", {}) or {}).items():
        if namespace not in {"us-gaap", "ifrs-full", "dei"}:
            continue
        for tag, fact in (namespace_facts or {}).items():
            label = fact.get("label") or tag
            text = _norm(f"{tag} {label} {fact.get('description','')}")
            score = 0
            for alias in aliases:
                a = _norm(alias)
                if a == _norm(tag) or a == _norm(label):
                    score = max(score, 100)
                elif a and a in text:
                    score = max(score, 70 + min(20, len(a)//4))
            if score:
                results.append((score, namespace, tag, fact))
    return sorted(results, key=lambda x: x[0], reverse=True)


SEC_FACT_ALIASES = {
    "revenue": [
        "RevenueFromContractWithCustomerExcludingAssessedTax", "SalesRevenueNet", "Revenues",
        "Revenue", "total revenues", "operating income",
    ],
    "op": [
        "OperatingIncomeLoss", "ProfitLossFromOperatingActivities", "operating profit before tax",
        "ProfitLossBeforeTax", "profit before tax",
    ],
    "ni": ["NetIncomeLoss", "ProfitLoss", "profit attributable to owners", "net profit"],
    "eps": ["EarningsPerShareDiluted", "DilutedEarningsLossPerShare", "diluted earnings per share"],
    "ocf": ["NetCashProvidedByUsedInOperatingActivities", "CashFlowsFromUsedInOperatingActivities"],
    "capex": ["PaymentsToAcquirePropertyPlantAndEquipment", "PurchaseOfPropertyPlantAndEquipment"],
    "depr": ["DepreciationDepletionAndAmortization", "DepreciationAmortisationAndImpairmentExpense"],
    "sbc": ["ShareBasedCompensation", "ShareBasedPaymentExpense"],
    "assets": ["Assets", "AssetsTotal"],
    "liabilities": ["Liabilities", "LiabilitiesTotal"],
    "equity": ["StockholdersEquity", "Equity", "EquityAttributableToOwnersOfParent"],
    "cash": ["CashAndCashEquivalentsAtCarryingValue", "CashAndCashEquivalents"],
    "debt": ["LongTermDebt", "Borrowings", "Debt"],
}


def sec_crossborder_history(facts: dict | None) -> tuple[dict, dict]:
    if not facts:
        return {}, {"source": "SEC cross-border XBRL", "years": [], "warnings": ["No Company Facts"]}
    out: dict[int, dict] = {}
    selected = {}
    for field, aliases in SEC_FACT_ALIASES.items():
        candidates = _fact_candidates(facts, aliases)
        if not candidates:
            continue
        score, namespace, tag, fact = candidates[0]
        selected[field] = f"{namespace}:{tag}"
        units = fact.get("units", {}) or {}
        preferred = None
        if field == "eps":
            for u in units:
                if "share" in u.lower():
                    preferred = u; break
        if preferred is None:
            preferred = next(iter(units), None)
        if preferred is None:
            continue
        best = {}
        for x in units.get(preferred, []):
            if x.get("form") not in ANNUAL_FORMS:
                continue
            val = x.get("val"); end = x.get("end")
            if val is None or not end:
                continue
            start = x.get("start")
            if start:
                try:
                    days = (pd.Timestamp(end) - pd.Timestamp(start)).days
                    if days < 250 or days > 450:
                        continue
                except Exception:
                    pass
            try:
                y = int(str(end)[:4]); val = float(val)
            except Exception:
                continue
            stamp = (str(x.get("filed") or ""), str(end))
            if y not in best or stamp > best[y][0]:
                best[y] = (stamp, val)
        for y, (_, val) in best.items():
            if field == "capex":
                val = abs(val)
            out.setdefault(y, {})[field] = val

    for y, d in out.items():
        if d.get("ni") is not None and d.get("eps") not in (None, 0):
            d["shares"] = d["ni"] / d["eps"]
        if d.get("ocf") is not None and d.get("capex") is not None:
            d["fcf"] = d["ocf"] - abs(d["capex"])
    return out, {"source": "SEC annual XBRL (10-K/20-F/40-F; US-GAAP/IFRS)", "years": sorted(out), "selected_tags": selected}


def _discover_ir_pages(ticker: str, info: dict | None) -> list[str]:
    pages = list(ISSUER_IR_PAGES.get(ticker.upper(), []))
    website = (info or {}).get("website")
    if website:
        base = website.rstrip("/")
        for suffix in (
            "/investors", "/investor-relations", "/investors/financials",
            "/investor-relations/financial-information", "/investors/results-and-reports",
        ):
            pages.append(base + suffix)
    return list(dict.fromkeys(pages))


def _download_links(page_url: str, timeout=20):
    if BeautifulSoup is None:
        return []
    r = requests.get(page_url, timeout=timeout, headers={"User-Agent":"Mozilla/5.0 EquityResearch/1.0"})
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")
    links = []
    for a in soup.find_all("a", href=True):
        href = urljoin(r.url, a.get("href"))
        text = _norm(a.get_text(" ", strip=True) + " " + href)
        if any(k in text for k in ("annual report", "financial report", "financial statements", "key figures", "results")):
            links.append(href)
    return list(dict.fromkeys(links))


def _extract_table_history(df: pd.DataFrame) -> dict:
    if df is None or df.empty:
        return {}
    data = df.copy()
    data.columns = [str(c) for c in data.columns]
    year_cols = {c:_year(c) for c in data.columns}
    year_cols = {c:y for c,y in year_cols.items() if y}
    if not year_cols:
        # Sometimes years are in first rows; promote a plausible row to header.
        for r in range(min(8, len(data))):
            vals = data.iloc[r].tolist()
            years = [_year(v) for v in vals]
            if sum(y is not None for y in years) >= 2:
                data.columns = [str(v) for v in vals]
                data = data.iloc[r+1:].copy()
                year_cols = {c:_year(c) for c in data.columns if _year(c)}
                break
    if not year_cols:
        return {}
    label_col = next((c for c in data.columns if c not in year_cols), data.columns[0])
    out = {}
    context = " ".join(str(x) for x in data.head(4).astype(str).values.flatten())
    scale = _scale_hint(context)
    for field, aliases in FIELD_ALIASES.items():
        for _, row in data.iterrows():
            label = _norm(row.get(label_col))
            if not label:
                continue
            if not any(_norm(a) in label or label in _norm(a) for a in aliases):
                continue
            for col, y in year_cols.items():
                val = _number(row.get(col))
                if val is not None:
                    # EPS is per share; do not scale.
                    out.setdefault(y, {})[field] = val if field == "eps" else val * scale
            break
    return out


def issuer_website_history(ticker: str, info: dict | None = None) -> tuple[dict, dict]:
    pages = _discover_ir_pages(ticker, info)
    warnings = []
    discovered = []
    merged = {}
    for page in pages:
        try:
            links = _download_links(page)
        except Exception as exc:
            warnings.append(f"{page}: {exc}")
            continue
        for link in links[:30]:
            if link in discovered:
                continue
            discovered.append(link)
            lower = link.lower().split("?")[0]
            try:
                if lower.endswith(".csv"):
                    r=requests.get(link,timeout=30,headers={"User-Agent":"Mozilla/5.0 EquityResearch/1.0"}); r.raise_for_status()
                    tables=[pd.read_csv(StringIO(r.text))]
                elif lower.endswith((".xlsx",".xls")):
                    r=requests.get(link,timeout=30,headers={"User-Agent":"Mozilla/5.0 EquityResearch/1.0"}); r.raise_for_status()
                    book=pd.ExcelFile(BytesIO(r.content)); tables=[]
                    for sheet in book.sheet_names[:30]:
                        try: tables.append(pd.read_excel(book,sheet_name=sheet,header=None))
                        except Exception: pass
                else:
                    continue
                for table in tables:
                    found=_extract_table_history(table)
                    for y,d in found.items():
                        merged.setdefault(y,{}).update({k:v for k,v in d.items() if v is not None})
            except Exception as exc:
                warnings.append(f"{link}: {exc}")
    return merged, {
        "source": "Issuer investor-relations downloadable financial tables",
        "pages": pages,
        "discovered_links": discovered[:50],
        "years": sorted(merged),
        "warnings": warnings[:10],
    }


def _merge_histories(*histories: tuple[dict, str]) -> tuple[dict, dict]:
    out = {}
    provenance = {}
    # Inputs are ordered highest priority first. Existing fields are not overwritten.
    for history, source in histories:
        for y, d in (history or {}).items():
            for field, value in (d or {}).items():
                if value is None:
                    continue
                if field not in out.setdefault(int(y), {}):
                    out[int(y)][field] = value
                    provenance.setdefault(int(y), {})[field] = source
    for y,d in out.items():
        if d.get("ni") is not None and d.get("eps") not in (None,0) and d.get("shares") is None:
            d["shares"] = d["ni"] / d["eps"]
        if d.get("ocf") is not None and d.get("capex") is not None and d.get("fcf") is None:
            d["fcf"] = d["ocf"] - abs(d["capex"])
    return out, provenance


def build_crossborder_history(ticker: str, info: dict | None, facts: dict | None) -> tuple[dict, dict]:
    website_hist, website_meta = issuer_website_history(ticker, info)
    sec_hist, sec_meta = sec_crossborder_history(facts)
    yf_hist, yf_meta = yfinance_statement_history(ticker)
    merged, provenance = _merge_histories(
        (website_hist, "Issuer IR"),
        (sec_hist, "SEC annual XBRL"),
        (yf_hist, "Yahoo annual statements"),
    )
    years = sorted(y for y,d in merged.items() if d.get("revenue") is not None or d.get("ni") is not None)
    if len(years) > 6:
        keep=set(years[-6:]); merged={y:d for y,d in merged.items() if y in keep}; provenance={y:d for y,d in provenance.items() if y in keep}
    meta = {
        "priority": ["Issuer investor relations", "SEC 10-K/20-F/40-F XBRL", "Yahoo annual statements fallback"],
        "issuer": website_meta,
        "sec": sec_meta,
        "yahoo": yf_meta,
        "provenance": provenance,
        "years": sorted(merged),
    }
    return merged, meta


def patch_financial_statements_from_history(wb, hist: dict, meta: dict | None = None):
    """Populate high-confidence basic rows when the SEC-only Financial Statements sheet is blank."""
    if not hist or "Financial Statements" not in wb.sheetnames:
        return
    ws=wb["Financial Statements"]
    years=sorted(hist)[-6:]
    # Income-statement header row is 6 in the standard sheet.
    for c in range(2,8):
        ws.cell(6,c).value=None
    start=8-len(years)
    for idx,y in enumerate(years,start):
        ws.cell(6,idx).value=y
    labels={str(ws.cell(r,1).value or "").strip():r for r in range(1,ws.max_row+1)}
    mapping={"Revenue":"revenue","Operating Income":"op","Pre-Tax Income":"op","Net Income":"ni","Diluted EPS":"eps"}
    for label,field in mapping.items():
        r=labels.get(label)
        if not r: continue
        for idx,y in enumerate(years,start):
            v=hist[y].get(field)
            if v is not None:
                ws.cell(r,idx).value=v if field=="eps" else v/1e9
    # Cash-flow rows where present.
    cash_map={"Depreciation & Amortization":"depr","Stock-Based Compensation":"sbc","Operating Cash Flow":"ocf","Capital Expenditures":"capex"}
    for label,field in cash_map.items():
        r=labels.get(label)
        if not r: continue
        for idx,y in enumerate(years,start):
            v=hist[y].get(field)
            if v is not None:
                ws.cell(r,idx).value=(-abs(v) if field=="capex" else v)/1e9
    # Basic balance-sheet rows: only years present in history.
    bs_map={"Cash & Cash Equivalents":"cash","Total Assets":"assets","Total Liabilities":"liabilities","Long-Term Debt":"debt","Stockholders' Equity":"equity"}
    for label,field in bs_map.items():
        r=labels.get(label)
        if not r: continue
        # Balance sheet in standard sheet generally has up to four annual columns B:E.
        bs_years=[y for y in years if hist[y].get(field) is not None][-4:]
        for c in range(2,6): ws.cell(23,c).value=None
        for idx,y in enumerate(bs_years,2):
            ws.cell(23,idx).value=y
            ws.cell(r,idx).value=hist[y][field]/1e9
    # Replace SEC-only note with transparent hierarchy.
    try:
        ws["A3"]="USD billions unless per-share data. Source hierarchy: issuer IR reports/tables → SEC annual XBRL (10-K/20-F/40-F; US-GAAP/IFRS) → Yahoo annual statements fallback. Missing fields remain blank rather than estimated."
    except Exception:
        pass


def write_source_hierarchy(wb, ticker: str, meta: dict):
    if "Historical Financials" in wb.sheetnames:
        ws=wb["Historical Financials"]
        ws["A24"]="Sources"
        ws["B24"]="Issuer IR → SEC 10-K/20-F/40-F XBRL → Yahoo annual statements fallback"
        ws["A25"]="Observed years"
        ws["B25"]=", ".join(str(y) for y in meta.get("years",[])) or "None"
        issuer=meta.get("issuer",{}) or {}
        ws["A26"]="Issuer IR page"
        ws["B26"]=(issuer.get("pages") or [None])[0]
    if "Company Data" in wb.sheetnames:
        ws=wb["Company Data"]
        ws["D5"]="Historical data: issuer IR / 20-F IFRS / Yahoo fallback"
        if ticker.upper() in ISSUER_IR_PAGES:
            ws["D6"]="Issuer IR: " + ISSUER_IR_PAGES[ticker.upper()][0]
