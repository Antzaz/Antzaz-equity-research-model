from __future__ import annotations

"""Evidence-gated supplier and supply-chain intelligence for equity research.

The module never invents a ranked supplier list. It scans issuer/regulatory evidence for named
supplier relationships, classifies the evidence, and writes at most 20 verified candidates to the
workbook. Importance and dependency scores are transparent research proxies, not disclosed spend.
"""

from dataclasses import dataclass, asdict
import re
from urllib.parse import urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill

from source_registry import issuer_sources

NAVY="17365D"; BLUE="2F75B5"; WHITE="FFFFFF"; GREY="666666"; LINK_GREEN="008000"
MAX_SUPPLIERS=20
RELATIONSHIP_TERMS=(
    "supplier","suppliers","vendor","vendors","supply agreement","supply chain",
    "manufactured by","manufacturing partner","foundry","contract manufacturer",
    "source from","sourced from","procure from","purchased from","sole source","single source",
)
RISK_TERMS={
    "Single / sole source":("sole source","single source","only supplier","single supplier"),
    "Capacity / availability":("capacity constraint","shortage","allocation","lead time","supply constraint"),
    "Geopolitical / geographic":("geopolitical","taiwan","china","export control","sanction","tariff"),
    "Commodity / input":("commodity","raw material","rare earth","copper","lithium","energy price"),
}
LEGAL_SUFFIX=r"(?:Inc\.?|Corp\.?|Corporation|Ltd\.?|Limited|PLC|plc|AG|SE|S\.A\.?|N\.V\.?|LLC|Co\.?)"
NAME_PATTERN=re.compile(
    rf"\b([A-Z][A-Za-z0-9&.'\-]+(?:\s+[A-Z][A-Za-z0-9&.'\-]+){{0,5}}(?:\s+{LEGAL_SUFFIX})?)\b"
)

STOP_NAMES={
    "United States","Annual Report","Form 10","Form 10 K","Item","Company","Group",
    "Supply Chain","Risk Factors","December","January","February","March","April","May",
    "June","July","August","September","October","November",
}

@dataclass
class SupplierEvidence:
    supplier: str
    relationship: str
    evidence_status: str
    importance_score: float
    dependency_score: float
    risk_flags: str
    evidence: str
    source: str
    source_type: str
    as_of: str | None = None

def _clean(value):
    return re.sub(r"\s+"," ",str(value or "")).strip()

def _clip(value,n=360):
    s=_clean(value)
    return s if len(s)<=n else s[:n-1].rsplit(" ",1)[0]+"…"

def _fetch_text(url,timeout=12):
    if not url: return ""
    try:
        r=requests.get(url,timeout=timeout,headers={"User-Agent":"Antzaz Equity Research educational research"})
        r.raise_for_status()
        soup=BeautifulSoup(r.content,"lxml")
        for tag in soup(["script","style","noscript"]): tag.decompose()
        return _clean(soup.get_text(" ",strip=True))
    except Exception:
        return ""

def _filing_sources(wb):
    out=[]
    if "Filings" not in wb.sheetnames: return out
    ws=wb["Filings"]
    for row in ws.iter_rows():
        vals=[c.value for c in row]
        urls=[str(v) for v in vals if isinstance(v,str) and v.startswith("http")]
        if not urls: continue
        label=" ".join(_clean(v) for v in vals[:4] if v)
        if any(x in label.upper() for x in ("10-K","20-F","40-F","ANNUAL")):
            out.append((urls[0],"Regulatory / annual filing"))
    return list(dict.fromkeys(out))[:3]

def _latest_regulatory_filing(ticker):
    """Resolve the latest annual filing directly from SEC for any US-listed issuer."""
    try:
        headers={"User-Agent":"Antzaz Equity Research educational research contact research@example.com"}
        tickers=requests.get("https://www.sec.gov/files/company_tickers.json",headers=headers,timeout=20).json()
        cik=None
        for item in tickers.values():
            if str(item.get("ticker","")).upper()==str(ticker or "").upper():
                cik=str(item["cik_str"]).zfill(10); break
        if not cik: return None
        subs=requests.get(f"https://data.sec.gov/submissions/CIK{cik}.json",headers=headers,timeout=20).json()
        recent=subs.get("filings",{}).get("recent",{})
        for form,acc,doc in zip(recent.get("form",[]),recent.get("accessionNumber",[]),recent.get("primaryDocument",[])):
            if form in {"10-K","20-F","40-F"}:
                return f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{acc.replace('-','')}/{doc}"
    except Exception:
        return None
    return None

def _candidate_sources(wb,ticker,info):
    sources=[]
    direct=_latest_regulatory_filing(ticker)
    if direct: sources.append((direct,"Regulatory / annual filing"))
    for url,kind in _filing_sources(wb): sources.append((url,kind))
    issuer=issuer_sources(ticker,(info or {}).get("website"))
    for key in ("annual_reports","filings","investor","financial_reports","company_website"):
        url=issuer.get(key)
        if url: sources.append((url,"Issuer-owned"))
    seen=set(); out=[]
    for row in sources:
        if row[0] in seen: continue
        seen.add(row[0]); out.append(row)
    return out[:6]

def _windows(text):
    low=text.lower(); spans=[]
    for term in RELATIONSHIP_TERMS:
        start=0
        while True:
            i=low.find(term,start)
            if i<0: break
            spans.append(_clean(text[max(0,i-220):min(len(text),i+520)]))
            start=i+len(term)
            if len(spans)>=120: return spans
    return spans

def _relationship(window):
    low=window.lower()
    if "sole source" in low or "single source" in low: return "Single / sole-source supplier"
    if "foundry" in low or "manufactured by" in low: return "Manufacturing / foundry"
    if "contract manufacturer" in low: return "Contract manufacturing"
    if "supply agreement" in low: return "Supply agreement"
    if "vendor" in low: return "Vendor"
    return "Supplier / input provider"

def _risk_flags(window):
    low=window.lower(); flags=[]
    for label,terms in RISK_TERMS.items():
        if any(t in low for t in terms): flags.append(label)
    return "; ".join(flags) or "No specific risk flag in retrieved evidence"

def _scores(window,source_type):
    low=window.lower()
    importance=35.0
    dependency=25.0
    if "material" in low or "significant" in low or "major" in low: importance+=20
    if "primary" in low or "principal" in low: importance+=15
    if "sole source" in low or "single source" in low:
        importance+=25; dependency+=55
    if "limited number" in low or "few suppliers" in low: dependency+=25
    if "long-term" in low or "multi-year" in low: importance+=10; dependency+=10
    if source_type.startswith("Regulatory"): importance+=5
    return min(100,importance),min(100,dependency)

def _extract_names(window):
    names=[]
    for m in NAME_PATTERN.finditer(window):
        name=_clean(m.group(1)).strip(" ,.;:")
        if len(name)<3 or name in STOP_NAMES: continue
        if name.split()[0] in {"Our","The","These","This","Such","We","Risk","Supply"}: continue
        names.append(name)
    return list(dict.fromkeys(names))

def collect_supplier_evidence(wb,ticker,info=None,max_suppliers=MAX_SUPPLIERS):
    """Return evidence-ranked supplier candidates. Empty is valid when evidence is insufficient."""
    company_words=set(re.findall(r"[a-z0-9]+",_clean((info or {}).get("longName") or ticker).lower()))
    found={}
    for url,source_type in _candidate_sources(wb,ticker,info or {}):
        text=_fetch_text(url)
        if not text: continue
        for window in _windows(text):
            rel=_relationship(window)
            importance,dependency=_scores(window,source_type)
            for name in _extract_names(window):
                norm=re.sub(r"[^a-z0-9]+"," ",name.lower()).strip()
                if not norm or set(norm.split()) & company_words: continue
                # Require the candidate to be reasonably close to an explicit supplier term.
                pos=window.lower().find(norm.split()[0])
                supplier_positions=[window.lower().find(t) for t in RELATIONSHIP_TERMS if window.lower().find(t)>=0]
                if pos<0 or not supplier_positions or min(abs(pos-p) for p in supplier_positions)>230: continue
                rec=SupplierEvidence(
                    supplier=name,
                    relationship=rel,
                    evidence_status="Disclosed" if source_type.startswith("Regulatory") else "Confirmed public evidence",
                    importance_score=importance,
                    dependency_score=dependency,
                    risk_flags=_risk_flags(window),
                    evidence=_clip(window),
                    source=url,
                    source_type=source_type,
                )
                prior=found.get(norm)
                if prior is None or (rec.importance_score+rec.dependency_score)>(prior.importance_score+prior.dependency_score):
                    found[norm]=rec
    rows=sorted(found.values(),key=lambda x:(x.importance_score+x.dependency_score,x.dependency_score),reverse=True)
    return rows[:max_suppliers]

def _fill(color): return PatternFill("solid",fgColor=color)

def ensure_supply_chain_intelligence(wb,ticker,info=None,records=None,max_suppliers=MAX_SUPPLIERS):
    """Create the workbook Supply Chain sheet and return a serializable result summary."""
    records=list(records) if records is not None else collect_supplier_evidence(wb,ticker,info or {},max_suppliers)
    if "Supply Chain" in wb.sheetnames: wb.remove(wb["Supply Chain"])
    ws=wb.create_sheet("Supply Chain"); ws.sheet_view.showGridLines=False
    for c in range(1,11): ws.cell(1,c).fill=_fill(NAVY)
    ws["A1"]=f"{ticker} — Supplier & Supply-Chain Intelligence"; ws["A1"].font=Font(bold=True,color=WHITE,size=18)
    ws["A3"]="Evidence-gated research only. The table shows up to 20 suppliers identifiable from public evidence; it is not a complete procurement ledger. Importance/dependency scores are transparent research proxies, not disclosed spend shares."
    ws.merge_cells("A3:J3"); ws["A3"].font=Font(italic=True,color=GREY); ws["A3"].alignment=Alignment(wrap_text=True)

    headers=["Rank","Supplier","Relationship","Evidence status","Importance / 100","Dependency / 100","Risk flags","Evidence","Source type","Source"]
    for c,h in enumerate(headers,1):
        cell=ws.cell(5,c,h); cell.fill=_fill(BLUE); cell.font=Font(bold=True,color=WHITE); cell.alignment=Alignment(wrap_text=True)
    for r,rec in enumerate(records,6):
        vals=[r-5,rec.supplier,rec.relationship,rec.evidence_status,rec.importance_score,rec.dependency_score,rec.risk_flags,rec.evidence,rec.source_type,rec.source]
        for c,v in enumerate(vals,1): ws.cell(r,c,v)
        ws.cell(r,8).alignment=Alignment(wrap_text=True,vertical="top")
        ws.cell(r,10).hyperlink=rec.source; ws.cell(r,10).font=Font(color=LINK_GREEN,underline="single")
    if not records:
        ws["A6"]="REVIEW — No named supplier relationships were verified automatically from the retrieved public evidence."
        ws.merge_cells("A6:J6"); ws["A6"].alignment=Alignment(wrap_text=True)

    risk_row=max(9,6+len(records)+2)
    ws.cell(risk_row,1,"Supply-chain risk summary")
    for c in range(1,11): ws.cell(risk_row,c).fill=_fill(NAVY); ws.cell(risk_row,c).font=Font(bold=True,color=WHITE)
    singles=sum("Single / sole" in r.relationship for r in records)
    geo=sum("Geopolitical" in r.risk_flags for r in records)
    highdep=sum(r.dependency_score>=70 for r in records)
    summary=[
        ("Verified supplier candidates",len(records),"Count of public-evidence relationships; not total supplier count."),
        ("High-dependency relationships",highdep,"Proxy dependency score >=70/100."),
        ("Single / sole-source evidence",singles,"Explicit wording found in retrieved evidence."),
        ("Geopolitical / geographic flags",geo,"Evidence text contained mapped geographic/geopolitical risk terms."),
    ]
    for rr,row in enumerate(summary,risk_row+1):
        for c,v in enumerate(row,1): ws.cell(rr,c,v)

    if records:
        chart=BarChart(); chart.type="bar"; chart.style=10; chart.title="Top supplier dependency signals"; chart.y_axis.title="Supplier"; chart.x_axis.title="Dependency proxy / 100"
        data=Reference(ws,min_col=6,min_row=5,max_row=5+len(records))
        cats=Reference(ws,min_col=2,min_row=6,max_row=5+len(records))
        chart.add_data(data,titles_from_data=True); chart.set_categories(cats); chart.height=8; chart.width=15
        ws.add_chart(chart,f"E{risk_row+1}")

    widths={"A":10,"B":30,"C":28,"D":24,"E":16,"F":16,"G":32,"H":65,"I":25,"J":55}
    for col,width in widths.items(): ws.column_dimensions[col].width=width
    ws.freeze_panes="A6"
    return {
        "ticker":ticker,
        "supplier_count":len(records),
        "status":"PUBLIC-EVIDENCE" if records else "REVIEW",
        "high_dependency_count":highdep,
        "single_source_count":singles,
        "geopolitical_flag_count":geo,
        "suppliers":[asdict(r) for r in records],
    }
