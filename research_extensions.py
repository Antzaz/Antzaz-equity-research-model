from __future__ import annotations

"""Research extensions and workbook de-duplication.

Adds a source-auditable people/leadership layer, employee-sentiment context, market-share
context, and a same-sector alternative screen. It also removes presentation sheets that
repeat information already available in the authoritative dashboard/analytics tabs.

The employee signal is deliberately scoped: a program satisfaction score is not presented
as a company-wide engagement score. External crowd-review websites are listed as optional
manual corroboration sources rather than scraped into the model.
"""

import math
from urllib.parse import urljoin

import requests
import yfinance as yf
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from market_context import market_share_record

NAVY = "17365D"
BLUE = "2F75B5"
WHITE = "FFFFFF"
LIGHT = "F5F9FC"
GOLD = "FFF2CC"
PALE_GREEN = "E2F0D9"
PALE_RED = "FCE4D6"
GREY = "666666"
INPUT_BLUE = "0000FF"
LINK_GREEN = "008000"
THIN = Side(style="thin", color="D9E1F2")
FMT_PCT = "0.0%;[Red](0.0%);-"
FMT_SCORE = "0.0"
FMT_MULT = "0.0x;[Red](0.0x);-"

TSM_SOURCES = {
    "annual": "https://investor.tsmc.com/static/annualReports/2025/english/index.html",
    "leadership": "https://www.tsmc.com/english/aboutTSMC/executives",
    "board": "https://investor.tsmc.com/english/board-of-directors",
    "culture": "https://esg.tsmc.com/en-US/articles/358",
    "sustainability": "https://esg.tsmc.com/en-US/ESG-data-hub/latest-sustainability-information?tab=overview",
}


def _fill(color):
    return PatternFill("solid", fgColor=color)


def _num(v, default=None):
    try:
        if isinstance(v, bool) or v in (None, ""):
            return default
        x = float(v)
        return x if math.isfinite(x) else default
    except Exception:
        return default


def _section(ws, row, title, end=10):
    for c in range(1, end + 1):
        ws.cell(row, c).fill = _fill(NAVY)
        ws.cell(row, c).font = Font(bold=True, color=WHITE, size=11)
    ws.cell(row, 1, title)


def _header(ws, row, start, end):
    for c in range(start, end + 1):
        cell = ws.cell(row, c)
        cell.fill = _fill(BLUE)
        cell.font = Font(bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=THIN)


def _safe_info(ticker: str, info: dict | None = None) -> dict:
    if info:
        return info
    try:
        return yf.Ticker(ticker).info or {}
    except Exception:
        return {}


def _official_source_urls(ticker: str, info: dict) -> dict:
    if ticker.upper() == "TSM":
        return dict(TSM_SOURCES)
    website = str(info.get("website") or "").rstrip("/")
    if not website:
        return {}
    return {
        "leadership": website + "/about/leadership",
        "governance": website + "/investors/corporate-governance",
        "sustainability": website + "/sustainability",
        "careers": website + "/careers",
    }


def _employee_signal(ticker: str, info: dict) -> dict:
    """Return a transparent worker-sentiment signal when a reliable official score exists."""
    # TSMC's latest easily verifiable numeric employee-related satisfaction figure is a
    # program-specific score, not the global employee engagement survey. Keep the scope.
    if ticker.upper() == "TSM":
        return {
            "score": 96.0,
            "scope": "Global Inclusive Workplace learning-program satisfaction; not a company-wide engagement score",
            "period": "2024-2025 program",
            "source": TSM_SOURCES["culture"],
            "status": "SCOPE-LIMITED",
            "evidence": "TSMC reports >100,500 participants and an average satisfaction score of 96 points for the inclusion-learning program.",
        }
    return {
        "score": None,
        "scope": "No comparable company-wide employee happiness/engagement score was automatically verified",
        "period": None,
        "source": _official_source_urls(ticker, info).get("sustainability"),
        "status": "REVIEW",
        "evidence": "Use the issuer sustainability/annual report first; corroborate manually with employee-review platforms if desired.",
    }


def _history_metrics(wb):
    if "Historical Financials" not in wb.sheetnames:
        return {}
    h = wb["Historical Financials"]
    pts = []
    for c in range(2, 8):
        y = h.cell(3, c).value
        rev = _num(h.cell(4, c).value)
        if isinstance(y, (int, float)) and rev and rev > 0:
            pts.append((int(y), rev))
    cagr = None
    if len(pts) >= 2:
        y0, v0 = pts[0]
        y1, v1 = pts[-1]
        n = max(1, y1 - y0)
        cagr = (v1 / v0) ** (1 / n) - 1 if v0 > 0 and v1 > 0 else None
    rev = _num(h["G4"].value)
    op = _num(h["G9"].value)
    ocf = _num(h["G14"].value)
    cap = _num(h["G15"].value)
    return {
        "revenue_cagr": cagr,
        "operating_margin": (op / rev) if rev and op is not None else None,
        "fcf_margin": ((ocf - cap) / rev) if rev and ocf is not None and cap is not None else None,
    }


def _leadership_proxy(wb, ticker: str, info: dict, employee: dict) -> tuple[float, list[tuple]]:
    """Transparent proxy score; intentionally not presented as a factual management rating."""
    hist = _history_metrics(wb)
    cagr = hist.get("revenue_cagr")
    opm = hist.get("operating_margin")
    fcfm = hist.get("fcf_margin")
    d = wb["Company Data"] if "Company Data" in wb.sheetnames else None
    net_debt = _num(d["B14"].value) if d else None

    execution = 50.0
    if cagr is not None and opm is not None:
        execution = max(0.0, min(100.0, 50 * min(1.0, max(0.0, cagr) / 0.20) + 50 * min(1.0, max(0.0, opm) / 0.40)))
    capital = 50.0
    if fcfm is not None:
        capital = 35 + 45 * min(1.0, max(0.0, fcfm) / 0.25)
        if net_debt is not None and net_debt < 0:
            capital += 20
        capital = min(100.0, capital)
    officers = info.get("companyOfficers") or []
    depth = min(100.0, 35 + 8 * len([x for x in officers if isinstance(x, dict) and x.get("name")]))
    if ticker.upper() == "TSM":
        depth = max(depth, 90.0)
    culture = _num(employee.get("score"), 50.0)
    governance = 70.0 if _official_source_urls(ticker, info).get("governance") or ticker.upper() == "TSM" else 50.0
    score = 0.30 * execution + 0.25 * capital + 0.15 * depth + 0.15 * culture + 0.15 * governance
    rows = [
        ("Execution track record", execution, "Revenue growth and operating-margin history; performance proxy, not direct causality"),
        ("Capital allocation / cash generation", capital, "FCF margin and net cash/debt profile"),
        ("Leadership depth / continuity", depth, "Public officer depth; TSM continuity supported by issuer leadership history"),
        ("Employee / culture signal", culture, employee.get("scope")),
        ("Governance disclosure", governance, "Availability of official governance/board disclosure"),
    ]
    return score, rows


def _rank_percentile(values: list[float], value: float, higher_better=True):
    clean = sorted(v for v in values if v is not None and math.isfinite(v))
    if not clean or value is None:
        return None
    if len(clean) == 1:
        return 50.0
    below = sum(v < value for v in clean)
    equal = sum(v == value for v in clean)
    pct = (below + 0.5 * max(1, equal - 1)) / (len(clean) - 1) * 100
    pct = max(0.0, min(100.0, pct))
    return pct if higher_better else 100.0 - pct


def _peer_rows(wb):
    if "Peer Comps" not in wb.sheetnames:
        return []
    ws = wb["Peer Comps"]
    rows = []
    for r in range(4, min(ws.max_row, 20) + 1):
        ticker = str(ws.cell(r, 2).value or "").strip().upper()
        if not ticker or ticker.startswith("REVIEW"):
            continue
        rows.append({
            "row": r,
            "company": ws.cell(r, 1).value,
            "ticker": ticker,
            "pe": _num(ws.cell(r, 3).value),
            "ev_ebitda": _num(ws.cell(r, 5).value),
            "growth": _num(ws.cell(r, 6).value),
            "margin": _num(ws.cell(r, 7).value),
            "roe": _num(ws.cell(r, 8).value),
            "market_share": _num(ws.cell(r, 13).value) if ws.max_column >= 13 else None,
        })
    return rows


def _alternative_screen(wb, target_ticker: str) -> dict:
    rows = _peer_rows(wb)
    if len(rows) < 2:
        return {"result": "No validated peer set available", "candidate": None}
    specs = [
        ("growth", True, 0.25),
        ("margin", True, 0.25),
        ("roe", True, 0.15),
        ("pe", False, 0.20),
        ("ev_ebitda", False, 0.15),
    ]
    for row in rows:
        weighted = 0.0
        used = 0.0
        for key, higher, weight in specs:
            value = row.get(key)
            vals = [x.get(key) for x in rows if x.get(key) is not None]
            if value is None or len(vals) < 2:
                continue
            pct = _rank_percentile(vals, value, higher)
            if pct is None:
                continue
            weighted += pct * weight
            used += weight
        row["screen_score"] = weighted / used if used else None
    target = next((x for x in rows if x["ticker"] == target_ticker.upper()), None)
    candidates = [x for x in rows if x["ticker"] != target_ticker.upper() and x.get("screen_score") is not None]
    if not target or target.get("screen_score") is None or not candidates:
        return {"result": "Peer metrics are incomplete; no robust alternative recommendation", "candidate": None, "target": target}
    best = max(candidates, key=lambda x: x["screen_score"])
    gap = best["screen_score"] - target["screen_score"]
    reasons = []
    for key, label, higher in [
        ("growth", "revenue growth", True), ("margin", "operating margin", True),
        ("roe", "ROE", True), ("pe", "forward P/E", False), ("ev_ebitda", "EV/EBITDA", False),
    ]:
        a, b = best.get(key), target.get(key)
        if a is None or b is None:
            continue
        better = a > b if higher else a < b
        if better:
            reasons.append(label)
    clearly_better = gap >= 8.0 and len(reasons) >= 2
    return {
        "result": (f"{best['ticker']} screens better on current public peer metrics" if clearly_better else "No clearly superior same-sector company on the current peer screen"),
        "candidate": best if clearly_better else None,
        "best": best,
        "target": target,
        "score_gap": gap,
        "reasons": reasons,
    }


def _write_people_leadership_sheet(wb, ticker: str, info: dict, employee: dict, leadership_score: float, leadership_rows: list[tuple], alt: dict):
    if "Leadership & Culture" in wb.sheetnames:
        wb.remove(wb["Leadership & Culture"])
    ws = wb.create_sheet("Leadership & Culture")
    ws.sheet_view.showGridLines = False
    for c in range(1, 11):
        ws.cell(1, c).fill = _fill(NAVY)
        ws.cell(2, c).fill = _fill(NAVY)
    ws["A1"] = f"{ticker} — Leadership, Workforce & Alternative Screen"
    ws["A1"].font = Font(bold=True, color=WHITE, size=18)
    ws["A3"] = "Leadership and employee evidence is source-scoped. The same-sector alternative is a quantitative research screen, not personalized investment advice."
    ws["A3"].font = Font(italic=True, color=GREY)
    ws["A3"].alignment = Alignment(wrap_text=True)

    _section(ws, 5, "Worker Happiness / Employee Experience Evidence")
    heads = ["Metric", "Value", "Scope", "Period", "Status", "Source"]
    for c, v in enumerate(heads, 1):
        ws.cell(6, c, v)
    _header(ws, 6, 1, 6)
    ws.cell(7, 1, "Worker happiness / satisfaction signal")
    ws.cell(7, 2, employee.get("score"))
    ws.cell(7, 2).number_format = "0.0"
    ws.cell(7, 3, employee.get("scope"))
    ws.cell(7, 4, employee.get("period"))
    ws.cell(7, 5, employee.get("status"))
    ws.cell(7, 6, employee.get("source"))
    ws.cell(8, 1, "Evidence")
    ws.cell(8, 2, employee.get("evidence"))
    ws.merge_cells("B8:F8")
    ws["B8"].alignment = Alignment(wrap_text=True)
    if employee.get("source"):
        ws["F7"].hyperlink = employee.get("source")
        ws["F7"].font = Font(color=LINK_GREEN, underline="single")

    _section(ws, 10, "Leadership Evidence Score — Transparent Proxy")
    ws["A11"] = "Composite proxy / 100"
    ws["B11"] = leadership_score
    ws["B11"].number_format = FMT_SCORE
    ws["C11"] = "Uses execution, capital allocation, leadership depth, culture and governance disclosure. Treat as a research organizer, not a factual management rating."
    ws.merge_cells("C11:F11")
    ws["C11"].alignment = Alignment(wrap_text=True)
    for c, v in enumerate(["Dimension", "Score / 100", "Evidence / Caveat"], 1):
        ws.cell(13, c, v)
    _header(ws, 13, 1, 3)
    for r, (name, score, note) in enumerate(leadership_rows, 14):
        ws.cell(r, 1, name)
        ws.cell(r, 2, score)
        ws.cell(r, 2).number_format = FMT_SCORE
        ws.cell(r, 3, note)
        ws.cell(r, 3).alignment = Alignment(wrap_text=True)

    _section(ws, 21, "Executive Team — Public Snapshot")
    for c, v in enumerate(["Name", "Title", "Age / Birth Year", "Reported Pay", "Source"], 1):
        ws.cell(22, c, v)
    _header(ws, 22, 1, 5)
    officers = info.get("companyOfficers") or []
    source_urls = _official_source_urls(ticker, info)
    leadership_source = source_urls.get("leadership") or source_urls.get("governance") or source_urls.get("annual")
    row = 23
    for officer in officers[:8]:
        if not isinstance(officer, dict):
            continue
        ws.cell(row, 1, officer.get("name"))
        ws.cell(row, 2, officer.get("title"))
        age = officer.get("age") or officer.get("yearBorn")
        ws.cell(row, 3, age)
        ws.cell(row, 4, officer.get("totalPay"))
        ws.cell(row, 5, leadership_source)
        if leadership_source:
            ws.cell(row, 5).hyperlink = leadership_source
            ws.cell(row, 5).font = Font(color=LINK_GREEN, underline="single")
        row += 1
    if row == 23:
        ws["A23"] = "Public officer data unavailable; use the official leadership / annual-report source below."

    _section(ws, 33, "Same-Sector Alternative Screen")
    for c, v in enumerate(["Item", "Result", "Score", "Why / Evidence"], 1):
        ws.cell(34, c, v)
    _header(ws, 34, 1, 4)
    target = alt.get("target") or {}
    best = alt.get("best") or {}
    candidate = alt.get("candidate") or {}
    ws.cell(35, 1, "Current company peer-screen score")
    ws.cell(35, 2, target.get("ticker") or ticker)
    ws.cell(35, 3, target.get("screen_score"))
    ws.cell(36, 1, "Best peer on current metrics")
    ws.cell(36, 2, best.get("ticker"))
    ws.cell(36, 3, best.get("screen_score"))
    ws.cell(36, 4, ", ".join(alt.get("reasons") or []))
    ws.cell(37, 1, "Research conclusion")
    ws.cell(37, 2, alt.get("result"))
    ws.merge_cells("B37:D37")
    ws["B37"].alignment = Alignment(wrap_text=True)
    if candidate:
        ws.cell(38, 1, "Candidate for deeper research")
        ws.cell(38, 2, candidate.get("ticker"))
        ws.cell(38, 3, candidate.get("screen_score"))
        ws.cell(38, 4, "Screen only: validate business quality, balance sheet, market structure, valuation and risks before preferring it.")
        ws.cell(38, 4).alignment = Alignment(wrap_text=True)

    _section(ws, 41, "Official / Manual Research Sources")
    for c, v in enumerate(["Topic", "Preferred Source"], 1):
        ws.cell(42, c, v)
    _header(ws, 42, 1, 2)
    sources = _official_source_urls(ticker, info)
    sources.setdefault("employee reviews - optional corroboration", "Glassdoor / Indeed / Comparably — manual review only; not auto-scraped")
    rr = 43
    for topic, url in sources.items():
        ws.cell(rr, 1, topic)
        ws.cell(rr, 2, url)
        if isinstance(url, str) and url.startswith("http"):
            ws.cell(rr, 2).hyperlink = url
            ws.cell(rr, 2).font = Font(color=LINK_GREEN, underline="single")
        rr += 1

    widths = {"A": 34, "B": 26, "C": 54, "D": 24, "E": 22, "F": 55}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A6"


def _dashboard_peer_direct(wb):
    if "Dashboard" not in wb.sheetnames or "Peer Comps" not in wb.sheetnames:
        return
    ws = wb["Dashboard"]
    specs = [
        (19, "C", False),
        (20, "D", False),
        (21, "E", False),
        (22, "F", True),
        (23, "G", True),
        (24, "H", True),
    ]
    for r, col, higher in specs:
        ws.cell(r, 2, f"='Peer Comps'!{col}4")
        ws.cell(r, 3, f'=IFERROR(MEDIAN(\'Peer Comps\'!{col}5:{col}9),"")')
        comp = ">" if higher else "<"
        good = "Better" if higher else "Attractive"
        bad = "Worse" if higher else "Premium"
        ws.cell(r, 4, f'=IF(OR(B{r}="",C{r}=""),"",IF(B{r}{comp}C{r},"{good}","{bad}"))')


def _append_dashboard_people(wb, ticker: str, employee: dict, leadership_score: float, alt: dict):
    if "Dashboard" not in wb.sheetnames:
        return
    ws = wb["Dashboard"]
    start = max(33, ws.max_row + 2)
    ws.cell(start, 1, "People, Leadership & Market Position")
    for c in range(1, 5):
        ws.cell(start, c).fill = _fill(NAVY)
        ws.cell(start, c).font = Font(bold=True, color=WHITE)
    rows = [
        ("Worker happiness / satisfaction signal", employee.get("score"), employee.get("scope")),
        ("Leadership evidence proxy / 100", leadership_score, "See Leadership & Culture for transparent component scores"),
    ]
    share = market_share_record(ticker)
    rows.append(("Comparable industry market share", share.get("share"), (share.get("basis") or "No comparable source") + ("; " + share.get("period", "") if share else "")))
    rows.append(("Same-sector alternative screen", (alt.get("candidate") or {}).get("ticker") or "None clearly superior", alt.get("result")))
    for i, (label, value, note) in enumerate(rows, start + 1):
        ws.cell(i, 1, label)
        ws.cell(i, 2, value)
        ws.cell(i, 3, note)
        ws.cell(i, 3).alignment = Alignment(wrap_text=True)
        if "market share" in label.lower() and isinstance(value, (int, float)):
            ws.cell(i, 2).number_format = FMT_PCT
        elif isinstance(value, (int, float)):
            ws.cell(i, 2).number_format = FMT_SCORE
    ws.column_dimensions["C"].width = max(ws.column_dimensions["C"].width or 0, 52)


def _append_investment_summary(wb, employee: dict, leadership_score: float, alt: dict):
    if "Investment Summary" not in wb.sheetnames:
        return
    ws = wb["Investment Summary"]
    row = ws.max_row + 2
    for c in range(1, 9):
        ws.cell(row, c).fill = _fill(NAVY)
        ws.cell(row, c).font = Font(bold=True, color=WHITE)
    ws.cell(row, 1, "People, Leadership & Competitive Position")
    values = [
        ("Worker happiness / employee signal", employee.get("score"), employee.get("scope")),
        ("Leadership evidence proxy / 100", leadership_score, "See Leadership & Culture"),
        ("Same-sector alternative", (alt.get("candidate") or {}).get("ticker") or "None clearly superior", alt.get("result")),
    ]
    for rr, (label, value, note) in enumerate(values, row + 1):
        ws.cell(rr, 1, label)
        ws.cell(rr, 2, value)
        ws.cell(rr, 3, note)
        ws.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=8)
        ws.cell(rr, 3).alignment = Alignment(wrap_text=True)


def _data_quality_checks(wb, ticker: str, employee: dict):
    if "Data Quality" not in wb.sheetnames:
        return
    ws = wb["Data Quality"]
    existing = {str(ws.cell(r, 1).value or "").strip(): r for r in range(1, ws.max_row + 1)}
    checks = [
        ("Market-share comparability", "PASS" if market_share_record(ticker) else "REVIEW", (market_share_record(ticker).get("basis") if market_share_record(ticker) else "No comparable public industry-share source mapped"), "Market share is populated only on a like-for-like industry basis; blanks are preferable to false precision."),
        ("Employee sentiment scope", "PASS" if employee.get("score") is not None else "REVIEW", employee.get("scope"), "Program satisfaction, engagement, and company-wide happiness are different measures and must be labeled by scope."),
    ]
    for name, status, observed, why in checks:
        r = existing.get(name) or ws.max_row + 1
        ws.cell(r, 1, name)
        ws.cell(r, 2, status)
        ws.cell(r, 3, observed)
        ws.cell(r, 4, why)
        ws.cell(r, 2).fill = _fill(PALE_GREEN if status == "PASS" else GOLD)
        ws.cell(r, 2).font = Font(bold=True)
        for c in range(1, 5):
            ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")


def _remove_duplicate_presentation_sheets(wb):
    """Remove only sheets that repeat an authoritative source without unique workflow value."""
    _dashboard_peer_direct(wb)
    for name in ("Visual Dashboard", "Comparative Analysis", "Valuation Cross-Checks"):
        if name in wb.sheetnames:
            wb.remove(wb[name])


def ensure_research_extensions(wb, ticker: str, info: dict | None = None):
    info = _safe_info(ticker, info)
    employee = _employee_signal(ticker, info)
    leadership_score, leadership_rows = _leadership_proxy(wb, ticker, info, employee)
    alt = _alternative_screen(wb, ticker)
    _write_people_leadership_sheet(wb, ticker, info, employee, leadership_score, leadership_rows, alt)
    _append_dashboard_people(wb, ticker, employee, leadership_score, alt)
    _append_investment_summary(wb, employee, leadership_score, alt)
    _data_quality_checks(wb, ticker, employee)
    _remove_duplicate_presentation_sheets(wb)
    return {
        "employee": employee,
        "leadership_score": leadership_score,
        "alternative": alt,
    }
