"""Financial-statement and business-segment analysis for the equity research workbook."""

import re
from io import StringIO

import requests

try:
    import pandas as pd
except Exception:
    pd = None

from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter

NAVY = "17365D"
BLUE = "2F75B5"
WHITE = "FFFFFF"
LIGHT = "F5F9FC"
PALE_BLUE = "D9EAF7"
GOLD = "FFF2CC"
INPUT_BLUE = "0000FF"
GREY = "666666"

FMT_BN = '#,##0.0;[Red](#,##0.0);-'
FMT_PCT = '0.0%;[Red](0.0%);-'
FMT_EPS = '$0.00;[Red]($0.00);-'
THIN = Side(style="thin", color="808080")


def _fill(color):
    return PatternFill("solid", fgColor=color)


def _merged_annual_series(facts, tags, preferred_unit=None):
    """Merge several alternative US-GAAP tags by fiscal year, preferring latest filed value."""
    if not facts:
        return {}
    gaap = facts.get("facts", {}).get("us-gaap", {})
    best = {}
    for tag in tags:
        fact = gaap.get(tag)
        if not fact:
            continue
        units = fact.get("units", {})
        unit = preferred_unit if preferred_unit in units else (next(iter(units)) if units else None)
        if not unit:
            continue
        for x in units.get(unit, []):
            if x.get("form") not in ("10-K", "10-K/A") or x.get("fp") != "FY":
                continue
            fy = x.get("fy")
            val = x.get("val")
            if fy is None or val is None:
                continue
            try:
                fy = int(fy)
                val = float(val)
            except Exception:
                continue
            stamp = str(x.get("filed") or "") + str(x.get("end") or "")
            if fy not in best or stamp >= best[fy][0]:
                best[fy] = (stamp, val)
    return {y: v for y, (_, v) in best.items()}


def _scale(v):
    return None if v is None else float(v) / 1e9


def _title(ws, text):
    ws.merge_cells("A1:H2")
    ws["A1"] = text
    ws["A1"].fill = _fill(NAVY)
    ws["A1"].font = Font(bold=True, color=WHITE, size=18)
    ws["A1"].alignment = Alignment(vertical="center")
    ws.sheet_view.showGridLines = False


def _section(ws, row, title, end_col="H"):
    ws.merge_cells(f"A{row}:{end_col}{row}")
    ws[f"A{row}"] = title
    ws[f"A{row}"].fill = _fill(NAVY)
    ws[f"A{row}"].font = Font(bold=True, color=WHITE, size=12)


def _header(ws, row, cols):
    for c in range(1, cols + 1):
        cell = ws.cell(row, c)
        cell.fill = _fill(BLUE)
        cell.font = Font(bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def ensure_financial_statements(wb, ticker, facts):
    """Create a clean, readable Income Statement / Balance Sheet / Cash Flow sheet from SEC Company Facts."""
    if "Financial Statements" in wb.sheetnames:
        wb.remove(wb["Financial Statements"])
    ws = wb.create_sheet("Financial Statements")
    _title(ws, f"{ticker} — Financial Statements")
    ws.merge_cells("A3:H3")
    ws["A3"] = "USD billions unless per-share data. Annual SEC Company Facts; missing tags are left blank rather than estimated."
    ws["A3"].font = Font(italic=True, color=GREY)

    income_map = [
        ("Revenue", ["RevenueFromContractWithCustomerExcludingAssessedTax", "SalesRevenueNet", "Revenues"], None),
        ("Cost of Revenue", ["CostOfRevenue", "CostOfGoodsAndServicesSold"], None),
        ("Gross Profit", ["GrossProfit"], None),
        ("Research & Development", ["ResearchAndDevelopmentExpense"], None),
        ("Sales & Marketing", ["SellingAndMarketingExpense", "MarketingExpense"], None),
        ("General & Administrative", ["GeneralAndAdministrativeExpense"], None),
        ("Operating Income", ["OperatingIncomeLoss"], None),
        ("Other Income / (Expense), Net", ["NonoperatingIncomeExpense", "OtherNonoperatingIncomeExpense"], None),
        ("Pre-Tax Income", ["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"], None),
        ("Income Taxes", ["IncomeTaxExpenseBenefit"], None),
        ("Net Income", ["NetIncomeLoss", "ProfitLoss"], None),
        ("Diluted EPS", ["EarningsPerShareDiluted"], "USD/shares"),
    ]

    series = {name: _merged_annual_series(facts, tags, unit) for name, tags, unit in income_map}
    years = sorted(series["Revenue"])[-6:] if series.get("Revenue") else []
    if not years and "Historical Financials" in wb.sheetnames:
        hist = wb["Historical Financials"]
        years = [int(hist.cell(3, c).value) for c in range(2, 8) if hist.cell(3, c).value]

    _section(ws, 5, "Income Statement")
    ws.cell(6, 1, "Metric")
    for j, y in enumerate(years, 2):
        ws.cell(6, j, y)
    _header(ws, 6, max(1, len(years) + 1))

    row = 7
    for name, _, unit in income_map:
        ws.cell(row, 1, name)
        for j, y in enumerate(years, 2):
            val = series.get(name, {}).get(y)
            if name == "Gross Profit" and val is None:
                rev = series.get("Revenue", {}).get(y)
                cost = series.get("Cost of Revenue", {}).get(y)
                if rev is not None and cost is not None:
                    val = rev - cost
            ws.cell(row, j, val if unit == "USD/shares" else _scale(val))
            ws.cell(row, j).number_format = FMT_EPS if unit == "USD/shares" else FMT_BN
        row += 1

    # Operating margin directly under operating income.
    op_row = 13
    ws.insert_rows(op_row + 1, 1)
    ws.cell(op_row + 1, 1, "Operating Margin")
    for j in range(2, 2 + len(years)):
        col = get_column_letter(j)
        ws.cell(op_row + 1, j, f'=IFERROR({col}{op_row}/{col}7,"")')
        ws.cell(op_row + 1, j).number_format = FMT_PCT
    ws.row_dimensions[op_row + 1].outlineLevel = 1

    # Balance-sheet tags.
    bs_map = [
        ("Cash & Cash Equivalents", ["CashAndCashEquivalentsAtCarryingValue"]),
        ("Marketable / Short-Term Securities", ["MarketableSecuritiesCurrent", "ShortTermInvestments"]),
        ("Accounts Receivable", ["AccountsReceivableNetCurrent"]),
        ("Other Current Assets", ["OtherCurrentAssets"]),
        ("Total Current Assets", ["AssetsCurrent"]),
        ("Property & Equipment, Net", ["PropertyPlantAndEquipmentNet"]),
        ("Goodwill", ["Goodwill"]),
        ("Total Assets", ["Assets"]),
        ("Accounts Payable", ["AccountsPayableCurrent"]),
        ("Deferred Revenue", ["ContractWithCustomerLiabilityCurrent", "DeferredRevenueCurrent"]),
        ("Total Current Liabilities", ["LiabilitiesCurrent"]),
        ("Long-Term Debt", ["LongTermDebtNoncurrent", "LongTermDebt"]),
        ("Total Liabilities", ["Liabilities"]),
        ("Stockholders' Equity", ["StockholdersEquity", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"]),
    ]
    bs_series = {name: _merged_annual_series(facts, tags) for name, tags in bs_map}
    bs_years = sorted(bs_series.get("Total Assets", {}))[-4:]

    bs_start = 23
    _section(ws, bs_start - 1, "Balance Sheet")
    ws.cell(bs_start, 1, "Metric")
    for j, y in enumerate(bs_years, 2):
        ws.cell(bs_start, j, y)
    _header(ws, bs_start, max(1, len(bs_years) + 1))
    row = bs_start + 1
    for name, _ in bs_map:
        ws.cell(row, 1, name)
        for j, y in enumerate(bs_years, 2):
            ws.cell(row, j, _scale(bs_series.get(name, {}).get(y)))
            ws.cell(row, j).number_format = FMT_BN
        if name in {"Total Current Assets", "Total Assets", "Total Current Liabilities", "Total Liabilities", "Stockholders' Equity"}:
            ws.cell(row, 1).font = Font(bold=True)
            for c in range(1, max(2, len(bs_years) + 2)):
                ws.cell(row, c).border = Border(bottom=THIN)
        row += 1

    # Cash-flow statement.
    cf_map = [
        ("Net Income", ["NetIncomeLoss", "ProfitLoss"]),
        ("Depreciation & Amortization", ["DepreciationDepletionAndAmortization", "DepreciationDepletionAndAmortizationPropertyPlantAndEquipment"]),
        ("Stock-Based Compensation", ["ShareBasedCompensation"]),
        ("Operating Cash Flow", ["NetCashProvidedByUsedInOperatingActivities"]),
        ("Capital Expenditures", ["PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsToAcquireProductiveAssets"]),
        ("Acquisitions", ["PaymentsToAcquireBusinessesNetOfCashAcquired", "PaymentsToAcquireBusinessesNetOfCashAndCashEquivalentsAcquired"]),
        ("Share Repurchases", ["PaymentsForRepurchaseOfCommonStock"]),
        ("Dividends", ["PaymentsOfDividends"]),
        ("Debt Issuance", ["ProceedsFromIssuanceOfLongTermDebt", "ProceedsFromIssuanceOfDebt"]),
        ("Debt Repayments", ["RepaymentsOfLongTermDebt", "RepaymentsOfDebt"]),
        ("Ending Cash", ["CashAndCashEquivalentsAtCarryingValue"]),
    ]
    cf_series = {name: _merged_annual_series(facts, tags) for name, tags in cf_map}
    cf_years = years[-6:]
    cf_start = max(row + 2, 49)
    _section(ws, cf_start, "Cash Flow Statement")
    ws.cell(cf_start + 1, 1, "Metric")
    for j, y in enumerate(cf_years, 2):
        ws.cell(cf_start + 1, j, y)
    _header(ws, cf_start + 1, max(1, len(cf_years) + 1))
    row = cf_start + 2
    row_lookup = {}
    for name, _ in cf_map:
        row_lookup[name] = row
        ws.cell(row, 1, name)
        for j, y in enumerate(cf_years, 2):
            val = cf_series.get(name, {}).get(y)
            # Cash outflows are shown as negative in the analyst view.
            if name in {"Capital Expenditures", "Acquisitions", "Share Repurchases", "Dividends", "Debt Repayments"} and val is not None:
                val = -abs(val)
            ws.cell(row, j, _scale(val))
            ws.cell(row, j).number_format = FMT_BN
        row += 1

    # FCF formula after OCF/capex.
    insert_at = row_lookup.get("Capital Expenditures", cf_start + 6) + 1
    ws.insert_rows(insert_at, 1)
    ws.cell(insert_at, 1, "Free Cash Flow")
    for j in range(2, 2 + len(cf_years)):
        col = get_column_letter(j)
        ocf_row = row_lookup["Operating Cash Flow"]
        capex_row = row_lookup["Capital Expenditures"]
        ws.cell(insert_at, j, f"={col}{ocf_row}+{col}{capex_row}")
        ws.cell(insert_at, j).number_format = FMT_BN
    ws.cell(insert_at, 1).font = Font(bold=True)
    for c in range(1, max(2, len(cf_years) + 2)):
        ws.cell(insert_at, c).border = Border(bottom=THIN)

    ws.column_dimensions["A"].width = 40
    for c in range(2, 8):
        ws.column_dimensions[get_column_letter(c)].width = 14
    ws.freeze_panes = "B7"
    return ws


SEGMENT_CONFIGS = {
    "GOOGL": {
        "segments": ["Google Services", "Google Cloud", "Other Bets"],
        "business_lines": ["Google Search & other", "YouTube ads", "Google Network", "Google subscriptions, platforms, and devices", "Google Cloud", "Other Bets"],
    },
    "GOOG": {
        "segments": ["Google Services", "Google Cloud", "Other Bets"],
        "business_lines": ["Google Search & other", "YouTube ads", "Google Network", "Google subscriptions, platforms, and devices", "Google Cloud", "Other Bets"],
    },
    "MSFT": {"segments": ["Productivity and Business Processes", "Intelligent Cloud", "More Personal Computing"], "business_lines": []},
    "AMZN": {"segments": ["North America", "International", "AWS"], "business_lines": []},
    "META": {"segments": ["Family of Apps", "Reality Labs"], "business_lines": []},
    "NVDA": {"segments": ["Compute & Networking", "Graphics"], "business_lines": []},
}


def _latest_10k_html(ticker, headers):
    tickers = requests.get("https://www.sec.gov/files/company_tickers.json", headers=headers, timeout=30).json()
    cik = None
    for item in tickers.values():
        if str(item.get("ticker", "")).upper() == ticker.upper():
            cik = str(item["cik_str"]).zfill(10)
            break
    if not cik:
        return None, None
    subs = requests.get(f"https://data.sec.gov/submissions/CIK{cik}.json", headers=headers, timeout=30).json()
    recent = subs.get("filings", {}).get("recent", {})
    for form, acc, doc in zip(recent.get("form", []), recent.get("accessionNumber", []), recent.get("primaryDocument", [])):
        if form == "10-K":
            url = f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{acc.replace('-', '')}/{doc}"
            html = requests.get(url, headers=headers, timeout=45).text
            return html, url
    return None, None


def _numbers_from_row(row):
    nums = []
    for val in list(row):
        s = str(val)
        # Parentheses mean negatives; strip commas and currency symbols.
        for token in re.findall(r"\(?-?\d[\d,]*(?:\.\d+)?\)?", s):
            neg = token.startswith("(") and token.endswith(")")
            clean = token.strip("()").replace(",", "")
            try:
                num = float(clean)
            except Exception:
                continue
            if 1900 <= num <= 2100:
                continue
            nums.append(-num if neg else num)
    return nums


def _extract_rows_from_tables(html, labels):
    if not html or pd is None:
        return {}
    try:
        tables = pd.read_html(StringIO(html))
    except Exception:
        return {}
    out = {label: [] for label in labels}
    for df in tables:
        for _, row in df.iterrows():
            text = " | ".join(str(v) for v in row.tolist())
            for label in labels:
                if label.lower() in text.lower():
                    nums = _numbers_from_row(row)
                    if len(nums) >= 2:
                        out[label].append(nums[-3:])
    return out


def _manual_segment_sheet(wb, ticker, source_url=None):
    if "Segment Analysis" in wb.sheetnames:
        wb.remove(wb["Segment Analysis"])
    ws = wb.create_sheet("Segment Analysis")
    _title(ws, f"{ticker} — Business & Segment Analysis")
    ws.merge_cells("A3:H3")
    ws["A3"] = "Automatic segment extraction was unavailable. Enter disclosed segment data in the yellow cells; do not infer undisclosed segments."
    ws["A3"].font = Font(italic=True, color=GREY)
    _section(ws, 5, "Manual Segment Input")
    headers = ["Segment / Business Line", "Year -2 Revenue", "Year -1 Revenue", "Latest Revenue", "Latest Growth", "Latest Operating Income", "Latest Op. Margin", "Source / Notes"]
    for c, v in enumerate(headers, 1):
        ws.cell(6, c, v)
    _header(ws, 6, 8)
    for r in range(7, 19):
        for c in range(1, 9):
            ws.cell(r, c).fill = _fill(GOLD)
            ws.cell(r, c).font = Font(color=INPUT_BLUE)
        for c in range(2, 7):
            ws.cell(r, c).number_format = FMT_BN
        ws.cell(r, 5).number_format = FMT_PCT
        ws.cell(r, 7).number_format = FMT_PCT
    ws["A21"] = "SEC source"
    ws["B21"] = source_url or ""
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["H"].width = 44
    return ws


def ensure_segment_analysis(wb, ticker, headers):
    """Build company business-segment analysis. Supported mega-tech tickers use latest 10-K table extraction; otherwise create manual input."""
    cfg = SEGMENT_CONFIGS.get(ticker.upper())
    html, source_url = _latest_10k_html(ticker, headers)
    if not cfg or not html:
        return _manual_segment_sheet(wb, ticker, source_url)

    labels = list(dict.fromkeys(cfg.get("segments", []) + cfg.get("business_lines", [])))
    extracted = _extract_rows_from_tables(html, labels)
    if not any(extracted.get(label) for label in labels):
        return _manual_segment_sheet(wb, ticker, source_url)

    if "Segment Analysis" in wb.sheetnames:
        wb.remove(wb["Segment Analysis"])
    ws = wb.create_sheet("Segment Analysis")
    _title(ws, f"{ticker} — Business & Segment Analysis")
    ws.merge_cells("A3:P3")
    ws["A3"] = "Business mix, growth and segment economics from the latest annual filing. USD billions."
    ws["A3"].font = Font(italic=True, color=GREY)

    segments = cfg.get("segments", [])
    _section(ws, 5, "Reported Operating Segments")
    headers_row = ["Segment", "Year -2 Revenue", "Year -1 Revenue", "Latest Revenue", "Latest Growth", "Latest Op. Income", "Latest Op. Margin"]
    for c, v in enumerate(headers_row, 1):
        ws.cell(6, c, v)
    _header(ws, 6, 7)

    row = 7
    successful = 0
    for label in segments:
        matches = extracted.get(label, [])
        revenue = matches[0] if matches else []
        op_income = matches[1] if len(matches) > 1 else []
        if len(revenue) < 2:
            continue
        rev = [x / 1000 for x in revenue[-3:]]
        while len(rev) < 3:
            rev.insert(0, None)
        latest_oi = (op_income[-1] / 1000) if op_income else None
        ws.cell(row, 1, label)
        for c, v in enumerate(rev, 2):
            ws.cell(row, c, v)
            ws.cell(row, c).number_format = FMT_BN
        if rev[-1] is not None and rev[-2]:
            ws.cell(row, 5, rev[-1] / rev[-2] - 1)
            ws.cell(row, 5).number_format = FMT_PCT
        ws.cell(row, 6, latest_oi)
        ws.cell(row, 6).number_format = FMT_BN
        if latest_oi is not None and rev[-1]:
            ws.cell(row, 7, latest_oi / rev[-1])
            ws.cell(row, 7).number_format = FMT_PCT
        row += 1
        successful += 1

    business_lines = cfg.get("business_lines", [])
    if business_lines:
        start = max(row + 2, 13)
        _section(ws, start, "Revenue by Business Line")
        for c, v in enumerate(["Business Line", "Year -2", "Year -1", "Latest", "Latest Growth", "Latest Mix"], 1):
            ws.cell(start + 1, c, v)
        _header(ws, start + 1, 6)
        r = start + 2
        latest_total = None
        if "Historical Financials" in wb.sheetnames:
            latest_total = wb["Historical Financials"]["G4"].value
        for label in business_lines:
            matches = extracted.get(label, [])
            revenue = matches[0] if matches else []
            if len(revenue) < 2:
                continue
            rev = [x / 1000 for x in revenue[-3:]]
            while len(rev) < 3:
                rev.insert(0, None)
            ws.cell(r, 1, label)
            for c, v in enumerate(rev, 2):
                ws.cell(r, c, v)
                ws.cell(r, c).number_format = FMT_BN
            if rev[-1] is not None and rev[-2]:
                ws.cell(r, 5, rev[-1] / rev[-2] - 1)
                ws.cell(r, 5).number_format = FMT_PCT
            if latest_total and rev[-1] is not None:
                ws.cell(r, 6, rev[-1] / latest_total)
                ws.cell(r, 6).number_format = FMT_PCT
            r += 1
        if r > start + 2:
            ws.conditional_formatting.add(
                f"E{start+2}:E{r-1}",
                ColorScaleRule(start_type="min", start_color="F8696B", mid_type="percentile", mid_value=50, mid_color="FFEB84", end_type="max", end_color="63BE7B")
            )
            # Clean latest-revenue chart.
            helper_col = 18
            ws.cell(5, helper_col, "Business Line")
            ws.cell(5, helper_col + 1, "Latest Revenue")
            hr = 6
            for rr in range(start + 2, r):
                ws.cell(hr, helper_col, ws.cell(rr, 1).value)
                ws.cell(hr, helper_col + 1, ws.cell(rr, 4).value)
                hr += 1
            chart = BarChart()
            chart.type = "bar"
            chart.style = 10
            chart.title = "Latest Revenue by Business Line"
            chart.height = 7.5
            chart.width = 13.0
            chart.legend = None
            chart.add_data(Reference(ws, min_col=helper_col + 1, min_row=5, max_row=hr - 1), titles_from_data=True)
            chart.set_categories(Reference(ws, min_col=helper_col, min_row=6, max_row=hr - 1))
            ws.add_chart(chart, "I15")

    ws["A29"] = "SEC source"
    ws["B29"] = source_url or ""
    ws["B29"].font = Font(color="008000")
    ws.column_dimensions["A"].width = 35
    for c in range(2, 8):
        ws.column_dimensions[get_column_letter(c)].width = 14
    ws.column_dimensions["I"].width = 30
    ws.column_dimensions["J"].width = 18
    for c in range(18, 22):
        ws.column_dimensions[get_column_letter(c)].hidden = True

    if successful == 0:
        wb.remove(ws)
        return _manual_segment_sheet(wb, ticker, source_url)
    return ws
