"""Extended visualization layer: core dashboard + business mix + advanced valuation snapshots."""

from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.chart import BarChart, Reference

from visualization import ensure_visual_dashboard as ensure_core_visual_dashboard

NAVY = "17365D"
BLUE = "2F75B5"
WHITE = "FFFFFF"
LIGHT = "F5F9FC"
LINK_GREEN = "008000"

FMT_PCT = '0.0%;[Red](0.0%);-'
FMT_PRICE = '$#,##0.00;[Red]($#,##0.00);-'
FMT_BN = '#,##0.0;[Red](#,##0.0);-'
FMT_MULT = '0.0x;[Red](0.0x);-'


def _fill(color):
    return PatternFill("solid", fgColor=color)


def _section(ws, start_cell, end_cell, title):
    rng = f"{start_cell}:{end_cell}"
    if rng not in {str(r) for r in ws.merged_cells.ranges}:
        ws.merge_cells(rng)
    cell = ws[start_cell]
    cell.value = title
    cell.fill = _fill(NAVY)
    cell.font = Font(bold=True, color=WHITE, size=11)


def ensure_visual_dashboard(wb, ticker=None):
    ws = ensure_core_visual_dashboard(wb, ticker)
    if ws is None:
        return None

    # Business-mix extension.
    if "Segment Analysis" in wb.sheetnames:
        seg = wb["Segment Analysis"]
        _section(ws, "A63", "H63", "Business Mix")
        ws["A64"], ws["B64"], ws["C64"], ws["D64"] = "Segment", "Year -2", "Year -1", "Latest"
        for c in range(1, 5):
            ws.cell(64, c).fill = _fill(BLUE)
            ws.cell(64, c).font = Font(bold=True, color=WHITE)
        dst = 65
        for src in range(7, 12):
            label = seg.cell(src, 1).value
            if not label:
                continue
            ws.cell(dst, 1, f"='Segment Analysis'!A{src}")
            ws.cell(dst, 2, f"='Segment Analysis'!B{src}")
            ws.cell(dst, 3, f"='Segment Analysis'!C{src}")
            ws.cell(dst, 4, f"='Segment Analysis'!D{src}")
            for c in range(2, 5):
                ws.cell(dst, c).number_format = FMT_BN
                ws.cell(dst, c).font = Font(color=LINK_GREEN)
            dst += 1
        if dst > 65:
            chart = BarChart()
            chart.type = "col"
            chart.style = 10
            chart.title = "Segment Revenue Growth"
            chart.height = 7.5
            chart.width = 13
            chart.legend.position = "b"
            chart.add_data(Reference(ws, min_col=2, max_col=4, min_row=64, max_row=dst - 1), titles_from_data=True)
            chart.set_categories(Reference(ws, min_col=1, min_row=65, max_row=dst - 1))
            ws.add_chart(chart, "A71")

    # Advanced valuation extension.
    if "Advanced Analytics" in wb.sheetnames:
        adv = wb["Advanced Analytics"]
        _section(ws, "I63", "P63", "Advanced Valuation")
        rows = [
            ("Reverse DCF Implied FCF CAGR", "='Advanced Analytics'!B38", FMT_PCT),
            ("Monte Carlo Median Value", "='Advanced Analytics'!J35", FMT_PRICE),
            ("Monte Carlo P10", "='Advanced Analytics'!J33", FMT_PRICE),
            ("Probability > Current Price", "='Advanced Analytics'!J38", FMT_PCT),
            ("Composite Investment Score", "='Advanced Analytics'!F42", "0.0"),
            ("Current Forward P/E", "='Company Data'!B15", FMT_MULT),
        ]
        ws["I64"], ws["J64"] = "Metric", "Result"
        for c in (9, 10):
            ws.cell(64, c).fill = _fill(BLUE)
            ws.cell(64, c).font = Font(bold=True, color=WHITE)
        for r, (label, formula, fmt) in enumerate(rows, 65):
            ws.cell(r, 9, label)
            ws.cell(r, 10, formula)
            ws.cell(r, 10).number_format = fmt
            ws.cell(r, 10).font = Font(bold=True, color=LINK_GREEN)

        ws.merge_cells("I73:P76")
        ws["I73"] = (
            "Use Financial Statements for clean reported numbers, Segment Analysis for business-line economics, "
            "and Advanced Analytics for expectations, earnings surprises, Monte Carlo valuation and scorecard diagnostics."
        )
        ws["I73"].fill = _fill(LIGHT)
        ws["I73"].alignment = Alignment(wrap_text=True, vertical="center")

    ws.column_dimensions["I"].width = 28
    ws.column_dimensions["J"].width = 18
    return ws
