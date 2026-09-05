from __future__ import annotations

"""Excel-safe chart helpers for generated research workbooks.

Excel normally excludes hidden chart-source cells from plots, and openpyxl chart axes can be
serialized ambiguously unless their positions are explicit. The primary chart generators use the
helpers in this module directly; ``apply_chart_compatibility_fix`` remains as a defensive post-pass
for older/reused workbooks.
"""

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font


HELPER_RANGES = {
    "ML & Quantitative Research": ("X", "AA", "X1:AA40"),
    "AI Growth Forecast": ("P", "Q", "P1:Q30"),
}


def make_chart_helpers_excel_visible(ws, first_col: str, last_col: str, helper_range: str) -> None:
    """Keep chart sources visible to Excel while making them unobtrusive to users."""
    start = ws[first_col + "1"].column
    end = ws[last_col + "1"].column
    for col_idx in range(start, end + 1):
        letter = ws.cell(1, col_idx).column_letter
        dim = ws.column_dimensions[letter]
        dim.hidden = False
        dim.width = 1.5
    for row in ws[helper_range]:
        for cell in row:
            cell.font = Font(name="Aptos", size=1, color="FFFFFF")


def is_horizontal_bar(chart) -> bool:
    """Detect chart orientation from the chart itself, never from list position."""
    direction = getattr(chart, "type", None)
    if not isinstance(direction, str):
        direction = getattr(chart, "barDir", None)
    value = getattr(direction, "val", direction)
    return str(value or "").lower() == "bar"


def configure_chart_for_excel(chart) -> None:
    """Make chart sources and axes explicit without changing chart-specific titles/scales."""
    chart.visible_cells_only = False
    x_axis = getattr(chart, "x_axis", None)
    y_axis = getattr(chart, "y_axis", None)
    if x_axis is None or y_axis is None:
        return

    x_axis.delete = False
    y_axis.delete = False
    x_axis.tickLblPos = "nextTo"
    y_axis.tickLblPos = "nextTo"

    if is_horizontal_bar(chart):
        # For openpyxl BarChart(type='bar'), x_axis is the category axis and y_axis the value axis.
        x_axis.axPos = "l"
        y_axis.axPos = "b"
    else:
        x_axis.axPos = "b"
        y_axis.axPos = "l"


def _repair_charts(ws) -> int:
    charts = list(getattr(ws, "_charts", []) or [])
    for chart in charts:
        configure_chart_for_excel(chart)
    return len(charts)


def apply_chart_compatibility_fix(workbook_path: str | Path) -> dict:
    path = Path(workbook_path)
    wb = load_workbook(path)
    touched: list[str] = []
    chart_count = 0

    for sheet_name, (first_col, last_col, helper_range) in HELPER_RANGES.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        make_chart_helpers_excel_visible(ws, first_col, last_col, helper_range)
        chart_count += _repair_charts(ws)
        touched.append(sheet_name)

    if touched:
        wb.save(path)
    return {
        "workbook": str(path),
        "chart_helper_sheets": touched,
        "charts_repaired": chart_count,
    }


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Repair ML/AI chart sources and axes for Excel compatibility.")
    parser.add_argument("workbook")
    args = parser.parse_args()
    print(apply_chart_compatibility_fix(args.workbook))
