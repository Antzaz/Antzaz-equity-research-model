from __future__ import annotations

"""Excel chart-compatibility post-pass for generated research workbooks.

The ML and AI chart writers keep compact helper tables outside the main analysis area. Excel
normally excludes hidden source cells from chart plots, which can make text category labels
vanish even when the chart values still render. This post-pass keeps those helper columns
technically visible but extremely narrow and visually unobtrusive, so Excel consistently retains
category labels and series data.
"""

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font


HELPER_RANGES = {
    "ML & Quantitative Research": ("X", "AA", "X1:AA40"),
    "AI Growth Forecast": ("P", "Q", "P1:Q30"),
}


def apply_chart_compatibility_fix(workbook_path: str | Path) -> dict:
    path = Path(workbook_path)
    wb = load_workbook(path)
    touched: list[str] = []

    for sheet_name, (first_col, last_col, helper_range) in HELPER_RANGES.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        start = ws[first_col + "1"].column
        end = ws[last_col + "1"].column
        for col_idx in range(start, end + 1):
            letter = ws.cell(1, col_idx).column_letter
            dim = ws.column_dimensions[letter]
            # Do not hide chart-source columns: Excel's default plotVisOnly behavior otherwise
            # suppresses their category labels. Keep them narrow instead.
            dim.hidden = False
            dim.width = 1.5
        for row in ws[helper_range]:
            for cell in row:
                cell.font = Font(name="Aptos", size=1, color="FFFFFF")
        touched.append(sheet_name)

    if touched:
        wb.save(path)
    return {"workbook": str(path), "chart_helper_sheets": touched}


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Keep ML/AI chart helper sources Excel-visible without clutter.")
    parser.add_argument("workbook")
    args = parser.parse_args()
    print(apply_chart_compatibility_fix(args.workbook))
