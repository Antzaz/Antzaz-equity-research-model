from __future__ import annotations

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

from machine_learning.ai_growth_output import _set_chart_text_categories as set_ai_categories
from machine_learning.workbook import _set_chart_text_categories as set_ml_categories


def _sample_chart(ws):
    ws["X2"] = "Category"
    ws["Y2"] = "Model"
    ws["Z2"] = "Baseline"
    ws["X3"] = "Alpha"
    ws["X4"] = "Beta"
    ws["Y3"] = 10
    ws["Y4"] = 20
    ws["Z3"] = 12
    ws["Z4"] = 18
    chart = BarChart()
    chart.add_data(Reference(ws, min_col=25, max_col=26, min_row=3, max_row=4), titles_from_data=False)
    return chart


def _assert_chart_refs(chart, expected_titles):
    assert len(chart.series) == len(expected_titles)
    for series, title in zip(chart.series, expected_titles):
        assert series.cat.strRef is not None
        assert series.cat.numRef is None
        assert series.cat.strRef.f == "'Chart Test'!$X$3:$X$4"
        assert series.tx is not None
        assert series.tx.v == title
        assert series.tx.strRef is None


def test_ml_chart_categories_are_explicit_text_refs_with_literal_series_names():
    wb = Workbook()
    ws = wb.active
    ws.title = "Chart Test"
    chart = _sample_chart(ws)
    set_ml_categories(chart, ws, 24, 3, 4, ["Forecast magnitude", "Typical historical error"])
    _assert_chart_refs(chart, ["Forecast magnitude", "Typical historical error"])


def test_ai_growth_chart_categories_are_explicit_text_refs_with_literal_series_names():
    wb = Workbook()
    ws = wb.active
    ws.title = "Chart Test"
    chart = _sample_chart(ws)
    set_ai_categories(chart, ws, 24, 3, 4, ["Model", "Baseline"])
    _assert_chart_refs(chart, ["Model", "Baseline"])
