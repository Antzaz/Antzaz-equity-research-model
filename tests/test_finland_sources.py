from __future__ import annotations

import finland_sources
import issuer_source_engine
from source_registry import full_source_catalog, market_sources


def _sample_esef_payload():
    return {
        "documentInfo": {"documentType": "https://xbrl.org/2021/xbrl-json"},
        "facts": {
            "rev": {
                "value": "1000000000",
                "dimensions": {
                    "concept": "ifrs-full:Revenue",
                    "entity": "lei:TEST",
                    "period": "2024-01-01T00:00:00/2025-01-01T00:00:00",
                    "unit": "iso4217:EUR",
                },
            },
            "op": {
                "value": "140000000",
                "dimensions": {
                    "concept": "ifrs-full:ProfitLossFromOperatingActivities",
                    "entity": "lei:TEST",
                    "period": "2024-01-01T00:00:00/2025-01-01T00:00:00",
                    "unit": "iso4217:EUR",
                },
            },
            "ni": {
                "value": "100000000",
                "dimensions": {
                    "concept": "ifrs-full:ProfitLoss",
                    "entity": "lei:TEST",
                    "period": "2024-01-01T00:00:00/2025-01-01T00:00:00",
                    "unit": "iso4217:EUR",
                },
            },
            "eps": {
                "value": "2.0",
                "dimensions": {
                    "concept": "ifrs-full:DilutedEarningsLossPerShare",
                    "entity": "lei:TEST",
                    "period": "2024-01-01T00:00:00/2025-01-01T00:00:00",
                    "unit": "iso4217:EUR/xbrli:shares",
                },
            },
            "assets": {
                "value": "2500000000",
                "dimensions": {
                    "concept": "ifrs-full:Assets",
                    "entity": "lei:TEST",
                    "period": "2025-01-01T00:00:00",
                    "unit": "iso4217:EUR",
                },
            },
            # A segment/dimensioned fact must never overwrite consolidated history.
            "segment_rev": {
                "value": "9000000000",
                "dimensions": {
                    "concept": "ifrs-full:Revenue",
                    "entity": "lei:TEST",
                    "period": "2024-01-01T00:00:00/2025-01-01T00:00:00",
                    "unit": "iso4217:EUR",
                    "test:SegmentAxis": "test:SegmentMember",
                },
            },
        },
    }


def test_finland_source_routing_is_specific_to_helsinki():
    assert finland_sources.is_finnish_ticker("KNEBV.HE")
    assert finland_sources.is_finnish_ticker("nokia.he")
    assert not finland_sources.is_finnish_ticker("NOK")
    sources=market_sources("KNEBV.HE")
    assert "finland_oam" in sources
    assert "xbrl_international_esef" in sources
    assert "statistics_finland" in sources
    assert market_sources("MSFT")=={}
    catalog=full_source_catalog("KNEBV.HE")
    assert catalog["market"]["bank_of_finland"]["env_key"] is None


def test_esef_parser_keeps_consolidated_annual_facts_only():
    hist,meta=finland_sources._extract_esef_report(_sample_esef_payload())
    assert hist[2024]["revenue"]==1_000_000_000
    assert hist[2024]["op"]==140_000_000
    assert hist[2024]["ni"]==100_000_000
    assert hist[2024]["eps"]==2.0
    assert hist[2024]["shares"]==50_000_000
    assert hist[2024]["assets"]==2_500_000_000
    assert meta["years"]==[2024]


def test_finland_esef_history_uses_free_json_filing(monkeypatch):
    monkeypatch.setattr(
        finland_sources,
        "_filing_json_urls",
        lambda ticker,info,limit=8: (
            ["https://filings.xbrl.org/test.json"],
            {"status":"PASS","matched_name":"test"},
        ),
    )
    monkeypatch.setattr(
        finland_sources,
        "_request_json",
        lambda url,params=None,timeout=20: _sample_esef_payload(),
    )
    hist,meta=finland_sources.finland_esef_history(
        "TEST.HE",
        {"longName":"Test Oyj"},
    )
    assert hist[2024]["revenue"]==1_000_000_000
    assert meta["status"]=="PASS"
    assert "Finnish OAM" in meta["source"]


def test_crossborder_history_prioritizes_finnish_esef(monkeypatch):
    monkeypatch.setattr(
        issuer_source_engine,
        "finland_esef_history",
        lambda ticker,info: (
            {2024:{"revenue":100.0,"ni":10.0}},
            {"status":"PASS","years":[2024]},
        ),
    )
    monkeypatch.setattr(
        issuer_source_engine,
        "issuer_website_history",
        lambda ticker,info: (
            {2024:{"revenue":90.0,"op":12.0}},
            {"source":"issuer"},
        ),
    )
    monkeypatch.setattr(
        issuer_source_engine,
        "sec_crossborder_history",
        lambda facts: ({},{"source":"sec"}),
    )
    monkeypatch.setattr(
        issuer_source_engine,
        "yfinance_statement_history",
        lambda ticker: (
            {2024:{"revenue":80.0,"op":11.0,"cash":5.0}},
            {"source":"yahoo"},
        ),
    )

    hist,meta=issuer_source_engine.build_crossborder_history(
        "TEST.HE",
        {"longName":"Test Oyj"},
        None,
    )
    assert hist[2024]["revenue"]==100.0
    assert hist[2024]["op"]==12.0
    assert hist[2024]["cash"]==5.0
    assert meta["priority"][0].startswith("Finnish OAM")


def test_guarded_filings_wrapper_forwards_company_info(monkeypatch):
    import safe_update_model
    from openpyxl import Workbook

    captured={}
    def fake_update_filings(wb,ticker,info=None):
        captured["ticker"]=ticker
        captured["info"]=info
        return "ok"

    monkeypatch.setattr(safe_update_model,"_ORIGINAL_UPDATE_FILINGS",fake_update_filings)
    wb=Workbook()
    wb.active.title="Filings"
    result=safe_update_model._safe_update_filings(
        wb,
        "YIT.HE",
        {"longName":"YIT Oyj"},
    )
    assert result=="ok"
    assert captured["ticker"]=="YIT.HE"
    assert captured["info"]["longName"]=="YIT Oyj"
