from __future__ import annotations

import ml_research as ml


class _FakeStore:
    def upsert_symbols(self, rows):
        return len(rows)

    def build_features(self, benchmark):
        return 150


def test_default_ml_search_targets_broad_twenty_year_history():
    args = ml.parser().parse_args(["GOOGL", "--no-workbook-write"])
    assert args.auto_history_limit == 500
    assert args.auto_history_years == 20
    assert args.auto_history_batch_size == 100
    assert args.max_universe == 50


def test_old_150_feature_row_database_is_not_declared_deep_ready(monkeypatch):
    symbols = [{"symbol": "GOOGL"}, {"symbol": "MSFT"}, {"symbol": "SPY"}]
    monkeypatch.setattr(ml, "_starter_universe", lambda *args, **kwargs: symbols)
    monkeypatch.setattr(ml, "_db_count", lambda store, table: 150 if table == "features" else 3)
    monkeypatch.setattr(ml, "_coverage_counts", lambda store, names: (3, 2))
    monkeypatch.setattr(
        ml,
        "_symbol_count",
        lambda store, table, symbol: 126 if table == "prices" else 2,
    )
    monkeypatch.setattr(ml, "yahoo_price_rows", lambda *args, **kwargs: iter(()))
    monkeypatch.setattr(ml, "yahoo_fundamental_rows", lambda *args, **kwargs: [])

    status = ml._auto_seed_history(
        _FakeStore(), "GOOGL", ["MSFT"], "SPY", limit=500, years=20, batch_size=100
    )
    assert status["attempted"] is True
    assert status["ready"] is False
    assert status["minimum_feature_rows"] >= 600
    assert "deepening" in status["note"]
