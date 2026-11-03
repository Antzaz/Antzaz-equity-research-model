from __future__ import annotations

import json
from pathlib import Path
import sys

import pandas as pd
import pytest

IR = Path(__file__).resolve().parents[1] / "institutional_research"
if str(IR) not in sys.path:
    sys.path.insert(0, str(IR))

from src import data
from src.export import write_outputs
import offline_portfolio_check


def _sample_close():
    dates = pd.bdate_range("2026-10-19", periods=10)
    return pd.DataFrame(
        {"AAA": list(range(100, 110)), "BBB": list(range(200, 210))},
        index=dates,
    )


def test_online_refresh_creates_durable_offline_quote_cache(tmp_path, monkeypatch):
    monkeypatch.setattr(data, "CACHE_DIR", tmp_path)
    monkeypatch.delenv("ANTZAZ_OFFLINE", raising=False)
    sample = _sample_close()
    monkeypatch.setattr(data, "_download_closes", lambda tickers, period: sample)
    prices = data.download_prices(["AAA", "BBB"], period="2y")
    assert prices.columns.tolist() == ["AAA", "BBB"]
    assert (tmp_path / "adjusted_closes.csv").is_file()
    assert data.price_cache_status()["cached_symbols"] == 2

    monkeypatch.setenv("ANTZAZ_OFFLINE", "1")
    monkeypatch.setattr(data, "_download_closes", lambda *args, **kwargs: pytest.fail("Offline contacted Yahoo"))
    cached = data.download_prices(["AAA", "BBB"])
    pd.testing.assert_frame_equal(cached, prices, check_freq=False)


def test_offline_refuses_to_invent_missing_quotes(tmp_path, monkeypatch):
    monkeypatch.setattr(data, "CACHE_DIR", tmp_path)
    monkeypatch.setenv("ANTZAZ_OFFLINE", "1")
    with pytest.raises(RuntimeError, match="Offline price cache"):
        data.download_prices(["MISSING"])


def test_live_provider_outage_uses_existing_local_cache(tmp_path, monkeypatch):
    monkeypatch.setattr(data, "CACHE_DIR", tmp_path)
    monkeypatch.delenv("ANTZAZ_OFFLINE", raising=False)
    monkeypatch.setattr(data, "_download_closes", lambda tickers, period: _sample_close())
    data.download_prices(["AAA", "BBB"])
    def fail(*args, **kwargs):
        raise ConnectionError("Network disconnected")
    monkeypatch.setattr(data, "_download_closes", fail)
    with pytest.warns(UserWarning, match="Live quotes unavailable"):
        cached = data.download_prices(["AAA", "BBB"])
    assert cached.iloc[-1]["AAA"] == 109


def test_offline_company_information_is_saved(tmp_path, monkeypatch):
    monkeypatch.setattr(data, "CACHE_DIR", tmp_path)
    monkeypatch.delenv("ANTZAZ_OFFLINE", raising=False)
    class FakeTicker:
        @property
        def info(self):
            return {"longName": "Example A", "currency": "USD"}
    monkeypatch.setattr(data.yf, "Ticker", lambda ticker: FakeTicker())
    assert data.fetch_info(["AAA"], pause=0)["AAA"]["currency"] == "USD"
    monkeypatch.setenv("ANTZAZ_OFFLINE", "1")
    monkeypatch.setattr(data.yf, "Ticker", lambda ticker: pytest.fail("Offline info contacted Yahoo"))
    assert data.fetch_info(["AAA"])["AAA"]["longName"] == "Example A"


def _sample_portfolio_folder(base: Path):
    base.mkdir(parents=True, exist_ok=True)
    (base / "portfolio.csv").write_text(
        "Ticker,Shares,AverageCost,ManualWeight,Notes\nAAA,2,100,,Test\nBBB,1,200,,Test\n",
        encoding="utf-8",
    )
    latest = base / "outputs" / "latest"
    latest.mkdir(parents=True)
    (latest / "summary.json").write_text(json.dumps({"portfolio": {}}), encoding="utf-8")
    pd.DataFrame({"Ticker": ["AAA", "BBB"], "Weight": [.5, .5]}).to_csv(
        latest / "holdings_analysis.csv", index=False
    )
    cache = base / ".offline_cache"
    cache.mkdir()
    _sample_close().to_csv(cache / "adjusted_closes.csv")


def test_readiness_check_matches_portfolio_and_cache(tmp_path, monkeypatch):
    _sample_portfolio_folder(tmp_path)
    monkeypatch.setattr(offline_portfolio_check.importlib.util, "find_spec", lambda name: object())
    ok, notes = offline_portfolio_check.check(tmp_path)
    assert ok
    assert any("Offline price history covers all 2" in line for line in notes)


def test_readiness_check_detects_changed_portfolio(tmp_path, monkeypatch):
    _sample_portfolio_folder(tmp_path)
    monkeypatch.setattr(offline_portfolio_check.importlib.util, "find_spec", lambda name: object())
    pd.DataFrame({"Ticker": ["AAA"], "Weight": [1.0]}).to_csv(
        tmp_path / "outputs" / "latest" / "holdings_analysis.csv", index=False
    )
    ok, notes = offline_portfolio_check.check(tmp_path)
    assert not ok
    assert "cached analysis" in " ".join(notes)


def test_offline_export_preserves_last_good_snapshot_on_error(tmp_path):
    out = tmp_path / "outputs"
    previous = out / "latest"
    previous.mkdir(parents=True)
    (previous / "summary.json").write_text('{"portfolio":{"old":true}}', encoding="utf-8")

    class FailingTable:
        def to_csv(self, *args, **kwargs):
            raise RuntimeError("Simulated write failure")

    with pytest.raises(RuntimeError, match="Simulated write failure"):
        write_outputs(out, {"bad_table": FailingTable()}, {"portfolio": {"new": True}})
    assert json.loads((previous / "summary.json").read_text(encoding="utf-8"))["portfolio"]["old"] is True


def test_standalone_launcher_never_refreshes_online():
    content = (IR / "start_dashboard.bat").read_text(encoding="utf-8").lower()
    assert "antzaz_offline=1" in content
    assert "pip install" not in content
    assert "python run_research.py" not in content



def test_random_forest_weekly_prediction_avoids_sklearn_feature_name_warnings(monkeypatch):
    import warnings
    import numpy as np
    from sklearn.ensemble import RandomForestRegressor
    from src import short_horizon_rf

    original_forest = RandomForestRegressor

    # Keep the test quick without changing production's 500-tree forecast.
    def small_forest(*args, **kwargs):
        kwargs["n_estimators"] = 8
        kwargs["n_jobs"] = 1
        return original_forest(*args, **kwargs)

    monkeypatch.setattr(short_horizon_rf, "RandomForestRegressor", small_forest)
    dates = pd.bdate_range("2024-01-02", periods=330)
    returns = np.sin(np.arange(330, dtype=float) / 13.0) * 0.004
    closes = pd.Series(100.0 * np.cumprod(1.0 + returns), index=dates)

    with warnings.catch_warnings(record=True) as recorded:
        warnings.simplefilter("always", UserWarning)
        forecast, metadata = short_horizon_rf._one_ticker_path(closes, horizon=5)

    feature_name_warnings = [
        str(w.message) for w in recorded
        if "feature names" in str(w.message).lower()
    ]
    assert feature_name_warnings == []
    assert metadata["status"] == "PASS"
    assert len(forecast) == 5
    assert all(np.isfinite(row["predicted_price"]) for row in forecast)
    assert all(row["predicted_price"] > 0 for row in forecast)



def test_first_launch_routes_to_setup_when_no_analysis():
    app_code = (IR / "app.py").read_text(encoding="utf-8")
    setup_page = IR / "pages" / "0_Portfolio_Setup.py"
    assert setup_page.exists()
    assert "HAS_ANALYSIS = " in app_code
    assert '"pages/0_Portfolio_Setup.py", "Setup & Holdings"' in app_code
    assert "not HAS_ANALYSIS" in app_code
    assert "Build / refresh portfolio analysis (online)" in setup_page.read_text(encoding="utf-8")


def test_windows_hub_launches_first_run_page_without_saved_outputs():
    hub = (IR.parent / "desktop_research_hub.py").read_text(encoding="utf-8")
    assert "No saved analytics yet. Opening Setup & Holdings" in hub
    assert 'if has_analysis:' in hub
    assert '"ANTZAZ_OFFLINE"] = "1"' in hub
    assert '"dashboard_startup.log"' in hub


def test_standalone_batch_opens_setup_when_analysis_missing():
    batch = (IR / "start_dashboard.bat").read_text(encoding="utf-8").lower()
    assert "the app will open setup" in batch
    assert 'if not exist "outputs\\latest\\summary.json"' in batch
    assert "python run_research.py" not in batch
    assert "set \"antzaz_offline=1\"" in batch


def test_dashboard_missing_analysis_has_recovery_instead_of_unhelpful_error():
    dashboard = (IR / "dashboard.py").read_text(encoding="utf-8")
    assert "has not generated" in dashboard
    assert "Open **Setup & Holdings**" in dashboard
    assert "No analysis outputs found. Run:" not in dashboard
