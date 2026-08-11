from __future__ import annotations

from pathlib import Path
import tempfile

import numpy as np
import pandas as pd

from machine_learning.history_store import HistoryStore


def test_history_store_builds_point_in_time_features():
    with tempfile.TemporaryDirectory() as td:
        store=HistoryStore(Path(td)/"history.sqlite")
        store.upsert_symbols([{"symbol":"AAA"},{"symbol":"SPY"}])
        dates=pd.bdate_range("2018-01-02","2024-12-31")
        aaa=100*np.cumprod(np.repeat(1.00035,len(dates)))
        spy=100*np.cumprod(np.repeat(1.00020,len(dates)))
        price_rows=[]
        for symbol,series in (("AAA",aaa),("SPY",spy)):
            for dt,px in zip(dates,series):
                price_rows.append({"symbol":symbol,"date":dt.date().isoformat(),"close":float(px),"adj_close":float(px),"volume":1_000_000,"source":"Yahoo Finance"})
        store.upsert_prices(price_rows)
        store.upsert_fundamentals([
            {"symbol":"AAA","fiscal_year":2019,"fiscal_date":"2019-12-31","available_date":"2020-04-29","period":"annual","revenue":1000,"operating_income":200,"net_income":120,"eps":1.2,"shares":100,"ocf":180,"capex":50,"fcf":130,"cash":100,"debt":200,"equity":600,"rd":80,"source":"SEC Company Facts"},
            {"symbol":"AAA","fiscal_year":2020,"fiscal_date":"2020-12-31","available_date":"2021-04-30","period":"annual","revenue":1100,"operating_income":230,"net_income":140,"eps":1.4,"shares":100,"ocf":205,"capex":55,"fcf":150,"cash":120,"debt":190,"equity":650,"rd":90,"source":"SEC Company Facts"},
            {"symbol":"AAA","fiscal_year":2021,"fiscal_date":"2021-12-31","available_date":"2022-04-30","period":"annual","revenue":1250,"operating_income":275,"net_income":165,"eps":1.65,"shares":100,"ocf":240,"capex":60,"fcf":180,"cash":140,"debt":180,"equity":700,"rd":100,"source":"SEC Company Facts"},
        ])
        n=store.build_features("SPY")
        assert n>=2
        frame=store.expected_return_frame(min_rows=1)
        assert not frame.empty
        assert (frame["target_date"]>frame["as_of"]).all()
        assert "price_to_sales" in frame.columns
        assert frame["revenue_growth"].notna().any()
        assert frame["target_excess_return_12m"].notna().all()


def test_upserts_are_idempotent_and_keys_are_not_stored():
    with tempfile.TemporaryDirectory() as td:
        store=HistoryStore(Path(td)/"history.sqlite")
        row={"symbol":"AAA","date":"2024-01-02","close":100,"adj_close":100,"volume":10,"source":"Yahoo Finance"}
        store.upsert_prices([row]); store.upsert_prices([row])
        assert store.table_counts()["prices"]==1
        with store.connect() as con:
            names=[r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")]
            assert "api_keys" not in names


if __name__=="__main__":
    test_history_store_builds_point_in_time_features()
    test_upserts_are_idempotent_and_keys_are_not_stored()
    print("persistent ML history tests passed")
