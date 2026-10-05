from __future__ import annotations

import numpy as np
import pandas as pd

from institutional_research.src.short_horizon_rf import _one_ticker_path


def test_rf_weekly_path_has_five_business_days():
    idx=pd.bdate_range("2024-01-02",periods=420)
    rng=np.random.default_rng(7)
    prices=pd.Series(100*np.cumprod(1+rng.normal(.0004,.012,len(idx))),index=idx)
    rows,meta=_one_ticker_path(prices,horizon=5)
    assert meta["status"]=="PASS"
    assert len(rows)==5
    assert all(float(r["predicted_price"])>0 for r in rows)
    dates=pd.to_datetime([r["date"] for r in rows])
    assert all(d.weekday()<5 for d in dates)
