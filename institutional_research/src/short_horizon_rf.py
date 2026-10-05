from __future__ import annotations

"""Short-horizon Random Forest price-path research for current portfolio holdings.

This module is deliberately separate from valuation and position sizing. It trains one model per
holding on lagged market features and recursively produces the next five business-day price path.
Forecasts are descriptive ML research signals, not target prices or trade instructions.
"""

from pathlib import Path
import json
import numpy as np
import pandas as pd
from sklearn.ensemble import RandomForestRegressor

from .data import download_prices

FEATURES=["ret_1","ret_2","ret_5","ret_10","ret_20","vol_5","vol_20","ma_gap_5","ma_gap_20","range_20"]
HORIZON_DAYS=5


def _features(close: pd.Series) -> pd.DataFrame:
    s=pd.to_numeric(close,errors="coerce").dropna()
    r=s.pct_change()
    out=pd.DataFrame(index=s.index)
    out["ret_1"]=r
    out["ret_2"]=s.pct_change(2)
    out["ret_5"]=s.pct_change(5)
    out["ret_10"]=s.pct_change(10)
    out["ret_20"]=s.pct_change(20)
    out["vol_5"]=r.rolling(5).std()
    out["vol_20"]=r.rolling(20).std()
    out["ma_gap_5"]=s/s.rolling(5).mean()-1
    out["ma_gap_20"]=s/s.rolling(20).mean()-1
    out["range_20"]=s/s.rolling(20).max()-1
    out["target"]=r.shift(-1)
    return out


def _one_ticker_path(close: pd.Series, horizon: int=HORIZON_DAYS) -> tuple[list[dict],dict]:
    hist=pd.to_numeric(close,errors="coerce").dropna()
    frame=_features(hist).dropna()
    if len(frame)<252:
        return [],{"status":"INSUFFICIENT_DATA","rows":int(len(frame))}
    train=frame.iloc[:-1].copy()
    model=RandomForestRegressor(
        n_estimators=500,max_depth=7,min_samples_leaf=5,max_features=.8,
        random_state=42,n_jobs=-1,
    )
    model.fit(train[FEATURES],train["target"])
    rolling=hist.copy(); rows=[]
    dates=pd.bdate_range(hist.index[-1]+pd.Timedelta(days=1),periods=horizon)
    for date in dates:
        feat=_features(rolling).iloc[-1][FEATURES]
        if feat.isna().any(): break
        tree_preds=np.array([est.predict(pd.DataFrame([feat],columns=FEATURES))[0] for est in model.estimators_],dtype=float)
        pred=float(np.median(tree_preds))
        # Bound recursive daily returns to the historical 1st/99th percentile to prevent
        # unstable compounding from an outlier tree/path.
        lo,hi=train["target"].quantile([.01,.99]).tolist()
        pred=float(np.clip(pred,lo,hi))
        price=float(rolling.iloc[-1]*(1+pred))
        rows.append({
            "date":pd.Timestamp(date).date().isoformat(),"predicted_price":price,
            "predicted_return":pred,
            "tree_p10_return":float(np.quantile(tree_preds,.10)),
            "tree_p90_return":float(np.quantile(tree_preds,.90)),
        })
        rolling.loc[pd.Timestamp(date)]=price
    meta={
        "status":"PASS" if len(rows)==horizon else "REVIEW",
        "training_rows":int(len(train)),"history_start":hist.index.min().date().isoformat(),
        "history_end":hist.index.max().date().isoformat(),"last_price":float(hist.iloc[-1]),
        "method":"RandomForestRegressor recursive 1-day returns; 5 business-day research path",
    }
    return rows,meta


def build_portfolio_rf_forecast(tickers:list[str],period:str="5y",horizon:int=HORIZON_DAYS) -> tuple[pd.DataFrame,dict]:
    tickers=list(dict.fromkeys(str(t).upper().strip() for t in tickers if str(t).strip()))
    if not tickers: return pd.DataFrame(),{"status":"NO_HOLDINGS"}
    prices=download_prices(tickers,period=period)
    all_rows=[]; meta={"status":"PASS","horizon_business_days":int(horizon),"models":{}}
    for ticker in tickers:
        if ticker not in prices.columns:
            meta["models"][ticker]={"status":"NO_PRICE_DATA"}; continue
        path,detail=_one_ticker_path(prices[ticker],horizon=horizon)
        meta["models"][ticker]=detail
        for row in path: all_rows.append({"Ticker":ticker,**row})
    if not all_rows: meta["status"]="INSUFFICIENT_DATA"
    return pd.DataFrame(all_rows),meta


def write_portfolio_rf_forecast(tickers:list[str],output_dir:Path,period:str="5y",horizon:int=HORIZON_DAYS):
    frame,meta=build_portfolio_rf_forecast(tickers,period=period,horizon=horizon)
    output_dir=Path(output_dir); output_dir.mkdir(parents=True,exist_ok=True)
    frame.to_csv(output_dir/"random_forest_next_week.csv",index=False)
    (output_dir/"random_forest_next_week.json").write_text(json.dumps(meta,indent=2),encoding="utf-8")
    return frame,meta
