from __future__ import annotations

from pathlib import Path
import numpy as np
import pandas as pd


def analyze_forecasts(path: str | Path) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        return pd.DataFrame()

    df = pd.read_csv(path, comment="#")
    if df.empty:
        return df

    for col in ["Forecast", "Actual"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["Error"] = df["Actual"] - df["Forecast"]
    df["ErrorPct"] = np.where(
        df["Forecast"].abs() > 1e-12,
        df["Error"] / df["Forecast"].abs(),
        np.nan,
    )
    df["AbsoluteErrorPct"] = df["ErrorPct"].abs()
    df["BiasDirection"] = np.where(
        df["Error"].isna(),
        "",
        np.where(df["Error"] > 0, "Under-forecast", np.where(df["Error"] < 0, "Over-forecast", "Exact")),
    )
    return df
