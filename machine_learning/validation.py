from __future__ import annotations

from dataclasses import dataclass
import numpy as np
import pandas as pd
from sklearn.base import clone
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score


@dataclass
class WalkForwardResult:
    predictions: pd.DataFrame
    metrics: dict[str, float | int | None]


def regression_metrics(y_true, y_pred) -> dict[str, float | int | None]:
    y = np.asarray(y_true, dtype=float)
    p = np.asarray(y_pred, dtype=float)
    mask = np.isfinite(y) & np.isfinite(p)
    y, p = y[mask], p[mask]
    if len(y) == 0:
        return {"n": 0, "mae": None, "rmse": None, "r2": None, "directional_accuracy": None}
    return {
        "n": int(len(y)),
        "mae": float(mean_absolute_error(y, p)),
        "rmse": float(mean_squared_error(y, p) ** 0.5),
        "r2": float(r2_score(y, p)) if len(y) >= 2 else None,
        "directional_accuracy": float(np.mean(np.sign(y) == np.sign(p))),
    }


def expanding_walk_forward(
    estimator,
    frame: pd.DataFrame,
    feature_cols: list[str],
    target_col: str,
    date_col: str = "as_of",
    target_date_col: str = "target_date",
    min_train: int = 24,
    step: int = 1,
) -> WalkForwardResult:
    """Purged expanding walk-forward validation.

    For forward targets such as next-12-month excess return, a row is trainable only after its
    target horizon has fully elapsed. This prevents a subtle look-ahead leak where a historical
    feature row is chronologically earlier than the test row but its realized 12-month return
    still contains prices from after the test decision date.

    Rows sharing the same as-of date are evaluated together so one company at a date cannot train
    on another company's outcome from that same decision date.
    """
    df = frame.copy()
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    has_target_date = target_date_col in df.columns
    if has_target_date:
        df[target_date_col] = pd.to_datetime(df[target_date_col], errors="coerce")
    df = df.dropna(subset=[date_col, target_col]).sort_values([date_col]).reset_index(drop=True)
    if df.empty:
        return WalkForwardResult(pd.DataFrame(), regression_metrics([], []))

    # Preserve the old `step` concept as a validation-sampling interval, while using complete
    # date groups to eliminate same-date cross-sectional leakage.
    candidate_starts = list(range(max(0, min_train), len(df), max(1, int(step))))
    test_dates=[]
    seen=set()
    for start in candidate_starts:
        dt=df.iloc[start][date_col]
        if pd.isna(dt) or dt in seen:
            continue
        seen.add(dt); test_dates.append(dt)

    rows = []
    for test_date in test_dates:
        test = df[df[date_col] == test_date]
        if has_target_date:
            train = df[(df[date_col] < test_date) & df[target_date_col].notna() & (df[target_date_col] <= test_date)]
        else:
            train = df[df[date_col] < test_date]
        if len(train) < min_train or test.empty:
            continue
        model = clone(estimator)
        model.fit(train[feature_cols], train[target_col])
        pred = model.predict(test[feature_cols])
        for i, (_, row) in enumerate(test.iterrows()):
            rows.append({
                "as_of": row[date_col],
                "actual": float(row[target_col]),
                "prediction": float(pred[i]),
                "train_rows": int(len(train)),
            })
    out = pd.DataFrame(rows)
    metrics = regression_metrics(out["actual"], out["prediction"]) if not out.empty else regression_metrics([], [])
    if not out.empty:
        metrics["min_train_rows"] = int(out["train_rows"].min())
        metrics["max_train_rows"] = int(out["train_rows"].max())
        metrics["purged_target_horizons"] = bool(has_target_date)
    return WalkForwardResult(out, metrics)


def assert_no_future_leakage(frame: pd.DataFrame, feature_date_col: str = "as_of", target_date_col: str = "target_date") -> None:
    if feature_date_col not in frame or target_date_col not in frame:
        return
    a = pd.to_datetime(frame[feature_date_col], errors="coerce")
    t = pd.to_datetime(frame[target_date_col], errors="coerce")
    bad = t.notna() & a.notna() & (t <= a)
    if bad.any():
        raise ValueError(f"Leakage guard failed: {int(bad.sum())} row(s) have target_date <= as_of")
