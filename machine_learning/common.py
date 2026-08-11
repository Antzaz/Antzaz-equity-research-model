from __future__ import annotations

from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any
import json
import math

import numpy as np
import pandas as pd

RANDOM_STATE = 42


def num(value: Any, default: float | None = None) -> float | None:
    try:
        if isinstance(value, bool) or value is None or value == "":
            return default
        x = float(value)
        return x if math.isfinite(x) else default
    except Exception:
        return default


def annualized_return(series: pd.Series, periods: int = 252) -> float | None:
    s = pd.Series(series).dropna()
    if len(s) < 2:
        return None
    total = float((1.0 + s).prod())
    if total <= 0:
        return None
    return total ** (periods / len(s)) - 1.0


def annualized_vol(series: pd.Series, periods: int = 252) -> float | None:
    s = pd.Series(series).dropna()
    if len(s) < 2:
        return None
    return float(s.std(ddof=1) * np.sqrt(periods))


def max_drawdown_from_prices(prices: pd.Series) -> float | None:
    p = pd.Series(prices).dropna()
    if len(p) < 2:
        return None
    running = p.cummax()
    dd = p / running - 1.0
    return float(dd.min())


def json_ready(value: Any) -> Any:
    if isinstance(value, dict):
        return {str(k): json_ready(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [json_ready(v) for v in value]
    if isinstance(value, pd.Timestamp):
        return value.isoformat()
    if isinstance(value, np.generic):
        return value.item()
    if isinstance(value, np.ndarray):
        return value.tolist()
    if isinstance(value, float) and not math.isfinite(value):
        return None
    return value


@dataclass
class MLResult:
    name: str
    status: str
    summary: str
    prediction: float | str | None = None
    confidence: str | None = None
    metrics: dict[str, Any] | None = None
    drivers: list[dict[str, Any]] | None = None
    details: dict[str, Any] | None = None

    def to_dict(self) -> dict[str, Any]:
        return json_ready(asdict(self))


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(json_ready(payload), indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
