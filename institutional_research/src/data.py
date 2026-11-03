from __future__ import annotations

"""Public market data with a durable LOCAL cache for the private Windows dashboard.

Online portfolio research refreshes the cached adjusted-close history and company information.
Offline dashboards always read that saved cache; they never contact Yahoo. Cache contents are
private local runtime files under institutional_research/.offline_cache/ (gitignored).
"""

from datetime import datetime, timezone
import json
import os
from pathlib import Path
import tempfile
import time
import warnings
from typing import Dict, Iterable

import pandas as pd
import yfinance as yf


CACHE_DIR = Path(__file__).resolve().parents[1] / ".offline_cache"
PRICE_CACHE = "adjusted_closes.csv"
INFO_CACHE = "company_info.json"


def offline_mode() -> bool:
    return os.getenv("ANTZAZ_OFFLINE", "").strip().lower() in {"1", "true", "yes"}


def _symbols(tickers: Iterable[str]) -> list[str]:
    return list(dict.fromkeys(str(t).upper().strip() for t in tickers if str(t).strip()))


def _normalize_frame(frame: pd.DataFrame) -> pd.DataFrame:
    if frame is None or frame.empty:
        return pd.DataFrame()
    result = frame.copy()
    result.index = pd.to_datetime(result.index, errors="coerce", utc=True)
    result = result[~result.index.isna()].sort_index()
    result.index = result.index.tz_localize(None).normalize()
    result.columns = [str(c).upper().strip() for c in result.columns]
    result = result.apply(pd.to_numeric, errors="coerce")
    result = result.loc[:, ~result.columns.duplicated(keep="last")]
    result = result[~result.index.duplicated(keep="last")]
    return result.dropna(how="all")


def cached_prices() -> pd.DataFrame:
    path = CACHE_DIR / PRICE_CACHE
    if not path.is_file() or path.stat().st_size == 0:
        return pd.DataFrame()
    try:
        return _normalize_frame(pd.read_csv(path, index_col=0, parse_dates=True))
    except (OSError, ValueError, pd.errors.ParserError) as exc:
        warnings.warn(f"Offline price cache could not be read: {exc}", stacklevel=2)
        return pd.DataFrame()


def _atomic_write(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = None
    try:
        with tempfile.NamedTemporaryFile("w", encoding="utf-8", dir=path.parent, suffix=".tmp", delete=False) as handle:
            tmp = Path(handle.name)
            handle.write(content)
        os.replace(tmp, path)
    finally:
        if tmp is not None and tmp.exists():
            tmp.unlink(missing_ok=True)


def _save_prices(new_prices: pd.DataFrame) -> None:
    incoming = _normalize_frame(new_prices)
    if incoming.empty:
        return
    prior = cached_prices()
    merged = incoming.combine_first(prior) if not prior.empty else incoming
    _atomic_write(CACHE_DIR / PRICE_CACHE, merged.to_csv(index_label="Date"))
    metadata = {
        "last_online_refresh_utc": datetime.now(timezone.utc).isoformat(),
        "latest_observed_market_date": merged.index.max().date().isoformat(),
        "cached_symbols": len(merged.columns),
    }
    _atomic_write(CACHE_DIR / "cache_status.json", json.dumps(metadata, indent=2))


def price_cache_status() -> dict:
    path = CACHE_DIR / "cache_status.json"
    if path.is_file():
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
        except (OSError, ValueError):
            pass
    cache = cached_prices()
    return {
        "latest_observed_market_date": cache.index.max().date().isoformat() if not cache.empty else None,
        "cached_symbols": len(cache.columns),
        "last_online_refresh_utc": None,
    }


def _download_closes(tickers: list[str], period: str) -> pd.DataFrame:
    raw = yf.download(
        tickers=tickers,
        period=period,
        auto_adjust=True,
        progress=False,
        threads=True,
        group_by="column",
    )
    if raw is None or raw.empty:
        raise RuntimeError("Yahoo Finance returned no price history")
    if isinstance(raw.columns, pd.MultiIndex):
        if "Close" not in raw.columns.get_level_values(0):
            raise RuntimeError("Price download did not contain adjusted Close data")
        close = raw["Close"].copy()
    else:
        if "Close" not in raw:
            raise RuntimeError("Price download did not contain adjusted Close data")
        close = raw[["Close"]].copy()
        close.columns = tickers[:1]
    if isinstance(close, pd.Series):
        close = close.to_frame(name=tickers[0])
    normalized = _normalize_frame(close)
    if normalized.empty:
        raise RuntimeError("Yahoo Finance returned no usable adjusted closes")
    return normalized


def download_prices(tickers: Iterable[str], period: str = "5y") -> pd.DataFrame:
    requested = _symbols(tickers)
    if not requested:
        raise ValueError("No tickers supplied")
    saved = cached_prices()
    if offline_mode():
        missing = [t for t in requested if t not in saved or saved[t].dropna().empty]
        if missing:
            raise RuntimeError(
                f"Offline price cache is missing {len(missing)} required security(ies). "
                "Connect to the internet and run Sync All Data once before going offline."
            )
        return saved[requested].dropna(how="all")

    new_prices = pd.DataFrame()
    last_error = None
    for attempt in range(2):
        try:
            new_prices = _download_closes(requested, period)
            break
        except Exception as exc:
            last_error = exc
            if attempt == 0:
                time.sleep(1.5)
    if not new_prices.empty:
        # Yahoo sometimes returns a partially populated batch. Retry missing series
        # individually before falling back to older cached prices.
        incomplete = [
            ticker for ticker in requested
            if ticker not in new_prices or new_prices[ticker].dropna().empty
        ]
        for ticker in incomplete:
            try:
                retry = _download_closes([ticker], period)
                new_prices = retry.combine_first(new_prices)
            except Exception as exc:
                warnings.warn(
                    f"Individual market-data retry did not recover one missing security: {exc}",
                    stacklevel=2,
                )
        try:
            _save_prices(new_prices)
        except OSError as exc:
            warnings.warn(f"Could not write local offline price cache: {exc}", stacklevel=2)
    else:
        warnings.warn(f"Live quotes unavailable, attempting locally saved prices: {last_error}", stacklevel=2)

    merged = (
        new_prices.combine_first(saved)
        if not new_prices.empty and not saved.empty
        else new_prices if not new_prices.empty else saved
    )
    missing = [t for t in requested if t not in merged or merged[t].dropna().empty]
    if missing:
        raise RuntimeError(
            f"Missing usable price history for {len(missing)} security(ies). "
            "Retry the online sync; historical values will not be fabricated."
        )
    if new_prices.empty or any(t not in new_prices or new_prices[t].dropna().empty for t in requested):
        warnings.warn("Some requested price series came from a previously saved local cache.", stacklevel=2)
    return merged[requested].dropna(how="all")


def _read_cached_info() -> dict[str, dict]:
    path = CACHE_DIR / INFO_CACHE
    if path.is_file():
        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
            return raw if isinstance(raw, dict) else {}
        except (OSError, ValueError):
            return {}
    return {}


def fetch_info(tickers: Iterable[str], pause: float = 0.05) -> Dict[str, dict]:
    requested = _symbols(tickers)
    saved = _read_cached_info()
    if offline_mode():
        return {ticker: saved.get(ticker, {}) for ticker in requested}

    out: dict[str, dict] = {}
    for ticker in requested:
        try:
            data = yf.Ticker(ticker).info or {}
            if data and isinstance(data, dict):
                saved[ticker] = data
                out[ticker] = data
            else:
                out[ticker] = saved.get(ticker, {})
        except Exception:
            out[ticker] = saved.get(ticker, {})
        if pause:
            time.sleep(pause)
    try:
        _atomic_write(CACHE_DIR / INFO_CACHE, json.dumps(saved, ensure_ascii=False, default=str))
    except OSError as exc:
        warnings.warn(f"Could not write offline company-info cache: {exc}", stacklevel=2)
    return out
