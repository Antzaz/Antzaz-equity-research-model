from __future__ import annotations

"""Verify a hosted ML refresh actually has usable, recent market observations.

Run after ml_history.py daily-refresh. This catches the previously silent case where
the command exited 0 despite market-data providers returning no usable observations.
No symbols, position data, or private learning records are printed.
"""

import argparse
from datetime import date, datetime, timezone
from pathlib import Path
import sqlite3


def inspect(db: Path, max_age_days: int = 7, min_coverage: float = 0.60) -> dict:
    if not db.is_file() or db.stat().st_size == 0:
        raise RuntimeError("ML learning database is missing or empty")
    with sqlite3.connect(db) as con:
        required = {r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        if not {"symbols", "prices", "features"}.issubset(required):
            raise RuntimeError("ML database schema is incomplete")
        registered = con.execute("SELECT COUNT(DISTINCT symbol) FROM symbols").fetchone()[0]
        latest = con.execute("SELECT MAX(date) FROM prices WHERE adj_close > 0 OR close > 0").fetchone()[0]
        recent_count = con.execute(
            "SELECT COUNT(DISTINCT symbol) FROM prices "
            "WHERE date >= date('now', ?) AND (adj_close > 0 OR close > 0)",
            (f"-{int(max_age_days)} days",),
        ).fetchone()[0]
        feature_count = con.execute("SELECT COUNT(*) FROM features").fetchone()[0]
        last_fetched = con.execute("SELECT MAX(fetched_at) FROM prices").fetchone()[0]
    if not latest or not registered:
        raise RuntimeError("ML market price history has no usable observations")
    last_date = date.fromisoformat(str(latest)[:10])
    age_days = (datetime.now(timezone.utc).date() - last_date).days
    if age_days < 0 or age_days > max_age_days:
        raise RuntimeError(f"ML market history is stale: latest trading date is {age_days} day(s) old")
    coverage = recent_count / registered
    if recent_count < 10 or coverage < min_coverage:
        raise RuntimeError(
            f"ML symbol coverage too low: {recent_count}/{registered} recent symbols "
            f"({coverage:.1%}; required {min_coverage:.0%})"
        )
    if feature_count == 0:
        raise RuntimeError("ML feature table is empty after refresh")
    if not last_fetched:
        raise RuntimeError("ML price data has no fetch timestamp")
    try:
        fetched = datetime.fromisoformat(str(last_fetched).replace("Z", "+00:00"))
        if fetched.tzinfo is None:
            fetched = fetched.replace(tzinfo=timezone.utc)
        fetch_age = (datetime.now(timezone.utc) - fetched).total_seconds() / 3600
        if fetch_age < -1 or fetch_age > 6:
            raise RuntimeError(f"ML price fetch is stale ({fetch_age:.1f} hours)")
    except ValueError as exc:
        raise RuntimeError("ML price fetch timestamp is invalid") from exc
    return {
        "status": "PASS",
        "latest_market_date": last_date.isoformat(),
        "recent_symbol_count": recent_count,
        "registered_symbol_count": registered,
        "recent_symbol_coverage": round(coverage, 4),
        "feature_rows": feature_count,
        "last_price_fetch_utc": str(last_fetched),
    }


def main(argv=None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--db", type=Path, default=Path("ml_data/ml_history.sqlite"))
    parser.add_argument("--max-age-days", type=int, default=7)
    parser.add_argument("--min-coverage", type=float, default=0.60)
    args = parser.parse_args(argv)
    status = inspect(args.db, args.max_age_days, args.min_coverage)
    import json
    print(json.dumps(status, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
