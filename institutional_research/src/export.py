from __future__ import annotations

from datetime import datetime
from pathlib import Path
import json
import shutil
import pandas as pd


def _json_default(value):
    if hasattr(value, "item"):
        return value.item()
    raise TypeError(f"Object of type {type(value).__name__} is not JSON serializable")


def write_outputs(
    base_output_dir: str | Path,
    tables: dict[str, pd.DataFrame],
    summaries: dict,
):
    base = Path(base_output_dir)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    snapshot = base / "snapshots" / timestamp
    latest = base / "latest"

    snapshot.mkdir(parents=True, exist_ok=True)
    staging = base / ".latest_staging"
    backup = base / ".latest_previous"
    # Recover a prior interrupted replacement before attempting a new one.
    if not latest.exists() and backup.exists():
        backup.rename(latest)
    if staging.exists():
        shutil.rmtree(staging)
    staging.mkdir(parents=True, exist_ok=True)

    # Export a complete new version before touching the currently working offline
    # dashboard snapshot. A failed refresh must not destroy the last good data.
    try:
        for name, df in tables.items():
            if df is None:
                continue
            # A schema-less DataFrame would write an unreadable empty CSV.
            if isinstance(df, pd.DataFrame) and len(df.columns) == 0:
                continue
            for target in [snapshot, staging]:
                df.to_csv(target / f"{name}.csv", index=False)

        for target in [snapshot, staging]:
            with open(target / "summary.json", "w", encoding="utf-8") as f:
                json.dump(summaries, f, indent=2, default=_json_default)

        if backup.exists():
            shutil.rmtree(backup)
        had_previous = latest.exists()
        if had_previous:
            latest.rename(backup)
        try:
            staging.rename(latest)
        except Exception:
            if had_previous and backup.exists() and not latest.exists():
                backup.rename(latest)
            raise
        if backup.exists():
            shutil.rmtree(backup)
    finally:
        if staging.exists():
            shutil.rmtree(staging)

    return snapshot, latest
