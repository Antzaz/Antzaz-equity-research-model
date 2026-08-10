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
    if latest.exists():
        shutil.rmtree(latest)
    latest.mkdir(parents=True, exist_ok=True)

    for name, df in tables.items():
        if df is None:
            continue
        # A completely schema-less DataFrame writes a zero-byte/blank CSV that
        # pandas cannot read back (EmptyDataError). Missing optional analyses
        # are represented by an absent file instead; DataFrames with defined
        # columns but zero rows still export their headers normally.
        if isinstance(df, pd.DataFrame) and len(df.columns) == 0:
            continue
        for target in [snapshot, latest]:
            df.to_csv(target / f"{name}.csv", index=False)

    for target in [snapshot, latest]:
        with open(target / "summary.json", "w", encoding="utf-8") as f:
            json.dump(summaries, f, indent=2, default=_json_default)

    return snapshot, latest
