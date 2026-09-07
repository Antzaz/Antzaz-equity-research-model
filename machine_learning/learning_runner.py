from __future__ import annotations

"""CLI entrypoint that runs the installed continual-learning extensions.

Use this instead of executing continual_learning.py directly so package-level integration
hooks (including 1M/3M/6M research horizons) are guaranteed to be installed first.
"""

import argparse
import json
from pathlib import Path

import machine_learning  # noqa: F401 - importing installs the learning extensions
from .history_store import HistoryStore
from . import continual_learning as cl


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Run governed continual learning across all configured horizons")
    parser.add_argument("--db", default=str(Path(__file__).resolve().parents[1] / "ml_data" / "ml_history.sqlite"))
    parser.add_argument("--benchmark", default=cl.DEFAULT_BENCHMARK)
    args = parser.parse_args(argv)

    store = HistoryStore(Path(args.db))
    summary = cl.run_continual_learning_cycle(store, args.benchmark)
    print(json.dumps(summary, indent=2, default=str))
    registry = cl.registry_frame(store)
    if not registry.empty:
        print(registry.to_string(index=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
