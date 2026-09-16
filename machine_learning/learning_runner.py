from __future__ import annotations

"""CLI entrypoint for the governed continual-learning research stack.

Use this instead of executing continual_learning.py directly so package-level integration
hooks are guaranteed to be installed first. The runner now covers:

- existing 12M expected-return and earnings-surprise feedback;
- research-only 1D/1W/1M/3M/6M return horizons;
- forward volatility and drawdown context models;
- live forecast maturation and champion/challenger governance;
- the bounded weekly autonomous challenger-research loop.
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
