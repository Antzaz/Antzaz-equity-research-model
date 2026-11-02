from __future__ import annotations

"""Validate an unattended research-agent batch and preserve its latest status.

The production research.py runner may return success with WARN agent results.
Warnings are allowed, but missing manifests, missing expected agents, FAIL statuses,
and stale outputs must never be reported as a successful monitoring cycle.
This checker deliberately avoids printing private symbols or holdings.
"""

import argparse
from datetime import datetime, timezone
import json
from pathlib import Path
import shutil


EXPECTED_AGENTS = {
    "Source Health Agent",
    "Filings & Financials Agent",
    "KPI / Earnings Agent",
    "Thesis Monitor Agent",
    "Research QA / Skeptic Agent",
}


def verify(root: Path, tickers: list[str], max_age_hours: float = 4.0) -> dict:
    if not tickers:
        raise RuntimeError("Agent monitor did not select any companies")
    now = datetime.now(timezone.utc)
    passes = warns = 0
    for ticker in tickers:
        candidates = list((root / ticker).glob("*/manifest.json"))
        if not candidates:
            raise RuntimeError("An agent research report is missing for a selected company")
        path = max(candidates, key=lambda p: p.stat().st_mtime)
        age_hours = (now - datetime.fromtimestamp(path.stat().st_mtime, timezone.utc)).total_seconds() / 3600
        if age_hours < -0.1 or age_hours > max_age_hours:
            raise RuntimeError("An agent report was not refreshed during the current monitoring cycle")
        data = json.loads(path.read_text(encoding="utf-8"))
        if data.get("ticker") != ticker:
            raise RuntimeError("Agent report does not match the selected company")
        results = data.get("results")
        if not isinstance(results, list):
            raise RuntimeError("Agent report is missing structured results")
        names = {x.get("name") for x in results if isinstance(x, dict)}
        if not EXPECTED_AGENTS.issubset(names):
            raise RuntimeError("One or more expected research agents were not executed")
        for result in results:
            if not isinstance(result, dict):
                raise RuntimeError("Invalid agent result")
            status = result.get("status")
            if status not in {"PASS", "WARN"}:
                raise RuntimeError("An agent failed or returned an unrecognized status")
            passes += status == "PASS"
            warns += status == "WARN"

    return {
        "status": "PASS",
        "companies_checked": len(tickers),
        "expected_agent_runs": len(tickers) * len(EXPECTED_AGENTS),
        "passing_agent_runs": passes,
        "warning_agent_runs": warns,
        "generated_utc": now.isoformat(),
    }


def main(argv=None) -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--run-root", type=Path, default=Path("research_runs"))
    p.add_argument("--tickers-file", type=Path, required=True)
    p.add_argument("--state-dir", type=Path, default=Path("research_data/agent_monitor"))
    p.add_argument("--max-age-hours", type=float, default=4.0)
    args = p.parse_args(argv)
    tickers = [x.strip() for x in args.tickers_file.read_text(encoding="utf-8").splitlines() if x.strip()]
    summary = verify(args.run_root, tickers, args.max_age_hours)
    args.state_dir.mkdir(parents=True, exist_ok=True)
    # The source directory is encrypted before GitHub Actions uploads it. Store the
    # detailed per-company manifests there, never as a public commit or an unencrypted artifact.
    for ticker in tickers:
        latest = max((args.run_root / ticker).glob("*/manifest.json"), key=lambda x: x.stat().st_mtime)
        target = args.state_dir / ticker / "latest_manifest.json"
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(latest, target)
    (args.state_dir / "latest_summary.json").write_text(json.dumps(summary, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(summary, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
