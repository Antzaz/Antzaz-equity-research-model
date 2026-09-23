from __future__ import annotations

from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_offline_decision_artifacts_are_not_exported_to_public_showcase():
    builder=(ROOT/"automation"/"build_showcase_snapshot.py").read_text(encoding="utf-8")
    fundamentals=(ROOT/"automation"/"build_public_fundamentals.py").read_text(encoding="utf-8")
    exporter=(ROOT/"automation"/"export_showcase.ps1").read_text(encoding="utf-8")

    private_markers=[
        "transaction_ledger.csv",
        "benchmark_sector_history.csv",
        "fundamental_scenarios.csv",
        "portfolio_decision_journal.csv",
        "Forecast Accountability",
        "Earnings & Revisions",
        "Capital Allocation",
        "Valuation History",
        "SOTP Framework",
        "Thesis Timeline",
        "forecast_history.json",
        "forecast_accuracy_summary.csv",
    ]

    # Public payload builders are intentionally whitelist-based and must not read the new
    # offline/private decision-accountability artifacts.
    for marker in private_markers:
        assert marker not in builder, f"public portfolio snapshot references private artifact: {marker}"
        assert marker not in fundamentals, f"public fundamentals references private artifact: {marker}"

    # The export script copies only the already-sanitized showcase bundle, not private inputs.
    for marker in private_markers:
        assert marker not in exporter, f"public showcase export references private artifact: {marker}"


def test_offline_private_inputs_are_gitignored():
    ignore=(ROOT/"institutional_research"/".gitignore").read_text(encoding="utf-8").splitlines()
    required={
        "transaction_ledger.csv",
        "benchmark_sector_history.csv",
        "fundamental_scenarios.csv",
        "portfolio_decision_journal.csv",
    }
    assert required.issubset(set(ignore))
