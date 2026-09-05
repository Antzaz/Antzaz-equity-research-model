from __future__ import annotations

"""Production runner extending safe_update_model with cross-sector integrity controls.

The guarded core model remains stable while this entry point installs business-model routing,
canonical accounting guards, conservative public-data recovery, scenario-assumption sanity checks,
commodity normalization, public AI-evidence enrichment, verified segment adapters, and a final
decision-first workbook layout.
"""

import advanced_analytics_v2
import decision_view_v2
import safe_update_model as safe

from ai_public_evidence import seed_public_ai_evidence
from canonical_statement_guard import apply_canonical_statement_guard, decorate_data_quality
from cross_sector_runtime import install_cross_sector_runtime
from commodity_valuation_v3 import (
    apply_commodity_normalization,
    commodity_base_value,
    commodity_monte_carlo,
    decorate_decision_and_quality,
)
from deal_analysis import ensure_deal_analysis
from google_segment_analysis import ensure_google_segment_analysis
from public_data_orchestrator import run_public_data_recovery
from scenario_integrity import repair_default_scenario_margin_paths
from visualization_v2 import ensure_visual_dashboard
from workbook_presentation import apply_workbook_presentation


# Install business-model routing after safe_update_model has installed its guarded wrappers, but
# before we capture/extend those wrappers below. update_model.main resolves these module globals
# at execution time, so the production research.py path receives the cross-sector behavior.
install_cross_sector_runtime(safe_module=safe, update_model_module=safe.update_model)

_ORIGINAL_DYNAMIC = safe.apply_dynamic_wacc
_ORIGINAL_DECISION = safe.ensure_decision_view
_ORIGINAL_VERIFIED_SEGMENT = safe._verified_segment
_ORIGINAL_RESEARCH_EXTENSIONS = safe.update_model.ensure_research_extensions
_ORIGINAL_FINANCIAL_STATEMENTS = safe.update_model.ensure_financial_statements
_ORIGINAL_SCENARIOS = safe.update_model.update_scenarios


def _commodity_aware_dynamic_wacc(wb, ticker, info=None):
    result = _ORIGINAL_DYNAMIC(wb, ticker, info or {})
    try:
        apply_commodity_normalization(wb, ticker, info or {})
    except Exception as exc:
        print(f"Warning: commodity valuation v3 post-WACC overlay failed: {exc}")
    return result


def _commodity_aware_decision(wb, ticker):
    result = _ORIGINAL_DECISION(wb, ticker)
    try:
        decorate_decision_and_quality(wb, ticker)
    except Exception as exc:
        print(f"Warning: commodity Decision View / quality decoration failed: {exc}")
    return result


def _verified_segment_with_alphabet(wb, ticker):
    if str(ticker or "").upper().strip() in {"GOOG", "GOOGL"}:
        return ensure_google_segment_analysis(wb, ticker)
    return _ORIGINAL_VERIFIED_SEGMENT(wb, ticker)


def _scenarios_with_integrity(wb, hist, info):
    """Preserve reported margins when legacy template bounds would silently replace them."""
    result = _ORIGINAL_SCENARIOS(wb, hist, info)
    ticker = ""
    try:
        ticker = str(wb["Company Data"]["B4"].value or "").upper().strip()
        guard = repair_default_scenario_margin_paths(wb, ticker, hist, info or {})
        if guard.get("changed"):
            setattr(wb, "_scenario_margin_guard", guard)
            print(f"Scenario margin integrity: {guard.get('reason')}")
    except Exception as exc:
        print(f"Warning: scenario margin integrity guard failed: {exc}")
    return result


def _financial_statements_with_canonical_guard(wb, ticker, facts):
    """Run existing statement repairs, exact SEC guards, then exhaust safe public fallbacks."""
    result = _ORIGINAL_FINANCIAL_STATEMENTS(wb, ticker, facts)
    try:
        guard = apply_canonical_statement_guard(wb, ticker, facts)
        setattr(wb, "_canonical_statement_guard", guard)
        corrections = guard.get("material_corrections") or []
        if guard.get("exact_cells_written"):
            print(
                f"Canonical statement guard: exact={guard['exact_cells_written']} cells; "
                f"history_sync={guard.get('history_sync', 0)}; material corrections={len(corrections)}"
            )
    except Exception as exc:
        print(f"Warning: SEC-first canonical statement guard failed: {exc}")

    # Only blanks are eligible here. Exact SEC/issuer values written above are protected.
    try:
        info = getattr(wb, "_wacc_info", {}) or {}
        recovery = run_public_data_recovery(wb, ticker, info)
        setattr(wb, "_public_data_recovery", recovery)
        if recovery.get("provider_cells_filled") or recovery.get("derived_cells_filled") or recovery.get("profile_rows_filled"):
            print(
                "Public-data recovery: "
                f"provider={recovery.get('provider_cells_filled', 0)} cells; "
                f"derived={recovery.get('derived_cells_filled', 0)}; "
                f"profile rows={recovery.get('profile_rows_filled', 0)}; "
                f"history sync={recovery.get('history_sync', 0)}"
            )
    except Exception as exc:
        print(f"Warning: conservative public-data recovery failed: {exc}")
    return result


def _research_extensions_with_deals(wb, ticker, info=None):
    """Run final extensions, rebuild the dashboard from final data, then consolidate presentation."""
    result = _ORIGINAL_RESEARCH_EXTENSIONS(wb, ticker, info)

    # research_extensions deliberately removes some legacy presentation tabs. Rebuild the modern
    # Visual Dashboard *after* the final financial/segment/valuation state so it cannot be stale.
    try:
        ensure_visual_dashboard(wb, ticker)
    except Exception as exc:
        print(f"Warning: final Visual Dashboard rebuild failed: {exc}")

    try:
        ensure_deal_analysis(wb, ticker, info or {})
    except Exception as exc:
        # A temporary news / EDGAR outage must never prevent the deterministic valuation model
        # from being produced. The deal sheet itself also degrades to an explicit no-data state.
        print(f"Warning: Deals & Transactions refresh failed: {exc}")
    try:
        decorate_data_quality(wb, ticker, getattr(wb, "_canonical_statement_guard", {}))
    except Exception as exc:
        print(f"Warning: canonical accounting Data Quality decoration failed: {exc}")
    scenario = getattr(wb, "_scenario_margin_guard", {}) or {}
    if scenario.get("changed") and "Data Quality" in wb.sheetnames:
        try:
            ws = wb["Data Quality"]
            label = "Scenario operating-margin integrity"
            row = next((r for r in range(1, ws.max_row + 1) if str(ws.cell(r, 1).value or "").strip() == label), ws.max_row + 1)
            ws.cell(row, 1, label); ws.cell(row, 2, "PASS")
            ws.cell(row, 3, scenario.get("reason")); ws.cell(row, 4, "Prevents arbitrary template margin caps from replacing observed economics.")
        except Exception:
            pass

    try:
        presentation = apply_workbook_presentation(wb, ticker)
        setattr(wb, "_workbook_presentation", presentation)
        print(
            f"Workbook presentation: visible={presentation.get('visible_sheets')}; "
            f"hidden={', '.join(presentation.get('hidden_tabs') or []) or 'none'}"
        )
    except Exception as exc:
        print(f"Warning: workbook presentation consolidation failed: {exc}")
    return result


# Patch valuation consumers before the build starts. Their functions resolve module globals at
# execution time, so the unshocked commodity base case and Monte Carlo use v3 while explicit
# stress shocks continue to use the normalized operating FCFF engine.
advanced_analytics_v2._base_value = commodity_base_value
advanced_analytics_v2._monte_carlo = commodity_monte_carlo
decision_view_v2._base_value = commodity_base_value

# safe_update_model resolves these globals at runtime for segment, scenarios, statements, WACC and
# final Decision View refreshes. These wrappers therefore apply generically to every searched issuer.
safe._verified_segment = _verified_segment_with_alphabet
safe.apply_dynamic_wacc = _commodity_aware_dynamic_wacc
safe.apply_commodity_normalization = apply_commodity_normalization
safe.ensure_decision_view = _commodity_aware_decision
safe.update_model.update_scenarios = _scenarios_with_integrity
safe.update_model.ensure_financial_statements = _financial_statements_with_canonical_guard
safe.update_model.ensure_research_extensions = _research_extensions_with_deals


def main():
    # Seed uncoded companies with conservative qualitative public AI evidence before any workbook
    # module builds the AI Impact Analysis sheet. Curated issuer packs remain authoritative.
    try:
        ticker = safe.update_model.get_ticker()
        # update_model.main() will call get_ticker() again; normalize argv so both calls agree.
        import sys
        if len(sys.argv) > 1:
            sys.argv[1] = ticker
        seeded = seed_public_ai_evidence(ticker)
        if seeded.get("seeded"):
            print(f"Public AI evidence: seeded {seeded['seeded']} qualitative row(s) for {ticker}")
    except Exception as exc:
        print(f"Warning: public AI evidence seeding failed: {exc}")
    safe.main()


if __name__ == "__main__":
    main()
