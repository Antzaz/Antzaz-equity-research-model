from __future__ import annotations

"""Production runner extending safe_update_model with commodity valuation v3.

The guarded core model remains unchanged for ordinary corporate issuers.  Before the build starts,
the cross-sector runtime installs business-model routing so banks, insurers, REITs, commodity
producers and other sectors either use an appropriate framework or explicitly gate unsuitable
industrial FCF/DCF outputs.

Commodity producers receive:
- v2 mid-cycle operating normalization,
- v3 independent equity-FCF cross-check,
- triangulated primary fair value,
- commodity-aware Monte Carlo,
- explicit Decision View / Data Quality disclosure.

The runner also installs narrow verified issuer adapters when generic parsers are known to be
brittle. Alphabet's segment adapter is one such case: it uses only company-reported 10-K values.
Every issuer also receives a source-backed Deals & Transactions sheet after the guarded research
extensions finish, so the deal monitor cannot be removed by the low-value-tab pruning pass.
"""

import advanced_analytics_v2
import decision_view_v2
import safe_update_model as safe

from cross_sector_runtime import install_cross_sector_runtime
from commodity_valuation_v3 import (
    apply_commodity_normalization,
    commodity_base_value,
    commodity_monte_carlo,
    decorate_decision_and_quality,
)
from deal_analysis import ensure_deal_analysis
from google_segment_analysis import ensure_google_segment_analysis


# Install business-model routing after safe_update_model has installed its guarded wrappers, but
# before we capture/extend those wrappers below.  update_model.main resolves these module globals
# at execution time, so the production research.py path receives the cross-sector behavior.
install_cross_sector_runtime(safe_module=safe, update_model_module=safe.update_model)

_ORIGINAL_DYNAMIC = safe.apply_dynamic_wacc
_ORIGINAL_DECISION = safe.ensure_decision_view
_ORIGINAL_VERIFIED_SEGMENT = safe._verified_segment
_ORIGINAL_RESEARCH_EXTENSIONS = safe.update_model.ensure_research_extensions


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


def _research_extensions_with_deals(wb, ticker, info=None):
    """Run all guarded extensions, then append the transaction monitor as a final research tab."""
    result = _ORIGINAL_RESEARCH_EXTENSIONS(wb, ticker, info)
    try:
        ensure_deal_analysis(wb, ticker, info or {})
    except Exception as exc:
        # A temporary news / EDGAR outage must never prevent the deterministic valuation model
        # from being produced.  The deal sheet itself also degrades to an explicit no-data state.
        print(f"Warning: Deals & Transactions refresh failed: {exc}")
    return result


# Patch valuation consumers before the build starts. Their functions resolve module globals at
# execution time, so the unshocked commodity base case and Monte Carlo use v3 while explicit
# stress shocks continue to use the normalized operating FCFF engine.
advanced_analytics_v2._base_value = commodity_base_value
advanced_analytics_v2._monte_carlo = commodity_monte_carlo
decision_view_v2._base_value = commodity_base_value

# safe_update_model resolves these globals at runtime for segment, initial/post-statements WACC
# and final Decision View refreshes. The Alphabet adapter is therefore used by both the main
# Segment Analysis pass and the later enrichment pass without changing other issuers.
safe._verified_segment = _verified_segment_with_alphabet
safe.apply_dynamic_wacc = _commodity_aware_dynamic_wacc
safe.apply_commodity_normalization = apply_commodity_normalization
safe.ensure_decision_view = _commodity_aware_decision
safe.update_model.ensure_research_extensions = _research_extensions_with_deals


def main():
    safe.main()


if __name__ == "__main__":
    main()
