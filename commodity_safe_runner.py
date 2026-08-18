from __future__ import annotations

"""Production runner extending safe_update_model with commodity valuation v3.

The guarded core model remains unchanged for non-commodity issuers. Commodity producers receive:
- v2 mid-cycle operating normalization,
- v3 independent equity-FCF cross-check,
- triangulated primary fair value,
- commodity-aware Monte Carlo,
- explicit Decision View / Data Quality disclosure.
"""

import advanced_analytics_v2
import decision_view_v2
import safe_update_model as safe

from commodity_valuation_v3 import (
    apply_commodity_normalization,
    commodity_base_value,
    commodity_monte_carlo,
    decorate_decision_and_quality,
)

_ORIGINAL_DYNAMIC = safe.apply_dynamic_wacc
_ORIGINAL_DECISION = safe.ensure_decision_view


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


# Patch valuation consumers before the build starts. Their functions resolve module globals at
# execution time, so the unshocked commodity base case and Monte Carlo use v3 while explicit
# stress shocks continue to use the normalized operating FCFF engine.
advanced_analytics_v2._base_value = commodity_base_value
advanced_analytics_v2._monte_carlo = commodity_monte_carlo
decision_view_v2._base_value = commodity_base_value

# safe_update_model resolves these globals at runtime for initial, post-statements and final WACC
# refreshes, so the commodity overlay cannot be lost later in the build.
safe.apply_dynamic_wacc = _commodity_aware_dynamic_wacc
safe.apply_commodity_normalization = apply_commodity_normalization
safe.ensure_decision_view = _commodity_aware_decision


def main():
    safe.main()


if __name__ == "__main__":
    main()
