from __future__ import annotations

"""Production runner that extends safe_update_model with commodity valuation v2.

safe_update_model already installs the core data-integrity and scoring patches. This runner
only changes the two commodity-sensitive hooks so every later WACC refresh is immediately
followed by the commodity normalization overlay before any downstream valuation consumer runs.
"""

import safe_update_model as safe
from commodity_valuation_v2 import apply_commodity_normalization

_ORIGINAL_DYNAMIC=safe.apply_dynamic_wacc


def _commodity_aware_dynamic_wacc(wb,ticker,info=None):
    result=_ORIGINAL_DYNAMIC(wb,ticker,info or {})
    try:
        apply_commodity_normalization(wb,ticker,info or {})
    except Exception as exc:
        print(f"Warning: commodity valuation v2 post-WACC overlay failed: {exc}")
    return result


# The safe wrapper's functions resolve these globals at execution time, so replacing them here
# fixes initial, post-statements and final WACC refresh ordering without duplicating the entire
# deterministic build pipeline.
safe.apply_dynamic_wacc=_commodity_aware_dynamic_wacc
safe.apply_commodity_normalization=apply_commodity_normalization


def main():
    safe.main()


if __name__=="__main__":
    main()
