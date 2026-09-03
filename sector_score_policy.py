from __future__ import annotations

"""Business-model safeguards for the reliability-aware score engine.

The underlying v3 score engine remains the corporate default.  This overlay removes dimensions
whose economics are inappropriate for banks, insurers, broker/dealers and REITs instead of
allowing an industrial FCF/net-debt DCF to influence an investment score.
"""

from business_model_registry import workbook_policy
import score_engine_v3 as base


def compute_score_bundle(wb, ticker=None, base_value=None, severe_value=None, current_price=None, mc_prob=None):
    bundle = base.compute_score_bundle(
        wb, ticker=ticker, base_value=base_value, severe_value=severe_value,
        current_price=current_price, mc_prob=mc_prob,
    )
    policy = workbook_policy(wb, ticker)
    dims = bundle.get("dimensions") or {}
    reasons = list((bundle.get("valuation_model_reliability") or {}).get("reasons") or [])

    if not policy.industrial_fcf_primary:
        if "FCF Quality" in dims:
            dims["FCF Quality"].update({
                "score": None,
                "status": f"Excluded — not primary for {policy.label}",
                "components": policy.primary_valuation,
                "formula": "Business-model gate: industrial FCF quality is not scored for this issuer type.",
            })

    if not policy.reverse_dcf_allowed:
        for key in ("Absolute Valuation", "Stress Robustness"):
            if key in dims:
                dims[key].update({
                    "score": None,
                    "status": f"Excluded — industrial DCF not primary for {policy.label}",
                    "components": f"Use {policy.primary_valuation} instead of a conventional enterprise-value DCF.",
                    "formula": "Business-model valuation gate.",
                })
        reasons.append(
            f"Business-model gate: conventional industrial DCF/reverse-DCF is not primary for {policy.label}. "
            f"Preferred framework: {policy.primary_valuation}."
        )
        bundle["valuation_model_reliability"] = {"status": "REVIEW", "reasons": list(dict.fromkeys(reasons))}

    # v3 leverage scoring uses corporate net debt / EBITDA / FCF.  That is not an appropriate
    # balance-sheet quality score for deposit-funded or insurance balance sheets.
    if policy.key in {"bank", "capital_markets", "insurance", "insurance_conglomerate"} and "Balance Sheet" in dims:
        dims["Balance Sheet"].update({
            "score": None,
            "status": f"Excluded — sector-specific capital analysis required ({policy.label})",
            "actual": "Industrial net-debt/EBITDA and net-debt/FCF are not used for this business model.",
            "benchmark": policy.primary_valuation,
            "formula": "Business-model gate: use regulatory capital, reserves/funding and book-value economics.",
        })

    bundle["business_model_policy"] = policy.to_dict()
    return base._recompute(bundle)


def advanced_scorecard(wb, current_price, forward_pe, base_value, severe_value):
    bundle = compute_score_bundle(
        wb, base_value=base_value, severe_value=severe_value, current_price=current_price
    )
    order = (
        "Growth", "Profitability", "FCF Quality", "Balance Sheet",
        "Absolute Valuation", "Relative Valuation", "Stress Robustness",
    )
    return [
        (
            key,
            bundle["dimensions"][key]["score"],
            bundle["dimensions"][key].get("formula", "") + " See Score Audit Trail for inputs and sources.",
        )
        for key in order
    ]
