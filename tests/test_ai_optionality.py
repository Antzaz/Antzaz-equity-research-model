from __future__ import annotations

import pandas as pd

from institutional_research.src.ai_optionality import (
    DEFAULT_SCENARIOS,
    apply_ai_overlay,
    build_overlay_table,
    normalize_scenarios,
    portfolio_ai_scenarios,
)


def test_scenario_probabilities_are_normalized():
    raw = pd.DataFrame([
        {"Scenario": "A", "Probability": 2.0, "IncrementalReturn": -0.10},
        {"Scenario": "B", "Probability": 3.0, "IncrementalReturn": 0.20},
    ])
    out = normalize_scenarios(raw)
    assert abs(out["Probability"].sum() - 1.0) < 1e-12
    assert abs(out.loc[0, "Probability"] - 0.4) < 1e-12


def test_zero_ai_exposure_leaves_existing_forecast_unchanged():
    result = apply_ai_overlay(
        0.12,
        0.0,
        DEFAULT_SCENARIOS,
        evidence_confidence=1.0,
        maturity_confidence=1.0,
        ai_skill_confidence=1.0,
    )
    assert result.adjusted_return == result.core_return == 0.12
    assert result.certainty_equivalent_ai == 0.0


def test_uncertainty_penalty_reduces_ai_overlay():
    no_penalty = apply_ai_overlay(
        0.10,
        1.0,
        DEFAULT_SCENARIOS,
        evidence_confidence=1.0,
        maturity_confidence=1.0,
        ai_skill_confidence=1.0,
        uncertainty_aversion=0.0,
    )
    conservative = apply_ai_overlay(
        0.10,
        1.0,
        DEFAULT_SCENARIOS,
        evidence_confidence=1.0,
        maturity_confidence=1.0,
        ai_skill_confidence=1.0,
        uncertainty_aversion=0.8,
    )
    assert conservative.adjusted_return < no_penalty.adjusted_return
    assert conservative.uncertainty_penalty > 0.0


def test_low_ai_specific_skill_limits_overlay_without_touching_core():
    low_skill = apply_ai_overlay(
        0.11,
        1.0,
        DEFAULT_SCENARIOS,
        evidence_confidence=1.0,
        maturity_confidence=1.0,
        ai_skill_confidence=0.10,
        uncertainty_aversion=0.0,
    )
    full_skill = apply_ai_overlay(
        0.11,
        1.0,
        DEFAULT_SCENARIOS,
        evidence_confidence=1.0,
        maturity_confidence=1.0,
        ai_skill_confidence=1.0,
        uncertainty_aversion=0.0,
    )
    assert low_skill.core_return == full_skill.core_return == 0.11
    assert abs(low_skill.certainty_equivalent_ai) < abs(full_skill.certainty_equivalent_ai)


def test_build_overlay_table_defaults_to_no_change_without_exposure():
    expected = pd.DataFrame({
        "Ticker": ["AAA", "BBB"],
        "BlendedMLExpectedReturn": [0.08, 0.14],
    })
    cfg = pd.DataFrame({"Ticker": ["AAA", "BBB"]})
    out = build_overlay_table(
        expected,
        cfg,
        core_column="BlendedMLExpectedReturn",
        scenarios=DEFAULT_SCENARIOS,
    )
    assert out["AIExposure"].eq(0.0).all()
    assert (out["CoreExpectedReturn"] == out["AIOverlayExpectedReturn"]).all()


def test_portfolio_ai_factor_exposure_is_weighted_and_does_not_change_weights():
    weights = pd.DataFrame({"Ticker": ["AAA", "BBB"], "Weight": [0.60, 0.40]})
    cfg = pd.DataFrame({"Ticker": ["AAA", "BBB"], "AIExposure": [1.0, 0.25]})
    out = portfolio_ai_scenarios(weights, cfg, DEFAULT_SCENARIOS)
    expected_factor = 0.60 * 1.0 + 0.40 * 0.25
    assert abs(out["PortfolioAIFactorExposure"].iloc[0] - expected_factor) < 1e-12
    assert abs(weights["Weight"].sum() - 1.0) < 1e-12
