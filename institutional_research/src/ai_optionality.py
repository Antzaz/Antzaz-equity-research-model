from __future__ import annotations

"""Standalone AI optionality / uncertainty overlay for portfolio research.

This module is deliberately independent from the project's existing machine-learning
models and portfolio optimizer.  It consumes an already-produced expected return and
shows how an analyst-defined AI scenario distribution could change that return after
explicit evidence, maturity, skill and uncertainty haircuts.

Nothing in this module retrains ML, writes predictions, or changes optimizer inputs.
"""

from dataclasses import dataclass
from typing import Iterable

import numpy as np
import pandas as pd


DEFAULT_SCENARIOS = pd.DataFrame(
    [
        {"Scenario": "AI disappointment", "Probability": 0.15, "IncrementalReturn": -0.08},
        {"Scenario": "Base adoption", "Probability": 0.45, "IncrementalReturn": 0.02},
        {"Scenario": "Strong adoption", "Probability": 0.30, "IncrementalReturn": 0.10},
        {"Scenario": "AI supercycle", "Probability": 0.10, "IncrementalReturn": 0.22},
    ]
)


@dataclass(frozen=True)
class AIOverlayResult:
    core_return: float
    exposure: float
    scenario_mean: float
    scenario_sigma: float
    evidence_confidence: float
    maturity_confidence: float
    ai_skill_confidence: float
    credibility: float
    uncertainty_aversion: float
    raw_ai_contribution: float
    uncertainty_penalty: float
    certainty_equivalent_ai: float
    adjusted_return: float

    def to_dict(self) -> dict[str, float]:
        return {
            "CoreExpectedReturn": self.core_return,
            "AIExposure": self.exposure,
            "AIScenarioMean": self.scenario_mean,
            "AIScenarioSigma": self.scenario_sigma,
            "EvidenceConfidence": self.evidence_confidence,
            "MaturityConfidence": self.maturity_confidence,
            "AISkillConfidence": self.ai_skill_confidence,
            "AICredibility": self.credibility,
            "RawAIContribution": self.raw_ai_contribution,
            "AIUncertaintyPenalty": self.uncertainty_penalty,
            "CertaintyEquivalentAI": self.certainty_equivalent_ai,
            "AIOverlayExpectedReturn": self.adjusted_return,
        }


def _clip01(value: float) -> float:
    try:
        x = float(value)
    except Exception:
        return 0.0
    if not np.isfinite(x):
        return 0.0
    return float(np.clip(x, 0.0, 1.0))


def normalize_scenarios(frame: pd.DataFrame | None = None) -> pd.DataFrame:
    """Validate and normalize a scenario table.

    Probabilities must be non-negative and contain positive total mass.  They are
    normalized to one so the Streamlit editor can tolerate small rounding differences.
    IncrementalReturn is the return contribution for a stock with AIExposure=1.0.
    """
    df = (frame.copy() if frame is not None else DEFAULT_SCENARIOS.copy())
    required = {"Scenario", "Probability", "IncrementalReturn"}
    missing = required.difference(df.columns)
    if missing:
        raise ValueError(f"Scenario table is missing columns: {sorted(missing)}")
    df = df[list(required)].copy()
    df["Scenario"] = df["Scenario"].astype(str).str.strip()
    df["Probability"] = pd.to_numeric(df["Probability"], errors="coerce")
    df["IncrementalReturn"] = pd.to_numeric(df["IncrementalReturn"], errors="coerce")
    df = df.dropna(subset=["Probability", "IncrementalReturn"])
    if df.empty:
        raise ValueError("Scenario table contains no usable rows")
    if (df["Probability"] < 0).any():
        raise ValueError("Scenario probabilities cannot be negative")
    total = float(df["Probability"].sum())
    if not np.isfinite(total) or total <= 0:
        raise ValueError("Scenario probabilities must have positive total mass")
    df["Probability"] = df["Probability"] / total
    return df.reset_index(drop=True)


def scenario_moments(frame: pd.DataFrame | None = None) -> tuple[float, float]:
    df = normalize_scenarios(frame)
    p = df["Probability"].to_numpy(dtype=float)
    x = df["IncrementalReturn"].to_numpy(dtype=float)
    mean = float(np.sum(p * x))
    variance = float(np.sum(p * np.square(x - mean)))
    return mean, float(np.sqrt(max(variance, 0.0)))


def apply_ai_overlay(
    core_return: float,
    exposure: float,
    scenarios: pd.DataFrame | None = None,
    *,
    evidence_confidence: float = 0.50,
    maturity_confidence: float = 0.50,
    ai_skill_confidence: float = 0.25,
    uncertainty_aversion: float = 0.40,
) -> AIOverlayResult:
    """Apply a bounded, transparent AI optionality overlay to an existing forecast.

    Formula:
        credibility = evidence * maturity * AI-specific skill
        raw AI contribution = exposure * E[AI scenario return]
        uncertainty penalty = aversion * abs(exposure) * sigma(AI scenarios)
        certainty-equivalent AI = credibility * (raw contribution - penalty)
        adjusted return = existing core return + certainty-equivalent AI

    The function intentionally does not cap the final result because the Streamlit page
    is a sensitivity tool; implausible analyst assumptions should remain visible rather
    than being silently hidden by a clamp.
    """
    core = float(core_return)
    exp = float(exposure)
    if not np.isfinite(core) or not np.isfinite(exp):
        raise ValueError("Core return and AI exposure must be finite")
    mean, sigma = scenario_moments(scenarios)
    evidence = _clip01(evidence_confidence)
    maturity = _clip01(maturity_confidence)
    skill = _clip01(ai_skill_confidence)
    credibility = evidence * maturity * skill
    aversion = max(0.0, float(uncertainty_aversion))
    raw = exp * mean
    penalty = aversion * abs(exp) * sigma
    ce = credibility * (raw - penalty)
    return AIOverlayResult(
        core_return=core,
        exposure=exp,
        scenario_mean=mean,
        scenario_sigma=sigma,
        evidence_confidence=evidence,
        maturity_confidence=maturity,
        ai_skill_confidence=skill,
        credibility=credibility,
        uncertainty_aversion=aversion,
        raw_ai_contribution=raw,
        uncertainty_penalty=penalty,
        certainty_equivalent_ai=ce,
        adjusted_return=core + ce,
    )


def build_overlay_table(
    expected_returns: pd.DataFrame,
    exposures: pd.DataFrame,
    *,
    core_column: str,
    scenarios: pd.DataFrame | None = None,
    default_evidence: float = 0.50,
    default_maturity: float = 0.50,
    default_skill: float = 0.25,
    uncertainty_aversion: float = 0.40,
) -> pd.DataFrame:
    """Return one transparent AI-overlay row per ticker."""
    if "Ticker" not in expected_returns or core_column not in expected_returns:
        raise ValueError(f"Expected-return data requires Ticker and {core_column}")
    exp = expected_returns[["Ticker", core_column]].copy()
    exp["Ticker"] = exp["Ticker"].astype(str).str.upper().str.strip()
    exp[core_column] = pd.to_numeric(exp[core_column], errors="coerce")

    cfg = exposures.copy() if exposures is not None else pd.DataFrame(columns=["Ticker"])
    if "Ticker" not in cfg:
        cfg["Ticker"] = []
    cfg["Ticker"] = cfg["Ticker"].astype(str).str.upper().str.strip()
    merged = exp.merge(cfg, on="Ticker", how="left")

    def col(name: str, default: float) -> pd.Series:
        if name not in merged:
            return pd.Series(default, index=merged.index, dtype=float)
        return pd.to_numeric(merged[name], errors="coerce").fillna(default)

    merged["AIExposure"] = col("AIExposure", 0.0)
    merged["EvidenceConfidence"] = col("EvidenceConfidence", default_evidence)
    merged["MaturityConfidence"] = col("MaturityConfidence", default_maturity)
    merged["AISkillConfidence"] = col("AISkillConfidence", default_skill)

    rows = []
    for _, r in merged.dropna(subset=[core_column]).iterrows():
        result = apply_ai_overlay(
            r[core_column],
            r["AIExposure"],
            scenarios,
            evidence_confidence=r["EvidenceConfidence"],
            maturity_confidence=r["MaturityConfidence"],
            ai_skill_confidence=r["AISkillConfidence"],
            uncertainty_aversion=uncertainty_aversion,
        )
        rows.append({"Ticker": r["Ticker"], **result.to_dict()})
    return pd.DataFrame(rows)


def portfolio_ai_scenarios(
    weights: pd.DataFrame,
    exposures: pd.DataFrame,
    scenarios: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """Calculate the portfolio return shock from each AI scenario.

    This is an exposure diagnostic only.  It does not alter portfolio weights.
    """
    if "Ticker" not in weights or "Weight" not in weights:
        raise ValueError("Weights require Ticker and Weight columns")
    w = weights[["Ticker", "Weight"]].copy()
    w["Ticker"] = w["Ticker"].astype(str).str.upper().str.strip()
    w["Weight"] = pd.to_numeric(w["Weight"], errors="coerce").fillna(0.0)
    e = exposures[[c for c in ["Ticker", "AIExposure"] if c in exposures.columns]].copy()
    if "Ticker" not in e:
        e = pd.DataFrame({"Ticker": [], "AIExposure": []})
    if "AIExposure" not in e:
        e["AIExposure"] = 0.0
    e["Ticker"] = e["Ticker"].astype(str).str.upper().str.strip()
    e["AIExposure"] = pd.to_numeric(e["AIExposure"], errors="coerce").fillna(0.0)
    joined = w.merge(e, on="Ticker", how="left")
    joined["AIExposure"] = joined["AIExposure"].fillna(0.0)
    factor_exposure = float((joined["Weight"] * joined["AIExposure"]).sum())

    scen = normalize_scenarios(scenarios)
    out = scen.copy()
    out["PortfolioAIShock"] = out["IncrementalReturn"] * factor_exposure
    out["PortfolioAIFactorExposure"] = factor_exposure
    return out


def sensitivity_grid(
    core_return: float,
    scenarios: pd.DataFrame | None = None,
    *,
    evidence_confidence: float,
    maturity_confidence: float,
    ai_skill_confidence: float,
    exposures: Iterable[float] = tuple(np.linspace(0.0, 1.5, 7)),
    aversions: Iterable[float] = tuple(np.linspace(0.0, 1.0, 6)),
) -> pd.DataFrame:
    rows = []
    for exposure in exposures:
        for aversion in aversions:
            r = apply_ai_overlay(
                core_return,
                float(exposure),
                scenarios,
                evidence_confidence=evidence_confidence,
                maturity_confidence=maturity_confidence,
                ai_skill_confidence=ai_skill_confidence,
                uncertainty_aversion=float(aversion),
            )
            rows.append({
                "AIExposure": float(exposure),
                "UncertaintyAversion": float(aversion),
                "AdjustedReturn": r.adjusted_return,
                "AIContribution": r.certainty_equivalent_ai,
            })
    return pd.DataFrame(rows)
