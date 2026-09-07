from __future__ import annotations

"""Portfolio-side continual-learning governance.

The portfolio optimizer already keeps ML separate from deterministic allocation mathematics. This
patch adds one more safety layer: the confidence-based ML blend is multiplied by the realized
out-of-sample influence multiplier maintained in machine_learning.model_registry.

No registry / legacy database => preserve the existing behavior. Once the continual-learning
schema exists, UNPROVEN models receive only a small fraction of their former influence,
CHALLENGER models earn partial influence, CHAMPION models can use the full confidence-adjusted cap,
and DEMOTED models receive zero expected-return influence.
"""

from pathlib import Path
import sqlite3

import numpy as np
import pandas as pd


_INSTALLED = False


def _registry_row(db_path: Path, model: str = "Expected 12M Excess Return") -> dict | None:
    if not db_path.exists():
        return None
    try:
        with sqlite3.connect(db_path) as con:
            exists = con.execute(
                "SELECT 1 FROM sqlite_master WHERE type='table' AND name='model_registry'"
            ).fetchone()
            if not exists:
                return None
            row = con.execute(
                """SELECT model,champion_version,status,matured_predictions,mae,baseline_mae,
                          skill_vs_baseline,directional_accuracy,information_coefficient,
                          calibration_score,drift_ratio,influence_multiplier,last_evaluated
                   FROM model_registry WHERE model=?""",
                (model,),
            ).fetchone()
        if not row:
            return None
        keys = [
            "model","champion_version","status","matured_predictions","mae","baseline_mae",
            "skill_vs_baseline","directional_accuracy","information_coefficient",
            "calibration_score","drift_ratio","influence_multiplier","last_evaluated",
        ]
        return dict(zip(keys, row))
    except Exception:
        return None


def apply_learning_governance(
    expected_inputs: pd.DataFrame,
    meta: dict,
    history_db: str | Path | None,
) -> tuple[pd.DataFrame, dict]:
    out = expected_inputs.copy()
    meta = dict(meta or {})
    if out.empty or history_db in (None, ""):
        return out, meta
    row = _registry_row(Path(history_db))
    if row is None:
        return out, meta

    try:
        multiplier = float(row.get("influence_multiplier"))
    except Exception:
        multiplier = 1.0
    multiplier = float(np.clip(multiplier, 0.0, 1.0))
    if "MLBlendWeight" in out:
        original = pd.to_numeric(out["MLBlendWeight"], errors="coerce").fillna(0.0)
        learned = original * multiplier
        out["PreLearningMLBlendWeight"] = original
        out["LearningInfluenceMultiplier"] = multiplier
        out["MLBlendWeight"] = learned
        if {"ClassicalExpectedReturn", "MLTotalReturn"}.issubset(out.columns):
            classical = pd.to_numeric(out["ClassicalExpectedReturn"], errors="coerce")
            ml_total = pd.to_numeric(out["MLTotalReturn"], errors="coerce")
            out["BlendedMLExpectedReturn"] = np.where(
                ml_total.notna(),
                (1 - learned) * classical + learned * ml_total,
                classical,
            )
    out["LearningStatus"] = str(row.get("status") or "UNKNOWN")
    out["LearningMaturedPredictions"] = int(row.get("matured_predictions") or 0)
    out["LearningSkillVsBaseline"] = row.get("skill_vs_baseline")
    out["LearningDirectionalAccuracy"] = row.get("directional_accuracy")
    out["LearningInformationCoefficient"] = row.get("information_coefficient")
    out["LearningDriftRatio"] = row.get("drift_ratio")
    out["LearningChampionVersion"] = row.get("champion_version")

    meta["continual_learning"] = {
        "model": row.get("model"),
        "status": row.get("status"),
        "champion_version": row.get("champion_version"),
        "matured_predictions": int(row.get("matured_predictions") or 0),
        "skill_vs_baseline": row.get("skill_vs_baseline"),
        "directional_accuracy": row.get("directional_accuracy"),
        "information_coefficient": row.get("information_coefficient"),
        "drift_ratio": row.get("drift_ratio"),
        "influence_multiplier": multiplier,
        "last_evaluated": row.get("last_evaluated"),
        "policy": (
            "Realized OOS evidence multiplies the confidence-based ML blend. DEMOTED=0; "
            "UNPROVEN is deliberately small; CHAMPION can use the full pre-existing cap."
        ),
    }
    return out, meta


def install() -> None:
    global _INSTALLED
    if _INSTALLED:
        return
    from . import portfolio_optimization as po

    original = po.expected_return_inputs

    def governed_expected_return_inputs(*args, **kwargs):
        out, meta = original(*args, **kwargs)
        history_db = kwargs.get("history_db")
        if history_db is None and len(args) >= 5:
            history_db = args[4]
        return apply_learning_governance(out, meta, history_db)

    po.expected_return_inputs = governed_expected_return_inputs
    _INSTALLED = True
