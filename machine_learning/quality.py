from __future__ import annotations

"""Evidence-quality gates for ML research outputs.

A model can fit and produce a prediction while still having weak out-of-sample evidence.
This module deliberately separates computational success from investment usefulness.
"""

from .common import MLResult


def _num(v):
    try: return float(v)
    except Exception: return None


def _wf(metrics):
    if not isinstance(metrics,dict): return {}
    return metrics.get("walk_forward") or metrics.get("hgb_walk_forward") or {}


def _baseline_note(wf: dict) -> str | None:
    model_da=_num(wf.get("directional_accuracy")); base_da=_num(wf.get("baseline_directional_accuracy"))
    model_mae=_num(wf.get("mae")); base_mae=_num(wf.get("baseline_mae"))
    bits=[]
    if model_da is not None and base_da is not None:
        bits.append(f"direction {model_da:.1%} vs simple baseline {base_da:.1%}")
    if model_mae is not None and base_mae is not None:
        bits.append(f"MAE {model_mae:.1%} vs baseline {base_mae:.1%}")
    return "; ".join(bits) if bits else None


def gate_ml_result(result: MLResult) -> MLResult:
    if result.status not in {"PASS","REVIEW","WEAK_SIGNAL"}:
        return result
    metrics=result.metrics or {}
    notes=[]

    if result.name=="Expected 12M Excess Return":
        wf=metrics.get("hgb_walk_forward") or {}
        r2=_num(wf.get("r2")); da=_num(wf.get("directional_accuracy")); n=int(wf.get("n") or metrics.get("training_rows") or 0)
        mae_edge=_num(wf.get("mae_improvement_vs_baseline")); da_edge=_num(wf.get("directional_accuracy_edge_vs_baseline"))
        if mae_edge is not None and da_edge is not None and mae_edge<=0 and da_edge<=0:
            result.status="WEAK_SIGNAL"; result.confidence="Low"
            notes.append("The walk-forward model did not beat the simple historical baseline on error or direction; ignore the point forecast.")
        elif (r2 is not None and r2<=0) and (da is None or da<.55):
            result.status="WEAK_SIGNAL"; result.confidence="Low"
            notes.append("Negative walk-forward R² and sub-55% directional accuracy; prediction is not investment-grade evidence.")
        elif n<80 or (da is not None and da<.58) or (r2 is not None and r2<.03) or (
            mae_edge is not None and da_edge is not None and mae_edge<.05 and da_edge<.03
        ):
            result.status="REVIEW"; result.confidence="Low" if n<50 else "Moderate"
            notes.append("Out-of-sample skill is limited; use only as a weak second opinion.")
        baseline=_baseline_note(wf)
        if baseline:
            notes.append("Baseline check: "+baseline+".")

    elif result.name=="Consensus / Earnings Surprise":
        wf=metrics.get("walk_forward") or {}
        r2=_num(wf.get("r2")); da=_num(wf.get("directional_accuracy")); n=int(wf.get("n") or metrics.get("training_rows") or 0)
        pred=abs(_num(result.prediction) or 0.0)
        mae_edge=_num(wf.get("mae_improvement_vs_baseline")); da_edge=_num(wf.get("directional_accuracy_edge_vs_baseline"))
        if pred>.50:
            result.status="WEAK_SIGNAL"; result.confidence="Low"
            notes.append(
                "Absolute predicted EPS surprise exceeds 50%; the point estimate is outlier-sensitive and should not be used as a numeric anchor."
            )
        elif mae_edge is not None and da_edge is not None and mae_edge<=0 and da_edge<=0:
            result.status="WEAK_SIGNAL"; result.confidence="Low"
            notes.append("High hit-rate alone is not enough: the model failed to beat the simple historical earnings-surprise baseline.")
        elif n<12 or (r2 is not None and r2<0):
            result.status="WEAK_SIGNAL"; result.confidence="Low"
            notes.append("Earnings-history sample or walk-forward fit is too weak to treat the forecast as a reliable signal.")
        elif n<24 or (da is not None and da<.58) or (
            mae_edge is not None and da_edge is not None and mae_edge<.05 and da_edge<.03
        ):
            result.status="REVIEW"; result.confidence="Low" if n<24 else "Moderate"
            notes.append("The earnings model adds only limited value over a simple historical rule; use as context, not a numeric anchor.")
        baseline=_baseline_note(wf)
        if baseline:
            notes.append("Baseline check: "+baseline+".")

    elif result.name=="Financial Anomaly Detection":
        years=metrics.get("years") or []
        n=len(years)
        if n<10:
            result.status="REVIEW"; result.confidence="Low"
            notes.append(f"Only {n} annual observations; anomaly percentile is a diligence flag, not a statistically robust conclusion.")

    elif result.name=="Market Regime Classifier":
        probs=metrics.get("regime_probabilities") or {}
        total=sum(max(0.0,_num(v) or 0.0) for v in probs.values())
        p=max([_num(v) or 0 for v in probs.values()] or [0])
        if total and abs(total-1.0)>.01:
            result.status="REVIEW"; result.confidence="Low"
            notes.append(f"Regime weights sum to {total:.1%}, not 100%; treat this run as invalid until the duplicate-label aggregation issue is regenerated.")
        elif p<.55:
            result.status="REVIEW"; result.confidence="Moderate" if int(metrics.get("monthly_rows") or 0)>=60 else "Low"
            notes.append("Cluster assignment is diffuse; regime label is contextual rather than a strong timing signal.")

    if notes:
        result.details=dict(result.details or {})
        existing=list(result.details.get("quality_gate_notes") or [])
        result.details["quality_gate_notes"]=existing+notes
        result.summary=result.summary+" Quality gate: "+" ".join(notes)
    return result


def gate_results(results):
    return [gate_ml_result(r) for r in results]
