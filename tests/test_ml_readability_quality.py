from __future__ import annotations

import numpy as np
import pandas as pd
from sklearn.linear_model import Ridge

from ai_growth_research import _clean_ai_evidence
from machine_learning.common import MLResult
from machine_learning.models import MarketRegimeModel, REGIME_FEATURES
from machine_learning.quality import gate_ml_result
from machine_learning.validation import expanding_walk_forward


def test_walk_forward_reports_leakage_safe_baseline_metrics():
    rng=np.random.default_rng(7)
    dates=pd.date_range("2012-01-01",periods=80,freq="90D")
    frame=pd.DataFrame({
        "as_of":dates,
        "target_date":dates+pd.Timedelta(days=365),
        "x":rng.normal(size=len(dates)),
    })
    frame["target"]=0.08*frame["x"]+rng.normal(0,.08,len(frame))
    result=expanding_walk_forward(Ridge(alpha=1.0),frame,["x"],"target",min_train=24,step=3)
    m=result.metrics
    assert m.get("n",0)>0
    assert m.get("baseline_mae") is not None
    assert m.get("baseline_directional_accuracy") is not None
    assert m.get("mae_improvement_vs_baseline") is not None
    assert m.get("directional_accuracy_edge_vs_baseline") is not None
    assert m.get("baseline_method")=="expanding training-set mean"


def test_high_hit_rate_without_baseline_edge_is_not_automatically_pass():
    res=MLResult(
        "Consensus / Earnings Surprise","PASS","Synthetic",
        prediction=.03,confidence="High",
        metrics={"walk_forward":{
            "n":60,"r2":.08,"directional_accuracy":.91,"mae":.025,
            "baseline_directional_accuracy":.92,"baseline_mae":.024,
            "mae_improvement_vs_baseline":-.04,
            "directional_accuracy_edge_vs_baseline":-.01,
        }},
    )
    gated=gate_ml_result(res)
    assert gated.status=="WEAK_SIGNAL"
    assert gated.confidence=="Low"


def test_incomplete_regime_weights_are_review_gated():
    res=MLResult(
        "Market Regime Classifier","PASS","Transition / mixed",confidence="High",
        metrics={"monthly_rows":228,"regime_probabilities":{
            "Transition / mixed":.365,"Growth / risk-on":.307,
            "Inflation / stagflation pressure":.273,"Risk-off / crisis":.001,
        }},
    )
    gated=gate_ml_result(res)
    assert gated.status=="REVIEW"
    assert gated.confidence=="Low"
    assert "sum to" in " ".join((gated.details or {}).get("quality_gate_notes",[]))


def test_market_regime_human_labels_preserve_all_cluster_probability():
    rng=np.random.default_rng(42)
    frame=pd.DataFrame(index=pd.date_range("2008-01-31",periods=180,freq="ME"))
    for c in REGIME_FEATURES:
        frame[c]=rng.normal(0,.12,len(frame))
    frame["equity_vol_3m"]=np.abs(frame["equity_vol_3m"])+.08
    result=MarketRegimeModel().fit_predict(frame)
    assert result.status=="PASS"
    probs=(result.metrics or {}).get("regime_probabilities") or {}
    assert abs(sum(probs.values())-1.0)<1e-10
    assert len((result.metrics or {}).get("cluster_probabilities") or [])==5


def test_ai_template_rows_do_not_create_false_company_specific_signal():
    rows=[
        {"kpi":"AI-attributed revenue","current":"Analyst input"},
        {"kpi":"AI product users / seats","current":"To be updated"},
        {
            "kpi":"AI-related payment fraud detection",
            "current":"Company reported: AI risk models reduced fraud losses 12%",
            "signal":"Positive",
            "investment_read_through":"AI is improving transaction risk controls.",
        },
    ]
    corpus="\n".join([
        "AI-attributed revenue | Analyst input",
        "AI product users / seats | To be updated",
        "AI-related payment fraud detection | Company reported: AI risk models reduced fraud losses 12%",
    ])
    clean_rows,clean_corpus,stats=_clean_ai_evidence(rows,corpus)
    assert len(clean_rows)==1
    assert stats["filtered_placeholder_rows"]==2
    assert "Analyst input" not in clean_corpus
    assert "12%" in clean_corpus
