from __future__ import annotations

import numpy as np
import pandas as pd
from sklearn.linear_model import Ridge

from machine_learning.common import MLResult
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
