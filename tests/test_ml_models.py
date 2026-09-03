from __future__ import annotations

import numpy as np
import pandas as pd

from machine_learning.models import (
    ExpectedReturnModel,EarningsSurpriseModel,FinancialAnomalyModel,MarketRegimeModel,
    AIImpactMLModel,PortfolioPositionSizingModel,EXPECTED_FEATURES,EARNINGS_FEATURES,REGIME_FEATURES,
)
from machine_learning.common import MLResult
from machine_learning.quality import gate_ml_result
import segment_source_engine as segment_sources

rng=np.random.default_rng(42)

n=70
dates=pd.date_range("2010-01-01",periods=n,freq="90D")
df=pd.DataFrame({"as_of":dates,"target_date":dates+pd.Timedelta(days=365)})
for c in EXPECTED_FEATURES: df[c]=rng.normal(0,.2,n)
df["target_excess_return_12m"]=.15*df["revenue_growth"]+.1*df["fcf_margin"]-.08*df["volatility_6m"]+rng.normal(0,.05,n)
cur={c:float(df[c].iloc[-1]) for c in EXPECTED_FEATURES}
r=ExpectedReturnModel().fit_predict(df,cur)
assert r.status=="PASS" and isinstance(r.prediction,float)
wf=(r.metrics or {}).get("hgb_walk_forward",{})
assert wf.get("n",0)>0
assert wf.get("purged_target_horizons") is True
assert wf.get("min_train_rows",0)>=24

n=20
ed=pd.date_range("2020-01-01",periods=n,freq="90D")
e=pd.DataFrame({"as_of":ed-pd.Timedelta(days=1),"target_date":ed})
for c in EARNINGS_FEATURES: e[c]=rng.normal(0,.1,n)
e["target_surprise"]=.3*e["prior_surprise_1q"]+rng.normal(0,.03,n)
r=EarningsSurpriseModel().fit_predict(e)
assert r.status=="PASS"
assert (r.metrics or {}).get("walk_forward",{}).get("purged_target_horizons") is True

# An extreme point estimate must not be presented as a normal PASS even if backtest metrics look acceptable.
extreme=MLResult(
    "Consensus / Earnings Surprise","PASS","Synthetic extreme surprise",
    prediction=1.07,confidence="High",
    metrics={"walk_forward":{"n":72,"r2":.16,"directional_accuracy":.68}},
)
gated=gate_ml_result(extreme)
assert gated.status=="WEAK_SIGNAL"
assert gated.confidence=="Low"
assert "exceeds 50%" in " ".join((gated.details or {}).get("quality_gate_notes",[]))

f=pd.DataFrame({"year":range(2018,2026)})
for c in ["revenue_growth","operating_margin","net_margin","fcf_margin","capex_to_revenue","rd_to_revenue","sbc_to_revenue"]:
    f[c]=rng.normal(.15,.03,len(f))
r=FinancialAnomalyModel().fit_predict(f)
assert r.status=="PASS" and isinstance(r.prediction,float)

m=pd.DataFrame(index=pd.date_range("2015-01-31",periods=100,freq="ME"))
for c in REGIME_FEATURES: m[c]=rng.normal(0,.1,len(m))
m["equity_vol_3m"]=np.abs(m["equity_vol_3m"])+.1
r=MarketRegimeModel().fit_predict(m)
assert r.status=="PASS" and isinstance(r.prediction,str)

ai=pd.DataFrame({
    "captured_at":pd.date_range("2024-01-01",periods=12,freq="90D"),
    "AI revenue":np.linspace(10,30,12)+rng.normal(0,.5,12),
    "AI users":np.linspace(100,500,12)+rng.normal(0,5,12),
    "AI backlog":np.linspace(20,70,12)+rng.normal(0,1,12),
})
r=AIImpactMLModel().fit_predict(ai)
assert r.status=="PASS"
not_ready=AIImpactMLModel().fit_predict(ai.head(3))
assert not_ready.status=="INSUFFICIENT_DATA"

rets=pd.DataFrame(rng.normal(.0004,.012,(400,4)),columns=list("ABCD"))
r=PortfolioPositionSizingModel().optimize(
    rets,{"A":.12,"B":.10,"C":.08,"D":.09},{"A":.25,"B":.25,"C":.25,"D":.25},max_weight=.4,
)
assert r.status=="PASS"
weights=(r.details or {}).get("weights",[])
assert abs(sum(x["suggested_weight"] for x in weights)-1)<1e-6

# Source-policy regression: even if a regulatory document has a numerically lower legacy
# priority, issuer-owned material must stay first. Chevron's fallback must also cite Chevron.
original_issuer=segment_sources.issuer_segment_documents
original_sec=segment_sources.sec_segment_documents
try:
    segment_sources.issuer_segment_documents=lambda ticker:[{"kind":"issuer","url":"https://issuer.example/report","priority":12,"issuer_owned":True}]
    segment_sources.sec_segment_documents=lambda ticker,headers:[{"kind":"regulator","url":"https://regulator.example/filing","priority":1,"issuer_owned":False}]
    docs=segment_sources.collect_segment_documents("TEST",{})
    assert docs[0]["kind"]=="issuer",docs
finally:
    segment_sources.issuer_segment_documents=original_issuer
    segment_sources.sec_segment_documents=original_sec
cvx_fallback=segment_sources.verified_fallback("CVX")
assert cvx_fallback.get("source","").startswith("https://www.chevron.com/")
assert {"Upstream","Downstream"}.issubset(set(cvx_fallback.get("segments",[])))

print("six ML models + quality gates + issuer-first segment source synthetic smoke test PASS")
