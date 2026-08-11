from __future__ import annotations

from typing import Any
import numpy as np
import pandas as pd
from scipy.optimize import minimize
from sklearn.cluster import KMeans
from sklearn.covariance import LedoitWolf
from sklearn.ensemble import HistGradientBoostingRegressor, IsolationForest, RandomForestRegressor
from sklearn.impute import SimpleImputer
from sklearn.inspection import permutation_importance
from sklearn.linear_model import ElasticNet, Ridge
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler

from .common import MLResult, RANDOM_STATE, num
from .validation import assert_no_future_leakage, expanding_walk_forward

EXPECTED_FEATURES=[
    "revenue_growth","operating_margin","net_margin","fcf_margin","capex_to_revenue","rd_to_revenue",
    "roe","net_debt_to_revenue","momentum_12m","momentum_6m","volatility_6m","drawdown_12m",
]


def _confidence_from_n(n:int, threshold_high:int=80, threshold_medium:int=40) -> str:
    return "High" if n>=threshold_high else "Moderate" if n>=threshold_medium else "Low"


def _feature_drivers(model, X: pd.DataFrame, y: pd.Series, feature_cols: list[str], top_n:int=6) -> list[dict[str,Any]]:
    try:
        imp=permutation_importance(model,X,y,n_repeats=5,random_state=RANDOM_STATE,scoring="neg_mean_absolute_error")
        pairs=sorted(zip(feature_cols,imp.importances_mean),key=lambda x:abs(x[1]),reverse=True)[:top_n]
        return [{"feature":k,"importance":float(v)} for k,v in pairs]
    except Exception:
        return []


class ExpectedReturnModel:
    name="Expected 12M Excess Return"
    feature_cols=EXPECTED_FEATURES

    def _estimators(self):
        hgb=Pipeline([
            ("imputer",SimpleImputer(strategy="median")),
            ("model",HistGradientBoostingRegressor(max_iter=180,max_leaf_nodes=12,learning_rate=.05,l2_regularization=.5,random_state=RANDOM_STATE)),
        ])
        elastic=Pipeline([
            ("imputer",SimpleImputer(strategy="median")),
            ("scale",StandardScaler()),
            ("model",ElasticNet(alpha=.02,l1_ratio=.25,max_iter=5000,random_state=RANDOM_STATE)),
        ])
        return hgb,elastic

    def fit_predict(self, frame: pd.DataFrame, current: dict[str,Any]) -> MLResult:
        if frame is None or len(frame)<30:
            return MLResult(self.name,"INSUFFICIENT_DATA",f"Need at least 30 point-in-time company-year observations; found {0 if frame is None else len(frame)}.",confidence="Low",details={"minimum_rows":30})
        df=frame.copy(); assert_no_future_leakage(df)
        for c in self.feature_cols:
            if c not in df: df[c]=np.nan
        hgb,elastic=self._estimators()
        min_train=max(24,min(40,len(df)//2)); step=max(1,len(df)//20)
        wf_h=expanding_walk_forward(hgb,df,self.feature_cols,"target_excess_return_12m",min_train=min_train,step=step)
        wf_e=expanding_walk_forward(elastic,df,self.feature_cols,"target_excess_return_12m",min_train=min_train,step=step)
        hgb.fit(df[self.feature_cols],df["target_excess_return_12m"]); elastic.fit(df[self.feature_cols],df["target_excess_return_12m"])
        x=pd.DataFrame([{c:num(current.get(c)) for c in self.feature_cols}])
        p_h=float(hgb.predict(x)[0]); p_e=float(elastic.predict(x)[0]); pred=.65*p_h+.35*p_e
        metrics={"hgb_walk_forward":wf_h.metrics,"elastic_walk_forward":wf_e.metrics,"training_rows":len(df),"hgb_prediction":p_h,"elastic_prediction":p_e}
        drivers=_feature_drivers(hgb,df[self.feature_cols],df["target_excess_return_12m"],self.feature_cols)
        return MLResult(self.name,"PASS",f"ML ensemble estimates {pred:.1%} 12-month excess return versus benchmark; use as a cross-check, not a target price.",prediction=pred,confidence=_confidence_from_n(len(df)),metrics=metrics,drivers=drivers)


EARNINGS_FEATURES=["prior_surprise_1q","prior_surprise_2q_avg","prior_surprise_4q_avg","surprise_vol_4q","price_momentum_3m","price_vol_3m","eps_estimate"]
class EarningsSurpriseModel:
    name="Consensus / Earnings Surprise"
    def fit_predict(self, frame: pd.DataFrame) -> MLResult:
        if frame is None or len(frame)<8:
            return MLResult(self.name,"INSUFFICIENT_DATA",f"Need at least 8 historical earnings observations after lag construction; found {0 if frame is None else len(frame)}.",confidence="Low")
        df=frame.copy(); assert_no_future_leakage(df)
        for c in EARNINGS_FEATURES:
            if c not in df: df[c]=np.nan
        model=Pipeline([
            ("imputer",SimpleImputer(strategy="median")),
            ("model",RandomForestRegressor(n_estimators=300,max_depth=5,min_samples_leaf=2,random_state=RANDOM_STATE,n_jobs=-1)),
        ])
        min_train=max(6,min(12,len(df)//2))
        wf=expanding_walk_forward(model,df,EARNINGS_FEATURES,"target_surprise",min_train=min_train,step=1)
        model.fit(df[EARNINGS_FEATURES],df["target_surprise"])
        latest=df.iloc[-1]
        next_x=pd.DataFrame([{c:latest.get(c) for c in EARNINGS_FEATURES}])
        next_x["prior_surprise_1q"]=latest.get("target_surprise")
        next_x["prior_surprise_2q_avg"]=df["target_surprise"].tail(2).mean()
        next_x["prior_surprise_4q_avg"]=df["target_surprise"].tail(4).mean()
        next_x["surprise_vol_4q"]=df["target_surprise"].tail(4).std()
        pred=float(model.predict(next_x)[0])
        drivers=_feature_drivers(model,df[EARNINGS_FEATURES],df["target_surprise"],EARNINGS_FEATURES)
        return MLResult(self.name,"PASS",f"Model-implied next earnings surprise is {pred:.1%}; positive means reported EPS above the prevailing estimate, subject to current-consensus availability.",prediction=pred,confidence=_confidence_from_n(len(df),24,12),metrics={"walk_forward":wf.metrics,"training_rows":len(df)},drivers=drivers)


ANOMALY_FEATURES=["revenue_growth","operating_margin","net_margin","fcf_margin","capex_to_revenue","rd_to_revenue","sbc_to_revenue"]
class FinancialAnomalyModel:
    name="Financial Anomaly Detection"
    def fit_predict(self, frame: pd.DataFrame) -> MLResult:
        if frame is None or len(frame)<5:
            return MLResult(self.name,"INSUFFICIENT_DATA",f"Need at least 5 annual observations; found {0 if frame is None else len(frame)}.",confidence="Low")
        df=frame.copy().sort_values("year")
        for c in ANOMALY_FEATURES:
            if c not in df: df[c]=np.nan
        pipe=Pipeline([("imputer",SimpleImputer(strategy="median")),("scale",StandardScaler())])
        X=pipe.fit_transform(df[ANOMALY_FEATURES])
        model=IsolationForest(n_estimators=400,contamination="auto",random_state=RANDOM_STATE)
        model.fit(X)
        scores=-model.score_samples(X)
        latest=float(scores[-1]); pct=float((scores<=latest).mean())
        z=np.abs(np.asarray(X[-1],dtype=float))
        drivers=sorted([{"feature":f,"standardized_deviation":float(v)} for f,v in zip(ANOMALY_FEATURES,z)],key=lambda x:x["standardized_deviation"],reverse=True)[:5]
        label="Elevated" if pct>=.8 else "Moderate" if pct>=.6 else "Normal"
        return MLResult(self.name,"PASS",f"Latest financial profile is at the {pct:.0%} anomaly percentile of the company's available history ({label}).",prediction=pct,confidence=_confidence_from_n(len(df),10,6),metrics={"years":df["year"].tolist(),"raw_anomaly_score":latest},drivers=drivers)


REGIME_FEATURES=["equity_6m","bond_6m","credit_6m","commodity_6m","dollar_6m","equity_vol_3m"]
class MarketRegimeModel:
    name="Market Regime Classifier"
    def fit_predict(self, frame: pd.DataFrame) -> MLResult:
        if frame is None or len(frame)<36:
            return MLResult(self.name,"INSUFFICIENT_DATA",f"Need at least 36 monthly market observations; found {0 if frame is None else len(frame)}.",confidence="Low")
        df=frame.dropna(subset=REGIME_FEATURES).copy()
        scaler=StandardScaler(); X=scaler.fit_transform(df[REGIME_FEATURES])
        model=KMeans(n_clusters=5,n_init=20,random_state=RANDOM_STATE).fit(X)
        centers=pd.DataFrame(scaler.inverse_transform(model.cluster_centers_),columns=REGIME_FEATURES)
        def label(row):
            eq=row["equity_6m"]; vol=row["equity_vol_3m"]; bond=row["bond_6m"]; cmd=row["commodity_6m"]; credit=row["credit_6m"]
            if eq<0 and vol>centers["equity_vol_3m"].median(): return "Risk-off / crisis"
            if eq<0 and bond>0: return "Recession / disinflation"
            if cmd>0 and bond<0 and eq<=centers["equity_6m"].median(): return "Inflation / stagflation pressure"
            if eq>0 and credit>0 and vol<=centers["equity_vol_3m"].median(): return "Growth / risk-on"
            return "Transition / mixed"
        names={i:label(centers.iloc[i]) for i in range(5)}
        current=X[-1]; d=np.linalg.norm(model.cluster_centers_-current,axis=1); probs=np.exp(-d); probs=probs/probs.sum()
        cluster=int(model.predict([current])[0]); current_name=names[cluster]
        prob_map={names[i]:float(probs[i]) for i in range(5)}
        return MLResult(self.name,"PASS",f"Current market regime is classified as {current_name} with distance-based confidence {probs[cluster]:.0%}.",prediction=current_name,confidence=_confidence_from_n(len(df),96,60),metrics={"monthly_rows":len(df),"regime_probabilities":prob_map,"cluster_centers":centers.to_dict(orient="records")})


class AIImpactMLModel:
    name="AI Impact ML"
    def fit_predict(self, kpis: pd.DataFrame, target_series: pd.Series | None = None) -> MLResult:
        n=0 if kpis is None else len(kpis)
        if kpis is None or n<8:
            return MLResult(self.name,"INSUFFICIENT_DATA",f"AI KPI history has {n} snapshot(s). At least 8 dated snapshots are required before fitting a company-specific ML relation; until then the deterministic AI economics bridge remains authoritative.",confidence="Low",details={"minimum_snapshots":8,"current_snapshots":n})
        df=kpis.copy().sort_values("captured_at")
        numeric=[c for c in df.columns if c!="captured_at" and pd.api.types.is_numeric_dtype(df[c])]
        if len(numeric)<2:
            return MLResult(self.name,"INSUFFICIENT_DATA","AI KPI history does not yet contain at least two numeric explanatory series.",confidence="Low")
        if target_series is None or len(target_series)!=len(df):
            monetary=[c for c in numeric if any(x in c.lower() for x in ("revenue","backlog","run-rate","income"))]
            target_col=monetary[0] if monetary else numeric[0]
            y=pd.to_numeric(df[target_col],errors="coerce").pct_change().shift(-1)
            features=[c for c in numeric if c!=target_col]
            target_name=f"next-change in {target_col}"
        else:
            y=pd.Series(target_series,index=df.index,dtype=float); features=numeric; target_name="provided AI economic target"
        work=df[features].copy(); work["target"]=y; work["as_of"]=df["captured_at"]; work=work.dropna(subset=["target"])
        if len(work)<6 or not features:
            return MLResult(self.name,"INSUFFICIENT_DATA","After constructing a forward AI target, fewer than six usable training rows remain.",confidence="Low")
        model=Pipeline([("imputer",SimpleImputer(strategy="median")),("scale",StandardScaler()),("model",Ridge(alpha=2.0))])
        model.fit(work[features],work["target"])
        pred=float(model.predict(df[features].tail(1))[0])
        coefs=model.named_steps["model"].coef_
        drivers=sorted([{"feature":f,"coefficient":float(c)} for f,c in zip(features,coefs)],key=lambda x:abs(x["coefficient"]),reverse=True)[:6]
        return MLResult(self.name,"PASS",f"Company-specific AI model estimates {pred:.1%} for {target_name}; interpret only alongside the deterministic AI economics bridge and KPI-source quality.",prediction=pred,confidence=_confidence_from_n(len(work),20,10),metrics={"training_rows":len(work),"features":features,"target":target_name},drivers=drivers)


class PortfolioPositionSizingModel:
    name="Portfolio ML / Position Sizing"
    def optimize(self, returns: pd.DataFrame, expected_returns: dict[str,float] | None, current_weights: dict[str,float] | None=None, max_weight:float=.25, risk_aversion:float=4.0) -> MLResult:
        if returns is None or returns.shape[0]<126 or returns.shape[1]<2:
            return MLResult(self.name,"INSUFFICIENT_DATA","Need at least 126 daily return rows and two portfolio assets.",confidence="Low")
        r=returns.dropna(how="all").fillna(0.0)
        assets=list(r.columns); cov=LedoitWolf().fit(r.values).covariance_*252
        hist_mu=(r.mean()*252).to_dict(); mu=np.array([num((expected_returns or {}).get(a),hist_mu.get(a,0.0)) or 0.0 for a in assets])
        n=len(assets); start=np.array([(current_weights or {}).get(a,1/n) for a in assets],dtype=float); start=np.clip(start,0,None); start=start/start.sum() if start.sum()>0 else np.repeat(1/n,n)
        bounds=[(0,min(max_weight,1.0)) for _ in assets]
        cons={"type":"eq","fun":lambda w:float(w.sum()-1.0)}
        def objective(w): return float(-(w@mu-risk_aversion*(w@cov@w)))
        result=minimize(objective,start,method="SLSQP",bounds=bounds,constraints=[cons],options={"maxiter":500,"ftol":1e-10})
        if not result.success:
            return MLResult(self.name,"FAIL",f"Position-sizing optimizer failed: {result.message}",confidence="Low")
        w=np.asarray(result.x); port_ret=float(w@mu); port_vol=float(np.sqrt(w@cov@w))
        rows=[{"ticker":a,"suggested_weight":float(x),"current_weight":float(start[i]),"weight_change":float(x-start[i]),"expected_return_input":float(mu[i])} for i,(a,x) in enumerate(zip(assets,w))]
        rows=sorted(rows,key=lambda x:x["suggested_weight"],reverse=True)
        return MLResult(self.name,"PASS",f"ML-assisted long-only sizing suggests an annual expected return input of {port_ret:.1%} at {port_vol:.1%} shrinkage-estimated volatility. No trades are executed.",prediction=port_ret,confidence=_confidence_from_n(len(r),756,252),metrics={"annualized_expected_return":port_ret,"annualized_volatility":port_vol,"risk_aversion":risk_aversion,"max_weight":max_weight,"optimizer_success":True},details={"weights":rows})
