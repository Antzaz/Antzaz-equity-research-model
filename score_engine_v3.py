from __future__ import annotations

"""Reliability-aware overlay for Score Engine v2.

The key principle is that weak valuation-model evidence should reduce score coverage and
confidence, not mechanically turn a profitable company into a zero. Capital-intensive
balance sheets are also judged with multiple leverage lenses rather than net-debt/FCF alone.
"""

import math
from score_engine_v2 import (
    compute_score_bundle as _v2, DIMENSION_WEIGHTS, QUALITY_WEIGHTS, VALUATION_WEIGHTS,
    weighted_available, clamp, num,
)


def _latest_history(wb):
    if 'Historical Financials' not in wb.sheetnames: return {}
    ws=wb['Historical Financials']; vals={}
    cols=[c for c in range(2,min(ws.max_column,8)+1) if isinstance(ws.cell(3,c).value,(int,float))]
    if not cols: return vals
    c=cols[-1]
    vals['revenue']=num(ws.cell(4,c).value); vals['operating_income']=num(ws.cell(9,c).value); vals['net_income']=num(ws.cell(11,c).value)
    vals['ocf']=num(ws.cell(14,c).value); vals['capex']=num(ws.cell(15,c).value)
    if vals['ocf'] is not None and vals['capex'] is not None: vals['fcf']=vals['ocf']-abs(vals['capex'])
    return vals

def _company_net_debt(wb):
    if 'Company Data' not in wb.sheetnames: return None,None,None
    ws=wb['Company Data']; cash=num(ws['B12'].value); debt=num(ws['B13'].value); net=(debt-cash) if debt is not None and cash is not None else num(ws['B14'].value)
    return cash,debt,net

def _cost_of_capital_value(wb,label):
    if 'Cost of Capital' not in wb.sheetnames: return None
    ws=wb['Cost of Capital']
    for r in range(1,ws.max_row+1):
        if str(ws.cell(r,1).value or '').strip().lower()==label.lower(): return num(ws.cell(r,2).value)
    return None

def _capital_structure_score(wb):
    hist=_latest_history(wb); _,_,net=_company_net_debt(wb)
    if net is None: return None,{"status":"Missing"}
    if net<=0:
        return 92.0,{"net_debt":net,"status":"Complete","formula":"Net-cash balance sheet receives 92/100 before any separate liquidity analysis."}
    op=hist.get('operating_income'); fcf=hist.get('fcf'); dw=_cost_of_capital_value(wb,'Debt weight')
    op_ratio=net/op if op not in (None,0) and op>0 else None
    fcf_ratio=net/abs(fcf) if fcf not in (None,0) else None
    # Operating-income capacity gets the highest weight because annual FCF can be distorted by
    # working capital, hedging collateral or large investment cycles in capital-intensive sectors.
    op_score=clamp(75-8*(op_ratio-3)) if op_ratio is not None else None
    fcf_score=clamp(70-5*(fcf_ratio-3)) if fcf_ratio is not None else None
    capital_score=clamp(90-100*dw) if dw is not None else None
    score=weighted_available([(op_score,.50),(fcf_score,.25),(capital_score,.25)])
    return score,{"net_debt":net,"net_debt_to_operating_income":op_ratio,"net_debt_to_fcf":fcf_ratio,"debt_weight":dw,"components":{"operating_capacity":op_score,"fcf_capacity":fcf_score,"market_capital_structure":capital_score},"status":"Complete" if all(x is not None for x in (op_score,fcf_score,capital_score)) else "Partial"}

def _dcf_reliability(wb,base_value):
    hist=_latest_history(wb); base=num(base_value)
    reasons=[]
    if base is None: reasons.append('Base DCF value is missing.')
    elif base<=0 and (hist.get('operating_income') or 0)>0 and (hist.get('net_income') or 0)>0:
        reasons.append('DCF produces non-positive equity value despite positive latest operating income and net income; forecast cash-flow/capex assumptions require review.')
    # A negative intrinsic value is not impossible, but it is low-confidence evidence for a currently
    # profitable listed company and should not be converted automatically into a 0/100 valuation score.
    return len(reasons)==0,reasons

def _recompute(bundle):
    dims=bundle['dimensions']
    quality=weighted_available([(dims[k]['score'],w) for k,w in QUALITY_WEIGHTS.items()])
    valuation=weighted_available([(dims[k]['score'],w) for k,w in VALUATION_WEIGHTS.items()])
    stress=dims['Stress Robustness']['score']; bayes=dims['Bayesian Skew']['score']
    risk=weighted_available([(stress,.70),(bayes,.30)])
    available=sum(DIMENSION_WEIGHTS[k] for k,d in dims.items() if d['score'] is not None)
    overall=(sum(dims[k]['score']*DIMENSION_WEIGHTS[k] for k in DIMENSION_WEIGHTS if dims[k]['score'] is not None)/available) if available else None
    bundle['category_scores']={'Business Quality':quality,'Valuation / Stock Attractiveness':valuation,'Downside / Scenario Risk':risk,'Overall Investment Score':overall}
    bundle['coverage']=available/100 if available else 0
    bundle['method_version']='Score Engine v3 — reliability-gated, sector-aware single source of truth'
    return bundle


def compute_score_bundle(wb,ticker=None,base_value=None,severe_value=None,current_price=None,mc_prob=None):
    bundle=_v2(wb,ticker=ticker,base_value=base_value,severe_value=severe_value,current_price=current_price,mc_prob=mc_prob)
    dims=bundle['dimensions']

    bal,detail=_capital_structure_score(wb)
    dims['Balance Sheet'].update({
        'score':bal,
        'actual':f"Net debt={detail.get('net_debt')}; net debt/operating income={detail.get('net_debt_to_operating_income')}; net debt/FCF={detail.get('net_debt_to_fcf')}; market debt weight={detail.get('debt_weight')}",
        'benchmark':'Multi-lens capital structure: operating earnings capacity 50%, FCF capacity 25%, market debt weight 25%. Net debt/FCF alone is not authoritative for utilities/industrials.',
        'formula':'50% operating-income leverage score + 25% FCF leverage score + 25% market capital-structure score; available components reweighted. Net cash receives a strong score.',
        'components':str(detail.get('components')),
        'source_cells':'Company Data cash/debt; Historical Financials operating income/FCF; Cost of Capital debt weight',
        'status':detail.get('status','Missing'),
    })

    reliable,reasons=_dcf_reliability(wb,base_value)
    bundle['valuation_model_reliability']={'status':'PASS' if reliable else 'REVIEW','reasons':reasons}
    if not reliable:
        dims['Absolute Valuation'].update({
            'score':None,'status':'Excluded — DCF reliability review',
            'components':'Not scored until the deterministic DCF produces economically interpretable equity value.',
            'formula':'Reliability gate: a non-positive DCF for a currently profitable company is treated as a model-review condition, not automatic 0/100 evidence.',
        })
        # Monte Carlo is generated from the same valuation engine; do not double-count the same broken premise.
        dims['Stress Robustness'].update({
            'score':None,'status':'Excluded — valuation engine review',
            'components':'Stress/Monte Carlo valuation evidence excluded because the base valuation engine failed the reliability gate.',
        })
    # Extreme multi-year CFO volatility can still be informative, but it lowers evidentiary confidence.
    if dims['FCF Quality']['score'] is not None and dims['FCF Quality']['score']<50:
        dims['FCF Quality']['status']='Partial — volatile cash conversion; inspect working-capital/hedging effects'

    return _recompute(bundle)


def advanced_scorecard(wb,current_price,forward_pe,base_value,severe_value):
    bundle=compute_score_bundle(wb,base_value=base_value,severe_value=severe_value,current_price=current_price)
    order=('Growth','Profitability','FCF Quality','Balance Sheet','Absolute Valuation','Relative Valuation','Stress Robustness')
    return [(k,bundle['dimensions'][k]['score'],bundle['dimensions'][k]['formula']+' See Score Audit Trail for inputs and sources.') for k in order]
