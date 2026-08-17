from __future__ import annotations

"""Commodity-company valuation v2.

The generic corporate DCF can materially overvalue integrated oil & gas companies when a
peak commodity year or acquisition step-up is treated as secular revenue/margin growth.
This overlay keeps the ordinary model intact for non-commodity companies and changes only
commodity producers.

For configured issuers such as Chevron the framework uses issuer long-range guidance as a
cross-check, not as a forced target. It explicitly separates:
1) near-term acquisition/cycle step-up,
2) normalization,
3) production + price/mix long-run growth,
4) mature/depleting terminal economics.

The WACC overlay is intentionally different from simply replacing CAPM: the calculated
company WACC remains visible, while a commodity-risk valuation floor is applied to the DCF
because historical equity beta alone does not capture commodity-price/cash-flow uncertainty.
"""

import math

import commodity_valuation as legacy

FMT_PCT='0.0%;[Red](0.0%);-'

# Explicit project policy. These are valuation assumptions, not reported facts.
# They can be challenged/edited in the workbook and are exposed in Commodity Valuation.
COMMODITY_POLICY={
    "default":{
        "bear_wacc_floor":.095,
        "base_wacc_floor":.080,
        "bull_wacc_floor":.0725,
        "bear_terminal":.010,
        "base_terminal":.015,
        "bull_terminal":.020,
        "reference_prices":"Bear/Base/Bull commodity normalization; issuer-specific price anchors when available",
    },
    "CVX":{
        "bear_wacc_floor":.095,
        "base_wacc_floor":.080,
        "bull_wacc_floor":.0725,
        "bear_terminal":.010,
        "base_terminal":.015,
        "bull_terminal":.020,
        "brent_bear":55.0,
        "brent_base":70.0,
        "brent_bull":85.0,
        # 2026 contains Hess/full-year integration + current-cycle effects. 2027 normalizes;
        # 2028+ reflects low-single-digit volume plus price/mix rather than a secular 8% CAGR.
        "bear_growth":[.08,-.12,.015,.010,.010,.010,.005,.005,0.0,0.0],
        "base_growth":[.15,-.08,.040,.035,.030,.030,.025,.025,.020,.020],
        "bull_growth":[.22,-.04,.060,.050,.045,.040,.035,.030,.025,.025],
    },
}


def _num(v,default=None):
    try:
        if isinstance(v,bool) or v in (None,""): return default
        x=float(v)
        return x if math.isfinite(x) else default
    except Exception:
        return default


def _policy(ticker):
    base=dict(COMMODITY_POLICY["default"])
    base.update(COMMODITY_POLICY.get(str(ticker).upper(),{}))
    return base


def _set_wacc_and_terminal(wb,ticker,meta):
    if "Three-Case Scenarios" not in wb.sheetnames: return meta
    ws=wb["Three-Case Scenarios"]; p=_policy(ticker)
    raw_base=_num(ws["C6"].value,.09)
    # Preserve a higher company-calculated WACC, but do not let a low beta alone create a
    # sub-8% base discount rate for a cyclical commodity producer.
    base=max(raw_base,p["base_wacc_floor"])
    bear=max(_num(ws["B6"].value,base+.015),p["bear_wacc_floor"],base+.015)
    bull=max(p["bull_wacc_floor"],min(_num(ws["D6"].value,base-.0075),base-.005))
    ws["B6"],ws["C6"],ws["D6"]=bear,base,bull
    ws["B7"],ws["C7"],ws["D7"]=p["bear_terminal"],p["base_terminal"],p["bull_terminal"]
    for cell in ("B6","C6","D6","B7","C7","D7"): ws[cell].number_format=FMT_PCT
    meta.update({
        "raw_calculated_wacc":raw_base,"bear_wacc":bear,"wacc":base,"bull_wacc":bull,
        "bear_terminal_growth":p["bear_terminal"],"terminal_growth":p["base_terminal"],
        "bull_terminal_growth":p["bull_terminal"],"commodity_policy":p,
    })
    return meta


def _set_cvx_operating_path(wb,meta):
    ws=wb["Three-Case Scenarios"]; hist=legacy._historical_cycle(wb); p=_policy("CVX")
    start=hist.get("latest_revenue")
    if not start: return meta
    bear_g=p["bear_growth"]; base_g=p["base_growth"]; bull_g=p["bull_growth"]
    cycle=hist["margin"]
    # Keep a modest 2026 benefit from Hess/current conditions, then normalize toward historical
    # mid-cycle profitability instead of allowing a >20% margin to persist for a decade.
    bear_m=[max(.06,cycle-.025),max(.06,cycle-.030)]+legacy._fade(max(.06,cycle-.025),max(.06,cycle-.020),8)
    base_m=[min(.18,cycle+.025),min(.17,cycle+.015)]+legacy._fade(cycle+.010,cycle,8)
    bull_m=[min(.22,cycle+.055),min(.21,cycle+.045)]+legacy._fade(min(.20,cycle+.035),min(.19,cycle+.025),8)
    base_revs=legacy._project_revenues(start,base_g); bear_revs=legacy._project_revenues(start,bear_g); bull_revs=legacy._project_revenues(start,bull_g)
    guidance=legacy.ISSUER_COMMODITY_GUIDANCE.get("CVX",{})
    if guidance:
        base_cap,base_nom=legacy._capex_ratios_from_nominal(base_revs,guidance["base_capex_2026"],guidance["base_capex_2027_2030"])
        bear_cap,bear_nom=legacy._capex_ratios_from_nominal(bear_revs,guidance["bear_capex_2026"],guidance["bear_capex_2027_2030"])
        bull_cap,bull_nom=legacy._capex_ratios_from_nominal(bull_revs,guidance["bull_capex_2026"],guidance["bull_capex_2027_2030"])
    else:
        normalized=hist["capex"]
        base_cap=[normalized]*10; bear_cap=[min(.25,normalized+.02)]*10; bull_cap=[max(.02,normalized-.015)]*10
        base_nom=[r*c for r,c in zip(base_revs,base_cap)]; bear_nom=[r*c for r,c in zip(bear_revs,bear_cap)]; bull_nom=[r*c for r,c in zip(bull_revs,bull_cap)]

    blocks=[list(range(2,12)),list(range(14,24)),list(range(26,36))]
    paths=[(bear_g,bear_m,bear_cap),(base_g,base_m,base_cap),(bull_g,bull_m,bull_cap)]
    for cols,(growths,margins,capex_rates) in zip(blocks,paths):
        for c,g,m,cp in zip(cols,growths,margins,capex_rates):
            # Long-run D&A is capped near capex intensity so accounting depreciation does not
            # mechanically create implausibly high perpetual FCF when sustaining capex is lower.
            da=min(hist["da"],cp+.015)
            ws.cell(12,c).value=g; ws.cell(14,c).value=m; ws.cell(18,c).value=da; ws.cell(20,c).value=cp
            for r in (12,14,18,20): ws.cell(r,c).number_format=FMT_PCT

    meta.update({
        "hist_margin":cycle,"hist_da":hist["da"],"hist_capex":hist["capex"],"hist_years":hist["years"],
        "forecast_years":list(range(2026,2036)),"bear_growth":bear_g,"base_growth":base_g,"bull_growth":bull_g,
        "bear_margin":bear_m,"base_margin":base_m,"bull_margin":bull_m,
        "bear_capex_rates":bear_cap,"base_capex_rates":base_cap,"bull_capex_rates":bull_cap,
        "base_capex_nominal":base_nom,"bear_capex_nominal":bear_nom,"bull_capex_nominal":bull_nom,
        "brent_bear":p["brent_bear"],"brent_base":p["brent_base"],"brent_bull":p["brent_bull"],
    })
    return meta


def _augment_proof(wb,ticker,meta):
    # Re-render the legacy proof using the final normalized paths, then add the assumptions that
    # explain why the valuation differs from the generic corporate DCF.
    try: legacy._render_proof(wb,ticker,meta)
    except Exception: return
    ws=wb["Commodity Valuation"]
    ws["A18"]="Commodity Risk / Price Normalization"
    for c,v in enumerate(["Assumption","Bear","Base","Bull","Unit","Purpose","Status","Comment"],1): ws.cell(19,c,v)
    p=meta.get("commodity_policy") or _policy(ticker)
    rows=[
        ("Normalized Brent",p.get("brent_bear"),p.get("brent_base"),p.get("brent_bull"),"$/bbl","Cycle anchor for Chevron scenario interpretation","MODEL","Not a spot-price forecast; normalized valuation anchor"),
        ("DCF WACC",meta.get("bear_wacc"),meta.get("wacc"),meta.get("bull_wacc"),"%","Reflect commodity cash-flow uncertainty beyond historical beta","MODEL","Calculated WACC remains disclosed separately"),
        ("Terminal growth",meta.get("bear_terminal_growth"),meta.get("terminal_growth"),meta.get("bull_terminal_growth"),"%","Mature/depleting commodity terminal policy","MODEL","Lower than generic corporate template"),
        ("Calculated CAPM/WACC before overlay",None,meta.get("raw_calculated_wacc"),None,"%","Audit trail","CALCULATED","Commodity floor changes valuation WACC; it does not hide CAPM output"),
    ]
    for r,row in enumerate(rows,20):
        for c,v in enumerate(row,1): ws.cell(r,c,v)
        for c in (2,3,4):
            if row[4]=="%" and isinstance(ws.cell(r,c).value,(int,float)): ws.cell(r,c).number_format=FMT_PCT

    ws["A43"]="Valuation Reconciliation"
    for c,v in enumerate(["Metric","Bear","Base","Bull","Current Price","Interpretation","Source","Status"],1): ws.cell(44,c,v)
    scen=wb["Three-Case Scenarios"]
    rows=[
        ("Intrinsic Value / Share","='Three-Case Scenarios'!B39","='Three-Case Scenarios'!C39","='Three-Case Scenarios'!D39","='Company Data'!B8","Commodity-normalized DCF output","Three-Case Scenarios","MODEL"),
        ("Upside / (Downside)","='Three-Case Scenarios'!B40","='Three-Case Scenarios'!C40","='Three-Case Scenarios'!D40","", "Compare normalized value with market price","Three-Case Scenarios","MODEL"),
    ]
    for r,row in enumerate(rows,45):
        for c,v in enumerate(row,1): ws.cell(r,c,v)
    for c in (2,3,4,5): ws.cell(45,c).number_format='$#,##0.00;[Red]($#,##0.00);-'
    for c in (2,3,4): ws.cell(46,c).number_format=FMT_PCT


def apply_commodity_normalization(wb,ticker,info=None):
    ticker=str(ticker).upper().strip()
    if not legacy.is_commodity_producer(wb,ticker): return {"applied":False,"ticker":ticker}
    # Start from v1 because it provides the generic commodity framework and issuer capex logic.
    meta=legacy.apply_commodity_normalization(wb,ticker,info or {}) or {"applied":True,"ticker":ticker}
    if ticker=="CVX": meta=_set_cvx_operating_path(wb,meta)
    meta=_set_wacc_and_terminal(wb,ticker,meta)
    setattr(wb,"_commodity_valuation",meta)
    _augment_proof(wb,ticker,meta)
    print(
        f"Commodity valuation v2: {ticker}; base WACC={meta.get('wacc',0):.2%}; "
        f"terminal={meta.get('terminal_growth',0):.2%}; "
        f"raw CAPM/WACC={meta.get('raw_calculated_wacc',0):.2%}"
    )
    return meta
