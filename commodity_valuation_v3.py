from __future__ import annotations

"""Commodity-company valuation v3 — triangulated cash-flow valuation.

v2 fixed the core Chevron problem by normalizing growth, margins, capex, WACC and terminal
growth. v3 adds an independent equity-cash-flow cross-check so the authoritative commodity
fair value is not determined by one DCF formulation.

For commodity producers:
- Operating FCFF DCF: v2 normalized revenue/margin/capex model, discounted at WACC.
- Equity FCF DCF: median recent OCF-capex, grown on issuer-guidance-aware paths and discounted
  at an equity required return. Net debt is NOT subtracted because reported OCF-capex is
  after financing interest under US GAAP.
- Primary fair value: median of available normalized methods (two methods = transparent
  equal-weight midpoint). Large method divergence is surfaced as REVIEW, not hidden.

This is a model framework, not a recommendation or a commodity-price forecast.
"""

import math
import random
import statistics

import advanced_analytics_v2 as advanced
import commodity_valuation_v2 as v2

FMT_PCT='0.0%;[Red](0.0%);-'
FMT_PRICE='$#,##0.00;[Red]($#,##0.00);-'

_GENERIC_BASE = advanced._base_value
_GENERIC_MC = advanced._monte_carlo

EQUITY_FCF_POLICY = {
    "default": {
        "bear_equity_discount_floor": .100,
        "base_equity_discount_floor": .085,
        "bull_equity_discount_floor": .0775,
        "bear_fcf_growth": [-.10, .00, .02, .02, .02, .02, .015, .01, .01, .01],
        "base_fcf_growth": [.08, .08, .07, .06, .05, .04, .035, .03, .02, .015],
        "bull_fcf_growth": [.12, .12, .10, .09, .08, .06, .05, .04, .03, .02],
    },
    "CVX": {
        # Chevron guides to >10% annual adjusted FCF growth over five years at nominal $70
        # Brent. Base uses 10%, not >10%, and then fades toward mature terminal economics.
        "base_fcf_growth": [.10, .10, .10, .10, .10, .06, .045, .03, .02, .015],
        "bear_fcf_growth": [-.12, .00, .03, .03, .03, .025, .02, .015, .01, .01],
        "bull_fcf_growth": [.15, .15, .15, .15, .15, .08, .06, .045, .03, .02],
    },
}


def _num(v, default=None):
    try:
        if isinstance(v, bool) or v in (None, ""):
            return default
        x = float(v)
        return x if math.isfinite(x) else default
    except Exception:
        return default


def _policy(ticker):
    base = dict(EQUITY_FCF_POLICY["default"])
    base.update(EQUITY_FCF_POLICY.get(str(ticker).upper(), {}))
    return base


def _ticker(wb):
    try:
        return str(wb["Company Data"]["B4"].value or "").upper().strip()
    except Exception:
        return ""


def is_commodity_workbook(wb):
    return v2.legacy.is_commodity_producer(wb, _ticker(wb))


def _historical_fcfs(wb, years=3):
    if "Historical Financials" not in wb.sheetnames:
        return []
    ws = wb["Historical Financials"]
    vals = []
    for c in range(2, min(ws.max_column, 8) + 1):
        year = _num(ws.cell(3, c).value)
        ocf = _num(ws.cell(14, c).value)
        capex = _num(ws.cell(15, c).value)
        if year and ocf is not None and capex is not None:
            fcf = ocf - abs(capex)
            if fcf > 0:
                vals.append((int(year), fcf))
    return vals[-years:]


def normalized_starting_equity_fcf(wb):
    vals = _historical_fcfs(wb, 3)
    if not vals:
        return None, []
    return statistics.median(v for _, v in vals), vals


def _cost_of_equity(wb):
    if "Cost of Capital" not in wb.sheetnames:
        return None
    ws = wb["Cost of Capital"]
    for r in range(1, ws.max_row + 1):
        label = str(ws.cell(r, 1).value or "").strip().lower()
        if label in {"cost of equity", "cost of equity (capm)"}:
            return _num(ws.cell(r, 2).value)
    return None


def equity_fcf_value(wb, scenario="base"):
    ticker = _ticker(wb)
    if not v2.legacy.is_commodity_producer(wb, ticker):
        return None
    start, history = normalized_starting_equity_fcf(wb)
    if start is None or start <= 0:
        return None
    p = _policy(ticker)
    growths = list(p[f"{scenario}_fcf_growth"])
    commodity = getattr(wb, "_commodity_valuation", {}) or {}
    terminal = _num(commodity.get(f"{scenario}_terminal_growth"))
    if terminal is None:
        terminal = _num(commodity.get("terminal_growth"), .015)
    floor = p[f"{scenario}_equity_discount_floor"]
    calculated = _cost_of_equity(wb)
    rate = max(_num(calculated, floor), floor)
    fcf = start
    pv = 0.0
    for i, g in enumerate(growths, 1):
        fcf *= 1.0 + g
        pv += fcf / ((1.0 + rate) ** i)
    terminal = min(max(0.0, terminal), rate - .01)
    terminal_value = fcf * (1.0 + terminal) / (rate - terminal)
    pv += terminal_value / ((1.0 + rate) ** len(growths))
    shares = _num(wb["Company Data"]["B9"].value) if "Company Data" in wb.sheetnames else None
    if not shares or shares <= 0:
        return None
    value = pv / shares
    meta = getattr(wb, "_commodity_valuation", {}) or {}
    meta.update({
        "normalized_equity_fcf_start": start,
        "equity_fcf_history": history,
        f"{scenario}_equity_discount_rate": rate,
        f"{scenario}_equity_fcf_value": value,
        f"{scenario}_equity_fcf_growth": growths,
    })
    setattr(wb, "_commodity_valuation", meta)
    return value


def _update_triangulation_meta(wb, operating, direct, primary):
    meta = getattr(wb, "_commodity_valuation", {}) or {}
    clean = [x for x in (operating, direct) if isinstance(x, (int, float)) and x > 0]
    divergence = None
    if len(clean) >= 2:
        divergence = max(clean) / min(clean) - 1.0
    meta.update({
        "operating_fcff_value": operating,
        "equity_fcf_value": direct,
        "primary_fair_value": primary,
        "method_divergence": divergence,
        "primary_method": "Median of normalized operating FCFF DCF and normalized equity FCF DCF",
    })
    setattr(wb, "_commodity_valuation", meta)
    _write_v3_proof(wb, meta)
    _sync_investment_summary(wb, primary)
    return meta


def primary_fair_value(wb):
    ticker = _ticker(wb)
    if not v2.legacy.is_commodity_producer(wb, ticker):
        return None
    operating = _GENERIC_BASE(wb)
    direct = equity_fcf_value(wb, "base")
    clean = [x for x in (operating, direct) if isinstance(x, (int, float)) and x > 0]
    if not clean:
        return None
    primary = statistics.median(clean)
    _update_triangulation_meta(wb, operating, direct, primary)
    return primary


def commodity_base_value(wb, growth_shock=0, margin_shock=0, capex_shock=0,
                         wacc_shock=0, tgr_shock=0):
    """Dispatch base valuation.

    The unshocked base case uses triangulation. Explicit stress/opportunity shocks remain on the
    normalized operating FCFF engine so shock semantics remain economically coherent.
    """
    if not is_commodity_workbook(wb):
        return _GENERIC_BASE(wb, growth_shock, margin_shock, capex_shock, wacc_shock, tgr_shock)
    if any(abs(float(x)) > 1e-12 for x in (growth_shock, margin_shock, capex_shock, wacc_shock, tgr_shock)):
        return _GENERIC_BASE(wb, growth_shock, margin_shock, capex_shock, wacc_shock, tgr_shock)
    value = primary_fair_value(wb)
    return value if value is not None else _GENERIC_BASE(wb)


def _direct_fcf_mc(wb, n=5000, seed=42):
    ticker = _ticker(wb)
    start, _ = normalized_starting_equity_fcf(wb)
    if not start:
        return []
    p = _policy(ticker)
    base_growth = list(p["base_fcf_growth"])
    commodity = getattr(wb, "_commodity_valuation", {}) or {}
    tgr0 = _num(commodity.get("terminal_growth"), .015)
    rate0 = max(_num(_cost_of_equity(wb), p["base_equity_discount_floor"]),
                p["base_equity_discount_floor"])
    shares = _num(wb["Company Data"]["B9"].value)
    if not shares:
        return []
    rng = random.Random(seed + 1701)
    out = []
    for _ in range(n):
        # Common cycle factor affects near-term FCF growth more than terminal years.
        cycle = rng.gauss(0.0, .025)
        rate = min(.12, max(.075, rate0 + rng.gauss(0.0, .006)))
        tgr = min(rate - .0125, max(0.0, tgr0 + rng.gauss(0.0, .0025)))
        fcf = start * max(.75, min(1.25, 1.0 + rng.gauss(0.0, .07)))
        pv = 0.0
        for i, g in enumerate(base_growth, 1):
            fade = 1.0 - .75 * ((i - 1) / max(1, len(base_growth) - 1))
            gg = max(-.20, min(.20, g + cycle * fade + rng.gauss(0.0, .01)))
            fcf *= 1.0 + gg
            pv += fcf / ((1.0 + rate) ** i)
        tv = fcf * (1.0 + tgr) / (rate - tgr)
        pv += tv / ((1.0 + rate) ** len(base_growth))
        out.append(pv / shares)
    return sorted(out)


def commodity_monte_carlo(wb, n=5000, seed=42):
    if not is_commodity_workbook(wb):
        return _GENERIC_MC(wb, n, seed)
    generic = sorted(_GENERIC_MC(wb, n, seed))
    direct = _direct_fcf_mc(wb, n, seed)
    if not generic:
        return direct
    if not direct:
        return generic
    count = min(len(generic), len(direct))
    # Quantile blend is deliberate: two independently normalized valuation formulations,
    # combined without pretending their simulation shocks are directly comparable.
    return sorted((generic[i] + direct[i]) / 2.0 for i in range(count))


def _sync_investment_summary(wb, primary):
    if primary is None or "Investment Summary" not in wb.sheetnames:
        return
    ws = wb["Investment Summary"]
    for r in range(1, ws.max_row + 1):
        label = str(ws.cell(r, 1).value or "").strip().lower()
        if label in {"base dcf / share", "base dcf fair value", "primary fair value / share"}:
            ws.cell(r, 2).value = primary
            ws.cell(r, 2).number_format = FMT_PRICE


def _write_v3_proof(wb, meta):
    if "Commodity Valuation" not in wb.sheetnames:
        return
    ws = wb["Commodity Valuation"]
    row = 53
    ws.cell(row, 1).value = "Authoritative Commodity Valuation — Triangulation"
    heads = ["Method", "Value / Share", "Discount Rate", "Starting Cash Flow", "Growth Basis",
             "Debt Treatment", "Role", "Audit Note"]
    for c, v in enumerate(heads, 1):
        ws.cell(row + 1, c).value = v
    direct_rate = meta.get("base_equity_discount_rate")
    start = meta.get("normalized_equity_fcf_start")
    rows = [
        ("Normalized Operating FCFF DCF", meta.get("operating_fcff_value"),
         meta.get("wacc"), None, "v2 mid-cycle revenue/margin/capex path",
         "Enterprise value less net debt", "Independent cross-check",
         "Discounted at commodity-normalized WACC."),
        ("Normalized Equity FCF DCF", meta.get("equity_fcf_value"),
         direct_rate, start, "10% through 2030 for CVX, then fade; issuer >10% guidance is not exceeded in base",
         "No net-debt subtraction", "Independent cross-check",
         "OCF-capex is after interest; discount at equity required return to avoid WACC/debt double count."),
        ("Primary Fair Value", meta.get("primary_fair_value"), None, None,
         "Median of available normalized methods", "Method-specific",
         "Authoritative Decision View value",
         "Two-method median equals equal-weight midpoint; method divergence is disclosed."),
        ("Method Divergence", meta.get("method_divergence"), None, None, "",
         "", "Confidence control", "REVIEW above 35%; FAIL above 100%."),
    ]
    for r, values in enumerate(rows, row + 2):
        for c, v in enumerate(values, 1):
            ws.cell(r, c).value = v
        if isinstance(ws.cell(r, 2).value, (int, float)):
            ws.cell(r, 2).number_format = FMT_PCT if values[0] == "Method Divergence" else FMT_PRICE
        if isinstance(ws.cell(r, 3).value, (int, float)):
            ws.cell(r, 3).number_format = FMT_PCT
    ws.cell(row + 7, 1).value = "Equity-FCF Accounting Rule"
    ws.cell(row + 7, 2).value = (
        "Reported OCF minus capex is treated as equity-style cash flow because US GAAP operating "
        "cash flow includes interest paid. It is discounted at an equity required return and net "
        "debt is not subtracted again."
    )
    if f"B{row+7}:H{row+7}" not in {str(x) for x in ws.merged_cells.ranges}:
        ws.merge_cells(start_row=row + 7, start_column=2, end_row=row + 7, end_column=8)


def decorate_decision_and_quality(wb, ticker):
    """Make the final workbook explicit about the authoritative commodity valuation."""
    if not v2.legacy.is_commodity_producer(wb, ticker):
        return
    primary = primary_fair_value(wb)
    meta = getattr(wb, "_commodity_valuation", {}) or {}
    if "Decision View" in wb.sheetnames:
        ws = wb["Decision View"]
        ws["A4"] = (
            "Valuation method: commodity-normalized triangulation — operating FCFF DCF plus "
            "normalized equity FCF DCF. The generic secular-growth DCF is not authoritative."
        )
        for r in range(1, ws.max_row + 1):
            label = str(ws.cell(r, 1).value or "").strip()
            if label == "Base DCF Upside":
                ws.cell(r, 1).value = "Commodity Fair Value Upside"
                if primary is not None:
                    price = _num(wb["Company Data"]["B8"].value)
                    ws.cell(r, 2).value = (primary / price - 1.0) if price else None
                if ws.max_column >= 7:
                    ws.cell(r, 7).value = "Commodity Valuation"
            elif label == "Base DCF Fair Value":
                ws.cell(r, 1).value = "Commodity Triangulated Fair Value"
                ws.cell(r, 2).value = primary
                if ws.max_column >= 7:
                    ws.cell(r, 7).value = "Commodity Valuation"
    if "Data Quality" in wb.sheetnames:
        ws = wb["Data Quality"]
        label = "Commodity valuation triangulation"
        existing = None
        for r in range(1, ws.max_row + 1):
            if str(ws.cell(r, 1).value or "").strip() == label:
                existing = r
                break
        r = existing or (ws.max_row + 1)
        div = _num(meta.get("method_divergence"))
        if primary is None:
            status, detail = "FAIL", "Commodity producer detected but no primary normalized fair value was produced."
        elif div is None:
            status, detail = "REVIEW", f"Primary commodity fair value {primary:.2f}; only one normalized method was available."
        elif div > 1.0:
            status, detail = "FAIL", f"Normalized valuation methods diverge by {div:.1%}; operating={meta.get('operating_fcff_value')}, equity FCF={meta.get('equity_fcf_value')}."
        elif div > .35:
            status, detail = "REVIEW", f"Normalized valuation methods diverge by {div:.1%}; primary={primary:.2f}. Treat fair value as a range, not a point estimate."
        else:
            status, detail = "PASS", f"Operating FCFF and equity-FCF methods are within {div:.1%}; primary fair value={primary:.2f}."
        ws.cell(r, 1).value = label
        ws.cell(r, 2).value = status
        ws.cell(r, 3).value = detail


def apply_commodity_normalization(wb, ticker, info=None):
    """Apply v2 normalization, then attach the independent equity-FCF method."""
    meta = v2.apply_commodity_normalization(wb, ticker, info or {}) or {"applied": False, "ticker": ticker}
    if not meta.get("applied"):
        return meta
    try:
        primary_fair_value(wb)
    except Exception as exc:
        meta = getattr(wb, "_commodity_valuation", meta) or meta
        meta["triangulation_error"] = repr(exc)
        setattr(wb, "_commodity_valuation", meta)
    return getattr(wb, "_commodity_valuation", meta)
