"""Multi-source analyst-consensus collection for the equity research workbook.

Provider policy
---------------
* Yahoo Finance is the zero-configuration public fallback.
* Financial Modeling Prep (FMP_API_KEY) adds multi-year revenue/EPS/EBIT/EBITDA
  estimates and price-target consensus when the user's plan permits the endpoints.
* Alpha Vantage (ALPHAVANTAGE_API_KEY) adds annual EPS/revenue estimates and
  estimate-revision evidence when available.
* Finnhub (FINNHUB_API_KEY) can add revenue, EPS, EBIT, EBITDA, FCF and capex
  estimates when the user's Estimates subscription permits those endpoints.

Aggregators overlap in their underlying analyst universes, so provider analyst counts are
NEVER summed.  The workbook uses the median of available provider consensus means as a
cross-provider reference and keeps every provider observation in a provenance table.
"""

from __future__ import annotations

import math
import os
import statistics
from datetime import datetime, timezone

import requests
try:
    import yfinance as yf
except Exception:  # pragma: no cover - optional runtime fallback
    yf = None

from currency_normalization import convert_financial_amount_to_quote

FMP_DOC = "https://site.financialmodelingprep.com/developer/docs/stable/financial-estimates"
FMP_PT_DOC = "https://site.financialmodelingprep.com/developer/docs/stable/price-target-consensus"
ALPHA_DOC = "https://www.alphavantage.co/documentation/#earnings-estimates"
FINNHUB_DOC = "https://finnhub.io/docs/api/insider-sentiment"
YAHOO_URL = "https://finance.yahoo.com/"


def _num(value, default=None):
    try:
        if isinstance(value, bool) or value in (None, ""):
            return default
        x = float(value)
        return x if math.isfinite(x) else default
    except Exception:
        return default


def _year(value):
    if isinstance(value, (int, float)):
        return int(value)
    text = str(value or "")
    for token in text.replace("/", "-").split("-"):
        if token.isdigit() and len(token) == 4:
            y = int(token)
            if 1990 <= y <= 2200:
                return y
    return None


def _record(provider, metric, fiscal_year, mean=None, low=None, high=None,
            analysts=None, source_url=None, note=None, raw_currency=None):
    return {
        "provider": provider,
        "metric": metric,
        "fiscal_year": fiscal_year,
        "mean": _num(mean),
        "low": _num(low),
        "high": _num(high),
        "analysts": _num(analysts),
        "source_url": source_url,
        "note": note or "",
        "raw_currency": raw_currency,
        "fetched_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%MZ"),
    }


def _future_years(latest_year, horizon=5):
    if latest_year is None:
        latest_year = datetime.now().year - 1
    return set(range(int(latest_year) + 1, int(latest_year) + horizon + 1))


def _convert_revenue(value, info):
    v = _num(value)
    if v is None:
        return None
    converted = convert_financial_amount_to_quote(v, info or {})
    if converted is None:
        return None
    # Consensus APIs normally return currency units, while workbook financials use billions.
    return converted / 1e9 if abs(converted) > 1e6 else converted


def _safe_get_json(url, params=None, headers=None, timeout=18):
    try:
        r = requests.get(url, params=params, headers=headers or {"User-Agent": "Antzaz Equity Research"}, timeout=timeout)
        r.raise_for_status()
        data = r.json()
        return data, None
    except Exception as exc:
        return None, str(exc)


def yahoo_records(ticker, latest_year, info=None):
    status = {"provider": "Yahoo Finance", "configured": True, "ok": False, "detail": "No usable estimate rows", "source_url": YAHOO_URL}
    if yf is None:
        status["detail"] = "yfinance unavailable"
        return [], status
    records = []
    try:
        t = yf.Ticker(ticker)
        if not info:
            try:
                info = t.info or {}
            except Exception:
                info = {}
        for metric, attr in (("Revenue", "revenue_estimate"), ("EPS", "earnings_estimate")):
            df = getattr(t, attr, None)
            if df is None or getattr(df, "empty", True):
                getter = getattr(t, f"get_{attr}", None)
                if getter:
                    try:
                        df = getter()
                    except Exception:
                        df = None
            if df is None or getattr(df, "empty", True):
                continue
            mapping = [("0y", latest_year + 1 if latest_year else None), ("+1y", latest_year + 2 if latest_year else None)]
            for idx, fy in mapping:
                if fy is None or idx not in df.index:
                    continue
                row = df.loc[idx]
                mean = _num(row.get("avg")); low = _num(row.get("low")); high = _num(row.get("high"))
                if metric == "Revenue":
                    mean = _convert_revenue(mean, info); low = _convert_revenue(low, info); high = _convert_revenue(high, info)
                rec = _record("Yahoo Finance", metric, fy, mean, low, high,
                              row.get("numberOfAnalysts"), YAHOO_URL,
                              "Public yfinance/Yahoo consensus snapshot; provider universe can overlap other aggregators.",
                              (info or {}).get("financialCurrency"))
                if rec["mean"] is not None:
                    records.append(rec)
        # EPS estimate revisions are retained as provider evidence, not blended with level consensus.
        trend = getattr(t, "eps_trend", None)
        if trend is not None and not getattr(trend, "empty", True):
            for idx, fy in (("0y", latest_year + 1 if latest_year else None), ("+1y", latest_year + 2 if latest_year else None)):
                if fy is None or idx not in trend.index:
                    continue
                row = trend.loc[idx]
                cur = _num(row.get("current")); d30 = _num(row.get("30daysAgo")); d90 = _num(row.get("90daysAgo"))
                if cur is not None:
                    rec = _record("Yahoo Finance", "EPS Revision", fy, cur, d30, d90, None, YAHOO_URL,
                                  "mean=current; low=30-days-ago; high=90-days-ago for revision tracking")
                    records.append(rec)
        status["ok"] = any(r["metric"] in {"Revenue", "EPS"} for r in records)
        status["detail"] = f"{len(records)} usable observation(s)" if records else status["detail"]
    except Exception as exc:
        status["detail"] = str(exc)
    return records, status


def fmp_records(ticker, latest_year, info=None):
    key = (os.getenv("FMP_API_KEY") or "").strip()
    status = {"provider": "Financial Modeling Prep", "configured": bool(key), "ok": False,
              "detail": "FMP_API_KEY not configured" if not key else "No usable estimate rows", "source_url": FMP_DOC}
    if not key:
        return [], status, None
    data, err = _safe_get_json(
        "https://financialmodelingprep.com/stable/analyst-estimates",
        {"symbol": ticker, "period": "annual", "page": 0, "limit": 10, "apikey": key},
    )
    if err or not isinstance(data, list):
        status["detail"] = err or f"Unexpected response: {type(data).__name__}"
        return [], status, None
    records = []
    allowed = _future_years(latest_year)
    for row in data:
        if not isinstance(row, dict):
            continue
        fy = _year(row.get("date"))
        if fy not in allowed:
            continue
        rev_mean = _convert_revenue(row.get("revenueAvg"), info)
        rev_low = _convert_revenue(row.get("revenueLow"), info)
        rev_high = _convert_revenue(row.get("revenueHigh"), info)
        if rev_mean is not None:
            records.append(_record("FMP", "Revenue", fy, rev_mean, rev_low, rev_high,
                                   row.get("numAnalystsRevenue"), FMP_DOC,
                                   "FMP Financial Estimates annual consensus"))
        for metric, stem, analysts_key in (
            ("EPS", "eps", "numAnalystsEps"),
            ("EBIT", "ebit", "numAnalystsRevenue"),
            ("EBITDA", "ebitda", "numAnalystsRevenue"),
            ("Net Income", "netIncome", "numAnalystsRevenue"),
        ):
            mean = _num(row.get(stem + "Avg")); low = _num(row.get(stem + "Low")); high = _num(row.get(stem + "High"))
            if metric != "EPS":
                mean = _convert_revenue(mean, info); low = _convert_revenue(low, info); high = _convert_revenue(high, info)
            if mean is not None:
                records.append(_record("FMP", metric, fy, mean, low, high, row.get(analysts_key), FMP_DOC,
                                       "FMP Financial Estimates annual consensus"))
    price_target = None
    pt, pt_err = _safe_get_json("https://financialmodelingprep.com/stable/price-target-consensus", {"symbol": ticker, "apikey": key})
    if not pt_err:
        row = pt[0] if isinstance(pt, list) and pt else pt if isinstance(pt, dict) else None
        if isinstance(row, dict):
            def first(*keys):
                for k in keys:
                    v = _num(row.get(k))
                    if v is not None:
                        return v
                return None
            price_target = {
                "provider": "FMP", "mean": first("targetConsensus", "consensus", "priceTargetAverage", "average"),
                "median": first("targetMedian", "median", "priceTargetMedian"),
                "low": first("targetLow", "low", "priceTargetLow"),
                "high": first("targetHigh", "high", "priceTargetHigh"),
                "source_url": FMP_PT_DOC,
            }
    status["ok"] = bool(records)
    status["detail"] = f"{len(records)} usable observation(s)" if records else status["detail"]
    return records, status, price_target


def alpha_vantage_records(ticker, latest_year, info=None):
    key = (os.getenv("ALPHAVANTAGE_API_KEY") or "").strip()
    status = {"provider": "Alpha Vantage", "configured": bool(key), "ok": False,
              "detail": "ALPHAVANTAGE_API_KEY not configured" if not key else "No usable estimate rows", "source_url": ALPHA_DOC}
    if not key:
        return [], status
    data, err = _safe_get_json("https://www.alphavantage.co/query",
                               {"function": "EARNINGS_ESTIMATES", "symbol": ticker, "apikey": key})
    if err or not isinstance(data, dict):
        status["detail"] = err or "Unexpected response"
        return [], status
    if data.get("Information") or data.get("Note") or data.get("Error Message"):
        status["detail"] = str(data.get("Information") or data.get("Note") or data.get("Error Message"))[:180]
        return [], status
    annual = None
    for key_name, value in data.items():
        if isinstance(value, list) and "annual" in str(key_name).lower():
            annual = value; break
    if annual is None:
        annual = [x for v in data.values() if isinstance(v, list) for x in v if isinstance(x, dict)]
    records = []
    allowed = _future_years(latest_year)
    for row in annual:
        if not isinstance(row, dict):
            continue
        fy = _year(row.get("fiscalDateEnding") or row.get("date") or row.get("fiscalYear") or row.get("year"))
        if fy not in allowed:
            continue
        aliases = {
            "Revenue": ("estimatedRevenueAvg", "estimatedRevenueLow", "estimatedRevenueHigh", "numberAnalystsEstimatedRevenue"),
            "EPS": ("estimatedEPSAvg", "estimatedEPSLow", "estimatedEPSHigh", "numberAnalystsEstimatedEPS"),
        }
        for metric, keys in aliases.items():
            mean = _num(row.get(keys[0])); low = _num(row.get(keys[1])); high = _num(row.get(keys[2])); analysts = _num(row.get(keys[3]))
            # Be tolerant of alternate API field spellings.
            if mean is None:
                stem = "revenue" if metric == "Revenue" else "eps"
                candidates = {str(k).lower(): v for k, v in row.items()}
                for k, v in candidates.items():
                    if stem in k and ("avg" in k or "mean" in k or "consensus" in k):
                        mean = _num(v); break
            if metric == "Revenue":
                mean = _convert_revenue(mean, info); low = _convert_revenue(low, info); high = _convert_revenue(high, info)
            if mean is not None:
                records.append(_record("Alpha Vantage", metric, fy, mean, low, high, analysts, ALPHA_DOC,
                                       "Alpha Vantage EARNINGS_ESTIMATES annual estimate"))
    status["ok"] = bool(records)
    status["detail"] = f"{len(records)} usable observation(s)" if records else status["detail"]
    return records, status


def finnhub_records(ticker, latest_year, info=None):
    key = (os.getenv("FINNHUB_API_KEY") or "").strip()
    status = {"provider": "Finnhub", "configured": bool(key), "ok": False,
              "detail": "FINNHUB_API_KEY not configured" if not key else "No usable estimate rows; Estimates access may be required",
              "source_url": FINNHUB_DOC}
    if not key:
        return [], status
    specs = [
        ("Revenue", "revenue", "revenueAvg", "revenueLow", "revenueHigh"),
        ("EPS", "eps", "epsAvg", "epsLow", "epsHigh"),
        ("EBIT", "ebit", "ebitAvg", "ebitLow", "ebitHigh"),
        ("EBITDA", "ebitda", "ebitdaAvg", "ebitdaLow", "ebitdaHigh"),
        ("FCF", "fcf", "fcfAvg", "fcfLow", "fcfHigh"),
        ("Capex", "capex", "capexAvg", "capexLow", "capexHigh"),
        ("OCF", "ocf", "ocfAvg", "ocfLow", "ocfHigh"),
    ]
    allowed = _future_years(latest_year)
    records = []
    errors = []
    for metric, endpoint, avg_key, low_key, high_key in specs:
        data, err = _safe_get_json(f"https://finnhub.io/api/v1/stock/{endpoint}-estimate",
                                   {"symbol": ticker, "freq": "annual", "token": key}, timeout=12)
        if err:
            errors.append(f"{metric}: {err}"); continue
        rows = data.get("data", []) if isinstance(data, dict) else []
        for row in rows:
            if not isinstance(row, dict):
                continue
            fy = _year(row.get("year") or row.get("period"))
            if fy not in allowed:
                continue
            mean = _num(row.get(avg_key)); low = _num(row.get(low_key)); high = _num(row.get(high_key))
            if metric not in {"EPS"}:
                mean = _convert_revenue(mean, info); low = _convert_revenue(low, info); high = _convert_revenue(high, info)
            if mean is not None:
                records.append(_record("Finnhub", metric, fy, mean, low, high, row.get("numberAnalysts"), FINNHUB_DOC,
                                       f"Finnhub {metric} Estimates annual consensus; subscription access may apply"))
    status["ok"] = bool(records)
    if records:
        status["detail"] = f"{len(records)} usable observation(s)"
    elif errors:
        status["detail"] = "; ".join(errors[:2])[:180]
    return records, status


def collect_consensus(ticker, latest_year, info=None):
    """Return provider observations, provider statuses, blended cross-provider consensus and price targets."""
    observations = []
    statuses = []
    price_targets = []

    recs, st = yahoo_records(ticker, latest_year, info); observations.extend(recs); statuses.append(st)
    recs, st, pt = fmp_records(ticker, latest_year, info); observations.extend(recs); statuses.append(st)
    if pt and any(_num(pt.get(k)) is not None for k in ("mean", "median", "low", "high")):
        price_targets.append(pt)
    recs, st = alpha_vantage_records(ticker, latest_year, info); observations.extend(recs); statuses.append(st)
    recs, st = finnhub_records(ticker, latest_year, info); observations.extend(recs); statuses.append(st)

    grouped = {}
    for rec in observations:
        if rec["metric"] == "EPS Revision" or rec.get("mean") is None:
            continue
        grouped.setdefault((rec["metric"], rec["fiscal_year"]), []).append(rec)

    blended = {}
    for key, rows in grouped.items():
        means = [r["mean"] for r in rows if r["mean"] is not None]
        lows = [r["low"] for r in rows if r["low"] is not None]
        highs = [r["high"] for r in rows if r["high"] is not None]
        counts = [r["analysts"] for r in rows if r["analysts"] is not None]
        if not means:
            continue
        blended[key] = {
            "mean": statistics.median(means),
            "low": min(lows) if lows else None,
            "high": max(highs) if highs else None,
            # Aggregator universes overlap: max is safer than summing.
            "analysts": max(counts) if counts else None,
            "provider_count": len({r["provider"] for r in rows}),
            "providers": ", ".join(sorted({r["provider"] for r in rows})),
            "dispersion": ((max(means) - min(means)) / abs(statistics.median(means))) if len(means) >= 2 and statistics.median(means) else None,
        }

    return {
        "observations": observations,
        "statuses": statuses,
        "blended": blended,
        "price_targets": price_targets,
    }
