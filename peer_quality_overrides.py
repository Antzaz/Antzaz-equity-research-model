"""Peer-comparison quality overrides.

Negative P/E and EV/EBITDA ratios are not economically interpretable valuation multiples for
loss-making companies.  They may be mathematically returned by a public provider, but showing
-27x P/E next to profitable peers implies a valuation comparison that does not exist.

This module keeps the underlying peer company and operating metrics while marking non-positive
P/E / EV-EBITDA as not meaningful. Blank cells are then naturally excluded from peer medians and
relative-valuation calculations.
"""

from __future__ import annotations

import dynamic_peer_engine as peers


def _positive_multiple(value):
    v=peers._num(value)
    return v if v is not None and v>0 else None


def _coverage(info):
    values=(
        _positive_multiple(info.get("forwardPE")),
        peers._num(info.get("enterpriseToRevenue")),
        _positive_multiple(info.get("enterpriseToEbitda")),
        peers._num(info.get("revenueGrowth")),
        peers._num(info.get("operatingMargins")),
        peers._num(info.get("returnOnEquity")),
    )
    return sum(v is not None for v in values)/len(values)


def _metric_row(symbol,info,sector,industry,method):
    notes=list(info.get("_metric_notes") or [])
    fpe_raw=peers._num(info.get("forwardPE")); ev_ebitda_raw=peers._num(info.get("enterpriseToEbitda"))
    fpe=_positive_multiple(fpe_raw); ev_ebitda=_positive_multiple(ev_ebitda_raw)
    if fpe_raw is not None and fpe is None:
        notes.append("forward P/E N/M because forward EPS is non-positive")
    if ev_ebitda_raw is not None and ev_ebitda is None:
        notes.append("EV/EBITDA N/M because EBITDA is non-positive")
    source_note="Live Yahoo fields" if not notes else "Live Yahoo + quality/fallback notes: "+"; ".join(dict.fromkeys(notes))
    return [
        info.get("longName") or info.get("shortName") or symbol,
        symbol,
        fpe,
        peers._num(info.get("enterpriseToRevenue")),
        ev_ebitda,
        peers._num(info.get("revenueGrowth")),
        peers._num(info.get("operatingMargins")),
        peers._num(info.get("returnOnEquity")),
        info.get("sector") or sector,
        info.get("industry") or industry,
        method,
        f"https://finance.yahoo.com/quote/{symbol}/",
        None,None,None,
        method,
        _coverage(info),
        source_note,
    ]


def install_peer_quality_overrides():
    peers._coverage=_coverage
    peers._metric_row=_metric_row
    return True
