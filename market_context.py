"""Industry market-share context used by peer research.

Market share is only populated when a public source measures the target and peer on a
comparable industry basis. It is intentionally left blank otherwise. Peer-set market-cap
share is calculated separately by dynamic_peer_engine and should never be confused with
industry market share.
"""

from __future__ import annotations


TRENDFORCE_FOUNDRY_Q4_2025 = "https://www.trendforce.com/presscenter/news/20260312-12965.html"

# 4Q25 pure-play foundry revenue share. TSMC and Samsung are stated directly by
# TrendForce. Other values are derived from TrendForce's disclosed company revenue divided
# by its disclosed top-10 foundry total (~US$46.3bn), and therefore are marked derived.
FOUNDRY_SHARE_Q4_2025 = {
    "TSM": {"share": 0.704, "basis": "4Q25 global foundry revenue share", "method": "Reported"},
    "2330.TW": {"share": 0.704, "basis": "4Q25 global foundry revenue share", "method": "Reported"},
    "SSNLF": {"share": 0.071, "basis": "4Q25 global foundry revenue share", "method": "Reported"},
    "SMICY": {"share": 2.49 / 46.3, "basis": "4Q25 top-10 foundry revenue share", "method": "Derived from disclosed revenue"},
    "UMC": {"share": 2.00 / 46.3, "basis": "4Q25 top-10 foundry revenue share", "method": "Derived from disclosed revenue"},
    "GFS": {"share": 1.80 / 46.3, "basis": "4Q25 top-10 foundry revenue share", "method": "Derived from disclosed revenue"},
    "TSEM": {"share": 0.44 / 46.3, "basis": "4Q25 top-10 foundry revenue share", "method": "Derived from disclosed revenue"},
}

# When Yahoo's broad industry label is too wide, use a business-model-specific seed list
# first. Every candidate is still revalidated by dynamic_peer_engine before it is used.
PREFERRED_PEERS = {
    "TSM": ["UMC", "GFS", "TSEM", "SMICY", "SSNLF"],
}


def preferred_peer_symbols(ticker: str, industry: str | None = None) -> list[str]:
    return list(PREFERRED_PEERS.get(str(ticker or "").upper(), []))


def market_share_record(symbol: str) -> dict:
    rec = dict(FOUNDRY_SHARE_Q4_2025.get(str(symbol or "").upper(), {}))
    if rec:
        rec["source"] = TRENDFORCE_FOUNDRY_Q4_2025
        rec["period"] = "4Q25"
    return rec
