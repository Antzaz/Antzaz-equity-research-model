"""Industry market-share context used by peer research.

Market share is only populated when a public source measures the target and peer on a
comparable industry basis. It is intentionally left blank otherwise. Peer-set market-cap
share is calculated separately by dynamic_peer_engine and should never be confused with
industry market share.
"""

from __future__ import annotations


TRENDFORCE_FOUNDRY_1Q_2026 = "https://www.trendforce.com/presscenter/news/20260612-13095.html"

# 1Q26 global foundry revenue shares reported directly by TrendForce.
FOUNDRY_SHARE_1Q_2026 = {
    "TSM": {"share": 0.720, "basis": "1Q26 global foundry revenue share", "method": "Reported"},
    "2330.TW": {"share": 0.720, "basis": "1Q26 global foundry revenue share", "method": "Reported"},
    "SSNLF": {"share": 0.065, "basis": "1Q26 global foundry revenue share", "method": "Reported"},
    "SMICY": {"share": 0.051, "basis": "1Q26 global foundry revenue share", "method": "Reported"},
    "UMC": {"share": 0.039, "basis": "1Q26 global foundry revenue share", "method": "Reported"},
    "GFS": {"share": 0.033, "basis": "1Q26 global foundry revenue share", "method": "Reported"},
    "TSEM": {"share": 0.008, "basis": "1Q26 global foundry revenue share", "method": "Reported"},
}

# When Yahoo's broad industry label is too wide, use a business-model-specific seed list
# first. Every candidate is still revalidated by dynamic_peer_engine before it is used.
PREFERRED_PEERS = {
    "TSM": ["UMC", "GFS", "TSEM", "SMICY", "SSNLF"],
}


def preferred_peer_symbols(ticker: str, industry: str | None = None) -> list[str]:
    return list(PREFERRED_PEERS.get(str(ticker or "").upper(), []))


def market_share_record(symbol: str) -> dict:
    rec = dict(FOUNDRY_SHARE_1Q_2026.get(str(symbol or "").upper(), {}))
    if rec:
        rec["source"] = TRENDFORCE_FOUNDRY_1Q_2026
        rec["period"] = "1Q26"
    return rec
