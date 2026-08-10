"""Auditable market-position context used by peer and company research.

Two concepts stay deliberately separate:
- ``market_share_record`` is eligible for Peer Comps only when one like-for-like external
  market definition covers the target and its peers.
- ``business_market_share_records`` provides broader business-position snapshots for the
  Dashboard. Those snapshots must not be mixed into peer scoring unless comparability is
  established.
"""

from __future__ import annotations

from source_registry import SPECIALIST_MARKET_SOURCES


# 1Q26 global foundry revenue shares reported by TrendForce. These are comparable within
# the foundry peer set and therefore may be used in Peer Comps.
FOUNDRY_SHARE_1Q_2026 = {
    "TSM": {"share": 0.720, "basis": "Global foundry revenue share", "method": "Reported", "period": "1Q26", "source_key": "foundry"},
    "2330.TW": {"share": 0.720, "basis": "Global foundry revenue share", "method": "Reported", "period": "1Q26", "source_key": "foundry"},
    "SSNLF": {"share": 0.065, "basis": "Global foundry revenue share", "method": "Reported", "period": "1Q26", "source_key": "foundry"},
    "SMICY": {"share": 0.051, "basis": "Global foundry revenue share", "method": "Reported", "period": "1Q26", "source_key": "foundry"},
    "UMC": {"share": 0.039, "basis": "Global foundry revenue share", "method": "Reported", "period": "1Q26", "source_key": "foundry"},
    "GFS": {"share": 0.033, "basis": "Global foundry revenue share", "method": "Reported", "period": "1Q26", "source_key": "foundry"},
    "TSEM": {"share": 0.008, "basis": "Global foundry revenue share", "method": "Reported", "period": "1Q26", "source_key": "foundry"},
}

# Business-position snapshots are useful for company research but are NOT automatically
# comparable with the accounting-industry peer set. Each record carries its own market
# definition, period and source.
BUSINESS_MARKET_SHARE_SNAPSHOTS = {
    "GOOGL": [
        {"metric": "Worldwide search-engine usage share", "share": 0.9127, "period": "Jun-2026", "basis": "Worldwide search-engine usage", "method": "Measured", "source_key": "search_engine"},
        {"metric": "Cloud infrastructure services share", "share": 0.14, "period": "1Q26", "basis": "Worldwide cloud infrastructure services", "method": "Estimated by specialist research", "source_key": "cloud_infrastructure"},
    ],
    "MSFT": [
        {"metric": "Worldwide Bing search-engine usage share", "share": 0.0468, "period": "Jun-2026", "basis": "Worldwide search-engine usage", "method": "Measured", "source_key": "search_engine"},
        {"metric": "Cloud infrastructure services share", "share": 0.21, "period": "1Q26", "basis": "Worldwide cloud infrastructure services", "method": "Estimated by specialist research", "source_key": "cloud_infrastructure"},
    ],
    "AMZN": [
        {"metric": "Cloud infrastructure services share", "share": 0.28, "period": "1Q26", "basis": "Worldwide cloud infrastructure services", "method": "Estimated by specialist research", "source_key": "cloud_infrastructure"},
    ],
    "AAPL": [
        {"metric": "Worldwide smartphone shipment share", "share": 0.210, "period": "1Q26", "basis": "Worldwide smartphone unit shipments", "method": "Estimated by specialist research", "source_key": "smartphone_shipments"},
    ],
}
BUSINESS_MARKET_SHARE_SNAPSHOTS["GOOG"] = BUSINESS_MARKET_SHARE_SNAPSHOTS["GOOGL"]

# Business-model-specific discovery seeds. Every candidate is still revalidated by the
# dynamic peer engine before it enters Peer Comps.
PREFERRED_PEERS = {
    "TSM": ["UMC", "GFS", "TSEM", "SMICY", "SSNLF"],
}


def preferred_peer_symbols(ticker: str, industry: str | None = None) -> list[str]:
    return list(PREFERRED_PEERS.get(str(ticker or "").upper(), []))


def _attach_source(record: dict) -> dict:
    rec = dict(record)
    src = SPECIALIST_MARKET_SOURCES.get(rec.pop("source_key", ""), {})
    if src:
        rec["source"] = src.get("url")
        rec["provider"] = src.get("provider")
        rec["source_type"] = src.get("source_type")
    return rec


def market_share_record(symbol: str) -> dict:
    """Return a peer-comparable industry market-share record, or blank if none is mapped."""
    rec = FOUNDRY_SHARE_1Q_2026.get(str(symbol or "").upper())
    return _attach_source(rec) if rec else {}


def business_market_share_records(symbol: str) -> list[dict]:
    """Return source-scoped company market-position snapshots for dashboard research."""
    symbol = str(symbol or "").upper().strip()
    rows = [_attach_source(x) for x in BUSINESS_MARKET_SHARE_SNAPSHOTS.get(symbol, [])]
    if symbol in {"TSM", "2330.TW"}:
        foundry = market_share_record(symbol)
        if foundry:
            rows.insert(0, {"metric": "Global foundry revenue share", **foundry})
    return rows
