from __future__ import annotations

"""Zero-configuration public AI-evidence enrichment for uncoded tickers.

Curated issuer packs remain superior. For other companies, recent public company/news metadata can
still prevent an entirely empty AI evidence page. Secondary headlines are deliberately scored
Neutral: they establish *research leads*, not company-reported KPI facts.
"""

from datetime import datetime, timezone
from typing import Any

import yfinance as yf

import ai_effect_analysis

AI_TERMS = (
    "artificial intelligence", "generative ai", "genai", " ai ", "ai-", "machine learning",
    "large language model", "llm", "copilot", "gpu", "accelerator", "data center", "datacenter",
    "automation", "autonomous", "inference", "foundation model",
)


def _contains_ai(text: Any) -> bool:
    low = " " + str(text or "").lower() + " "
    return any(term in low for term in AI_TERMS)


def _news_fields(item: dict) -> tuple[str, str, str, str]:
    content = item.get("content") if isinstance(item.get("content"), dict) else item
    title = str(content.get("title") or item.get("title") or "").strip()
    provider_obj = content.get("provider") if isinstance(content.get("provider"), dict) else {}
    provider = str(provider_obj.get("displayName") or content.get("publisher") or item.get("publisher") or "Public news").strip()
    pub = str(content.get("pubDate") or item.get("providerPublishTime") or "").strip()
    url = ""
    for obj in (content.get("canonicalUrl"), content.get("clickThroughUrl")):
        if isinstance(obj, dict) and obj.get("url"):
            url = str(obj.get("url")); break
    if not url:
        url = str(content.get("link") or item.get("link") or "")
    if pub.isdigit():
        try:
            pub = datetime.fromtimestamp(int(pub), tz=timezone.utc).date().isoformat()
        except Exception:
            pass
    return title, provider, pub[:10], url


def build_public_ai_evidence(ticker: str, limit: int = 5) -> list[tuple]:
    ticker = str(ticker or "").upper().strip(); rows: list[tuple] = []
    try:
        t = yf.Ticker(ticker)
        info = t.info or {}
    except Exception:
        t = None; info = {}

    summary = str(info.get("longBusinessSummary") or "").strip()
    if summary and _contains_ai(summary):
        website = str(info.get("website") or "") or f"https://finance.yahoo.com/quote/{ticker}/profile/"
        snippet = summary[:420].replace("\n", " ")
        rows.append((
            "AI relevance in public company profile", None, "Qualitative company-profile evidence", "Neutral",
            "Public profile mentions AI/automation-related activity. Treat as context until issuer earnings/filings provide a measurable KPI. " + snippet,
            datetime.now(timezone.utc).date().isoformat(), website, "Public company profile / secondary",
        ))

    if t is not None:
        try:
            news = t.news or []
        except Exception:
            news = []
        for item in news:
            if not isinstance(item, dict):
                continue
            title, provider, pub, url = _news_fields(item)
            if not title or not _contains_ai(title):
                continue
            rows.append((
                f"AI-related public news — {provider}", None, "Qualitative public-source research lead", "Neutral",
                title + ". Verify material claims against issuer filings, earnings releases or investor presentations before upgrading the signal.",
                pub or datetime.now(timezone.utc).date().isoformat(), url, "Secondary public news",
            ))
            if len(rows) >= limit:
                break
    return rows[:limit]


def seed_public_ai_evidence(ticker: str) -> dict[str, Any]:
    ticker = str(ticker or "").upper().strip()
    if ticker in ai_effect_analysis.EVIDENCE_PACKS:
        return {"ticker": ticker, "seeded": 0, "mode": "curated pack retained"}
    rows = build_public_ai_evidence(ticker)
    if rows:
        ai_effect_analysis.EVIDENCE_PACKS[ticker] = rows
    return {"ticker": ticker, "seeded": len(rows), "mode": "public qualitative fallback" if rows else "no relevant public AI evidence found"}
