"""Deterministic source-registry and reachability checks for research runs."""

from __future__ import annotations

from typing import Any

import requests

from source_registry import issuer_sources, specialist_sources
from .core import AgentContext, AgentResult, utc_now_iso


class SourceHealthAgent:
    name = "Source Health Agent"

    def __init__(self, timeout: int = 6, max_issuer_checks: int = 4):
        self.timeout = timeout
        self.max_issuer_checks = max_issuer_checks

    def _check(self, label: str, url: str, source_type: str) -> dict[str, Any]:
        result = {"label": label, "url": url, "source_type": source_type, "reachable": False, "status_code": None}
        try:
            response = requests.get(
                url,
                timeout=self.timeout,
                allow_redirects=True,
                stream=True,
                headers={"User-Agent": "Mozilla/5.0 AntzazEquityResearch/1.0"},
            )
            result["status_code"] = response.status_code
            result["resolved_url"] = response.url
            result["reachable"] = response.status_code < 400
            response.close()
        except Exception as exc:
            result["error"] = str(exc)[:300]
        return result

    def run(self, ctx: AgentContext) -> AgentResult:
        issuer = issuer_sources(ctx.ticker)
        checks = []
        for label, url in list(issuer.items())[: self.max_issuer_checks]:
            checks.append(self._check(label, url, "issuer-owned"))

        # Specialist sources are catalogued for auditability, but only one is pinged per run
        # to keep research latency bounded. Their market definitions remain separate from
        # company filings and from each other.
        specialist = specialist_sources()
        if specialist:
            label, source = next(iter(specialist.items()))
            checks.append(self._check(label, source["url"], source.get("source_type", "specialist")))

        reachable = sum(1 for x in checks if x.get("reachable"))
        payload = {
            "ticker": ctx.ticker,
            "generated_at": utc_now_iso(),
            "registered_issuer_sources": issuer,
            "registered_specialist_sources": specialist,
            "checks": checks,
            "reachable_checks": reachable,
            "important_note": "A failed reachability probe may reflect anti-bot controls or transient network issues; it does not prove the source is invalid.",
        }
        artifact = ctx.write_json("source_health.json", payload)
        ctx.shared["source_health"] = payload

        if not issuer:
            status = "WARN"
            summary = "No explicit issuer source registry is mapped for this ticker; SEC/regulatory and existing fallback logic remain available."
        elif reachable:
            status = "PASS"
            summary = f"{len(issuer)} issuer source(s) registered; {reachable}/{len(checks)} bounded reachability probe(s) succeeded."
        else:
            status = "WARN"
            summary = f"{len(issuer)} issuer source(s) registered, but the bounded web probes were blocked or unavailable."
        return AgentResult(self.name, status, summary, payload, [str(artifact)])
