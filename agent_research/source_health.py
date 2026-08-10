"""Deterministic source-registry and bounded reachability checks for research runs."""

from __future__ import annotations

import os
from typing import Any

import requests

from source_registry import (
    consensus_sources,
    issuer_sources,
    professional_framework_sources,
    regulator_macro_sources,
    specialist_sources,
)
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
        specialist = specialist_sources()
        consensus = consensus_sources()
        professional = professional_framework_sources()
        regulator_macro = regulator_macro_sources()

        checks = []
        for label, url in list(issuer.items())[: self.max_issuer_checks]:
            checks.append(self._check(label, url, "issuer-owned"))

        # Ping only a small deterministic subset to keep research runs quick. All registered
        # sources remain in source_health.json even when they are not probed on every run.
        if specialist:
            label, source = next(iter(specialist.items()))
            checks.append(self._check(label, source["url"], source.get("source_type", "specialist")))
        if regulator_macro:
            label, source = next(iter(regulator_macro.items()))
            checks.append(self._check(label, source["url"], source.get("source_type", "regulator/macro")))

        consensus_config = {}
        for key, source in consensus.items():
            env_key = source.get("env_key")
            consensus_config[key] = {
                "provider": source.get("provider"),
                "env_key": env_key,
                "configured": True if not env_key else bool((os.getenv(str(env_key)) or "").strip()),
                "purpose": source.get("purpose"),
                "url": source.get("url"),
            }

        reachable = sum(1 for x in checks if x.get("reachable"))
        payload = {
            "ticker": ctx.ticker,
            "generated_at": utc_now_iso(),
            "registered_issuer_sources": issuer,
            "registered_regulator_macro_sources": regulator_macro,
            "registered_specialist_sources": specialist,
            "registered_consensus_sources": consensus,
            "consensus_provider_configuration": consensus_config,
            "registered_professional_framework_sources": professional,
            "checks": checks,
            "reachable_checks": reachable,
            "important_note": (
                "A failed reachability probe may reflect anti-bot controls or transient network issues; it does not prove the source is invalid. "
                "Optional consensus providers require their listed API key and sometimes paid endpoint access."
            ),
        }
        artifact = ctx.write_json("source_health.json", payload)
        ctx.shared["source_health"] = payload

        configured_consensus = sum(1 for x in consensus_config.values() if x.get("configured"))
        if not issuer:
            status = "WARN"
            summary = (
                f"No explicit issuer source registry is mapped for this ticker; {configured_consensus}/{len(consensus_config)} consensus source(s) "
                "are configured and regulator/fallback logic remains available."
            )
        elif reachable:
            status = "PASS"
            summary = (
                f"{len(issuer)} issuer source(s) registered; {reachable}/{len(checks)} bounded reachability probe(s) succeeded; "
                f"{configured_consensus}/{len(consensus_config)} consensus provider(s) configured/zero-config."
            )
        else:
            status = "WARN"
            summary = (
                f"{len(issuer)} issuer source(s) registered, but bounded probes were blocked/unavailable; "
                f"{configured_consensus}/{len(consensus_config)} consensus provider(s) configured/zero-config."
            )
        return AgentResult(self.name, status, summary, payload, [str(artifact)])
