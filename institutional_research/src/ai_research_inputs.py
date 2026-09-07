from __future__ import annotations

"""Derive conservative AI-sensitivity suggestions from existing dated research evidence.

The output is an analyst-input proxy, not a statistical beta and not an ML prediction.  It is
used only to keep the standalone AI Optionality page from defaulting to a visually inert all-zero
state when the repository already contains AI KPI evidence for a company.
"""

from pathlib import Path
import json
import re

import numpy as np
import pandas as pd


AI_TERMS = (
    "artificial intelligence", "generative ai", " ai ", "ai-", "ai ", "gemini", "copilot",
    "inference", "gpu", "accelerator", "foundation model", "large language model", "llm",
    "data center", "datacenter", "ai cloud", "cloud ai", "machine learning",
)
SIGNAL_VALUES = {
    "very strong": 0.95,
    "strong": 0.82,
    "positive": 0.70,
    "moderate": 0.58,
    "neutral": 0.50,
    "mixed": 0.45,
    "weak": 0.28,
    "key risk": 0.18,
    "risk": 0.25,
}


def _clip(value, lo=0.0, hi=1.0):
    try:
        x = float(value)
        return float(np.clip(x, lo, hi)) if np.isfinite(x) else lo
    except Exception:
        return lo


def _signal_value(text: str) -> float:
    raw = str(text or "").lower()
    for key, value in SIGNAL_VALUES.items():
        if key in raw:
            return value
    return 0.50


def _ai_relevant(text: str) -> bool:
    raw = " " + re.sub(r"\s+", " ", str(text or "").lower()) + " "
    return any(term in raw for term in AI_TERMS)


def _latest_kpi_rows(path: Path) -> list[dict]:
    if not path.exists():
        return []
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        snapshots = payload.get("snapshots") or []
        if not snapshots:
            return []
        latest = sorted(snapshots, key=lambda x: str(x.get("captured_at") or ""))[-1]
        return [dict(x) for x in (latest.get("kpis") or []) if isinstance(x, dict)]
    except Exception:
        return []


def _source_text(base: Path) -> str:
    source_dir = base / "ai_sources"
    if not source_dir.exists():
        return ""
    parts = []
    for path in sorted(source_dir.rglob("*")):
        if not path.is_file() or path.suffix.lower() not in {".txt", ".md", ".json"}:
            continue
        try:
            text = path.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            continue
        if text.strip():
            parts.append(text[:25000])
    return "\n".join(parts)


def suggested_ai_config(repo_root: Path, tickers: list[str]) -> pd.DataFrame:
    """Return one editable, conservative AI-sensitivity suggestion per ticker.

    Exposure magnitude comes from the density and strength of AI-specific evidence already stored
    in research_data.  Weak and strong signals can both imply material AI sensitivity; direction is
    handled by the scenario distribution, not by this exposure proxy.
    """
    root = Path(repo_root)
    rows = []
    for ticker in dict.fromkeys(str(t).upper().strip() for t in tickers if str(t).strip()):
        base = root / "research_data" / ticker
        kpis = _latest_kpi_rows(base / "kpi_history.json")
        relevant = []
        for row in kpis:
            text = " | ".join(
                str(row.get(k) or "")
                for k in ("kpi", "signal", "investment_read_through", "unit_comparison", "data_type")
            )
            if _ai_relevant(text):
                relevant.append((row, text))

        corpus = _source_text(base)
        corpus_relevant = _ai_relevant(corpus) if corpus else False
        evidence_count = len(relevant) + (1 if corpus_relevant else 0)
        if evidence_count <= 0:
            rows.append({
                "Ticker": ticker,
                "AIExposure": 0.0,
                "EvidenceConfidence": 0.20,
                "MaturityConfidence": 0.20,
                "AISkillConfidence": 0.25,
                "ResearchSource": "No dated AI-specific KPI/source evidence found",
                "ResearchEvidenceCount": 0,
            })
            continue

        signal_values = [_signal_value(row.get("signal")) for row, _ in relevant]
        # Distance from neutral measures economic intensity, regardless of whether the evidence is
        # currently positive or negative. Scenario returns determine direction later.
        intensity = float(np.mean([abs(v - 0.50) * 2.0 for v in signal_values])) if signal_values else 0.35
        coverage = min(1.0, evidence_count / 6.0)
        source_bonus = 0.10 if corpus_relevant else 0.0
        exposure = _clip((0.30 + 1.10 * intensity) * (0.45 + 0.55 * coverage) + source_bonus, 0.0, 1.50)
        evidence_confidence = _clip(0.25 + 0.55 * coverage + source_bonus, 0.20, 0.90)
        maturity_confidence = _clip(0.20 + 0.10 * evidence_count, 0.20, 0.80)

        rows.append({
            "Ticker": ticker,
            "AIExposure": exposure,
            "EvidenceConfidence": evidence_confidence,
            "MaturityConfidence": maturity_confidence,
            "AISkillConfidence": 0.25,
            "ResearchSource": "Existing dated AI KPI/source evidence",
            "ResearchEvidenceCount": int(evidence_count),
        })
    return pd.DataFrame(rows)
