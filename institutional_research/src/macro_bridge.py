from __future__ import annotations

"""Strict read-only contract for sanitized point-in-time Macro Lab context.

No macro regime is allowed to overwrite security valuations, historical returns,
ML labels, optimization weights, or make trading decisions.
"""

from datetime import date, datetime, timezone
import json
import math
from pathlib import Path

VALID_REGIMES = {
    "Goldilocks", "Reflation", "Disinflationary Slowdown", "Stagflation"
}
EXPECTED_FIELDS = {
    "schema_version", "status", "as_of", "generated_utc", "regime",
    "growth_momentum", "inflation_momentum", "methodology_version",
    "source", "interpretation"
}


def validate_macro_context(payload: dict, *, today: date | None = None, max_age_days: int = 95) -> dict:
    """Allowlisted schema prevents accidental portfolio/private-data publication."""
    if not isinstance(payload, dict) or set(payload) != EXPECTED_FIELDS:
        raise ValueError("Unexpected macro bridge schema or non-public fields")
    if payload["schema_version"] != 1 or payload["status"] != "PASS":
        raise ValueError("Macro context is not a validated v1 snapshot")
    if payload["regime"] not in VALID_REGIMES:
        raise ValueError("Unrecognized macro regime")
    today = today or datetime.now(timezone.utc).date()
    as_of = date.fromisoformat(payload["as_of"])
    if as_of > today:
        raise ValueError("Future-dated macro context")
    # Exported observations must be from a completed calendar month.
    if as_of.month == today.month and as_of.year == today.year:
        raise ValueError("Incomplete calendar-month macro context")
    if (today - as_of).days > max_age_days:
        raise ValueError("Stale macro context")
    generated = datetime.fromisoformat(payload["generated_utc"].replace("Z", "+00:00"))
    if generated.tzinfo is None or generated.astimezone(timezone.utc).date() > today:
        raise ValueError("Macro context generated timestamp invalid or in the future")
    for key in ("growth_momentum", "inflation_momentum"):
        if isinstance(payload[key], bool) or not isinstance(payload[key], (int, float)):
            raise ValueError(f"Invalid {key}")
        if not math.isfinite(float(payload[key])):
            raise ValueError(f"Nonfinite {key}")
    for key in ("methodology_version", "source", "interpretation"):
        if not isinstance(payload[key], str) or not payload[key].strip() or len(payload[key]) > 400:
            raise ValueError(f"Invalid {key}")
    return payload.copy()


def load_macro_context(path: Path, *, today: date | None = None) -> dict:
    """Fail open to UNAVAILABLE for optional context, never guess the current regime."""
    try:
        return validate_macro_context(json.loads(Path(path).read_text(encoding="utf-8")), today=today)
    except (OSError, ValueError, KeyError, TypeError, json.JSONDecodeError) as exc:
        return {
            "status": "UNAVAILABLE",
            "reason": f"{type(exc).__name__}: {exc}",
            "source": "Optional Global Macro Lab context; no model values modified",
        }
