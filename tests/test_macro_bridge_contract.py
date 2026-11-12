from datetime import date
import json

import pytest

from institutional_research.src.macro_bridge import validate_macro_context, load_macro_context, select_macro_context


GOOD = {
    "schema_version": 1, "status": "PASS",
    "as_of": "2026-10-31", "generated_utc": "2026-11-12T08:00:00+00:00",
    "regime": "Goldilocks", "growth_momentum": 0.2,
    "inflation_momentum": -0.1, "methodology_version": "macro-regime-v0.2-pit-expanded",
    "source": "Global Macro Lab; FRED/ALFRED point-in-time initial releases",
    "interpretation": "Descriptive context only"
}


def test_accepts_latest_completed_month():
    assert validate_macro_context(GOOD, today=date(2026, 11, 12))["regime"] == "Goldilocks"


def test_blocks_sensitive_fields_and_future_dates():
    with pytest.raises(ValueError, match="schema"):
        validate_macro_context({**GOOD, "shares": 10}, today=date(2026, 11, 12))
    with pytest.raises(ValueError, match="Future"):
        validate_macro_context({**GOOD, "as_of": "2026-12-31"}, today=date(2026, 11, 12))


def test_missing_or_stale_context_remains_unavailable(tmp_path):
    assert load_macro_context(tmp_path / "missing.json")["status"] == "UNAVAILABLE"
    path = tmp_path / "macro.json"
    path.write_text(json.dumps(GOOD), encoding="utf-8")
    assert load_macro_context(path, today=date(2027, 4, 1))["status"] == "UNAVAILABLE"


def test_prefers_newer_valid_context(tmp_path):
    old = {**GOOD, "as_of": "2026-09-30"}
    a = tmp_path / "offline.json"
    b = tmp_path / "published.json"
    a.write_text(json.dumps(old), encoding="utf-8")
    b.write_text(json.dumps(GOOD), encoding="utf-8")
    assert select_macro_context([a, b], today=date(2026, 11, 12))["as_of"] == "2026-10-31"
