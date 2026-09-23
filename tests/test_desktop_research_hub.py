from __future__ import annotations

from pathlib import Path
import sys
import pytest

ROOT=Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0,str(ROOT))

import desktop_research_hub as hub


@pytest.mark.parametrize("raw,expected",[
    ("googl","GOOGL"),
    (" BRK.B ","BRK.B"),
    ("sie.de","SIE.DE"),
    ("rds-a","RDS-A"),
])
def test_normalize_ticker(raw,expected):
    assert hub.normalize_ticker(raw)==expected


@pytest.mark.parametrize("raw",["","ABC DEF","$SPY","ABCDEFGHIJK"])
def test_normalize_ticker_rejects_invalid(raw):
    with pytest.raises(ValueError):
        hub.normalize_ticker(raw)


def test_requirement_fingerprint_changes_with_file_contents(tmp_path):
    a=tmp_path/"a.txt"; b=tmp_path/"b.txt"
    a.write_text("one",encoding="utf-8"); b.write_text("two",encoding="utf-8")
    first=hub.requirement_fingerprint([a,b])
    b.write_text("three",encoding="utf-8")
    second=hub.requirement_fingerprint([a,b])
    assert first!=second
