import sqlite3

import pytest

from automation.ml_learning_state import pack, unpack, _verify_database


def seeded_db(path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(path) as db:
        db.execute("CREATE TABLE prices(symbol TEXT, price REAL)")
        db.execute("CREATE TABLE features(id INTEGER)")
        db.execute("CREATE TABLE predictions(id INTEGER)")
        db.execute("INSERT INTO prices VALUES('SAMPLE', 1.25)")


def test_encrypted_learning_state_round_trip(tmp_path, monkeypatch):
    monkeypatch.setenv("LIVE_BUNDLE_PASSWORD", "test-only-password")
    db = tmp_path / "ml_data" / "ml_history.sqlite"
    seeded_db(db)
    archive = tmp_path / "snapshot.zip"
    pack(db.parent, archive, "LIVE_BUNDLE_PASSWORD")
    restored = tmp_path / "restored"
    unpack(archive, restored, "LIVE_BUNDLE_PASSWORD")
    _verify_database(restored / "ml_data" / "ml_history.sqlite")
    with sqlite3.connect(restored / "ml_data" / "ml_history.sqlite") as con:
        assert con.execute("SELECT price FROM prices").fetchone()[0] == 1.25


def test_missing_database_cannot_overwrite_history(tmp_path, monkeypatch):
    monkeypatch.setenv("LIVE_BUNDLE_PASSWORD", "test-only-password")
    archive = tmp_path / "snapshot.zip"
    with pytest.raises(RuntimeError, match="missing or empty"):
        pack(tmp_path / "not_a_database", archive, "LIVE_BUNDLE_PASSWORD")
    assert not archive.exists()


def test_corrupt_database_rejected(tmp_path):
    db = tmp_path / "ml_data" / "ml_history.sqlite"
    db.parent.mkdir()
    db.write_text("not a sqlite database")
    with pytest.raises(sqlite3.DatabaseError):
        _verify_database(db)
