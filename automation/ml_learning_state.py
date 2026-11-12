from __future__ import annotations

"""Carry a consistent, encrypted SQLite learning snapshot across stateless CI jobs.

Fail closed on a missing/corrupt database. Never upload a fresh empty archive on
a failed refresh, silently resetting previously learned outcomes.
"""

import argparse
import os
from pathlib import Path, PurePosixPath
import sqlite3
from tempfile import TemporaryDirectory

import pyzipper


def _password(name: str) -> bytes:
    value = (os.getenv(name) or "").strip()
    if not value:
        raise SystemExit(f"Missing required environment variable: {name}")
    return value.encode("utf-8")


def _verify_database(path: Path) -> None:
    if not path.is_file() or path.stat().st_size == 0:
        raise RuntimeError("Refusing to persist missing or empty ML historical database")
    with sqlite3.connect(f"file:{path.resolve().as_posix()}?mode=ro", uri=True) as con:
        if con.execute("PRAGMA quick_check").fetchone()[0] != "ok":
            raise RuntimeError("ML historical database failed SQLite integrity check")
        tables = {r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        if not {"prices", "features", "predictions"}.issubset(tables):
            raise RuntimeError("ML historical database missing core tables")


def pack(source: Path, output: Path, password_env: str) -> None:
    source = source.resolve()
    db = source / "ml_history.sqlite"
    _verify_database(db)
    output = output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory() as temp_dir:
        staged = Path(temp_dir) / "ml_history.sqlite"
        with sqlite3.connect(db) as origin, sqlite3.connect(staged) as target:
            origin.backup(target)  # Includes any outstanding WAL pages.
        _verify_database(staged)
        with pyzipper.AESZipFile(
            output, "w", compression=pyzipper.ZIP_DEFLATED,
            encryption=pyzipper.WZ_AES,
        ) as zf:
            zf.setpassword(_password(password_env))
            zf.setencryption(pyzipper.WZ_AES, nbits=256)
            zf.write(staged, "ml_data/ml_history.sqlite")
            count = 1
            for path in source.rglob("*"):
                if not path.is_file() or path.is_symlink() or path.resolve() == db:
                    continue
                if path.name.endswith(("-wal", "-shm")):
                    continue
                zf.write(path, (Path("ml_data") / path.relative_to(source)).as_posix())
                count += 1
    print(f"Encrypted consistent ML learning state: {count} file(s) -> {output}")


def unpack(archive: Path, destination: Path, password_env: str) -> None:
    archive = archive.resolve()
    destination = destination.resolve()
    if not archive.is_file():
        raise FileNotFoundError(f"Learning-state archive not found: {archive}")
    with TemporaryDirectory() as temp_dir:
        with pyzipper.AESZipFile(archive, "r") as zf:
            zf.setpassword(_password(password_env))
            for name in zf.namelist():
                parts = PurePosixPath(name).parts
                if not parts or parts[0] != "ml_data" or any(p in (".", "..") for p in parts):
                    raise RuntimeError("Untrusted path in encrypted learning archive")
            zf.extractall(temp_dir)
        restored = Path(temp_dir) / "ml_data"
        _verify_database(restored / "ml_history.sqlite")
        destination.mkdir(parents=True, exist_ok=True)
        target = destination / "ml_data"
        target.mkdir(parents=True, exist_ok=True)
        # The newly checked SQLite snapshot replaces only matching files in
        # this runner. The source artifact is never modified.
        for item in restored.rglob("*"):
            if item.is_file():
                relative = item.relative_to(restored)
                dest = target / relative
                dest.parent.mkdir(parents=True, exist_ok=True)
                item.replace(dest)
    print(f"Restored verified ML learning state from {archive}")


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description="Pack/unpack durable, encrypted ML learning state")
    sub = p.add_subparsers(dest="command", required=True)

    a = sub.add_parser("pack")
    a.add_argument("--source", default="ml_data")
    a.add_argument("--output", required=True)
    a.add_argument("--password-env", default="LIVE_BUNDLE_PASSWORD")

    b = sub.add_parser("unpack")
    b.add_argument("--input", required=True)
    b.add_argument("--destination", default=".")
    b.add_argument("--password-env", default="LIVE_BUNDLE_PASSWORD")

    c = sub.add_parser("verify")
    c.add_argument("--source", default="ml_data")

    args = p.parse_args(argv)
    if args.command == "pack":
        pack(Path(args.source), Path(args.output), args.password_env)
    elif args.command == "unpack":
        unpack(Path(args.input), Path(args.destination), args.password_env)
    else:
        _verify_database(Path(args.source) / "ml_history.sqlite")
        print("ML historical database integrity: PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
