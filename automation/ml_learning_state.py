from __future__ import annotations

"""Encrypt/decrypt the persistent ML learning database for GitHub Actions.

The repository is public and ``ml_data/`` is intentionally gitignored.  Daily hosted runners are
therefore stateless unless the SQLite learning store is carried between runs.  This helper packages
the directory into a password-protected AES zip using the existing LIVE_BUNDLE_PASSWORD secret.

Nothing is committed to Git.  The encrypted archive is intended only for private workflow artifacts.
"""

import argparse
import os
from pathlib import Path
import shutil

import pyzipper


def _password(name: str) -> bytes:
    value = (os.getenv(name) or "").strip()
    if not value:
        raise SystemExit(f"Missing required environment variable: {name}")
    return value.encode("utf-8")


def pack(source: Path, output: Path, password_env: str) -> None:
    source = source.resolve()
    output = output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    if not source.exists():
        source.mkdir(parents=True, exist_ok=True)
    with pyzipper.AESZipFile(
        output,
        "w",
        compression=pyzipper.ZIP_DEFLATED,
        encryption=pyzipper.WZ_AES,
    ) as zf:
        zf.setpassword(_password(password_env))
        zf.setencryption(pyzipper.WZ_AES, nbits=256)
        files = 0
        for path in source.rglob("*"):
            if not path.is_file():
                continue
            if path.name.endswith(("-wal", "-shm")):
                # WAL/SHM files are transient; SQLite is closed before packing in the workflow.
                continue
            arcname = Path(source.name) / path.relative_to(source)
            zf.write(path, arcname.as_posix())
            files += 1
    print(f"Encrypted ML learning state: {files} file(s) -> {output}")


def unpack(archive: Path, destination: Path, password_env: str) -> None:
    archive = archive.resolve()
    destination = destination.resolve()
    if not archive.exists():
        raise SystemExit(f"Learning-state archive not found: {archive}")
    destination.mkdir(parents=True, exist_ok=True)
    with pyzipper.AESZipFile(archive, "r") as zf:
        zf.setpassword(_password(password_env))
        zf.extractall(destination)
    print(f"Restored ML learning state from {archive}")


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description="Pack/unpack encrypted ML learning state")
    sub = p.add_subparsers(dest="command", required=True)

    a = sub.add_parser("pack")
    a.add_argument("--source", default="ml_data")
    a.add_argument("--output", required=True)
    a.add_argument("--password-env", default="LIVE_BUNDLE_PASSWORD")

    b = sub.add_parser("unpack")
    b.add_argument("--input", required=True)
    b.add_argument("--destination", default=".")
    b.add_argument("--password-env", default="LIVE_BUNDLE_PASSWORD")

    args = p.parse_args(argv)
    if args.command == "pack":
        pack(Path(args.source), Path(args.output), args.password_env)
    else:
        unpack(Path(args.input), Path(args.destination), args.password_env)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
