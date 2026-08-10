from __future__ import annotations

"""Download and decrypt the latest private live-data artifact from GitHub Actions.

Designed for Streamlit Cloud or another hosted dashboard. Secrets stay outside the
repository; the repository contains only the generic loader code.
"""

from io import BytesIO
import json
from pathlib import Path
import shutil
import zipfile

try:
    import pyzipper
except ImportError as exc:
    raise ImportError(
        "The private live portal requires pyzipper to decrypt its AES bundle. "
        "From the repository root run: "
        "python -m pip install -r .\\institutional_research\\requirements.txt"
    ) from exc

import requests


def _headers(token: str) -> dict:
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "Antzaz-live-research-dashboard",
    }


def _safe_extract_encrypted(zip_bytes: bytes, password: str, target: Path) -> dict:
    target = target.resolve()
    if target.exists():
        shutil.rmtree(target)
    target.mkdir(parents=True, exist_ok=True)

    with pyzipper.AESZipFile(BytesIO(zip_bytes), "r") as zf:
        zf.setpassword(password.encode("utf-8"))
        for member in zf.infolist():
            dest = (target / member.filename).resolve()
            if target not in dest.parents and dest != target:
                raise ValueError("Unsafe path in encrypted live-data bundle")
        zf.extractall(target)

    manifest_path = target / "manifest.json"
    if not manifest_path.exists():
        raise ValueError("Encrypted live-data bundle is missing manifest.json")
    return json.loads(manifest_path.read_text(encoding="utf-8"))


def download_latest_live_bundle(
    repository: str,
    token: str,
    password: str,
    target: str | Path,
    workflow_file: str = "daily-portfolio-refresh.yml",
    artifact_name: str = "private-live-data",
) -> dict:
    """Hydrate target with the newest successful daily/manual live-data artifact."""
    if not repository or "/" not in repository:
        raise ValueError("repository must use owner/name format")
    if not token:
        raise ValueError("GitHub token is required")
    if not password:
        raise ValueError("Live bundle password is required")

    base = f"https://api.github.com/repos/{repository}"
    h = _headers(token)
    runs_url = f"{base}/actions/workflows/{workflow_file}/runs?status=success&per_page=10"
    runs_resp = requests.get(runs_url, headers=h, timeout=30)
    runs_resp.raise_for_status()
    runs = runs_resp.json().get("workflow_runs", [])
    if not runs:
        raise RuntimeError("No successful daily portfolio refresh run was found")

    chosen = None
    chosen_artifact = None
    for run in runs:
        artifacts_url = f"{base}/actions/runs/{run['id']}/artifacts?per_page=100"
        r = requests.get(artifacts_url, headers=h, timeout=30)
        r.raise_for_status()
        artifacts = r.json().get("artifacts", [])
        matches = [a for a in artifacts if a.get("name") == artifact_name and not a.get("expired")]
        if matches:
            chosen = run
            chosen_artifact = matches[0]
            break
    if chosen_artifact is None:
        raise RuntimeError("No non-expired private-live-data artifact was found")

    download = requests.get(chosen_artifact["archive_download_url"], headers=h, timeout=60)
    download.raise_for_status()

    # upload-artifact wraps the encrypted ZIP in an outer artifact ZIP.
    with zipfile.ZipFile(BytesIO(download.content), "r") as outer:
        candidates = [n for n in outer.namelist() if n.lower().endswith(".zip")]
        if not candidates:
            raise RuntimeError("GitHub artifact does not contain the encrypted live-data ZIP")
        encrypted = outer.read(candidates[0])

    manifest = _safe_extract_encrypted(encrypted, password, Path(target))
    manifest["github_run_id"] = chosen.get("id")
    manifest["github_run_created_at"] = chosen.get("created_at")
    manifest["github_run_updated_at"] = chosen.get("updated_at")
    return manifest
