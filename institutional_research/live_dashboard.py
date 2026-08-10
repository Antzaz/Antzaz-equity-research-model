from __future__ import annotations

"""Hosted entrypoint for the private portfolio dashboard.

Streamlit Cloud should point to this file rather than dashboard.py. On startup it
hydrates outputs/latest from the newest encrypted daily GitHub Actions bundle, then
hands control to the existing dashboard unchanged.
"""

from pathlib import Path
import shutil

import streamlit as st

from src.live_data import download_latest_live_bundle

BASE = Path(__file__).resolve().parent
CACHE = BASE / ".live_cache"
LOCAL_OUT = BASE / "outputs" / "latest"


def _secret(name: str, default=None):
    try:
        group = st.secrets["live_data"]
        return group.get(name, default)
    except Exception:
        return default


def hydrate() -> dict:
    repository = _secret("repository", "Antzaz/Antzaz-equity-research-model")
    token = _secret("github_token")
    password = _secret("bundle_password")
    workflow = _secret("workflow_file", "daily-portfolio-refresh.yml")
    manifest = download_latest_live_bundle(
        repository=repository,
        token=token,
        password=password,
        target=CACHE,
        workflow_file=workflow,
    )
    source = CACHE / "portfolio_outputs"
    if not source.exists():
        raise RuntimeError("Live bundle does not contain portfolio_outputs")
    if LOCAL_OUT.exists():
        shutil.rmtree(LOCAL_OUT)
    LOCAL_OUT.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(source, LOCAL_OUT)
    return manifest


try:
    manifest = hydrate()
except Exception as exc:
    st.set_page_config(page_title="Institutional Portfolio Research", layout="wide")
    st.title("Institutional Portfolio Research Dashboard")
    st.error("Could not load the latest encrypted daily research bundle.")
    st.caption(str(exc))
    st.info(
        "Configure the live_data secrets for this hosted app and make sure the Daily private portfolio refresh workflow has completed successfully."
    )
    st.stop()

# dashboard.py owns the page configuration and all rendering.
import dashboard  # noqa: E402,F401
