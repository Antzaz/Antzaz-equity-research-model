from __future__ import annotations

"""Hosted/private research portal with a local fallback.

Hosted mode hydrates the newest encrypted GitHub Actions bundle and exposes Portfolio,
Alpha Analysis, and Company Research. Local mode can use an existing outputs/latest
folder without requiring GitHub credentials, which keeps day-to-day development simple.
"""

import os
from pathlib import Path
import runpy
import shutil

import streamlit as st

from src.live_data import download_latest_live_bundle

BASE = Path(__file__).resolve().parent
ROOT = BASE.parent
CACHE = BASE / ".live_cache"
LOCAL_OUT = BASE / "outputs" / "latest"


def _secret(name: str, default=None):
    """Read Streamlit secrets first, then supported environment-variable fallbacks."""
    try:
        group = st.secrets["live_data"]
        value = group.get(name)
        if value not in (None, ""):
            return value
    except Exception:
        pass

    env_names = {
        "repository": "LIVE_DATA_REPOSITORY",
        "github_token": "GITHUB_TOKEN",
        "bundle_password": "LIVE_BUNDLE_PASSWORD",
        "workflow_file": "LIVE_DATA_WORKFLOW_FILE",
    }
    env_value = os.getenv(env_names.get(name, "")) if name in env_names else None
    return env_value if env_value not in (None, "") else default


def _local_outputs_available() -> bool:
    return LOCAL_OUT.exists() and (LOCAL_OUT / "summary.json").exists()


@st.cache_resource(ttl=3600)
def hydrate() -> dict:
    repository = _secret("repository", "Antzaz/Antzaz-equity-research-model")
    token = _secret("github_token")
    password = _secret("bundle_password")
    workflow = _secret("workflow_file", "daily-portfolio-refresh.yml")

    # Local development should not require GitHub credentials when the user has already
    # generated portfolio outputs on this machine.
    if (not token or not password) and _local_outputs_available():
        return {
            "generated_utc": "local outputs/latest",
            "data_mode": "local",
            "company_research_available": False,
        }

    missing = []
    if not token:
        missing.append("github_token / GITHUB_TOKEN")
    if not password:
        missing.append("bundle_password / LIVE_BUNDLE_PASSWORD")
    if missing:
        raise RuntimeError(
            "Missing live-data credentials: " + ", ".join(missing) + ". "
            "For local Portfolio/Alpha use, run run_research.py first. For hosted use, "
            "configure the [live_data] Streamlit secrets."
        )

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
    manifest["data_mode"] = "encrypted_live_bundle"
    manifest["company_research_available"] = True
    return manifest


def _run_streamlit_script(path: Path):
    """Run an existing Streamlit page after the portal has set page configuration."""
    original = st.set_page_config
    st.set_page_config = lambda *args, **kwargs: None
    try:
        runpy.run_path(str(path), run_name="__main__")
    finally:
        st.set_page_config = original


st.set_page_config(page_title="Private Investment Research Portal", layout="wide")

try:
    manifest = hydrate()
except Exception as exc:
    st.title("Private Investment Research Portal")
    st.error("Could not load research data.")
    st.caption(str(exc))
    st.markdown(
        "**Hosted Streamlit:** add `[live_data]` secrets for `repository`, `github_token`, "
        "`bundle_password`, and optionally `workflow_file`.  \n"
        "**Local Streamlit:** run `python institutional_research/run_research.py` first; "
        "Portfolio and Alpha will then open without GitHub credentials."
    )
    st.stop()

live_mode = manifest.get("data_mode") == "encrypted_live_bundle"
views = ["Portfolio Dashboard", "Alpha Analysis"]
if manifest.get("company_research_available"):
    views.append("Company Research")

view = st.radio(
    "Research view",
    views,
    horizontal=True,
    key="research_portal_view",
)

if live_mode:
    st.caption(
        f"Latest private bundle: {manifest.get('generated_utc', 'unknown')} · "
        "Daily universe is derived from the private portfolio holdings file."
    )
else:
    st.caption(
        "Local mode · using institutional_research/outputs/latest. "
        "Portfolio and Alpha are available without GitHub secrets. Company Research requires the encrypted live bundle."
    )

if view == "Portfolio Dashboard":
    _run_streamlit_script(BASE / "dashboard.py")
elif view == "Alpha Analysis":
    _run_streamlit_script(BASE / "pages" / "1_Alpha_Analysis.py")
else:
    _run_streamlit_script(ROOT / "equity_live_dashboard.py")
