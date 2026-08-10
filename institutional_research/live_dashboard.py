from __future__ import annotations

"""Hosted private research portal.

This entrypoint hydrates the newest encrypted daily GitHub Actions bundle and exposes
Portfolio, Alpha Analysis, and Company Research from one private Streamlit app. It uses
an explicit top-level selector instead of relying on Streamlit's implicit pages/ sidebar,
so Alpha remains accessible even when sidebar page navigation is hidden or unavailable.
"""

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
    try:
        group = st.secrets["live_data"]
        return group.get(name, default)
    except Exception:
        return default


@st.cache_resource(ttl=3600)
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


def _run_streamlit_script(path: Path):
    """Run an existing Streamlit page after the portal has set page configuration.

    Existing standalone pages call st.set_page_config themselves. Within the portal that
    configuration has already been set, so temporarily make child calls a no-op.
    """
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
    st.error("Could not load the latest encrypted daily research bundle.")
    st.caption(str(exc))
    st.info(
        "Configure the live_data secrets for this hosted app and make sure the Daily private portfolio refresh workflow has completed successfully."
    )
    st.stop()

view = st.radio(
    "Research view",
    ["Portfolio Dashboard", "Alpha Analysis", "Company Research"],
    horizontal=True,
    key="research_portal_view",
)

st.caption(
    f"Latest private bundle: {manifest.get('generated_utc', 'unknown')} · "
    "Daily universe is derived from the private portfolio holdings file."
)

if view == "Portfolio Dashboard":
    _run_streamlit_script(BASE / "dashboard.py")
elif view == "Alpha Analysis":
    _run_streamlit_script(BASE / "pages" / "1_Alpha_Analysis.py")
else:
    _run_streamlit_script(ROOT / "equity_live_dashboard.py")
