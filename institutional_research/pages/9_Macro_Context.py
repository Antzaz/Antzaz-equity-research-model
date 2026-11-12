"""Read-only, point-in-time context from the separate Global Macro Lab."""

from pathlib import Path
import json

import streamlit as st

from src.macro_bridge import select_macro_context

st.title("Global Macro Research Context")
st.caption("Connected macroeconomic context; never a stand-alone buy/sell signal.")

base = Path(__file__).resolve().parents[1]
payload = select_macro_context([base / ".offline_cache" / "macro_context.json", base.parent / "research_data" / "macro_context.json"])

if payload.get("status") != "PASS":
    st.info(
        "No recent, validated Global Macro Lab context is cached. "
        "When online, run `python automation/sync_macro_context.py` "
        "from the project root, then reload this page."
    )
else:
    a, b, c = st.columns(3)
    a.metric("Latest completed-month regime", payload["regime"])
    b.metric("Regime as of", payload["as_of"])
    c.metric("Data methodology", payload["methodology_version"])
    st.write(
        "Three-month standardized momentum scores:",
        {"Growth": payload["growth_momentum"], "Inflation": payload["inflation_momentum"]},
    )
    st.caption(payload["source"])

st.warning(
    "Macro context is observational research. It does not alter DCF/WACC inputs, "
    "ML training labels, position sizing, portfolio weights, or trades."
)

snapshot = base / "outputs" / "latest" / "summary.json"
if snapshot.is_file():
    try:
        saved = json.loads(snapshot.read_text(encoding="utf-8")).get("macro_context", {})
        if saved.get("status") == "PASS":
            st.caption(
                f"Portfolio analysis captured regime {saved['regime']} as of "
                f"{saved['as_of']}; refresh the research analysis to incorporate newer context."
            )
    except (OSError, ValueError, KeyError):
        pass
