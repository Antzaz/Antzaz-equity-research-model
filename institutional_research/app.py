from __future__ import annotations

# Explicit Streamlit multipage navigation, with a first-run fallback when no
# local private analysis has been generated. Comments are not rendered in the UI.

from pathlib import Path

import streamlit as st


BASE = Path(__file__).resolve().parent

HAS_ANALYSIS = all(
    (BASE / "outputs" / "latest" / filename).is_file()
    for filename in ("summary.json", "holdings_analysis.csv")
)

PAGES = [
    ("pages/0_Portfolio_Setup.py", "Setup & Holdings", "🗂️", not HAS_ANALYSIS),
    ("dashboard.py", "Portfolio Dashboard", "📊", HAS_ANALYSIS),
    ("pages/1_Alpha_Analysis.py", "Alpha Analysis", "📈", False),
    ("pages/2_Portfolio_Optimization.py", "Portfolio Optimization", "⚖️", False),
    ("pages/3_Model_Learning.py", "Model Learning", "🧠", False),
    ("pages/4_AI_Optionality.py", "AI Optionality & Uncertainty", "🤖", False),
    ("pages/5_Realized_Portfolio.py", "Realized Portfolio", "🧾", False),
    ("pages/6_Decision_Loop.py", "Research Decision Loop", "🔁", False),
    ("pages/7_RF_Weekly_Forecast.py", "Next-Week RF Forecast", "🌲", False),
    ("pages/8_Stock_Prices.py", "Stock Prices & Purchase Costs", "💹", False),
]

missing = [relative for relative, *_ in PAGES if not (BASE / relative).exists()]
if missing:
    st.error("Streamlit navigation is incomplete. Missing page(s): " + ", ".join(missing))
    st.stop()

research_pages = [
    st.Page(
        relative,
        title=title,
        icon=icon,
        default=default,
    )
    for relative, title, icon, default in PAGES
]

navigation = st.navigation({"Institutional Research": research_pages}, position="sidebar", expanded=True)
navigation.run()
