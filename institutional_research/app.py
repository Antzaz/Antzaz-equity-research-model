from __future__ import annotations

"""Explicit multipage launcher for the institutional research Streamlit app.

Using st.navigation avoids relying on Streamlit's implicit pages/ discovery and makes
new research modules visible in a stable, intentional sidebar order.
"""

from pathlib import Path

import streamlit as st


BASE = Path(__file__).resolve().parent

PAGES = [
    ("dashboard.py", "Portfolio Dashboard", "📊", True),
    ("pages/1_Alpha_Analysis.py", "Alpha Analysis", "📈", False),
    ("pages/2_Portfolio_Optimization.py", "Portfolio Optimization", "⚖️", False),
    ("pages/3_Model_Learning.py", "Model Learning", "🧠", False),
    ("pages/4_AI_Optionality.py", "AI Optionality & Uncertainty", "🤖", False),
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
