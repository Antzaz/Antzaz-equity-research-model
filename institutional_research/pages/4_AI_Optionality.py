from __future__ import annotations

"""Standalone AI optionality page.

The renderer lives under src/ so navigation stays thin and the AI overlay remains isolated from
existing machine-learning models and portfolio-optimizer code.
"""

from src.ai_optionality_dashboard import render


render()
