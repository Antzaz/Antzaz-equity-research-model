"""Institutional research package initialization.

The portfolio governance patch is optional at package-import time so lightweight utilities
(e.g. reverse DCF/public-data builders) do not require SciPy/scikit-learn. Full offline
portfolio runs install the complete requirements and still activate the governance patch.
"""

try:
    from .continual_patch import install as _install_continual_portfolio_governance
except ModuleNotFoundError as exc:
    if exc.name not in {"scipy", "sklearn"}:
        raise
else:
    _install_continual_portfolio_governance()
