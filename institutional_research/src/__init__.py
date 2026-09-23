"""Institutional research package initialization.

The continual portfolio-governance patch is optional at package-import time. Lightweight
utilities such as reverse DCF/public-data builders do not require SciPy or scikit-learn;
full offline portfolio runs install the complete requirements and activate the patch.
"""

try:
    from .continual_patch import install as _install_continual_portfolio_governance
    _install_continual_portfolio_governance()
except ModuleNotFoundError as exc:
    if exc.name not in {"scipy", "sklearn"}:
        raise
