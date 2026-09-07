"""Public imports for the ML research layer.

Presentation and evidence-quality changes are intentionally kept behind the same model API so
normal equity-research searches receive the stronger validation without changing call sites.

The continual-learning hooks are installed here because ml_research.py imports the public model
package before it constructs HistoryStore.  The hooks do not change model mathematics or execute
trades: they close matured forecasts, persist gated training diagnostics and refresh the local
champion/challenger registry whenever the persistent point-in-time store is used.
"""

from .common import MLResult
from .models import (
    ExpectedReturnModel,
    EarningsSurpriseModel,
    FinancialAnomalyModel,
    MarketRegimeModel,
    AIImpactMLModel,
    PortfolioPositionSizingModel,
)
from .ai_growth import (
    AISignalSnapshot,
    GrowthForecast,
    LightGBMGrowthForecaster,
    ai_adjustments,
    expectations_gap,
)
from .continual_learning import install_hooks as _install_continual_learning_hooks
from .short_horizon_patch import install_short_horizon_learning as _install_short_horizon_learning

# Install the base feedback hooks first, then extend only the learning/governance layer with
# research-only 1M/3M/6M return horizons.  The existing 12M model and optimizer contract remain intact.
_install_continual_learning_hooks()
_install_short_horizon_learning()

__all__=[
    "MLResult","ExpectedReturnModel","EarningsSurpriseModel","FinancialAnomalyModel",
    "MarketRegimeModel","AIImpactMLModel","PortfolioPositionSizingModel",
    "AISignalSnapshot","GrowthForecast","LightGBMGrowthForecaster","ai_adjustments","expectations_gap",
]
