"""Public imports for the ML research layer.

Presentation and evidence-quality changes are intentionally kept behind the same model API so
normal equity-research searches receive the stronger validation without changing call sites.
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

__all__=[
    "MLResult","ExpectedReturnModel","EarningsSurpriseModel","FinancialAnomalyModel",
    "MarketRegimeModel","AIImpactMLModel","PortfolioPositionSizingModel",
    "AISignalSnapshot","GrowthForecast","LightGBMGrowthForecaster","ai_adjustments","expectations_gap",
]
