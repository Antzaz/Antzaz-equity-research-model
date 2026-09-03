from __future__ import annotations

"""Business-model policy registry for cross-sector equity research.

The project should not force every company through an industrial-company template.  This module
classifies the issuer from ticker/sector/industry/name and returns a conservative research policy.
Explicit ticker overrides are reserved for genuinely unusual structures; ordinary companies are
routed by business model so newly searched tickers inherit the correct safeguards automatically.
"""

from dataclasses import asdict, dataclass
import re
from typing import Any


@dataclass(frozen=True)
class BusinessModelPolicy:
    key: str
    label: str
    statement_profile: str
    primary_valuation: str
    reverse_dcf_allowed: bool
    industrial_fcf_primary: bool
    commodity_normalization: bool = False
    notes: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


POLICIES = {
    "default": BusinessModelPolicy(
        "default", "Standard corporate", "default",
        "DCF + earnings / cash-flow multiples", True, True,
        notes="Conventional operating-company cash-flow analysis is appropriate when statement coverage is reliable.",
    ),
    "digital_platform": BusinessModelPolicy(
        "digital_platform", "Digital platform / internet services", "default",
        "DCF + FCF / earnings multiples", True, True,
    ),
    "software": BusinessModelPolicy(
        "software", "Software / cloud", "default",
        "DCF + EV/FCF + EV/Sales / earnings cross-check", True, True,
    ),
    "semiconductor": BusinessModelPolicy(
        "semiconductor", "Semiconductor / hardware", "default",
        "DCF + normalized earnings / FCF + cycle-aware multiples", True, True,
    ),
    "bank": BusinessModelPolicy(
        "bank", "Bank / deposit-taking financial institution", "bank",
        "P/TBV + ROTCE/ROE + normalized P/E + dividend/capital return", False, False,
        notes="Industrial free cash flow, net debt and enterprise-value DCF are not primary bank valuation concepts.",
    ),
    "insurance": BusinessModelPolicy(
        "insurance", "Insurance / reinsurance", "insurance",
        "P/B + normalized ROE + normalized earnings + float/reserve economics", False, False,
        notes="Claims reserves, investment assets and regulatory capital make industrial FCF/reverse-DCF a poor primary framework.",
    ),
    "insurance_conglomerate": BusinessModelPolicy(
        "insurance_conglomerate", "Insurance-led operating conglomerate", "berkshire",
        "Sum-of-parts + normalized operating earnings + book-value / investment economics", False, False,
    ),
    "reit": BusinessModelPolicy(
        "reit", "Real estate investment trust", "reit",
        "NAV + P/AFFO + P/FFO + property-level cash economics", False, False,
        notes="GAAP depreciation and recurring property capex make industrial FCF/reverse-DCF less decision-useful than AFFO/FFO/NAV.",
    ),
    "commodity": BusinessModelPolicy(
        "commodity", "Commodity / natural-resource producer", "default",
        "Mid-cycle FCF / NAV + EV/EBITDA + commodity-price sensitivity", True, True, True,
        "Peak-cycle spot earnings should not be extrapolated as secular growth.",
    ),
    "utility": BusinessModelPolicy(
        "utility", "Regulated / integrated utility", "default",
        "Rate-base / regulated DCF + P/E + EV/EBITDA", True, True,
        notes="High leverage and capex are often structural; interpret FCF with the regulated investment cycle.",
    ),
    "pharma": BusinessModelPolicy(
        "pharma", "Pharmaceuticals", "default",
        "DCF + normalized P/E + pipeline / patent-cliff scenario analysis", True, True,
    ),
    "biotech": BusinessModelPolicy(
        "biotech", "Biotechnology", "default",
        "Pipeline rNPV / probability-weighted DCF + cash runway", True, True,
        notes="Pre-profit biotech requires probability-weighted clinical/pipeline valuation rather than mechanical historical extrapolation.",
    ),
    "healthcare": BusinessModelPolicy(
        "healthcare", "Healthcare services / devices", "default",
        "DCF + normalized earnings / FCF multiples", True, True,
    ),
    "industrial": BusinessModelPolicy(
        "industrial", "Industrial / capital goods", "default",
        "DCF + EV/EBITDA + normalized earnings / FCF", True, True,
    ),
    "consumer": BusinessModelPolicy(
        "consumer", "Consumer / retail", "default",
        "DCF + earnings / FCF multiples + unit economics", True, True,
    ),
    "payments": BusinessModelPolicy(
        "payments", "Payments / financial technology network", "default",
        "DCF + earnings / FCF multiples", True, True,
        notes="Financial-sector classification alone does not make an asset-light payments network a bank.",
    ),
    "capital_markets": BusinessModelPolicy(
        "capital_markets", "Broker / capital-markets financial institution", "bank",
        "P/B + normalized ROE + earnings + capital / funding analysis", False, False,
        notes="Balance-sheet funding and regulatory capital make industrial enterprise-value DCF a weak primary framework.",
    ),
    "asset_manager": BusinessModelPolicy(
        "asset_manager", "Asset / wealth manager", "default",
        "Earnings / FCF + AUM / fee-rate economics", True, True,
    ),
}


EXPLICIT = {
    "BRK.A": "insurance_conglomerate", "BRK-A": "insurance_conglomerate",
    "BRK.B": "insurance_conglomerate", "BRK-B": "insurance_conglomerate",
    "GOOG": "digital_platform", "GOOGL": "digital_platform",
}


def _text(*values: Any) -> str:
    return " ".join(str(v or "") for v in values).lower()


def _has(text: str, terms: tuple[str, ...]) -> bool:
    return any(term in text for term in terms)


def get_business_model_policy(
    ticker: str,
    sector: str | None = None,
    industry: str | None = None,
    name: str | None = None,
) -> BusinessModelPolicy:
    t = str(ticker or "").upper().strip()
    if t in EXPLICIT:
        return POLICIES[EXPLICIT[t]]

    sec = str(sector or "").strip().lower()
    ind = str(industry or "").strip().lower()
    blob = _text(sector, industry, name)

    # Real estate / REIT must be tested before generic financial-sector rules.
    if sec == "real estate" or _has(blob, ("reit", "real estate investment trust")):
        return POLICIES["reit"]

    if sec == "financial services" or _has(blob, ("banking", "insurance", "brokerage", "asset management")):
        if _has(ind, ("banks", "bank -", "banking", "savings", "mortgage finance")) or re.search(r"\bbank\b", ind):
            return POLICIES["bank"]
        if _has(ind, ("insurance", "reinsurance")) or _has(blob, ("property & casualty insurance", "life insurance", "insurance -")):
            return POLICIES["insurance"]
        if _has(ind, ("capital markets", "financial conglomerates", "broker", "securities")):
            return POLICIES["capital_markets"]
        if _has(ind, ("asset management", "wealth management")):
            return POLICIES["asset_manager"]
        if _has(ind, ("credit services", "payment", "financial data")) or _has(blob, ("payments network", "payment processing")):
            return POLICIES["payments"]

    if sec == "energy" or _has(ind, ("oil & gas", "coal", "uranium", "energy minerals")):
        return POLICIES["commodity"]
    if sec == "basic materials" and _has(ind, ("gold", "copper", "steel", "aluminum", "mining", "metals", "silver")):
        return POLICIES["commodity"]
    if sec == "utilities" or _has(ind, ("utilities -", "regulated electric", "regulated gas", "independent power")):
        return POLICIES["utility"]

    if sec == "technology":
        if _has(ind, ("semiconductor", "computer hardware", "electronic components")):
            return POLICIES["semiconductor"]
        if _has(ind, ("software", "information technology services")):
            return POLICIES["software"]

    if sec == "communication services" and _has(ind, ("internet content", "interactive media", "advertising agencies")):
        return POLICIES["digital_platform"]

    if sec == "healthcare":
        if _has(ind, ("biotechnology", "biotech")):
            return POLICIES["biotech"]
        if _has(ind, ("drug manufacturers", "pharmaceutical", "generic drugs")):
            return POLICIES["pharma"]
        return POLICIES["healthcare"]

    if sec == "industrials":
        return POLICIES["industrial"]
    if sec in {"consumer cyclical", "consumer defensive"}:
        return POLICIES["consumer"]

    return POLICIES["default"]


def workbook_policy(wb, ticker: str | None = None) -> BusinessModelPolicy:
    try:
        ws = wb["Company Data"]
        t = ticker or ws["B4"].value
        return get_business_model_policy(t, ws["B6"].value, ws["B7"].value, ws["B5"].value)
    except Exception:
        return get_business_model_policy(ticker or "")


def reverse_dcf_applicability_message(policy: BusinessModelPolicy) -> str:
    if policy.reverse_dcf_allowed:
        return "Applicable as a diagnostic when cash-flow data and valuation inputs are reliable."
    return f"N/M for {policy.label}: industrial reverse DCF is not a primary framework. Prefer {policy.primary_valuation}."
