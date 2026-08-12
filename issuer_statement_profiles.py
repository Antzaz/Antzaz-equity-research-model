from __future__ import annotations

"""Issuer/business-model-specific financial-statement profiles.

The workbook deliberately does not force banks, insurance conglomerates, or IFRS foreign
issuers through the same industrial-company line-item template.  Each profile only maps
reported/public line items; unsupported issuer-specific rows remain blank.
"""

from copy import deepcopy

import full_financial_statements_v2 as generic


def _row(label, yf_names=(), sec_tags=(), unit="money", aliases=()):
    return (label, list(yf_names), list(sec_tags), unit, tuple(aliases))


def _insert_before(rows, target, additions):
    out = []
    inserted = False
    for row in rows:
        if row[0] == target and not inserted:
            out.extend(additions)
            inserted = True
        out.append(row)
    if not inserted:
        out.extend(additions)
    return out


def _replace(rows, label, replacement):
    return [replacement if row[0] == label else row for row in rows]


def _drop(rows, labels):
    labels = set(labels)
    return [row for row in rows if row[0] not in labels]


# Corporate technology profiles retain the standardized three statements but expose operating
# expense lines that are decision-useful and explicitly reported by the issuers/providers.
GOOGL_INCOME = _insert_before(
    deepcopy(generic.INCOME),
    "Total Operating Expenses",
    [
        _row("Sales & Marketing", ["Selling And Marketing Expense", "Sales And Marketing"], ["SellingAndMarketingExpense"], aliases=("Sales & Marketing",)),
        _row("General & Administrative", ["General And Administrative Expense", "General And Administrative"], ["GeneralAndAdministrativeExpense"], aliases=("General & Administrative",)),
    ],
)

AMZN_INCOME = _insert_before(
    deepcopy(generic.INCOME),
    "Total Operating Expenses",
    [
        _row("Fulfillment", ["Fulfillment Expense", "Fulfillment"], ["FulfillmentExpense"], aliases=("Fulfillment",)),
        _row("Technology & Infrastructure", ["Technology And Infrastructure Expense", "Technology And Content Expense"], ["TechnologyAndContentExpense"], aliases=("Technology & Content",)),
        _row("Sales & Marketing", ["Sales And Marketing Expense", "Marketing Expense"], ["MarketingExpense"], aliases=("Sales & Marketing",)),
        _row("General & Administrative", ["General And Administrative Expense"], ["GeneralAndAdministrativeExpense"], aliases=("General & Administrative",)),
    ],
)

NVDA_INCOME = _replace(
    deepcopy(generic.INCOME),
    "Research & Development",
    _row("Research & Development", ["Research And Development"], ["ResearchAndDevelopmentExpense"], aliases=("R&D",)),
)


# IFRS rows use namespace-qualified tags where common IFRS XBRL names are stable. Yahoo annual
# statement aliases remain present because some foreign-filer Company Facts endpoints are sparse.
TSM_INCOME = [
    _row("Revenue", ["Total Revenue", "Operating Revenue"], ["ifrs-full:Revenue"], aliases=("Net Revenue",)),
    _row("Cost of Revenue", ["Cost Of Revenue"], ["ifrs-full:CostOfSales"], aliases=("Cost of Sales",)),
    _row("Gross Profit", ["Gross Profit"], ["ifrs-full:GrossProfit"]),
    _row("Research & Development", ["Research And Development"], ["ifrs-full:ResearchAndDevelopmentExpense"]),
    _row("Selling, General & Administrative", ["Selling General And Administration"], ["ifrs-full:DistributionCosts", "ifrs-full:AdministrativeExpense"]),
    _row("Other Operating Income / (Expense)", ["Other Operating Expenses", "Other Operating Income"], ["ifrs-full:OtherIncome", "ifrs-full:OtherExpenseByFunction"]),
    _row("Total Operating Expenses", ["Operating Expense"], ["ifrs-full:DistributionCosts", "ifrs-full:AdministrativeExpense"]),
    _row("Operating Income", ["Operating Income"], ["ifrs-full:ProfitLossFromOperatingActivities"]),
    _row("Interest Income", ["Interest Income Non Operating"], ["ifrs-full:InterestRevenueExpense"]),
    _row("Interest Expense", ["Interest Expense Non Operating"], ["ifrs-full:FinanceCosts"]),
    _row("Non-Operating Income / (Expense)", ["Other Non Operating Income Expenses", "Total Other Finance Cost"], ["ifrs-full:OtherIncomeExpenseFromNonoperatingActivities"]),
    _row("Pre-Tax Income", ["Pretax Income"], ["ifrs-full:ProfitLossBeforeTax"]),
    _row("Income Taxes", ["Tax Provision"], ["ifrs-full:IncomeTaxExpenseContinuingOperations"]),
    _row("Net Income", ["Net Income", "Net Income Common Stockholders"], ["ifrs-full:ProfitLoss"]),
    _row("Net Income Attributable to Parent", ["Net Income Common Stockholders"], ["ifrs-full:ProfitLossAttributableToOwnersOfParent"]),
    _row("Basic EPS", ["Basic EPS"], ["ifrs-full:BasicEarningsLossPerShare"], "eps"),
    _row("Diluted EPS", ["Diluted EPS"], ["ifrs-full:DilutedEarningsLossPerShare"], "eps"),
    _row("Basic Weighted Average Shares (bn)", ["Basic Average Shares"], ["ifrs-full:WeightedAverageShares"], "shares"),
    _row("Diluted Weighted Average Shares (bn)", ["Diluted Average Shares"], ["ifrs-full:AdjustedWeightedAverageShares"], "shares"),
]

TSM_BALANCE = [
    _row("Cash & Cash Equivalents", ["Cash And Cash Equivalents"], ["ifrs-full:CashAndCashEquivalents"]),
    _row("Financial Assets / Short-Term Investments", ["Other Short Term Investments", "Financial Assets"], ["ifrs-full:CurrentFinancialAssetsAtFairValueThroughProfitOrLoss", "ifrs-full:OtherCurrentFinancialAssets"]),
    _row("Accounts Receivable", ["Accounts Receivable"], ["ifrs-full:TradeAndOtherCurrentReceivables"]),
    _row("Inventory", ["Inventory"], ["ifrs-full:Inventories"]),
    _row("Other Current Assets", ["Other Current Assets"], ["ifrs-full:OtherCurrentAssets"]),
    _row("Total Current Assets", ["Current Assets"], ["ifrs-full:CurrentAssets"]),
    _row("Property & Equipment, Net", ["Net PPE"], ["ifrs-full:PropertyPlantAndEquipment"]),
    _row("Right-of-Use Assets", ["Operating Lease Right Of Use Asset"], ["ifrs-full:RightofuseAssets"]),
    _row("Goodwill", ["Goodwill"], ["ifrs-full:Goodwill"]),
    _row("Other Intangible Assets", ["Other Intangible Assets"], ["ifrs-full:IntangibleAssetsOtherThanGoodwill"]),
    _row("Long-Term Investments / Financial Assets", ["Investments And Other Financial Assets"], ["ifrs-full:NoncurrentFinancialAssetsAtFairValueThroughOtherComprehensiveIncome", "ifrs-full:OtherNoncurrentFinancialAssets"]),
    _row("Other Non-Current Assets", ["Other Non Current Assets"], ["ifrs-full:OtherNoncurrentAssets"]),
    _row("Total Non-Current Assets", ["Total Non Current Assets"], ["ifrs-full:NoncurrentAssets"]),
    _row("Total Assets", ["Total Assets"], ["ifrs-full:Assets"]),
    _row("Accounts Payable", ["Accounts Payable"], ["ifrs-full:TradeAndOtherCurrentPayables"]),
    _row("Current Debt", ["Current Debt", "Current Debt And Capital Lease Obligation"], ["ifrs-full:CurrentBorrowings"]),
    _row("Current Lease Liabilities", ["Current Capital Lease Obligation"], ["ifrs-full:CurrentLeaseLiabilities"]),
    _row("Other Current Liabilities", ["Other Current Liabilities"], ["ifrs-full:OtherCurrentLiabilities"]),
    _row("Total Current Liabilities", ["Current Liabilities"], ["ifrs-full:CurrentLiabilities"]),
    _row("Long-Term Debt", ["Long Term Debt"], ["ifrs-full:NoncurrentBorrowings"]),
    _row("Non-Current Lease Liabilities", ["Long Term Capital Lease Obligation"], ["ifrs-full:NoncurrentLeaseLiabilities"]),
    _row("Deferred Tax Liabilities", ["Trade And Other Payables Non Current"], ["ifrs-full:DeferredTaxLiabilities"]),
    _row("Other Non-Current Liabilities", ["Other Non Current Liabilities"], ["ifrs-full:OtherNoncurrentLiabilities"]),
    _row("Total Non-Current Liabilities", ["Total Non Current Liabilities Net Minority Interest"], ["ifrs-full:NoncurrentLiabilities"]),
    _row("Total Liabilities", ["Total Liabilities Net Minority Interest"], ["ifrs-full:Liabilities"]),
    _row("Share Capital", ["Common Stock"], ["ifrs-full:IssuedCapital"]),
    _row("Additional Paid-In Capital / Capital Surplus", ["Additional Paid In Capital"], ["ifrs-full:SharePremium"]),
    _row("Retained Earnings", ["Retained Earnings"], ["ifrs-full:RetainedEarnings"]),
    _row("Other Reserves / AOCI", ["Gains Losses Not Affecting Retained Earnings"], ["ifrs-full:OtherReserves"]),
    _row("Noncontrolling Interest", ["Minority Interest"], ["ifrs-full:NoncontrollingInterests"]),
    _row("Equity Attributable to Parent", ["Stockholders Equity"], ["ifrs-full:EquityAttributableToOwnersOfParent"]),
    _row("Total Equity", ["Total Equity Gross Minority Interest"], ["ifrs-full:Equity"]),
    _row("Total Liabilities & Equity", [], ["ifrs-full:EquityAndLiabilities"]),
]

TSM_CASH = deepcopy(generic.CASH)

SIE_INCOME = [
    _row("Revenue", ["Total Revenue", "Operating Revenue"], ["ifrs-full:Revenue"]),
    _row("Cost of Revenue", ["Cost Of Revenue"], ["ifrs-full:CostOfSales"]),
    _row("Gross Profit", ["Gross Profit"], ["ifrs-full:GrossProfit"]),
    _row("Research & Development", ["Research And Development"], ["ifrs-full:ResearchAndDevelopmentExpense"]),
    _row("Selling, General & Administrative", ["Selling General And Administration"], ["ifrs-full:DistributionCosts", "ifrs-full:AdministrativeExpense"]),
    _row("Other Operating Income / (Expense)", ["Other Operating Expenses", "Other Operating Income"], ["ifrs-full:OtherIncome", "ifrs-full:OtherExpenseByFunction"]),
    _row("Total Operating Expenses", ["Operating Expense"], []),
    _row("Operating Income", ["Operating Income", "EBIT"], ["ifrs-full:ProfitLossFromOperatingActivities"]),
    _row("Income from Investments / Equity Method", ["Other Non Operating Income Expenses"], ["ifrs-full:ShareOfProfitLossOfAssociatesAndJointVenturesAccountedForUsingEquityMethod"]),
    _row("Interest Income", ["Interest Income Non Operating"], ["ifrs-full:InterestRevenueExpense"]),
    _row("Interest Expense", ["Interest Expense Non Operating"], ["ifrs-full:FinanceCosts"]),
    _row("Pre-Tax Income", ["Pretax Income"], ["ifrs-full:ProfitLossBeforeTax"]),
    _row("Income Taxes", ["Tax Provision"], ["ifrs-full:IncomeTaxExpenseContinuingOperations"]),
    _row("Income from Continuing Operations", ["Net Income Continuous Operations"], ["ifrs-full:ProfitLossFromContinuingOperations"]),
    _row("Discontinued Operations", ["Net Income Discontinuous Operations"], ["ifrs-full:ProfitLossFromDiscontinuedOperations"]),
    _row("Net Income", ["Net Income", "Net Income Common Stockholders"], ["ifrs-full:ProfitLoss"]),
    _row("Net Income Attributable to Parent", ["Net Income Common Stockholders"], ["ifrs-full:ProfitLossAttributableToOwnersOfParent"]),
    _row("Basic EPS", ["Basic EPS"], ["ifrs-full:BasicEarningsLossPerShare"], "eps"),
    _row("Diluted EPS", ["Diluted EPS"], ["ifrs-full:DilutedEarningsLossPerShare"], "eps"),
    _row("Basic Weighted Average Shares (bn)", ["Basic Average Shares"], ["ifrs-full:WeightedAverageShares"], "shares"),
    _row("Diluted Weighted Average Shares (bn)", ["Diluted Average Shares"], ["ifrs-full:AdjustedWeightedAverageShares"], "shares"),
]
SIE_BALANCE = deepcopy(TSM_BALANCE)
SIE_CASH = deepcopy(generic.CASH)


# A bank's income statement and balance sheet are economically different from an industrial
# company's.  In particular, gross profit, inventory and industrial working capital are not the
# appropriate primary structure for JPMorgan.
BANK_INCOME = [
    _row("Interest Income", ["Interest Income", "Interest Income Non Operating"], ["InterestAndDividendIncomeOperating", "InterestIncomeExpenseNonoperatingNet"]),
    _row("Interest Expense", ["Interest Expense", "Interest Expense Non Operating"], ["InterestExpenseNonOperating"]),
    _row("Net Interest Income", ["Net Interest Income"], ["InterestIncomeExpenseNonoperatingNet"]),
    _row("Noninterest Revenue", ["Non Interest Income", "Other Non Interest Income"], ["NoninterestIncome"]),
    _row("Total Net Revenue", ["Total Revenue", "Operating Revenue"], ["Revenues"], aliases=("Revenue", "Net Revenue")),
    _row("Provision for Credit Losses", ["Provision For Loan Losses", "Credit Losses Provision"], ["ProvisionForLoanLeaseAndOtherLosses"]),
    _row("Compensation Expense", ["Compensation And Benefits"], ["LaborAndRelatedExpense"]),
    _row("Occupancy Expense", ["Occupancy And Equipment"], ["OccupancyNet"]),
    _row("Technology / Communications Expense", ["Technology Expense"], ["TechnologyAndCommunicationsExpense"]),
    _row("Professional & Outside Services", ["Professional Expense And Contract Services Expense"], ["ProfessionalAndContractServicesExpense"]),
    _row("Marketing Expense", ["Marketing Expense"], ["MarketingExpense"]),
    _row("Other Noninterest Expense", ["Other Non Interest Expense"], ["OtherNoninterestExpense"]),
    _row("Total Noninterest Expense", ["Non Interest Expense", "Operating Expense"], ["NoninterestExpense"]),
    _row("Income Before Tax", ["Pretax Income"], ["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest"]),
    _row("Income Taxes", ["Tax Provision"], ["IncomeTaxExpenseBenefit"]),
    _row("Net Income", ["Net Income", "Net Income Common Stockholders"], ["NetIncomeLoss"]),
    _row("Preferred Dividends", ["Preferred Stock Dividends"], ["PreferredStockDividendsAndOtherAdjustments"]),
    _row("Net Income Available to Common", ["Net Income Common Stockholders"], ["NetIncomeLossAvailableToCommonStockholdersBasic"]),
    _row("Basic EPS", ["Basic EPS"], ["EarningsPerShareBasic"], "eps"),
    _row("Diluted EPS", ["Diluted EPS"], ["EarningsPerShareDiluted"], "eps"),
    _row("Basic Weighted Average Shares (bn)", ["Basic Average Shares"], ["WeightedAverageNumberOfSharesOutstandingBasic"], "shares"),
    _row("Diluted Weighted Average Shares (bn)", ["Diluted Average Shares"], ["WeightedAverageNumberOfDilutedSharesOutstanding"], "shares"),
]

BANK_BALANCE = [
    _row("Cash & Due from Banks", ["Cash Cash Equivalents And Federal Funds Sold", "Cash And Cash Equivalents"], ["CashAndDueFromBanks"]),
    _row("Deposits with Banks / Fed Funds Sold", ["Federal Funds Sold And Securities Purchase Under Agreements To Resell"], ["FederalFundsSoldAndSecuritiesPurchasedUnderAgreementsToResell"]),
    _row("Trading Assets", ["Trading Securities", "Trading Assets"], ["TradingAssets"]),
    _row("Investment Securities", ["Investment Securities", "Available For Sale Securities"], ["InvestmentSecurities"]),
    _row("Loans, Gross", ["Gross Loans", "Loans Receivable"], ["LoansAndLeasesReceivableBeforeAllowance"]),
    _row("Allowance for Credit Losses", ["Allowance For Doubtful Accounts Receivable"], ["AllowanceForLoanAndLeaseLosses"]),
    _row("Loans, Net", ["Net Loan", "Loans Receivable"], ["LoansAndLeasesReceivableNetReportedAmount"]),
    _row("Accrued Interest & Receivables", ["Accounts Receivable"], ["AccountsNotesAndLoansReceivableNetCurrent"]),
    _row("Premises & Equipment, Net", ["Net PPE"], ["PropertyPlantAndEquipmentNet"]),
    _row("Goodwill", ["Goodwill"], ["Goodwill"]),
    _row("Other Intangible Assets", ["Other Intangible Assets"], ["FiniteLivedIntangibleAssetsNet"]),
    _row("Other Assets", ["Other Assets"], ["OtherAssetsNoncurrent"]),
    _row("Total Assets", ["Total Assets"], ["Assets"]),
    _row("Deposits", ["Total Deposits", "Deposits"], ["Deposits"]),
    _row("Federal Funds Purchased / Repos", ["Federal Funds Purchased And Securities Sold Under Agreement To Repurchase"], ["FederalFundsPurchasedAndSecuritiesSoldUnderAgreementsToRepurchase"]),
    _row("Trading Liabilities", ["Trading Liabilities"], ["TradingLiabilities"]),
    _row("Accounts Payable / Other Liabilities", ["Payables And Accrued Expenses", "Other Liabilities"], ["AccountsPayableAndOtherAccruedLiabilitiesCurrent"]),
    _row("Short-Term Borrowings", ["Current Debt", "Other Short Term Borrowings"], ["ShortTermBorrowings"]),
    _row("Long-Term Debt", ["Long Term Debt"], ["LongTermDebt"]),
    _row("Other Non-Current Liabilities", ["Other Non Current Liabilities"], ["OtherLiabilitiesNoncurrent"]),
    _row("Total Liabilities", ["Total Liabilities Net Minority Interest"], ["Liabilities"]),
    _row("Preferred Stock", ["Preferred Stock"], ["PreferredStocksIncludingAdditionalPaidInCapital"]),
    _row("Common Stock", ["Common Stock"], ["CommonStocksIncludingAdditionalPaidInCapital"]),
    _row("Additional Paid-In Capital", ["Additional Paid In Capital"], ["AdditionalPaidInCapital"]),
    _row("Retained Earnings", ["Retained Earnings"], ["RetainedEarningsAccumulatedDeficit"]),
    _row("Accumulated Other Comprehensive Income", ["Gains Losses Not Affecting Retained Earnings"], ["AccumulatedOtherComprehensiveIncomeLossNetOfTax"]),
    _row("Treasury Stock", ["Treasury Stock"], ["TreasuryStockValue"]),
    _row("Common Stockholders' Equity", ["Stockholders Equity"], ["StockholdersEquity"]),
    _row("Total Stockholders' Equity", ["Total Equity Gross Minority Interest"], ["StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"]),
    _row("Total Liabilities & Equity", [], ["LiabilitiesAndStockholdersEquity"]),
]
BANK_CASH = _insert_before(
    deepcopy(generic.CASH),
    "Operating Cash Flow",
    [
        _row("Provision for Credit Losses", ["Provision For Loan Losses"], ["ProvisionForLoanLeaseAndOtherLosses"]),
        _row("Change in Trading Assets / Liabilities", ["Change In Trading Asset", "Change In Trading Liabilities"], []),
        _row("Change in Loans", ["Change In Loans"], []),
        _row("Change in Deposits", ["Change In Deposits"], []),
    ],
)


BRK_INCOME = [
    _row("Insurance Premiums Earned", ["Insurance Revenue", "Premiums Earned"], ["InsurancePremiumsEarned"]),
    _row("Sales & Service Revenues", ["Sales And Service Revenue"], ["SalesRevenueNet"]),
    _row("Leasing Revenues", ["Lease Income"], ["OperatingLeaseIncome"]),
    _row("Interest, Dividend & Investment Income", ["Interest Income", "Investment Income"], ["InvestmentIncomeInterestAndDividend"]),
    _row("Insurance & Other Revenue", ["Total Revenue"], ["Revenues"]),
    _row("Railroad Revenue", ["Railroad Revenue"], []),
    _row("Utilities & Energy Revenue", ["Utility Revenue"], ["RegulatedAndUnregulatedOperatingRevenue"]),
    _row("Total Revenues", ["Total Revenue", "Operating Revenue"], ["Revenues"], aliases=("Revenue",)),
    _row("Insurance Losses & Loss Adjustment Expenses", ["Policyholder Benefits And Claims Payable"], ["PolicyholderBenefitsAndClaimsPayable"]),
    _row("Insurance Underwriting Expenses", ["Insurance And Claims"], []),
    _row("Cost of Sales & Services", ["Cost Of Revenue"], ["CostOfGoodsAndServicesSold"]),
    _row("Railroad Operating Expenses", ["Railroad Operating Expenses"], []),
    _row("Utilities & Energy Operating Expenses", ["Utility Operating Expense"], []),
    _row("Selling, General & Administrative", ["Selling General And Administration"], ["SellingGeneralAndAdministrativeExpense"]),
    _row("Total Operating Expenses", ["Operating Expense"], ["OperatingExpenses"]),
    _row("Operating Earnings / Income", ["Operating Income"], ["OperatingIncomeLoss"]),
    _row("Investment Gains / (Losses)", ["Other Non Operating Income Expenses"], ["InvestmentGainsLosses"]),
    _row("Equity Method Earnings / (Losses)", ["Otherunder Preferred Stock Dividend"], ["IncomeLossFromEquityMethodInvestments"]),
    _row("Pre-Tax Income", ["Pretax Income"], ["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest"]),
    _row("Income Taxes", ["Tax Provision"], ["IncomeTaxExpenseBenefit"]),
    _row("Net Earnings", ["Net Income"], ["NetIncomeLoss"]),
    _row("Noncontrolling Interests", ["Minority Interest"], ["NetIncomeLossAttributableToNoncontrollingInterest"]),
    _row("Net Earnings Attributable to Berkshire", ["Net Income Common Stockholders"], ["NetIncomeLossAvailableToCommonStockholdersBasic"]),
    _row("Basic EPS", ["Basic EPS"], ["EarningsPerShareBasic"], "eps"),
    _row("Diluted EPS", ["Diluted EPS"], ["EarningsPerShareDiluted"], "eps"),
    _row("Basic Weighted Average Shares (bn)", ["Basic Average Shares"], ["WeightedAverageNumberOfSharesOutstandingBasic"], "shares"),
    _row("Diluted Weighted Average Shares (bn)", ["Diluted Average Shares"], ["WeightedAverageNumberOfDilutedSharesOutstanding"], "shares"),
]

BRK_BALANCE = [
    _row("Cash & Cash Equivalents", ["Cash And Cash Equivalents"], ["CashAndCashEquivalentsAtCarryingValue"]),
    _row("U.S. Treasury Bills / Short-Term Investments", ["Other Short Term Investments"], ["ShortTermInvestments"]),
    _row("Fixed-Maturity Securities", ["Available For Sale Securities"], ["AvailableForSaleSecuritiesDebtSecurities"]),
    _row("Equity Securities", ["Equity Securities"], ["EquitySecuritiesFvNi"]),
    _row("Loans & Finance Receivables", ["Accounts Receivable"], ["FinancingReceivableExcludingAccruedInterestAfterAllowanceForCreditLoss"]),
    _row("Inventory", ["Inventory"], ["InventoryNet"]),
    _row("Other Current / Operating Assets", ["Other Current Assets"], ["OtherCurrentAssets"]),
    _row("Property & Equipment, Net", ["Net PPE"], ["PropertyPlantAndEquipmentNet"]),
    _row("Railroad, Utilities & Energy PPE", ["Net PPE"], []),
    _row("Goodwill", ["Goodwill"], ["Goodwill"]),
    _row("Other Intangible Assets", ["Other Intangible Assets"], ["FiniteLivedIntangibleAssetsNet"]),
    _row("Other Assets", ["Other Non Current Assets"], ["OtherAssetsNoncurrent"]),
    _row("Total Assets", ["Total Assets"], ["Assets"]),
    _row("Insurance Losses & Loss Adjustment Liabilities", ["Policyholder Benefits And Claims Payable"], ["PolicyholderBenefitsAndClaimsPayable"]),
    _row("Unearned Premiums", ["Current Deferred Revenue"], ["UnearnedPremiums"]),
    _row("Accounts Payable & Accrued Liabilities", ["Payables And Accrued Expenses"], ["AccountsPayableAndOtherAccruedLiabilitiesCurrent"]),
    _row("Short-Term Debt", ["Current Debt"], ["ShortTermBorrowings"]),
    _row("Long-Term Debt / Borrowings", ["Long Term Debt"], ["LongTermDebt"]),
    _row("Deferred Tax Liabilities", ["Trade And Other Payables Non Current"], ["DeferredTaxLiabilitiesNoncurrent"]),
    _row("Other Liabilities", ["Other Non Current Liabilities"], ["OtherLiabilitiesNoncurrent"]),
    _row("Total Liabilities", ["Total Liabilities Net Minority Interest"], ["Liabilities"]),
    _row("Common Stock", ["Common Stock"], ["CommonStocksIncludingAdditionalPaidInCapital"]),
    _row("Additional Paid-In Capital", ["Additional Paid In Capital"], ["AdditionalPaidInCapital"]),
    _row("Retained Earnings", ["Retained Earnings"], ["RetainedEarningsAccumulatedDeficit"]),
    _row("Accumulated Other Comprehensive Income", ["Gains Losses Not Affecting Retained Earnings"], ["AccumulatedOtherComprehensiveIncomeLossNetOfTax"]),
    _row("Treasury Stock", ["Treasury Stock"], ["TreasuryStockValue"]),
    _row("Berkshire Shareholders' Equity", ["Stockholders Equity"], ["StockholdersEquity"]),
    _row("Noncontrolling Interests", ["Minority Interest"], ["MinorityInterest"]),
    _row("Total Equity", ["Total Equity Gross Minority Interest"], ["StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"]),
    _row("Total Liabilities & Equity", [], ["LiabilitiesAndStockholdersEquity"]),
]
BRK_CASH = _insert_before(
    deepcopy(generic.CASH),
    "Operating Cash Flow",
    [
        _row("Investment Gains / Losses (Non-Cash Adjustment)", ["Operating Gains Losses"], ["GainLossOnInvestments"]),
        _row("Equity-Method / Impairment Adjustments", ["Other Non Cash Items"], ["IncomeLossFromEquityMethodInvestments"]),
    ],
)


PROFILES = {
    "default": {
        "name": "Standard corporate (US GAAP / structured provider)",
        "income": deepcopy(generic.INCOME), "balance": deepcopy(generic.BALANCE), "cash": deepcopy(generic.CASH),
        "canonical_revenue": "Revenue", "derive_fcf": True, "balance_net_debt": True,
        "min_structure": (20, 35, 25), "min_mapped": (12, 18, 14),
    },
    "google": {
        "name": "Alphabet / technology corporate", "income": GOOGL_INCOME,
        "balance": deepcopy(generic.BALANCE), "cash": deepcopy(generic.CASH),
        "canonical_revenue": "Revenue", "derive_fcf": True, "balance_net_debt": True,
        "min_structure": (22, 35, 25), "min_mapped": (13, 18, 14),
    },
    "amazon": {
        "name": "Amazon / technology & infrastructure corporate", "income": AMZN_INCOME,
        "balance": deepcopy(generic.BALANCE), "cash": deepcopy(generic.CASH),
        "canonical_revenue": "Revenue", "derive_fcf": True, "balance_net_debt": True,
        "min_structure": (24, 35, 25), "min_mapped": (13, 18, 14),
    },
    "nvidia": {
        "name": "NVIDIA / semiconductor corporate", "income": NVDA_INCOME,
        "balance": deepcopy(generic.BALANCE), "cash": deepcopy(generic.CASH),
        "canonical_revenue": "Revenue", "derive_fcf": True, "balance_net_debt": True,
        "min_structure": (20, 35, 25), "min_mapped": (13, 18, 14),
    },
    "tsm": {
        "name": "TSMC / IFRS foreign private issuer", "income": TSM_INCOME, "balance": TSM_BALANCE, "cash": TSM_CASH,
        "canonical_revenue": "Revenue", "derive_fcf": True, "balance_net_debt": True,
        "min_structure": (18, 30, 25), "min_mapped": (11, 16, 13),
    },
    "siemens": {
        "name": "Siemens / IFRS industrial", "income": SIE_INCOME, "balance": SIE_BALANCE, "cash": SIE_CASH,
        "canonical_revenue": "Revenue", "derive_fcf": True, "balance_net_debt": True,
        "min_structure": (20, 30, 25), "min_mapped": (11, 16, 13),
    },
    "bank": {
        "name": "Bank / financial institution", "income": BANK_INCOME, "balance": BANK_BALANCE, "cash": BANK_CASH,
        "canonical_revenue": "Total Net Revenue", "derive_fcf": False, "balance_net_debt": False,
        "min_structure": (20, 28, 25), "min_mapped": (10, 14, 10),
    },
    "berkshire": {
        "name": "Insurance / operating conglomerate", "income": BRK_INCOME, "balance": BRK_BALANCE, "cash": BRK_CASH,
        "canonical_revenue": "Total Revenues", "derive_fcf": False, "balance_net_debt": False,
        "min_structure": (24, 28, 25), "min_mapped": (10, 14, 11),
    },
}

TICKER_PROFILE = {
    "GOOGL": "google", "GOOG": "google",
    "AMZN": "amazon",
    "NVDA": "nvidia",
    "TSM": "tsm",
    "SIE.DE": "siemens",
    "JPM": "bank",
    "BRK.B": "berkshire", "BRK-B": "berkshire", "BRK.A": "berkshire", "BRK-A": "berkshire",
}


def get_statement_profile(ticker: str):
    key = TICKER_PROFILE.get(str(ticker).upper().strip(), "default")
    profile = deepcopy(PROFILES[key])
    profile["key"] = key
    return profile
