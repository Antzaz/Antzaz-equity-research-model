"""Central source registry for equity research.

Keep issuer-owned, regulator, consensus, specialist-industry and professional-methodology
sources in one place so workbook, agents and source-health checks do not maintain duplicate
hard-coded lists. Financial statements still follow the primary-source hierarchy:
issuer disclosures -> regulator/XBRL -> transparent market-data fallback.

Consensus aggregators and professional research are secondary evidence. They must never
silently overwrite company-reported facts or be presented as proprietary institutional data.
"""

from __future__ import annotations

from typing import Iterable

ISSUER_SOURCES = {
    "GOOGL": {"investor":"https://abc.xyz/investor/","earnings":"https://abc.xyz/investor/earnings/","annual_reports":"https://abc.xyz/investor/annual-reports/","governance":"https://abc.xyz/investor/board-and-governance/","workplace_governance":"https://abc.xyz/investor/board-and-governance/ldicc/default.aspx","additional_information":"https://abc.xyz/investor/additional-information/default.aspx"},
    "MSFT": {"investor":"https://www.microsoft.com/en-us/investor","investor_information":"https://www.microsoft.com/en-us/investor/investor-information","annual_reports":"https://www.microsoft.com/en-us/investor/annual-reports","filings":"https://www.microsoft.com/en-us/investor/sec-filings"},
    "AMZN": {"investor":"https://ir.aboutamazon.com/","quarterly_results":"https://ir.aboutamazon.com/quarterly-results/default.aspx","filings":"https://ir.aboutamazon.com/sec-filings/default.aspx","annual_reports":"https://ir.aboutamazon.com/annual-reports-proxies-and-shareholder-letters/default.aspx"},
    "META": {"investor":"https://investor.atmeta.com/home/default.aspx","financials":"https://investor.atmeta.com/financials/default.aspx"},
    "NVDA": {"investor":"https://investor.nvidia.com/","financial_reports":"https://investor.nvidia.com/financial-info/financial-reports/default.aspx","annual_reports":"https://investor.nvidia.com/financial-info/annual-reports-and-proxies/default.aspx","filings":"https://investor.nvidia.com/financial-info/sec-filings/default.aspx"},
    "AAPL": {"investor":"https://investor.apple.com/investor-relations/","filings":"https://investor.apple.com/sec-filings/default.aspx","governance":"https://investor.apple.com/leadership-and-governance/default.aspx"},
    "TSM": {"investor":"https://investor.tsmc.com/english","financial_reports":"https://investor.tsmc.com/english/financial-reports","quarterly_results":"https://investor.tsmc.com/english/quarterly-results","annual_reports":"https://investor.tsmc.com/static/annualReports/2025/english/index.html","filings":"https://investor.tsmc.com/english/sec-filings","leadership":"https://www.tsmc.com/english/aboutTSMC/executives","governance":"https://investor.tsmc.com/english/board-of-directors","workplace":"https://esg.tsmc.com/en-US/articles/358","sustainability":"https://esg.tsmc.com/en-US/ESG-data-hub/latest-sustainability-information?tab=overview"},
    "UBS": {"investor":"https://www.ubs.com/global/en/investor-relations.html","annual_reports":"https://www.ubs.com/global/en/investor-relations/financial-information/annual-reporting.html","quarterly_results":"https://www.ubs.com/global/en/investor-relations/financial-information/quarterly-reporting.html","business_structure":"https://www.ubs.com/global/en/our-firm/governance/ubs-group-ag/organization-structure.html"},
}

TICKER_ALIASES={"GOOG":"GOOGL","2330.TW":"TSM"}

REGULATOR_AND_MACRO_SOURCES={
    "sec_companyfacts":{"provider":"SEC EDGAR / Company Facts","url":"https://data.sec.gov/","purpose":"Primary US filing/XBRL facts and filings","source_type":"Regulator / primary filing"},
    "sec_edgar":{"provider":"SEC EDGAR","url":"https://www.sec.gov/edgar/search/","purpose":"10-K, 10-Q, 8-K, proxy and other filed disclosures","source_type":"Regulator / primary filing"},
    "fred":{"provider":"Federal Reserve Bank of St. Louis FRED","url":"https://fred.stlouisfed.org/","purpose":"Rates, inflation, labor, credit and macroeconomic time series","source_type":"Official / public macro data"},
    "bls":{"provider":"U.S. Bureau of Labor Statistics","url":"https://www.bls.gov/","purpose":"Labor, wage and productivity context","source_type":"Official statistics"},
    "bea":{"provider":"U.S. Bureau of Economic Analysis","url":"https://www.bea.gov/","purpose":"GDP, corporate profits and industry economic context","source_type":"Official statistics"},
}

CONSENSUS_DATA_SOURCES={
    "yahoo":{"provider":"Yahoo Finance / yfinance","url":"https://finance.yahoo.com/","purpose":"Zero-configuration public analyst revenue/EPS snapshot and revisions","source_type":"Public market-data aggregator","env_key":None},
    "fmp":{"provider":"Financial Modeling Prep","url":"https://site.financialmodelingprep.com/developer/docs/stable/financial-estimates","purpose":"Revenue, EPS, EBIT, EBITDA and other analyst estimates; price-target consensus","source_type":"API consensus aggregator","env_key":"FMP_API_KEY"},
    "alpha_vantage":{"provider":"Alpha Vantage","url":"https://www.alphavantage.co/documentation/#earnings-estimates","purpose":"Annual/quarterly revenue and EPS estimates, analyst count and revision history","source_type":"API consensus aggregator","env_key":"ALPHAVANTAGE_API_KEY"},
    "finnhub":{"provider":"Finnhub","url":"https://finnhub.io/docs/api/insider-sentiment","purpose":"Optional revenue, EPS, EBIT, EBITDA, FCF, OCF and capex estimate feeds when plan access permits","source_type":"API estimates / alternative-data provider","env_key":"FINNHUB_API_KEY"},
    "stockanalysis_validation":{"provider":"StockAnalysis.com","url":"https://stockanalysis.com/","purpose":"Human cross-check for public forecast/analyst pages when API coverage is thin","source_type":"Secondary public validation; not automated primary source","env_key":None},
    "marketbeat_validation":{"provider":"MarketBeat","url":"https://www.marketbeat.com/","purpose":"Human cross-check for earnings estimates, analyst ratings and target ranges","source_type":"Secondary public validation; not automated primary source","env_key":None},
}

SPECIALIST_MARKET_SOURCES={
    "foundry":{"provider":"TrendForce","url":"https://www.trendforce.com/presscenter/news/20260612-13095.html","purpose":"Comparable global foundry revenue share","source_type":"Specialist industry research"},
    "search_engine":{"provider":"StatCounter","url":"https://gs.statcounter.com/search-engine-market-share/","purpose":"Search-engine usage share by geography/device","source_type":"Specialist web-usage measurement"},
    "cloud_infrastructure":{"provider":"Synergy Research Group","url":"https://www.srgresearch.com/articles/cloud-market-annual-revenue-run-rate-topped-half-a-trillion-dollars-in-q1-as-growth-surge-continues","purpose":"Cloud infrastructure services market share","source_type":"Specialist industry research"},
    "smartphone_shipments":{"provider":"IDC","url":"https://www.idc.com/promo/smartphone-market-share/","purpose":"Worldwide smartphone shipment share","source_type":"Specialist industry research"},
    "semiconductor_market":{"provider":"WSTS","url":"https://www.wsts.org/","purpose":"Semiconductor market size and industry-cycle context","source_type":"Industry association / market statistics"},
    "pc_shipments":{"provider":"IDC","url":"https://www.idc.com/promo/pc-market-share/","purpose":"PC shipment and vendor-share context","source_type":"Specialist industry research"},
    "mobile_os":{"provider":"StatCounter","url":"https://gs.statcounter.com/os-market-share/mobile/worldwide","purpose":"Mobile operating-system usage-share context","source_type":"Specialist web-usage measurement"},
}

PROFESSIONAL_FRAMEWORK_SOURCES={
    "mckinsey_ai":{"provider":"McKinsey Global Institute","url":"https://www.mckinsey.com/capabilities/tech-and-ai/our-insights/the-economic-potential-of-generative-ai-the-next-productivity-frontier","purpose":"Bottom-up GenAI use-case value, task/function exposure and measurable revenue/cost outcomes","source_type":"Professional research methodology"},
    "goldman_ai_capex":{"provider":"Goldman Sachs Research","url":"https://www.goldmansachs.com/insights/articles/why-ai-companies-may-invest-more-than-500-billion-in-2026","purpose":"AI capex, productivity-beneficiary and capex-to-revenue investor framework","source_type":"Sell-side/public research summary"},
    "goldman_ai_valuation":{"provider":"Goldman Sachs Research","url":"https://www.goldmansachs.com/insights/articles/how-much-could-ai-boost-us-stocks","purpose":"Translate productivity assumptions into EPS growth and valuation scenarios","source_type":"Sell-side/public research summary"},
    "blackrock_ai":{"provider":"BlackRock Investment Institute","url":"https://www.blackrock.com/corporate/insights/blackrock-investment-institute/publications/outlook","purpose":"AI value capture, scarcity, financing and capex-vs-revenue order-of-magnitude framework","source_type":"Asset-manager public research"},
    "morgan_cap":{"provider":"Morgan Stanley Investment Management / Counterpoint Global","url":"https://www.morganstanley.com/im/en-us/institutional-investor/insights/series/consilient-observer.html","purpose":"ROIC, capital allocation, competitive-advantage period, expectations and valuation frameworks","source_type":"Asset-manager public research"},
    "morgan_bayes":{"provider":"Morgan Stanley Investment Management / Counterpoint Global","url":"https://www.morganstanley.com/im/en-fi/institutional-investor/insights/consilient-observer/bayes-and-base-rates-2.html","purpose":"Bayesian updating and historical reference-class base rates","source_type":"Asset-manager public research"},
    "damodaran":{"provider":"Aswath Damodaran / NYU Stern","url":"https://pages.stern.nyu.edu/~adamodar/","purpose":"Valuation, implied growth, excess-return and corporate-finance methodology reference","source_type":"Academic / practitioner methodology"},
}


def canonical_ticker(ticker:str)->str:
    symbol=str(ticker or "").upper().strip(); return TICKER_ALIASES.get(symbol,symbol)

def issuer_sources(ticker:str,website:str|None=None)->dict[str,str]:
    key=canonical_ticker(ticker); out=dict(ISSUER_SOURCES.get(key,{}))
    if website:
        base=str(website).strip().rstrip("/")
        if base and base not in out.values(): out.setdefault("company_website",base)
    return out

def _selected_pages(ticker:str,keys:Iterable[str],website:str|None=None)->list[str]:
    sources=issuer_sources(ticker,website); pages=[sources[k] for k in keys if sources.get(k)]; return list(dict.fromkeys(pages))
def investor_pages(ticker:str,website:str|None=None)->list[str]:
    return _selected_pages(ticker,("investor","investor_information","financial_reports","quarterly_results","earnings","annual_reports","filings","company_website"),website)
def segment_pages(ticker:str,website:str|None=None)->list[str]:
    return _selected_pages(ticker,("quarterly_results","financial_reports","annual_reports","business_structure","investor","company_website"),website)
def specialist_sources()->dict[str,dict[str,str]]: return {k:dict(v) for k,v in SPECIALIST_MARKET_SOURCES.items()}
def consensus_sources()->dict[str,dict[str,str|None]]: return {k:dict(v) for k,v in CONSENSUS_DATA_SOURCES.items()}
def professional_framework_sources()->dict[str,dict[str,str]]: return {k:dict(v) for k,v in PROFESSIONAL_FRAMEWORK_SOURCES.items()}
def regulator_macro_sources()->dict[str,dict[str,str]]: return {k:dict(v) for k,v in REGULATOR_AND_MACRO_SOURCES.items()}
def full_source_catalog(ticker:str|None=None,website:str|None=None)->dict[str,dict]:
    return {"issuer":issuer_sources(ticker,website) if ticker else {},"regulator_macro":regulator_macro_sources(),"consensus":consensus_sources(),"specialist":specialist_sources(),"professional_frameworks":professional_framework_sources()}
