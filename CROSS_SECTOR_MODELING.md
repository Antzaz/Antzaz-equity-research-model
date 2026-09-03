# Cross-Sector Modeling Contract

The production equity-research path is designed around a simple rule:

> A company either receives an economically appropriate research framework, or an unsuitable metric is explicitly marked REVIEW / N/M. Missing or non-comparable evidence is never replaced with a fabricated number.

## Production entry points

Use either:

```powershell
python .\commodity_safe_runner.py TICKER
```

for the deterministic production workbook, or:

```powershell
python .\research.py TICKER
```

for the same guarded workbook plus source/filing/KPI/thesis/QA agents.

`update_model.py` remains the low-level legacy generator used internally by the guarded runtime. New cross-sector investment work should use one of the production entry points above so business-model policy, statement routing, score gates and final quality checks are installed.

## Business-model routing

`business_model_registry.py` classifies issuers primarily from sector and industry rather than a ticker whitelist. Explicit ticker overrides are reserved for genuinely unusual structures.

| Business model | Primary research / valuation lens | Industrial reverse DCF |
| --- | --- | --- |
| Bank | P/TBV, ROTCE/ROE, normalized P/E, capital return | N/M |
| Insurance / reinsurance | P/B, normalized ROE/earnings, reserve/float economics | N/M |
| Insurance-led conglomerate | SOTP, operating earnings, book/investment economics | N/M |
| REIT | NAV, P/AFFO, P/FFO, property cash economics | N/M |
| Commodity producer | Mid-cycle FCF/NAV, EV/EBITDA, commodity sensitivity | Allowed with normalization |
| Utility | Rate-base/regulated DCF, P/E, EV/EBITDA | Allowed with sector interpretation |
| Software/cloud | DCF, EV/FCF, EV/Sales / earnings | Allowed |
| Semiconductor/hardware | DCF, normalized FCF/earnings, cycle-aware multiples | Allowed |
| Pharma | DCF, P/E, pipeline/patent-cliff scenarios | Allowed |
| Biotech | Pipeline rNPV / probability-weighted DCF, cash runway | Allowed with pipeline caveat |
| Industrial | DCF, EV/EBITDA, normalized earnings/FCF | Allowed |
| Consumer | DCF, earnings/FCF multiples, unit economics | Allowed |
| Payments network | DCF, earnings/FCF multiples | Allowed; not treated as a bank merely because sector is Financial Services |

## Statement profiles

The guarded runtime registers Company Data sector/industry before the full statement renderer executes.

- Known issuers can retain richer verified profiles (Alphabet, Amazon, NVIDIA, TSMC, Siemens, JPMorgan, Berkshire).
- Newly searched banks are routed to the bank statement taxonomy automatically.
- Newly searched insurers reuse a conservative insurance-oriented statement taxonomy with industrial FCF disabled.
- REITs retain corporate statements but industrial FCF derivation is disabled as a primary metric.
- Ordinary operating companies retain the standard corporate profile.

Unmapped statement rows remain blank. A structurally complete template is not evidence that every line item was disclosed; Data Quality separately reports mapped-row depth.

## Score-engine safeguards

The score engine excludes dimensions that are economically inappropriate instead of assigning a false low/high score:

- banks / insurers: industrial FCF Quality, corporate net-debt leverage score, absolute DCF valuation and DCF-derived stress valuation are excluded;
- REITs: industrial FCF Quality and conventional DCF-derived valuation/stress dimensions are excluded;
- conventional companies retain the normal reliability-gated score engine;
- commodity producers retain the commodity normalization/triangulation layer.

Available dimensions are reweighted by the existing reliability-aware score engine. Exclusion is not a neutral 50/100 assumption.

## Segment contract

Official segment economics have priority. If numeric segment extraction fails:

1. issuer/SEC discovered segment names are retained;
2. verified issuer adapters are used where a known parser weakness has been validated (for example Alphabet);
3. if the final Segment Analysis still lacks a useful taxonomy, the issuer/product grouping from Company Data is used as a descriptive-only fallback;
4. financial cells remain blank unless the issuer actually disclosed the values.

This prevents a failed HTML parser from turning into either an empty/stale segment tab or fabricated product economics.

## AI / ML safeguards

The AI Growth layer always allows evidence extraction and the fundamental revenue-growth forecast when data is sufficient.

For banks, insurers, broker/dealers and REITs:

- next-FY industrial FCF growth is marked NOT_APPLICABLE;
- the AI-adjusted FCF forecast is N/M;
- reverse-DCF implied FCF growth is N/M;
- the expectations gap is N/M;
- the output states the sector-appropriate primary framework instead.

This avoids turning a cross-sector ML feature into a financially inappropriate valuation conclusion.

## Public portfolio safeguards

The institutional/public reverse-DCF engine and the recruiter-safe fundamentals exporter use the same business-model registry. Public Equity Research therefore does not display a conventional reverse-DCF hurdle for a bank, insurer or REIT simply because a data provider supplies a `freeCashflow` field.

## Testing

`.github/workflows/cross-sector-contract-smoke.yml` has two layers:

1. offline contract tests covering banks, insurance, REITs, commodities, utilities, software, semiconductors, pharma, biotech, industrials, consumer companies, payments and digital platforms;
2. representative live production workbook builds for MSFT, AMD, BAC, PGR, PLD, CVX, NEE, LLY, CAT and WMT.

The live contract requires core sheets, final Data Quality policy, appropriate statement-profile routing, sector-safe valuation gating and no broken `#REF!` formulas. Segment financial coverage is allowed to be REVIEW when the issuer does not disclose/expose data in a reliably parsable form.

## What this contract does not claim

No automated system can guarantee full numeric coverage for every listed company. Issuer disclosure formats, foreign filings, data-provider availability and unusual corporate structures vary. The guarantee is about **failure behavior**: unsupported or economically inappropriate analysis should become REVIEW / N/M with an explanation, rather than silently producing a plausible-looking but wrong result.
