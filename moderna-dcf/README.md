# Moderna (MRNA) DCF Model

A discounted cash flow valuation of Moderna, built in Python and delivered as a fully formula-driven Excel workbook. Every number in the spreadsheet is either a documented input or a live formula, so you can change any assumption in Excel and the valuation updates.

**Model date:** October 5, 2026 · **Share price used:** $190.01 (Oct 2, 2026 close) · **Units:** US$ millions unless stated

## Results

| Scenario | Value per share | vs. $190.01 | 2040E revenue | Terminal value % of EV |
|---|---|---|---|---|
| Bear | $12.84 | −93% | $4.9bn | 208%* |
| **Base** | **$44.89** | **−76%** | **$10.6bn** | **52%** |
| Bull | $98.70 | −48% | $18.7bn | 46% |

WACC 11.7% (CAPM: 5.28% risk-free, 1.30 beta, 5.0% equity risk premium). *In the bear case near-term cash burn makes PV of explicit-period cash flows negative, so terminal value exceeds 100% of EV.

**Takeaway:** at ~$75bn market cap, the stock prices in more than this model's bull case: roughly $13bn of 2040 risk-adjusted intismeran revenue to Moderna is still not enough at an ~12% discount rate. Getting near $190 needs some combination of a much lower discount rate, intismeran working across many more tumor types, or value from the rest of the pipeline that isn't modeled here. The sensitivity grid on the DCF tab shows how much the discount rate alone moves the answer.

## How the model works

**Revenue** has two engines:

1. **Respiratory vaccines** (COVID, RSV, flu, flu+COVID combo, norovirus): FY2025 actual of $1,944mm, +8% in 2026 (company targets up to 10%), then scenario growth rates for 2027–30, 2031–35 and 2036–40.
2. **Intismeran autogene** (individualized cancer therapy with Merck): Moderna's share under the 50/50 profit split, modeled per indication (adjuvant melanoma, NSCLC, RCC/bladder/other) with a launch year, a linear six-year ramp to peak, and a probability of approval. Revenue is risk-adjusted (peak × ramp × probability).

**Costs:** vaccine gross margin ramps from ~62% in 2026 to a scenario target by 2030; intismeran carries its own cost-of-sales rate; R&D and SG&A follow 2026 guidance, a 2027 step-down, then the greater of a scenario floor or a percent of revenue.

**Taxes:** an estimated $12bn of usable NOLs shields income, capped at 80% of taxable income per year (post-2017 US federal rule), then 21%.

**Free cash flow:** EBIT − cash taxes + D&A − capex − increase in working capital. Stock-based compensation stays in opex as a real cost.

**Valuation:** 2027–2040 cash flows discounted with a mid-year convention from the valuation date, Gordon-growth terminal value after 2040, plus net cash. Net cash uses Moderna's guided year-end 2026 cash ($4.7–5.2bn midpoint, already net of the $950mm July 2026 litigation payment) less $600mm of term-loan debt. 2026 is shown for reference but not valued, because its remaining burn is already in that cash figure.

## Workbook tabs

| Tab | What's on it |
|---|---|
| **DCF** | Valuation summary, equity bridge, diagnostics, WACC × terminal-growth sensitivity |
| **Inputs** | Scenario selector, market data, WACC build, Bear/Base/Bull levers, operating assumptions |
| **Model** | 2025A–2040E revenue build, P&L, NOL tax schedule, free cash flow, discounting |
| **Sources** | Every source and method note |

Blue text = hard-coded input, black = formula, green = link from another sheet, yellow = key lever. Hover over cells with a red corner to see their source.

**To switch scenarios:** change `Inputs!C6` to 1 (Bear), 2 (Base) or 3 (Bull). To force a discount rate, type it into the WACC override cell.

## Running it

```bash
pip install -r requirements.txt
python build_model.py              # writes Moderna_DCF.xlsx
python build_model.py -o my.xlsx   # custom output path
```

The workbook is set to recalculate on open, so Excel shows the numbers immediately. The committed `Moderna_DCF.xlsx` is already calculated (Base case) so it also previews correctly on GitHub and in other viewers.

## Key assumptions to challenge

| Lever | Bear | Base | Bull |
|---|---|---|---|
| Vaccine growth 2027–30 | 2% | 8% | 15% |
| Vaccine gross margin (2030+) | 55% | 65% | 72% |
| Melanoma peak to Moderna / probability | $2.5bn / 80% | $4.0bn / 90% | $5.5bn / 95% |
| NSCLC peak to Moderna / probability | $2.0bn / 35% | $5.0bn / 50% | $10.0bn / 60% |
| RCC, bladder & other peak / probability | $0.5bn / 30% | $1.5bn / 45% | $3.3bn / 55% |
| R&D floor | $1.6bn | $2.4bn | $2.8bn |
| Terminal growth | 1.5% | 2.5% | 3.0% |

Bull-case intismeran peaks are anchored to William Blair's 2040 estimates (Aug 2026).

## Sources

- [Moderna Q2 2026 results and 2026 financial framework (SEC 8-K)](https://www.sec.gov/Archives/edgar/data/0001682852/000168285226000147/exhibit9912026q2pressrelea.htm)
- [Moderna Q4 / FY2025 results (SEC 8-K)](https://www.sec.gov/Archives/edgar/data/1682852/000168285226000015/exhibit9912025q4pressrelea.htm)
- [Moderna J.P. Morgan 2026 update, 2027 opex and 2028 breakeven target (SEC 8-K)](https://www.sec.gov/Archives/edgar/data/1682852/000168285226000007/exhibit991-01122026.htm)
- [MRNA price and market cap, Oct 2, 2026](https://www.ad-hoc-news.de/boerse/news/corporate-news/moderna-stock-gained-0-57-percent-as-nasdaq-100-entry-nears/70226496)
- [10-year Treasury yield, Oct 2, 2026](https://stockmarketwatch.com/bonds/10-year-treasury-yield)
- [William Blair intismeran estimates (Benzinga)](https://www.benzinga.com/analyst-stock-ratings/upgrades/26/08/61333408/modernas-intismeran-close-to-keytruda-style-oncology-opportunity-analyst)

## Disclaimer

For education and analysis only, not investment advice. Forecasts, probabilities and the NOL balance are illustrative assumptions. Key near-term catalysts (full Phase 3 melanoma data at ESMO in late October, Q3 earnings Oct 29, Analyst Day Nov 12, 2026) could change the inputs materially.
