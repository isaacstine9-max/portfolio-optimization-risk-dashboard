# Portfolio Optimization & Risk Dashboard

Python tools for building, comparing and stress-testing investment portfolios, plus a full discounted cash flow (DCF) model for valuing a single company.

## What's inside

| File | What it does |
|---|---|
| `portfolio_optimization_project.py` | Portfolio optimization and risk analysis across 8 ETFs (Modern Portfolio Theory) |
| `DCF_Model.ipynb` | Jupyter notebook: 5-year DCF valuation with sensitivity, scenario and Monte Carlo analysis |
| `requirements.txt` | Python packages needed to run both |

### Portfolio optimization (`portfolio_optimization_project.py`)

- Downloads daily adjusted prices from 2015 to today with `yfinance` for a diversified ETF universe:
  SPY (US large cap), QQQ (US growth/tech), IWM (US small cap), EFA (developed international), EEM (emerging markets), TLT (long-term Treasuries), GLD (gold), VNQ (US real estate).
- Plots growth of $100 per asset and the correlation matrix of daily returns.
- Finds two optimal long-only, fully invested portfolios with SciPy (SLSQP):
  - **Maximum Sharpe ratio**: the best return per unit of risk
  - **Minimum volatility**: the lowest-risk mix
- Simulates 20,000 random portfolios to map the risk/return opportunity set (efficient frontier).
- Compares Equal Weight, Max Sharpe and Min Volatility on:
  growth of $1, annualized return and volatility, Sharpe and Sortino ratios, maximum drawdown, 95% historical Value at Risk (VaR), 95% expected shortfall, and 63-day rolling volatility.
- Assumptions: 4% risk-free rate, 252 trading days per year.

### DCF valuation (`DCF_Model.ipynb`)

Values a company as the present value of its future free cash flows. All inputs live in the `BASE` dictionary in section 1; change them and re-run, and everything else updates. Figures are in $ millions.

1. **Inputs**: revenue, growth path, EBIT margins, tax, D&A, capex, working capital, CAPM inputs (risk-free rate, beta, equity risk premium), debt, cash, shares, share price, terminal growth.
2. **Model engine**: builds unlevered free cash flow for 5 years, discounts it at WACC (mid-year convention), adds a Gordon-growth terminal value, and bridges enterprise value to equity value per share.
3. **Base-case results**: projection table and valuation summary (WACC, EV, equity value, value per share vs market price).
4. **Charts**: revenue and growth, free cash flow build, where enterprise value comes from, EV-to-equity bridge.
5. **Sensitivity heatmaps**: value per share across WACC × terminal growth, WACC × margin, and growth × margin.
6. **Tornado chart**: which single input moves value the most.
7. **Scenarios**: bear, base and bull cases with a probability-weighted value.
8. **Reverse DCF**: the WACC, terminal growth, margin or growth the current share price implies.
9. **Monte Carlo**: 10,000 random draws of the key inputs, giving a distribution of value per share and the probability it exceeds the market price.
10. **Export**: saves projections, summary and a sensitivity table to `dcf_output.xlsx`.

## Setup

Requires Python 3.10+.

```bash
git clone https://github.com/isaacstine9-max/portfolio-optimization-risk-dashboard.git
cd portfolio-optimization-risk-dashboard

python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate

pip install -r requirements.txt
```

## Running

**DCF model**

```bash
jupyter notebook DCF_Model.ipynb
```

Then choose Run → Run All Cells. You can also upload the notebook to Google Colab and use Runtime → Run all.

**Portfolio optimization**

The script was exported from a Colab notebook and uses notebook-style `display()` calls, so run it with IPython (installed with Jupyter):

```bash
ipython portfolio_optimization_project.py
```

Or paste it into a Jupyter or Colab notebook. It needs an internet connection to download prices.

## Notes

- Portfolio weights are optimized and evaluated on the same price history (in-sample), so backtest results will look better than live performance would.
- DCF inputs are illustrative estimates. Replace them with figures from the company's filings before drawing conclusions.
- This is an educational project, not investment advice.

## Roadmap

- Interactive dashboard UI (e.g. Streamlit)
- Automated checks on every push with GitHub Actions

## Authors

- [@isaacstine9-max](https://github.com/isaacstine9-max)
- [@delectablee](https://github.com/delectablee)

## License

MIT; see [LICENSE](LICENSE).
