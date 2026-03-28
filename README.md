# Dividend Trend Signal Engine

A quantitative screening tool for dividend-paying stocks that combines yield analysis, trend regression, and payout sustainability metrics to generate systematic investment signals.

## Overview

This engine performs fundamental and technical analysis on dividend-paying stocks, with specialized handling for REITs and asset-heavy business trusts. It processes historical financial data to evaluate dividend safety, identify yield traps, and generate actionable signals based on both valuation and trend metrics.

## Features

### Fundamental Analysis
- **REIT/Trust-Aware Metrics**: Automatically uses Funds From Operations (FFO) or Operating Cash Flow (OCF) for REITs and asset-heavy entities; defaults to EPS for regular corporations
- **Payout Sustainability**: Calculates payout ratios with appropriate per-share metrics and flags unsustainable distributions

### Technical Analysis
- **Yield Z-Score**: Blended 5-year and 10-year mean-reversion signal for identifying historically attractive yields
- **Trend Regression**: Log-linear regression over 252-day rolling windows to determine trend direction and strength
- **Price Deviation**: Z-score calculation measuring current price distance from trend line

### Risk Detection
- **Dividend Trap Identification**: Flags stocks with unsustainable payouts (>100% of FFO/OCF for REITs, >90% for regular companies) combined with dividend cuts
- **Drawdown Analysis**: Tracks 1-year maximum drawdown to assess recovery potential
- **52-Week High/Low Proximity**: Measures relative price position for mean reversion signals

### Signal Generation
The engine produces seven distinct signal types based on combined valuation and trend metrics:

| Signal | Description |
|--------|-------------|
| **STRONG BUY** | Deep value with yield Z >1.5, price Z <-1.0, within 20% of 52W low, trend intact |
| **BUY** | Attractive yield with Z >1.0, within 30% of 52W low, trend intact |
| **CONTRARIAN BUY** | Very cheap yield (Z >2.0) with broken trend, within 40% of 52W low |
| **OPPORTUNISTIC BUY** | Growing yield with positive Z-score |
| **TRIM** | Yield compressed (Z <-1.5), historically expensive |
| **CAUTION** | High payout ratio (>95%) but no active dividend cut |
| **AVOID** | Dividend trap detected (cut and/or unsustainable payout) |

### Output
- **Excel Report**: Color-coded analysis with auto-adjusted column widths
- **Color Legend**: Green (BUY), Red (AVOID/CAUTION), Yellow (HOLD), Orange (TRIM)
- **Sorted Output**: Results ranked by signal strength, price Z-score, and yield Z-score

## Quick Start

```bash
# Install dependencies
pip install pandas numpy yfinance openpyxl python-dateutil

# Prepare input: sti_component.xlsx with columns [name, code, industry]
# Run the engine
python dividend_trend_signal_engine.py

# Output: dividend_trend_signal_output.xlsx
```

## Disclaimer
This content is for informational and research purposes only and does not constitute financial advice. Individuals should conduct their own due diligence prior to making any investment decision.
