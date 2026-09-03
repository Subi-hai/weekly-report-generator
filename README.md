# Weekly Market Report Generator

A small automation tool that turns a list of tickers into a formatted Excel report: a summary table with conditional formatting (green for gains, red for losses), plus a per-ticker price chart. Built to demonstrate the kind of report automation and Excel/VBA work used in investment operations and business process roles.

## What it does

- Pulls recent daily price and volume data for any list of stock/ETF tickers (via [yfinance](https://pypi.org/project/yfinance/))
- Builds a "Summary" sheet with last close, weekly % change (conditionally formatted green/red), and average volume
- Builds one sheet per ticker with the raw price history and a line chart
- Ships with a VBA macro (`AutoFormatReport.bas`) that does the same formatting step natively inside Excel, for cases where the data is already pasted in from a terminal or internal export instead of pulled via API

## Files

| File | Purpose |
|---|---|
| `weekly_report_generator.py` | Main script. Pulls live data and builds the formatted `.xlsx` report. |
| `demo_with_mock_data.py` | Runs the same report-building logic with synthetic data, no API access needed. Useful for a quick preview. |
| `AutoFormatReport.bas` | Excel VBA macro that formats a raw data dump into the same summary layout, entirely inside Excel. |
| `requirements.txt` | Python dependencies. |

## Usage

### Python version (live data)

```bash
pip install -r requirements.txt
python weekly_report_generator.py AAPL MSFT TLT GLD
```

This produces `weekly_report.xlsx` with a Summary tab and one tab per ticker.

### Demo version (no internet required)

```bash
python demo_with_mock_data.py
```

Produces `weekly_report_demo.xlsx` using synthetic price series, so you can see the formatting and charts without hitting a live API.

### VBA version (data already in Excel)

1. Open Excel, press `Alt+F11` to open the VBA editor.
2. `File > Import File...` and select `AutoFormatReport.bas`.
3. Add a sheet named `RawData` with columns `Date | Ticker | Close | Volume`.
4. Run the `BuildSummary` macro. It builds a formatted `Summary` sheet with the same conditional formatting as the Python version.

## Example output

The Summary sheet auto-colors weekly performance:

| Ticker | Last Close | Weekly % Change | Avg Volume |
|---|---|---|---|
| AAPL | $237.77 | +4.28% (green) | 5,030,975 |
| TLT | $85.70 | -6.84% (red) | 5,118,076 |

Each ticker also gets its own tab with a price trend chart.

## Why I built this

I wanted a hands-on way to practice the kind of automation and reporting work involved in business process and financial operations roles, connecting a data source, standardizing the output, and formatting it into something a stakeholder could open and immediately read, in both Python and native Excel VBA.
