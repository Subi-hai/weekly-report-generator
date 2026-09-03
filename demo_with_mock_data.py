"""
Demo runner: builds the same report as weekly_report_generator.py, but with
synthetic price data instead of a live Yahoo Finance pull. Useful for
previewing the report's formatting/charts without network access, and as a
quick sanity check that build_report() works correctly.
"""

import random
from datetime import datetime, timedelta

import pandas as pd

from weekly_report_generator import build_report

random.seed(7)


def make_mock_series(start_price, days=20, drift=0.0, vol=0.015):
    dates = [datetime.today() - timedelta(days=days - i) for i in range(days)]
    prices = [start_price]
    for _ in range(days - 1):
        change = random.gauss(drift, vol)
        prices.append(max(prices[-1] * (1 + change), 0.01))
    volumes = [random.randint(2_000_000, 8_000_000) for _ in range(days)]
    return pd.DataFrame({"Date": dates, "Close": prices, "Volume": volumes})


mock_data = {
    "AAPL": make_mock_series(228.0, drift=0.004),
    "MSFT": make_mock_series(415.0, drift=0.002),
    "TLT": make_mock_series(92.0, drift=-0.003),
    "GLD": make_mock_series(245.0, drift=0.006),
}

build_report(mock_data, output_path="weekly_report_demo.xlsx")
