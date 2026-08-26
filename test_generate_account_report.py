import unittest
from datetime import date

import pandas as pd

from generate_account_report import (
    PriceQuote,
    build_cash_flow_benchmark,
    calculate_benchmark_return,
    calculate_xirr,
)


class BenchmarkReturnTests(unittest.TestCase):
    def test_calculates_price_return(self):
        start = PriceQuote(symbol="VOO", price=400.0, source="test")
        end = PriceQuote(symbol="VOO", price=440.0, source="test")

        self.assertAlmostEqual(calculate_benchmark_return(start, end), 0.10)

    def test_missing_quote_returns_none(self):
        end = PriceQuote(symbol="VOO", price=440.0, source="test")

        self.assertIsNone(calculate_benchmark_return(None, end))

    def test_nonpositive_start_price_returns_none(self):
        start = PriceQuote(symbol="VOO", price=0.0, source="test")
        end = PriceQuote(symbol="VOO", price=440.0, source="test")

        self.assertIsNone(calculate_benchmark_return(start, end))


class MoneyWeightedReturnTests(unittest.TestCase):
    def test_xirr_for_one_year_return(self):
        flows = [
            (date(2025, 1, 1), -100.0),
            (date(2026, 1, 1), 110.0),
        ]

        self.assertAlmostEqual(calculate_xirr(flows), 0.10, places=8)

    def test_xirr_requires_both_cash_flow_signs(self):
        self.assertIsNone(calculate_xirr([(date(2025, 1, 1), -100.0)]))

    def test_xirr_requires_elapsed_time(self):
        flows = [
            (date(2025, 1, 1), -100.0),
            (date(2025, 1, 1), 110.0),
        ]

        self.assertIsNone(calculate_xirr(flows))

    def test_cash_flow_matched_benchmark_uses_each_deposit_date(self):
        trades = pd.DataFrame(
            [
                {
                    "Date": date(2025, 1, 1),
                    "Symbol": "CASH",
                    "Price": 1.0,
                    "Qty": 100.0,
                    "Comm Fee": 0.0,
                },
                {
                    "Date": date(2025, 7, 1),
                    "Symbol": "CASH",
                    "Price": 1.0,
                    "Qty": 100.0,
                    "Comm Fee": 0.0,
                },
            ]
        )
        quotes = {
            date(2025, 1, 1): PriceQuote("VOO", 100.0, "test"),
            date(2025, 7, 1): PriceQuote("VOO", 105.0, "test"),
        }
        end_quote = PriceQuote("VOO", 110.0, "test")

        result = build_cash_flow_benchmark(
            trades,
            date(2026, 1, 1),
            quotes,
            end_quote,
        )

        self.assertIsNotNone(result)
        expected_value = (100.0 / 100.0 + 100.0 / 105.0) * 110.0
        self.assertAlmostEqual(result.ending_value, expected_value)
        self.assertAlmostEqual(result.pnl, expected_value - 200.0)
        self.assertIsNotNone(result.money_weighted_return_pct)


if __name__ == "__main__":
    unittest.main()
