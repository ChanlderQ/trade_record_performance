import unittest
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd

from generate_account_report import (
    PriceQuote,
    build_daily_performance,
    build_cash_flow_benchmark,
    build_excess_return_curve,
    calculate_benchmark_return,
    calculate_performance_metrics,
    calculate_xirr,
    write_nav_chart,
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


class TimeWeightedReturnTests(unittest.TestCase):
    def test_daily_twr_reconstructs_value_and_drawdown(self):
        trades = pd.DataFrame(
            [
                {
                    "Date": date(2025, 1, 2),
                    "Symbol": "CASH",
                    "Price": 1.0,
                    "Qty": 100.0,
                    "Comm Fee": 0.0,
                },
                {
                    "Date": date(2025, 1, 2),
                    "Symbol": "ABC",
                    "Price": 100.0,
                    "Qty": 1.0,
                    "Comm Fee": 0.0,
                },
            ]
        )
        prices = {
            "ABC": {
                date(2025, 1, 2): 100.0,
                date(2025, 1, 3): 110.0,
                date(2025, 1, 6): 99.0,
            },
            "VOO": {
                date(2025, 1, 2): 400.0,
                date(2025, 1, 3): 404.0,
                date(2025, 1, 6): 402.0,
            },
        }

        result = build_daily_performance(
            trades,
            date(2025, 1, 6),
            prices,
            0.0,
        )

        self.assertAlmostEqual(result.history["Ending Value"].iloc[-1], 99.0)
        self.assertAlmostEqual(result.time_weighted_return_pct, -0.01)
        self.assertAlmostEqual(result.max_drawdown_pct, -0.10)
        self.assertIsNotNone(result.sharpe_ratio)

    def test_metrics_chain_daily_returns(self):
        history = pd.DataFrame(
            {
                "Daily Return": [0.0, 0.10, -0.10, 0.05],
                "TWR Index": [100.0, 110.0, 99.0, 103.95],
                "Trading Day": [True, True, True, True],
            },
            index=[
                date(2025, 1, 2),
                date(2025, 1, 3),
                date(2025, 1, 6),
                date(2025, 1, 7),
            ],
        )

        result = calculate_performance_metrics(history, 0.0)

        self.assertAlmostEqual(result.time_weighted_return_pct, 0.0395)
        self.assertAlmostEqual(result.max_drawdown_pct, -0.10)
        self.assertIsNotNone(result.sharpe_ratio)

    def test_writes_png_nav_chart(self):
        curve = pd.DataFrame(
            {"TWR Index": [100.0, 105.0]},
            index=[date(2025, 1, 2), date(2025, 1, 3)],
        )
        with TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "nav.png"
            excess = pd.DataFrame(
                {"Excess Return": [0.0, 0.05]},
                index=curve.index,
            )
            write_nav_chart(
                output,
                {"Combined": curve, "VOO": curve},
                excess,
            )

            self.assertTrue(output.read_bytes().startswith(b"\x89PNG\r\n\x1a\n"))

    def test_excess_curve_is_twr_minus_benchmark_return(self):
        dates = [date(2025, 1, 2), date(2025, 1, 3), date(2025, 1, 6)]
        account = pd.DataFrame(
            {"TWR Index": [100.0, 110.0, 120.0]},
            index=dates,
        )
        benchmark = pd.DataFrame(
            {"TWR Index": [100.0, 105.0, 110.0]},
            index=dates,
        )

        excess = build_excess_return_curve(account, benchmark)

        self.assertAlmostEqual(excess["Excess Return"].iloc[0], 0.0)
        self.assertAlmostEqual(excess["Excess Return"].iloc[1], 0.05)
        self.assertAlmostEqual(excess["Excess Return"].iloc[2], 0.10)


if __name__ == "__main__":
    unittest.main()
