import unittest

from generate_account_report import PriceQuote, calculate_benchmark_return


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


if __name__ == "__main__":
    unittest.main()
