import os
import unittest
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
from urllib.error import HTTPError

import generate_account_report as report


class TiingoTests(unittest.TestCase):
    def test_key_precedence(self):
        with TemporaryDirectory() as directory:
            env = Path(directory) / '.env'
            env.write_text('TIINGO_API_KEY=file-key\n')
            with patch.dict(os.environ, {'TIINGO_API_KEY': 'env-key'}):
                self.assertEqual(report.get_api_key('cli-key', env, 'TIINGO_API_KEY'), 'cli-key')
                self.assertEqual(report.get_api_key(None, env, 'TIINGO_API_KEY'), 'env-key')
            with patch.dict(os.environ, {}, clear=True):
                self.assertEqual(report.get_api_key(None, env, 'TIINGO_API_KEY'), 'file-key')

    def test_weekend_uses_previous_close_and_not_adjusted_close(self):
        payload = [
            {'date': '2026-09-25T00:00:00Z', 'close': 100, 'adjClose': 90},
            {'date': '2026-09-28T00:00:00Z', 'close': 110},
        ]
        with patch.object(report, 'read_json_url', return_value=payload) as request:
            quote = report.fetch_tiingo_historical_quote('BRK.B', date(2026, 9, 27), 'key')
        self.assertEqual(quote.price, 100)
        self.assertEqual(quote.price_time, '2026-09-25')
        self.assertIn('/brk-b/prices?', request.call_args.args[0])
        self.assertIn('endDate=2026-09-27', request.call_args.args[0])

    def test_invalid_payloads_fail_for_fallback(self):
        for payload in ({'detail': 'invalid token'}, [], [{'date': 'bad', 'close': 1}],
                        [{'date': '2026-09-25', 'close': float('nan')}]):
            with self.subTest(payload=payload), patch.object(report, 'read_json_url', return_value=payload):
                with self.assertRaises(RuntimeError):
                    report.fetch_tiingo_quote('VOO', 'key')

    def test_fallback_for_quotes_and_series(self):
        with TemporaryDirectory() as directory:
            cache = Path(directory)
            quote = report.PriceQuote('VOO', 100, 'Tiingo EOD', '2026-09-25')
            with patch.object(report, 'fetch_fmp_quote', side_effect=RuntimeError('failed')), \
                 patch.object(report, 'fetch_tiingo_quote', return_value=quote) as tiingo, \
                 patch.object(report, 'fetch_yahoo_quote') as yahoo:
                self.assertEqual(report.load_latest_prices(['VOO'], None, cache, False, 'key')['VOO'], quote)
                tiingo.assert_called_once_with('VOO', 'key')
                yahoo.assert_not_called()
            with patch.object(report, 'fetch_fmp_historical_series', side_effect=RuntimeError('failed')), \
                 patch.object(report, 'fetch_tiingo_prices', side_effect=RuntimeError('failed')), \
                 patch.object(report, 'fetch_yahoo_historical_series', return_value={date(2026, 9, 25): 100}):
                self.assertEqual(report.load_historical_series(['VOO'], date(2026, 9, 25), date(2026, 9, 27), None, cache, False, 'key')['VOO'], {date(2026, 9, 25): 100})

    def test_http_error_does_not_expose_key(self):
        with patch.object(report, 'read_json_url', side_effect=HTTPError('https://example.com?token=secret', 401, 'Unauthorized', {}, None)):
            with self.assertRaises(RuntimeError) as error:
                report.fetch_tiingo_quote('VOO', 'secret')
        self.assertNotIn('secret', str(error.exception))


if __name__ == '__main__':
    unittest.main()
