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
            with patch.object(report, 'market_today', return_value=date(2026, 9, 25)), \
                 patch.object(report, 'fetch_tiingo_quote', return_value=quote) as tiingo, \
                 patch.object(report, 'fetch_tiingo_realtime_quote') as realtime:
                self.assertEqual(report.load_latest_prices(['VOO'], None, cache, False, 'key')['VOO'], quote)
                tiingo.assert_called_once_with('VOO', 'key')
                realtime.assert_not_called()
            with patch.object(report, 'fetch_fmp_historical_series', side_effect=RuntimeError('failed')), \
                 patch.object(report, 'fetch_tiingo_prices', side_effect=RuntimeError('failed')), \
                 patch.object(report, 'fetch_yahoo_historical_series', return_value={date(2026, 9, 25): 100}):
                self.assertEqual(report.load_historical_series(['VOO'], date(2026, 9, 25), date(2026, 9, 27), None, cache, False, 'key')['VOO'], {date(2026, 9, 25): 100})

    def test_missing_today_is_supplemented_with_realtime(self):
        with TemporaryDirectory() as directory:
            history = {'VOO': {date(2026, 10, 2): 100}}
            quote = report.PriceQuote('VOO', 101, 'Tiingo realtime reference', '2026-10-05T16:25:00-04:00')
            with patch.object(report, 'market_today', return_value=date(2026, 10, 5)), \
                 patch.object(report, 'fetch_tiingo_realtime_quote', return_value=quote), \
                 patch.object(report, 'fetch_yahoo_quote') as yahoo:
                actual = report.load_latest_prices(['VOO'], None, Path(directory), False, 'key', history)['VOO']
                self.assertEqual(actual, quote)
                yahoo.assert_not_called()
            report.supplement_price_history(history, {'VOO': actual})
            self.assertEqual(history['VOO'], {date(2026, 10, 2): 100, date(2026, 10, 5): 101})

    def test_today_eod_is_preserved_without_realtime(self):
        with TemporaryDirectory() as directory, \
             patch.object(report, 'market_today', return_value=date(2026, 10, 5)), \
             patch.object(report, 'fetch_tiingo_realtime_quote') as realtime:
            actual = report.fetch_latest_quote('VOO', None, Path(directory), False, 'key', {date(2026, 10, 5): 100})
            self.assertEqual(actual.price, 100)
            self.assertEqual(actual.price_time, '2026-10-05')
            realtime.assert_not_called()

    def test_weekend_snapshot_does_not_create_today_or_replace_eod(self):
        with TemporaryDirectory() as directory, \
             patch.object(report, 'market_today', return_value=date(2026, 10, 4)), \
             patch.object(report, 'fetch_tiingo_realtime_quote', return_value=report.PriceQuote('VOO', 101, 'realtime', '2026-10-03T00:10:00Z')):
            history = {'VOO': {date(2026, 10, 2): 100}}
            actual = report.fetch_latest_quote('VOO', None, Path(directory), False, 'key', history['VOO'])
            report.supplement_price_history(history, {'VOO': actual})
            self.assertEqual(actual.price, 100)
            self.assertEqual(history['VOO'], {date(2026, 10, 2): 100})

    def test_realtime_failure_uses_fallback(self):
        with TemporaryDirectory() as directory, \
             patch.object(report, 'market_today', return_value=date(2026, 10, 5)), \
             patch.object(report, 'fetch_tiingo_realtime_quote', side_effect=RuntimeError('unavailable')), \
             patch.object(report, 'fetch_fmp_quote', return_value=report.PriceQuote('VOO', 101, 'FMP', '2026-10-05 16:00:00')):
            actual = report.fetch_latest_quote('VOO', None, Path(directory), False, 'key', {date(2026, 10, 2): 100})
            self.assertEqual(actual.source, 'FMP')

    def test_all_realtime_sources_fail_keeps_eod_with_warning(self):
        with TemporaryDirectory() as directory, \
             patch.object(report, 'market_today', return_value=date(2026, 10, 5)), \
             patch.object(report, 'fetch_tiingo_realtime_quote', side_effect=RuntimeError('unavailable')), \
             patch.object(report, 'fetch_fmp_quote', side_effect=RuntimeError('unavailable')), \
             patch.object(report, 'fetch_yahoo_quote', side_effect=RuntimeError('unavailable')):
            with self.assertWarns(UserWarning):
                actual = report.fetch_latest_quote('VOO', None, Path(directory), False, 'key', {date(2026, 10, 2): 100})
            self.assertEqual(actual.price_time, '2026-10-02')

    def test_realtime_price_and_market_timezone(self):
        payload = [{'ticker': 'VOO', 'tngoLast': 101, 'prevClose': 99, 'timestamp': '2026-10-06T00:10:00Z'}]
        with patch.object(report, 'read_json_url', return_value=payload):
            quote = report.fetch_tiingo_realtime_quote('VOO', 'key')
        self.assertEqual(quote.price, 101)
        self.assertEqual(report.quote_date(quote), date(2026, 10, 5))
        self.assertEqual(report.valuation_date_from_quotes({'VOO': quote}, date(2026, 10, 6)), date(2026, 10, 5))

    def test_realtime_invalid_payloads_fail(self):
        for payload in (None, [], {'detail': 'denied'},
                        [{'ticker': 'VOO', 'tngoLast': None, 'timestamp': '2026-10-05T20:00:00Z'}],
                        [{'ticker': 'VOO', 'tngoLast': float('nan'), 'timestamp': '2026-10-05T20:00:00Z'}],
                        [{'ticker': 'VOO', 'tngoLast': 101, 'timestamp': None}],
                        [{'ticker': 'VOO', 'tngoLast': 101, 'timestamp': '2026-10-05'}]):
            with self.subTest(payload=payload), patch.object(report, 'read_json_url', return_value=payload):
                with self.assertRaises(RuntimeError):
                    report.fetch_tiingo_realtime_quote('VOO', 'key')

    def test_supplement_preserves_existing_daily_close(self):
        history = {'VOO': {date(2026, 10, 5): 100}}
        report.supplement_price_history(history, {'VOO': report.PriceQuote('VOO', 101, 'realtime', '2026-10-05T16:25:00-04:00')})
        self.assertEqual(history['VOO'][date(2026, 10, 5)], 100)

    def test_http_error_does_not_expose_key(self):
        with patch.object(report, 'read_json_url', side_effect=HTTPError('https://example.com?token=secret', 401, 'Unauthorized', {}, None)):
            with self.assertRaises(RuntimeError) as error:
                report.fetch_tiingo_quote('VOO', 'secret')
        self.assertNotIn('secret', str(error.exception))


if __name__ == '__main__':
    unittest.main()
