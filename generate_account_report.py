#!/usr/bin/env python3
"""Generate account performance reports from the trade record workbook.

The workbook is expected to contain two sheets:
  - UOB
  - IB

Latest prices are downloaded from Financial Modeling Prep (FMP) first. If FMP has
no API key or a quote request fails, Yahoo Finance is used as a fallback.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import ssl
from collections import defaultdict, deque
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

import pandas as pd

try:
    import certifi
except ImportError:  # pragma: no cover - optional local dependency
    certifi = None


SHEETS = {
    "UOB": "UOB",
    "IB": "IB",
}

REQUIRED_COLUMNS = ["Date", "Symbol", "Price", "Qty", "Comm Fee", "Trade Value"]
CASH_SYMBOL = "CASH"
BENCHMARK_SYMBOL = "VOO"
TRADING_DAYS_PER_YEAR = 252
FMP_QUOTE_URL = "https://financialmodelingprep.com/stable/quote"
FMP_HISTORY_URL = "https://financialmodelingprep.com/stable/historical-price-eod/full"
YAHOO_CHART_URL = "https://query1.finance.yahoo.com/v8/finance/chart/{symbol}"


def ssl_context() -> ssl.SSLContext | None:
    if certifi is None:
        return None
    return ssl.create_default_context(cafile=certifi.where())


@dataclass
class Lot:
    qty: float
    unit_cost: float


@dataclass
class AccountReport:
    account: str
    trades: pd.DataFrame
    positions: pd.DataFrame
    realized_pnl: float
    unrealized_pnl: float
    total_pnl: float
    invested_capital: float
    net_cash_flow: float
    cash_balance: float
    ending_value: float
    return_pct: float | None
    annualized_return_pct: float | None
    money_weighted_return_pct: float | None
    time_weighted_return_pct: float | None
    max_drawdown_pct: float | None
    sharpe_ratio: float | None
    benchmark_start_price: float | None
    benchmark_end_price: float | None
    benchmark_return_pct: float | None
    relative_to_benchmark_pct: float | None
    matched_benchmark_ending_value: float | None
    matched_benchmark_pnl: float | None
    matched_benchmark_return_pct: float | None
    matched_benchmark_money_weighted_return_pct: float | None
    relative_money_weighted_return_pct: float | None
    ytd_total_pnl: float | None
    first_trade_date: date | None
    last_valuation_date: date | None


@dataclass
class PriceQuote:
    symbol: str
    price: float
    source: str
    price_time: str | None = None
    fetched_at: str | None = None


@dataclass
class CashFlowBenchmark:
    ending_value: float
    pnl: float
    return_pct: float | None
    money_weighted_return_pct: float | None


@dataclass
class DailyPerformance:
    history: pd.DataFrame
    time_weighted_return_pct: float | None
    max_drawdown_pct: float | None
    sharpe_ratio: float | None


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate separate and combined performance reports from the trade workbook."
    )
    parser.add_argument(
        "--input",
        default="James_Trade_Records.xlsx",
        help=(
            "Path to the trade record workbook. "
            "Defaults to James_Trade_Records.xlsx in the current directory."
        ),
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Path for the Markdown report. Defaults to account_report_YYYY-MM-DD.md.",
    )
    parser.add_argument(
        "--as-of",
        default=None,
        help=(
            "Optional valuation date in YYYY-MM-DD format. "
            "If omitted, latest quotes are used. If provided, prices are as of "
            "the closest trading day on or before this date."
        ),
    )
    parser.add_argument(
        "--api-key",
        default=None,
        help="FMP API key. Defaults to FMP_API_KEY from environment or .env.",
    )
    parser.add_argument(
        "--cache-dir",
        default=".price_cache",
        help="Directory for cached latest-price responses.",
    )
    parser.add_argument(
        "--force-refresh",
        action="store_true",
        help="Deprecated: latest prices are refreshed on every run.",
    )
    parser.add_argument(
        "--allow-cache-fallback",
        action="store_true",
        help="Use cached prices only if both FMP and Yahoo live quote requests fail.",
    )
    parser.add_argument(
        "--risk-free-rate",
        type=float,
        default=0.0,
        help=(
            "Annual risk-free rate used for the Sharpe ratio, expressed as a "
            "decimal. Defaults to 0.0."
        ),
    )
    return parser.parse_args()


def load_dotenv(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    if not path.exists():
        return values

    for raw_line in path.read_text().splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip("\"'")
        if key:
            values[key] = value
    return values


def get_api_key(cli_api_key: str | None, dotenv_path: Path) -> str | None:
    if cli_api_key:
        return cli_api_key
    if os.environ.get("FMP_API_KEY"):
        return os.environ["FMP_API_KEY"]
    return load_dotenv(dotenv_path).get("FMP_API_KEY")


def money(value: float | None) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "n/a"
    return f"${value:,.2f}"


def pct(value: float | None) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "n/a"
    return f"{value:.2%}"


def ratio(value: float | None) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "n/a"
    return f"{value:.2f}"


def read_trades(workbook_path: Path) -> dict[str, pd.DataFrame]:
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    result: dict[str, pd.DataFrame] = {}
    for sheet_name, account_name in SHEETS.items():
        df = pd.read_excel(workbook_path, sheet_name=sheet_name)
        missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
        if missing:
            raise ValueError(f"Sheet {sheet_name!r} is missing columns: {missing}")

        df = df[REQUIRED_COLUMNS].copy()
        df = df.dropna(subset=["Date", "Symbol", "Price", "Qty"])
        df["Date"] = pd.to_datetime(df["Date"]).dt.date
        df["Symbol"] = df["Symbol"].astype(str).str.upper().str.strip()
        for col in ["Price", "Qty", "Comm Fee", "Trade Value"]:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
        df = df[df["Qty"] != 0].sort_values(["Date", "Symbol"]).reset_index(drop=True)
        df["Account"] = account_name
        result[account_name] = df
    return result


def is_cash_symbol(symbol: str) -> bool:
    return str(symbol).upper().strip() == CASH_SYMBOL


def read_json_url(url: str) -> Any:
    request = Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126 Safari/537.36"
            )
        },
    )
    with urlopen(request, timeout=30, context=ssl_context()) as response:
        return json.loads(response.read().decode("utf-8"))


def fetch_fmp_quote(symbol: str, api_key: str | None) -> PriceQuote:
    if not api_key:
        raise RuntimeError("FMP API key is not set")

    query = urlencode({"symbol": symbol, "apikey": api_key})
    url = f"{FMP_QUOTE_URL}?{query}"
    try:
        payload = read_json_url(url)
    except HTTPError as exc:
        raise RuntimeError(f"FMP HTTP error for {symbol}: {exc.code} {exc.reason}") from exc
    except URLError as exc:
        raise RuntimeError(f"FMP network error for {symbol}: {exc.reason}") from exc

    if not isinstance(payload, list) or not payload:
        raise RuntimeError(f"FMP returned no quote for {symbol}")
    quote = payload[0]
    if "price" not in quote or quote["price"] is None:
        raise RuntimeError(f"FMP quote for {symbol} did not include price")

    price_time = None
    if quote.get("timestamp"):
        price_time = datetime.fromtimestamp(int(quote["timestamp"])).strftime("%Y-%m-%d %H:%M:%S")
    return PriceQuote(
        symbol=symbol,
        price=float(quote["price"]),
        source="FMP",
        price_time=price_time,
        fetched_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )


def fetch_yahoo_quote(symbol: str) -> PriceQuote:
    query = urlencode({"range": "5d", "interval": "1d"})
    url = f"{YAHOO_CHART_URL.format(symbol=symbol)}?{query}"
    try:
        payload = read_json_url(url)
    except HTTPError as exc:
        raise RuntimeError(f"Yahoo Finance HTTP error for {symbol}: {exc.code} {exc.reason}") from exc
    except URLError as exc:
        raise RuntimeError(f"Yahoo Finance network error for {symbol}: {exc.reason}") from exc

    chart = payload.get("chart", {}) if isinstance(payload, dict) else {}
    error = chart.get("error")
    if error:
        raise RuntimeError(f"Yahoo Finance error for {symbol}: {error}")
    results = chart.get("result") or []
    if not results:
        raise RuntimeError(f"Yahoo Finance returned no quote for {symbol}")

    result = results[0]
    meta = result.get("meta", {})
    price = meta.get("regularMarketPrice")
    timestamp = meta.get("regularMarketTime")

    if price is None:
        closes = (
            result.get("indicators", {})
            .get("quote", [{}])[0]
            .get("close", [])
        )
        valid_closes = [close for close in closes if close is not None]
        if valid_closes:
            price = valid_closes[-1]

    if price is None:
        raise RuntimeError(f"Yahoo Finance quote for {symbol} did not include price")

    price_time = None
    if timestamp:
        price_time = datetime.fromtimestamp(int(timestamp)).strftime("%Y-%m-%d %H:%M:%S")
    return PriceQuote(
        symbol=symbol,
        price=float(price),
        source="Yahoo Finance",
        price_time=price_time,
        fetched_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )


def latest_price_on_or_before(
    rows: list[tuple[date, float]],
    target_date: date,
) -> tuple[date, float] | None:
    valid_rows = [(row_date, price) for row_date, price in rows if row_date <= target_date]
    if not valid_rows:
        return None
    return sorted(valid_rows, key=lambda item: item[0])[-1]


def fetch_fmp_historical_quote(
    symbol: str,
    target_date: date,
    api_key: str | None,
) -> PriceQuote:
    if not api_key:
        raise RuntimeError("FMP API key is not set")

    from_date = target_date - timedelta(days=10)
    query = urlencode(
        {
            "symbol": symbol,
            "from": from_date.isoformat(),
            "to": target_date.isoformat(),
            "apikey": api_key,
        }
    )
    url = f"{FMP_HISTORY_URL}?{query}"
    try:
        payload = read_json_url(url)
    except HTTPError as exc:
        raise RuntimeError(f"FMP historical HTTP error for {symbol}: {exc.code} {exc.reason}") from exc
    except URLError as exc:
        raise RuntimeError(f"FMP historical network error for {symbol}: {exc.reason}") from exc

    if not isinstance(payload, list) or not payload:
        raise RuntimeError(f"FMP returned no historical prices for {symbol}")

    rows = []
    for item in payload:
        if "date" in item and "close" in item and item["close"] is not None:
            rows.append((pd.to_datetime(item["date"]).date(), float(item["close"])))
    dated_price = latest_price_on_or_before(rows, target_date)
    if dated_price is None:
        raise RuntimeError(f"FMP historical prices for {symbol} did not include a price on or before {target_date}")
    price_date, price = dated_price

    return PriceQuote(
        symbol=symbol,
        price=price,
        source="FMP historical",
        price_time=price_date.isoformat(),
        fetched_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )


def fetch_yahoo_historical_quote(symbol: str, target_date: date) -> PriceQuote:
    start_dt = datetime.combine(target_date - timedelta(days=10), datetime.min.time())
    end_dt = datetime.combine(target_date + timedelta(days=1), datetime.min.time())
    query = urlencode(
        {
            "period1": int(start_dt.timestamp()),
            "period2": int(end_dt.timestamp()),
            "interval": "1d",
        }
    )
    url = f"{YAHOO_CHART_URL.format(symbol=symbol)}?{query}"
    try:
        payload = read_json_url(url)
    except HTTPError as exc:
        raise RuntimeError(f"Yahoo Finance historical HTTP error for {symbol}: {exc.code} {exc.reason}") from exc
    except URLError as exc:
        raise RuntimeError(f"Yahoo Finance historical network error for {symbol}: {exc.reason}") from exc

    chart = payload.get("chart", {}) if isinstance(payload, dict) else {}
    error = chart.get("error")
    if error:
        raise RuntimeError(f"Yahoo Finance historical error for {symbol}: {error}")
    results = chart.get("result") or []
    if not results:
        raise RuntimeError(f"Yahoo Finance returned no historical prices for {symbol}")

    result = results[0]
    timestamps = result.get("timestamp") or []
    closes = (
        result.get("indicators", {})
        .get("quote", [{}])[0]
        .get("close", [])
    )
    rows = []
    for timestamp, close in zip(timestamps, closes):
        if close is None:
            continue
        row_date = datetime.fromtimestamp(int(timestamp)).date()
        rows.append((row_date, float(close)))
    dated_price = latest_price_on_or_before(rows, target_date)
    if dated_price is None:
        raise RuntimeError(f"Yahoo Finance historical prices for {symbol} did not include a price on or before {target_date}")
    price_date, price = dated_price

    return PriceQuote(
        symbol=symbol,
        price=price,
        source="Yahoo Finance historical",
        price_time=price_date.isoformat(),
        fetched_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )


def fetch_fmp_historical_series(
    symbol: str,
    start_date: date,
    end_date: date,
    api_key: str | None,
) -> dict[date, float]:
    if not api_key:
        raise RuntimeError("FMP API key is not set")

    query = urlencode(
        {
            "symbol": symbol,
            "from": start_date.isoformat(),
            "to": end_date.isoformat(),
            "apikey": api_key,
        }
    )
    url = f"{FMP_HISTORY_URL}?{query}"
    try:
        payload = read_json_url(url)
    except HTTPError as exc:
        raise RuntimeError(
            f"FMP historical-series HTTP error for {symbol}: {exc.code} {exc.reason}"
        ) from exc
    except URLError as exc:
        raise RuntimeError(
            f"FMP historical-series network error for {symbol}: {exc.reason}"
        ) from exc

    if not isinstance(payload, list) or not payload:
        raise RuntimeError(f"FMP returned no historical series for {symbol}")
    prices = {
        pd.to_datetime(item["date"]).date(): float(item["close"])
        for item in payload
        if item.get("date") and item.get("close") is not None
    }
    if not prices:
        raise RuntimeError(f"FMP historical series for {symbol} had no closing prices")
    return prices


def fetch_yahoo_historical_series(
    symbol: str,
    start_date: date,
    end_date: date,
) -> dict[date, float]:
    start_dt = datetime.combine(start_date, datetime.min.time())
    end_dt = datetime.combine(end_date + timedelta(days=1), datetime.min.time())
    query = urlencode(
        {
            "period1": int(start_dt.timestamp()),
            "period2": int(end_dt.timestamp()),
            "interval": "1d",
        }
    )
    url = f"{YAHOO_CHART_URL.format(symbol=symbol)}?{query}"
    try:
        payload = read_json_url(url)
    except HTTPError as exc:
        raise RuntimeError(
            f"Yahoo Finance historical-series HTTP error for {symbol}: "
            f"{exc.code} {exc.reason}"
        ) from exc
    except URLError as exc:
        raise RuntimeError(
            f"Yahoo Finance historical-series network error for {symbol}: {exc.reason}"
        ) from exc

    chart = payload.get("chart", {}) if isinstance(payload, dict) else {}
    error = chart.get("error")
    if error:
        raise RuntimeError(f"Yahoo Finance historical-series error for {symbol}: {error}")
    results = chart.get("result") or []
    if not results:
        raise RuntimeError(f"Yahoo Finance returned no historical series for {symbol}")

    result = results[0]
    timestamps = result.get("timestamp") or []
    closes = result.get("indicators", {}).get("quote", [{}])[0].get("close", [])
    prices = {
        datetime.fromtimestamp(int(timestamp)).date(): float(close)
        for timestamp, close in zip(timestamps, closes)
        if close is not None
    }
    if not prices:
        raise RuntimeError(
            f"Yahoo Finance historical series for {symbol} had no closing prices"
        )
    return prices


def fetch_historical_series(
    symbol: str,
    start_date: date,
    end_date: date,
    api_key: str | None,
    cache_dir: Path,
    allow_cache_fallback: bool,
) -> dict[date, float]:
    cache_dir.mkdir(parents=True, exist_ok=True)
    cache_file = cache_dir / (
        f"{symbol}_{start_date.isoformat()}_{end_date.isoformat()}_series.json"
    )

    errors = []
    try:
        prices = fetch_fmp_historical_series(symbol, start_date, end_date, api_key)
    except RuntimeError as exc:
        errors.append(str(exc))
        try:
            prices = fetch_yahoo_historical_series(symbol, start_date, end_date)
        except RuntimeError as yahoo_exc:
            errors.append(str(yahoo_exc))
            if cache_file.exists() and allow_cache_fallback:
                cached = json.loads(cache_file.read_text())
                return {
                    pd.to_datetime(row["date"]).date(): float(row["price"])
                    for row in cached["prices"]
                }
            raise RuntimeError(
                f"Could not fetch historical series for {symbol}: {'; '.join(errors)}"
            ) from yahoo_exc

    cache_file.write_text(
        json.dumps(
            {
                "symbol": symbol,
                "start_date": start_date.isoformat(),
                "end_date": end_date.isoformat(),
                "prices": [
                    {"date": price_date.isoformat(), "price": price}
                    for price_date, price in sorted(prices.items())
                ],
            },
            indent=2,
            sort_keys=True,
        )
    )
    return prices


def fetch_latest_quote(
    symbol: str,
    api_key: str | None,
    cache_dir: Path,
    allow_cache_fallback: bool,
) -> PriceQuote:
    cache_dir.mkdir(parents=True, exist_ok=True)
    cache_file = cache_dir / f"{symbol}_latest.json"

    errors = []
    try:
        quote = fetch_fmp_quote(symbol, api_key)
    except RuntimeError as exc:
        errors.append(str(exc))
        try:
            quote = fetch_yahoo_quote(symbol)
        except RuntimeError as yahoo_exc:
            errors.append(str(yahoo_exc))
            if cache_file.exists() and allow_cache_fallback:
                cached = PriceQuote(**json.loads(cache_file.read_text()))
                cached.source = f"cached {cached.source}"
                return cached
            raise RuntimeError(f"Could not fetch latest price for {symbol}: {'; '.join(errors)}") from yahoo_exc

    cache_file.write_text(json.dumps(quote.__dict__, indent=2, sort_keys=True))
    return quote


def fetch_historical_quote(
    symbol: str,
    target_date: date,
    api_key: str | None,
    cache_dir: Path,
    allow_cache_fallback: bool,
) -> PriceQuote:
    cache_dir.mkdir(parents=True, exist_ok=True)
    cache_file = cache_dir / f"{symbol}_{target_date.isoformat()}_historical.json"

    errors = []
    try:
        quote = fetch_fmp_historical_quote(symbol, target_date, api_key)
    except RuntimeError as exc:
        errors.append(str(exc))
        try:
            quote = fetch_yahoo_historical_quote(symbol, target_date)
        except RuntimeError as yahoo_exc:
            errors.append(str(yahoo_exc))
            if cache_file.exists() and allow_cache_fallback:
                cached = PriceQuote(**json.loads(cache_file.read_text()))
                cached.source = f"cached {cached.source}"
                return cached
            raise RuntimeError(
                f"Could not fetch historical price for {symbol} on {target_date}: {'; '.join(errors)}"
            ) from yahoo_exc

    cache_file.write_text(json.dumps(quote.__dict__, indent=2, sort_keys=True))
    return quote


def load_latest_prices(
    symbols: list[str],
    api_key: str | None,
    cache_dir: Path,
    allow_cache_fallback: bool,
) -> dict[str, PriceQuote]:
    return {
        symbol: fetch_latest_quote(symbol, api_key, cache_dir, allow_cache_fallback)
        for symbol in symbols
    }


def load_historical_prices(
    symbols: list[str],
    target_date: date,
    api_key: str | None,
    cache_dir: Path,
    allow_cache_fallback: bool,
) -> dict[str, PriceQuote]:
    return {
        symbol: fetch_historical_quote(
            symbol,
            target_date,
            api_key,
            cache_dir,
            allow_cache_fallback,
        )
        for symbol in symbols
    }


def load_historical_series(
    symbols: list[str],
    start_date: date,
    end_date: date,
    api_key: str | None,
    cache_dir: Path,
    allow_cache_fallback: bool,
) -> dict[str, dict[date, float]]:
    return {
        symbol: fetch_historical_series(
            symbol,
            start_date,
            end_date,
            api_key,
            cache_dir,
            allow_cache_fallback,
        )
        for symbol in symbols
    }


def calculate_fifo(
    trades: pd.DataFrame,
    through_date: date,
) -> tuple[float, float, dict[str, float], dict[str, float]]:
    lots: dict[str, deque[Lot]] = defaultdict(deque)
    realized_pnl = 0.0
    invested_capital = 0.0

    day_trades = trades[trades["Date"] <= through_date]
    for _, row in day_trades.iterrows():
        symbol = row["Symbol"]
        if is_cash_symbol(symbol):
            continue
        qty = float(row["Qty"])
        price = float(row["Price"])
        fee = float(row["Comm Fee"])

        if qty > 0:
            gross_cost = qty * price
            total_cost = gross_cost + fee
            invested_capital += total_cost
            lots[symbol].append(Lot(qty=qty, unit_cost=total_cost / qty))
            continue

        sell_qty = -qty
        gross_proceeds = sell_qty * price
        net_proceeds = gross_proceeds - fee
        remaining_to_match = sell_qty
        matched_cost = 0.0

        while remaining_to_match > 1e-9 and lots[symbol]:
            lot = lots[symbol][0]
            matched_qty = min(remaining_to_match, lot.qty)
            matched_cost += matched_qty * lot.unit_cost
            lot.qty -= matched_qty
            remaining_to_match -= matched_qty
            if lot.qty <= 1e-9:
                lots[symbol].popleft()

        if remaining_to_match > 1e-9:
            raise ValueError(
                f"{through_date}: sell quantity for {symbol} exceeds existing FIFO position"
            )
        realized_pnl += net_proceeds - matched_cost

    open_qty: dict[str, float] = defaultdict(float)
    open_cost: dict[str, float] = defaultdict(float)
    for symbol, symbol_lots in lots.items():
        for lot in symbol_lots:
            open_qty[symbol] += lot.qty
            open_cost[symbol] += lot.qty * lot.unit_cost

    return realized_pnl, invested_capital, dict(open_qty), dict(open_cost)


def calculate_cash_metrics(trades: pd.DataFrame, through_date: date) -> tuple[float, float]:
    day_trades = trades[trades["Date"] <= through_date]
    net_cash_flow = 0.0
    security_cash_effect = 0.0

    for _, row in day_trades.iterrows():
        symbol = row["Symbol"]
        qty = float(row["Qty"])
        price = float(row["Price"])
        fee = float(row["Comm Fee"])

        if is_cash_symbol(symbol):
            net_cash_flow += qty * price - fee
        elif qty > 0:
            security_cash_effect -= qty * price + fee
        else:
            security_cash_effect += (-qty) * price - fee

    return net_cash_flow, net_cash_flow + security_cash_effect


def external_cash_flows(
    trades: pd.DataFrame,
    through_date: date,
) -> list[tuple[date, float]]:
    cash_rows = trades[
        (trades["Date"] <= through_date) & trades["Symbol"].map(is_cash_symbol)
    ]
    flows_by_date: dict[date, float] = defaultdict(float)
    for _, row in cash_rows.iterrows():
        amount = float(row["Qty"]) * float(row["Price"]) - float(row["Comm Fee"])
        flows_by_date[row["Date"]] += amount
    return [
        (flow_date, amount)
        for flow_date, amount in sorted(flows_by_date.items())
        if abs(amount) > 1e-9
    ]


def xnpv(rate: float, cash_flows: list[tuple[date, float]]) -> float:
    if rate <= -1:
        raise ValueError("XNPV rate must be greater than -100%")
    start_date = min(flow_date for flow_date, _ in cash_flows)
    return sum(
        amount / (1 + rate) ** ((flow_date - start_date).days / 365)
        for flow_date, amount in cash_flows
    )


def calculate_xirr(cash_flows: list[tuple[date, float]]) -> float | None:
    flows_by_date: dict[date, float] = defaultdict(float)
    for flow_date, amount in cash_flows:
        flows_by_date[flow_date] += amount
    flows = [
        (flow_date, amount)
        for flow_date, amount in sorted(flows_by_date.items())
        if abs(amount) > 1e-9
    ]
    if not flows:
        return None
    if flows[0][0] == flows[-1][0]:
        return None
    amounts = [amount for _, amount in flows]
    if not any(amount < 0 for amount in amounts) or not any(
        amount > 0 for amount in amounts
    ):
        return None

    # Newton's method matches the conventional 10% starting guess used by XIRR.
    rate = 0.10
    start_date = flows[0][0]
    for _ in range(100):
        value = xnpv(rate, flows)
        derivative = sum(
            -((flow_date - start_date).days / 365)
            * amount
            / (1 + rate) ** (((flow_date - start_date).days / 365) + 1)
            for flow_date, amount in flows
        )
        if abs(value) <= 1e-7:
            return rate
        if abs(derivative) <= 1e-12:
            break
        next_rate = rate - value / derivative
        if not math.isfinite(next_rate) or next_rate <= -0.999999:
            break
        rate = next_rate

    # Fall back to a bracket search when Newton's method is unstable.
    candidates = [
        -0.9999,
        -0.99,
        -0.9,
        -0.75,
        -0.5,
        -0.25,
        0.0,
        0.1,
        0.25,
        0.5,
        1.0,
        2.0,
        5.0,
        10.0,
        100.0,
        1000.0,
    ]
    previous_rate = candidates[0]
    previous_value = xnpv(previous_rate, flows)
    for candidate_rate in candidates[1:]:
        candidate_value = xnpv(candidate_rate, flows)
        if candidate_value == 0:
            return candidate_rate
        if previous_value * candidate_value < 0:
            low, high = previous_rate, candidate_rate
            low_value = previous_value
            for _ in range(200):
                midpoint = (low + high) / 2
                midpoint_value = xnpv(midpoint, flows)
                if abs(midpoint_value) <= 1e-7:
                    return midpoint
                if low_value * midpoint_value <= 0:
                    high = midpoint
                else:
                    low = midpoint
                    low_value = midpoint_value
            return (low + high) / 2
        previous_rate = candidate_rate
        previous_value = candidate_value
    return None


def calculate_account_xirr(
    trades: pd.DataFrame,
    valuation_date: date,
    ending_value: float,
) -> float | None:
    investor_flows = [
        (flow_date, -amount)
        for flow_date, amount in external_cash_flows(trades, valuation_date)
    ]
    investor_flows.append((valuation_date, ending_value))
    return calculate_xirr(investor_flows)


def build_cash_flow_benchmark(
    trades: pd.DataFrame,
    valuation_date: date,
    cash_flow_quotes: dict[date, PriceQuote],
    end_quote: PriceQuote | None,
) -> CashFlowBenchmark | None:
    flows = external_cash_flows(trades, valuation_date)
    if not flows or end_quote is None:
        return None

    units = 0.0
    investor_flows = []
    net_cash_flow = 0.0
    for flow_date, amount in flows:
        quote = cash_flow_quotes.get(flow_date)
        if quote is None or quote.price <= 0:
            return None
        units += amount / quote.price
        net_cash_flow += amount
        investor_flows.append((flow_date, -amount))

    ending_value = units * end_quote.price
    pnl = ending_value - net_cash_flow
    return_base = net_cash_flow if net_cash_flow > 1e-9 else None
    return_pct = pnl / return_base if return_base else None
    investor_flows.append((valuation_date, ending_value))
    return CashFlowBenchmark(
        ending_value=ending_value,
        pnl=pnl,
        return_pct=return_pct,
        money_weighted_return_pct=calculate_xirr(investor_flows),
    )


def calculate_performance_metrics(
    history: pd.DataFrame,
    annual_risk_free_rate: float,
) -> DailyPerformance:
    if history.empty:
        return DailyPerformance(history, None, None, None)

    time_weighted_return_pct = float(history["TWR Index"].iloc[-1] / 100 - 1)
    running_peak = history["TWR Index"].cummax()
    drawdowns = history["TWR Index"] / running_peak - 1
    max_drawdown_pct = float(drawdowns.min())

    trading_returns = history.loc[
        history["Trading Day"],
        "Daily Return",
    ]
    sharpe_ratio = None
    if len(trading_returns) >= 2:
        daily_risk_free_rate = (
            (1 + annual_risk_free_rate) ** (1 / TRADING_DAYS_PER_YEAR) - 1
        )
        excess_returns = trading_returns - daily_risk_free_rate
        volatility = float(excess_returns.std(ddof=1))
        if volatility > 1e-12:
            sharpe_ratio = (
                float(excess_returns.mean())
                / volatility
                * math.sqrt(TRADING_DAYS_PER_YEAR)
            )

    return DailyPerformance(
        history=history,
        time_weighted_return_pct=time_weighted_return_pct,
        max_drawdown_pct=max_drawdown_pct,
        sharpe_ratio=sharpe_ratio,
    )


def build_daily_performance(
    trades: pd.DataFrame,
    valuation_date: date,
    price_history: dict[str, dict[date, float]],
    annual_risk_free_rate: float,
) -> DailyPerformance:
    eligible_trades = trades[trades["Date"] <= valuation_date].copy()
    if eligible_trades.empty:
        return DailyPerformance(pd.DataFrame(), None, None, None)

    start_date = eligible_trades["Date"].min()
    benchmark_dates = {
        price_date
        for price_date in price_history.get(BENCHMARK_SYMBOL, {})
        if start_date <= price_date <= valuation_date
    }
    transaction_dates = set(eligible_trades["Date"])
    valuation_dates = sorted(
        benchmark_dates | transaction_dates | {start_date, valuation_date}
    )

    price_frame = pd.DataFrame(
        {
            symbol: pd.Series(prices, dtype=float)
            for symbol, prices in price_history.items()
        }
    ).sort_index()
    price_frame = price_frame.reindex(valuation_dates).ffill()

    trades_by_date = {
        trade_date: rows
        for trade_date, rows in eligible_trades.groupby("Date", sort=True)
    }
    quantities: dict[str, float] = defaultdict(float)
    cash_balance = 0.0
    previous_value: float | None = None
    twr_index = 100.0
    rows = []

    for valuation_day in valuation_dates:
        external_flow = 0.0
        for _, trade in trades_by_date.get(
            valuation_day,
            pd.DataFrame(),
        ).iterrows():
            symbol = trade["Symbol"]
            qty = float(trade["Qty"])
            trade_price = float(trade["Price"])
            fee = float(trade["Comm Fee"])
            if is_cash_symbol(symbol):
                flow = qty * trade_price - fee
                cash_balance += flow
                external_flow += flow
            else:
                cash_balance -= qty * trade_price + fee
                quantities[symbol] += qty

        market_value = 0.0
        for symbol, qty in quantities.items():
            if abs(qty) <= 1e-9:
                continue
            if symbol not in price_frame.columns:
                raise RuntimeError(f"No daily price history for {symbol}")
            price = price_frame.at[valuation_day, symbol]
            if pd.isna(price):
                raise RuntimeError(
                    f"No daily price for {symbol} on or before {valuation_day}"
                )
            market_value += qty * float(price)

        ending_value = cash_balance + market_value
        if previous_value is None:
            if external_flow <= 1e-9:
                raise RuntimeError(
                    f"Cannot start TWR on {valuation_day}: a positive initial "
                    "CASH contribution is required"
                )
            daily_return = (ending_value - external_flow) / external_flow
            twr_index *= 1 + daily_return
        else:
            return_base = previous_value + external_flow
            if return_base <= 1e-9:
                raise RuntimeError(
                    f"Cannot calculate TWR on {valuation_day}: beginning value plus "
                    f"external cash flow is not positive"
                )
            daily_return = (ending_value - previous_value - external_flow) / return_base
            twr_index *= 1 + daily_return

        rows.append(
            {
                "Date": valuation_day,
                "Cash Flow": external_flow,
                "Cash Balance": cash_balance,
                "Market Value": market_value,
                "Ending Value": ending_value,
                "Daily Return": daily_return,
                "TWR Index": twr_index,
                "Trading Day": valuation_day in benchmark_dates,
            }
        )
        previous_value = ending_value

    history = pd.DataFrame(rows).set_index("Date")
    return calculate_performance_metrics(history, annual_risk_free_rate)


def build_benchmark_curve(
    prices: dict[date, float],
    start_date: date,
    end_date: date,
) -> pd.DataFrame:
    dated_prices = sorted(
        (price_date, price)
        for price_date, price in prices.items()
        if price_date <= end_date
    )
    start_price = latest_price_on_or_before(dated_prices, start_date)
    if start_price is None:
        return pd.DataFrame()
    start_price_date, base_price = start_price
    rows = [
        {
            "Date": price_date,
            "TWR Index": price / base_price * 100,
        }
        for price_date, price in dated_prices
        if start_price_date <= price_date <= end_date
    ]
    return pd.DataFrame(rows).set_index("Date")


def build_excess_return_curve(
    account_curve: pd.DataFrame,
    benchmark_curve: pd.DataFrame,
) -> pd.DataFrame:
    if account_curve.empty or benchmark_curve.empty:
        return pd.DataFrame()
    combined_index = account_curve.index.union(benchmark_curve.index).sort_values()
    benchmark_aligned = (
        benchmark_curve["TWR Index"]
        .reindex(combined_index)
        .ffill()
        .reindex(account_curve.index)
    )
    excess_return = (
        account_curve["TWR Index"] / 100 - benchmark_aligned / 100
    ).dropna()
    return pd.DataFrame({"Excess Return": excess_return})


def write_nav_chart(
    output_path: Path,
    curves: dict[str, pd.DataFrame],
    excess_return_curve: pd.DataFrame | None = None,
) -> None:
    usable_curves = {
        name: curve.dropna(subset=["TWR Index"])
        for name, curve in curves.items()
        if not curve.empty
    }
    usable_curves = {name: curve for name, curve in usable_curves.items() if not curve.empty}
    if not usable_curves:
        return

    usable_excess = (
        excess_return_curve.dropna(subset=["Excess Return"])
        if excess_return_curve is not None and not excess_return_curve.empty
        else pd.DataFrame()
    )

    # Matplotlib needs a writable cache in sandboxed and CI environments.
    chart_cache = Path(os.environ.get("TMPDIR", "/tmp")) / "performance-chart-cache"
    chart_cache.mkdir(parents=True, exist_ok=True)
    os.environ.setdefault("MPLCONFIGDIR", str(chart_cache / "matplotlib"))
    os.environ.setdefault("XDG_CACHE_HOME", str(chart_cache))

    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.dates as mdates
    import matplotlib.pyplot as plt
    import matplotlib.ticker as mtick

    figure, left_axis = plt.subplots(figsize=(12, 6.4), dpi=150)
    colors = ["#2563eb", "#16a34a", "#f97316", "#7c3aed"]

    for index, (name, curve) in enumerate(usable_curves.items()):
        left_axis.plot(
            pd.to_datetime(curve.index),
            curve["TWR Index"],
            color=colors[index % len(colors)],
            linewidth=2.2,
            label=name,
        )

    left_axis.set_title("TWR Net Value and Excess Return", loc="left", fontsize=16, weight="bold")
    left_axis.set_ylabel("Net value (initial capital = 100)")
    left_axis.grid(axis="both", color="#e5e7eb", linewidth=0.8)
    left_axis.set_axisbelow(True)
    date_locator = mdates.AutoDateLocator(minticks=4, maxticks=7)
    left_axis.xaxis.set_major_locator(date_locator)
    left_axis.xaxis.set_major_formatter(mdates.ConciseDateFormatter(date_locator))

    handles, labels = left_axis.get_legend_handles_labels()
    if not usable_excess.empty:
        right_axis = left_axis.twinx()
        right_axis.plot(
            pd.to_datetime(usable_excess.index),
            usable_excess["Excess Return"],
            color="#dc2626",
            linewidth=2.2,
            linestyle="--",
            label="Excess return",
        )
        right_axis.axhline(0, color="#fecaca", linewidth=1, linestyle="--")
        right_axis.set_ylabel("Cumulative excess return", color="#dc2626")
        right_axis.tick_params(axis="y", colors="#dc2626")
        right_axis.spines["right"].set_color("#dc2626")
        right_axis.yaxis.set_major_formatter(mtick.PercentFormatter(1.0))
        right_handles, right_labels = right_axis.get_legend_handles_labels()
        handles += right_handles
        labels += right_labels

    left_axis.legend(handles, labels, loc="upper left", ncol=len(labels), frameon=False)
    figure.tight_layout()
    figure.savefig(output_path, format="png", dpi=150, facecolor="white", bbox_inches="tight")
    plt.close(figure)


def process_trades_until(
    trades: pd.DataFrame,
    through_date: date,
    prices_on_date: pd.Series,
) -> tuple[float, float, float, dict[str, float], dict[str, float], float]:
    realized_pnl, invested_capital, open_qty, open_cost = calculate_fifo(trades, through_date)
    market_value = 0.0
    for symbol, qty in open_qty.items():
        if qty == 0:
            continue
        if symbol not in prices_on_date or pd.isna(prices_on_date[symbol]):
            raise RuntimeError(f"No market price for {symbol} on or before {through_date}")
        market_value += qty * float(prices_on_date[symbol])

    unrealized_pnl = market_value - sum(open_cost.values())
    return (
        realized_pnl,
        unrealized_pnl,
        invested_capital,
        dict(open_qty),
        dict(open_cost),
        market_value,
    )


def build_account_report(
    account: str,
    trades: pd.DataFrame,
    prices: pd.DataFrame,
    quotes: dict[str, PriceQuote],
    ytd_start_prices: pd.DataFrame,
    ytd_start_date: date,
    benchmark_start_quote: PriceQuote | None,
    benchmark_end_quote: PriceQuote | None,
    benchmark_cash_flow_quotes: dict[date, PriceQuote],
    daily_performance: DailyPerformance,
) -> AccountReport:
    first_trade = trades["Date"].min() if not trades.empty else None

    if trades.empty or len(prices.index) == 0:
        return AccountReport(
            account=account,
            trades=trades,
            positions=pd.DataFrame(),
            realized_pnl=0.0,
            unrealized_pnl=0.0,
            total_pnl=0.0,
            invested_capital=0.0,
            net_cash_flow=0.0,
            cash_balance=0.0,
            ending_value=0.0,
            return_pct=None,
            annualized_return_pct=None,
            money_weighted_return_pct=None,
            time_weighted_return_pct=None,
            max_drawdown_pct=None,
            sharpe_ratio=None,
            benchmark_start_price=None,
            benchmark_end_price=None,
            benchmark_return_pct=None,
            relative_to_benchmark_pct=None,
            matched_benchmark_ending_value=None,
            matched_benchmark_pnl=None,
            matched_benchmark_return_pct=None,
            matched_benchmark_money_weighted_return_pct=None,
            relative_money_weighted_return_pct=None,
            ytd_total_pnl=None,
            first_trade_date=None,
            last_valuation_date=None,
        )

    valuation_date = prices.index[0].date()
    last_price_row = prices.iloc[0]
    (
        realized,
        unrealized,
        invested,
        open_qty,
        open_cost,
        market_value,
    ) = process_trades_until(trades, valuation_date, last_price_row)
    net_cash_flow, cash_balance = calculate_cash_metrics(trades, valuation_date)
    ending_value = cash_balance + market_value
    total_pnl = realized + unrealized
    return_base = net_cash_flow if net_cash_flow > 1e-9 else invested
    return_pct = total_pnl / return_base if return_base else None
    annualized_return_pct = annualize(return_pct, first_trade, valuation_date)
    money_weighted_return_pct = calculate_account_xirr(
        trades,
        valuation_date,
        ending_value,
    )
    benchmark_return_pct = calculate_benchmark_return(
        benchmark_start_quote,
        benchmark_end_quote,
    )
    relative_to_benchmark_pct = (
        return_pct - benchmark_return_pct
        if return_pct is not None and benchmark_return_pct is not None
        else None
    )
    matched_benchmark = build_cash_flow_benchmark(
        trades,
        valuation_date,
        benchmark_cash_flow_quotes,
        benchmark_end_quote,
    )
    matched_benchmark_xirr = (
        matched_benchmark.money_weighted_return_pct if matched_benchmark else None
    )
    relative_money_weighted_return_pct = (
        money_weighted_return_pct - matched_benchmark_xirr
        if money_weighted_return_pct is not None
        and matched_benchmark_xirr is not None
        else None
    )
    ytd_total_pnl = calculate_ytd_total_pnl(
        trades,
        ytd_start_date,
        ytd_start_prices,
        total_pnl,
    )

    positions = build_positions(open_qty, open_cost, last_price_row, quotes)
    return AccountReport(
        account=account,
        trades=trades,
        positions=positions,
        realized_pnl=realized,
        unrealized_pnl=unrealized,
        total_pnl=total_pnl,
        invested_capital=invested,
        net_cash_flow=net_cash_flow,
        cash_balance=cash_balance,
        ending_value=ending_value,
        return_pct=return_pct,
        annualized_return_pct=annualized_return_pct,
        money_weighted_return_pct=money_weighted_return_pct,
        time_weighted_return_pct=daily_performance.time_weighted_return_pct,
        max_drawdown_pct=daily_performance.max_drawdown_pct,
        sharpe_ratio=daily_performance.sharpe_ratio,
        benchmark_start_price=(
            benchmark_start_quote.price if benchmark_start_quote else None
        ),
        benchmark_end_price=(
            benchmark_end_quote.price if benchmark_end_quote else None
        ),
        benchmark_return_pct=benchmark_return_pct,
        relative_to_benchmark_pct=relative_to_benchmark_pct,
        matched_benchmark_ending_value=(
            matched_benchmark.ending_value if matched_benchmark else None
        ),
        matched_benchmark_pnl=(matched_benchmark.pnl if matched_benchmark else None),
        matched_benchmark_return_pct=(
            matched_benchmark.return_pct if matched_benchmark else None
        ),
        matched_benchmark_money_weighted_return_pct=matched_benchmark_xirr,
        relative_money_weighted_return_pct=relative_money_weighted_return_pct,
        ytd_total_pnl=ytd_total_pnl,
        first_trade_date=first_trade,
        last_valuation_date=valuation_date,
    )


def calculate_ytd_total_pnl(
    trades: pd.DataFrame,
    ytd_start_date: date,
    ytd_start_prices: pd.DataFrame,
    current_total_pnl: float,
) -> float:
    if ytd_start_prices.empty:
        start_price_row = pd.Series(dtype=float)
    else:
        start_price_row = ytd_start_prices.iloc[0]
    realized, unrealized, _, _, _, _ = process_trades_until(
        trades,
        ytd_start_date,
        start_price_row,
    )
    return current_total_pnl - (realized + unrealized)


def build_positions(
    open_qty: dict[str, float],
    open_cost: dict[str, float],
    prices_on_date: pd.Series,
    quotes: dict[str, PriceQuote],
) -> pd.DataFrame:
    rows = []
    for symbol, qty in sorted(open_qty.items()):
        if abs(qty) <= 1e-9:
            continue
        latest_price = float(prices_on_date[symbol])
        cost_basis = open_cost.get(symbol, 0.0)
        market_value = qty * latest_price
        rows.append(
            {
                "Symbol": symbol,
                "Open Qty": qty,
                "Avg Buy Cost": cost_basis / qty if qty else None,
                "Latest Price": latest_price,
                "Cost Basis": cost_basis,
                "Market Value": market_value,
                "Unrealized P&L": market_value - cost_basis,
            }
        )
    return pd.DataFrame(rows)


def annualize(return_pct: float | None, start: date | None, end: date | None) -> float | None:
    if return_pct is None or start is None or end is None:
        return None
    days = max((end - start).days, 1)
    if return_pct <= -1:
        return None
    return (1 + return_pct) ** (365 / days) - 1


def calculate_benchmark_return(
    start_quote: PriceQuote | None,
    end_quote: PriceQuote | None,
) -> float | None:
    if start_quote is None or end_quote is None or start_quote.price <= 0:
        return None
    return end_quote.price / start_quote.price - 1


def combine_reports(
    reports: list[AccountReport],
    benchmark_start_quote: PriceQuote | None,
    benchmark_end_quote: PriceQuote | None,
    combined_trades: pd.DataFrame,
    benchmark_cash_flow_quotes: dict[date, PriceQuote],
    daily_performance: DailyPerformance,
) -> dict[str, Any]:
    first_trade_dates = [
        report.first_trade_date for report in reports if report.first_trade_date
    ]
    last_valuation_dates = [
        report.last_valuation_date for report in reports if report.last_valuation_date
    ]
    first_trade = min(first_trade_dates) if first_trade_dates else None
    last_date = max(last_valuation_dates) if last_valuation_dates else None
    total_pnl = sum(report.total_pnl for report in reports)
    invested = sum(report.invested_capital for report in reports)
    net_cash_flow = sum(report.net_cash_flow for report in reports)
    cash_balance = sum(report.cash_balance for report in reports)
    ending_value = sum(report.ending_value for report in reports)
    return_base = net_cash_flow if net_cash_flow > 1e-9 else invested
    return_pct = total_pnl / return_base if return_base else None
    money_weighted_return_pct = (
        calculate_account_xirr(combined_trades, last_date, ending_value)
        if last_date
        else None
    )
    benchmark_return_pct = calculate_benchmark_return(
        benchmark_start_quote,
        benchmark_end_quote,
    )
    matched_benchmark = (
        build_cash_flow_benchmark(
            combined_trades,
            last_date,
            benchmark_cash_flow_quotes,
            benchmark_end_quote,
        )
        if last_date
        else None
    )
    matched_benchmark_xirr = (
        matched_benchmark.money_weighted_return_pct if matched_benchmark else None
    )

    return {
        "realized_pnl": sum(report.realized_pnl for report in reports),
        "unrealized_pnl": sum(report.unrealized_pnl for report in reports),
        "total_pnl": total_pnl,
        "invested_capital": invested,
        "net_cash_flow": net_cash_flow,
        "cash_balance": cash_balance,
        "ending_value": ending_value,
        "return_pct": return_pct,
        "annualized_return_pct": annualize(return_pct, first_trade, last_date),
        "money_weighted_return_pct": money_weighted_return_pct,
        "time_weighted_return_pct": daily_performance.time_weighted_return_pct,
        "max_drawdown_pct": daily_performance.max_drawdown_pct,
        "sharpe_ratio": daily_performance.sharpe_ratio,
        "benchmark_start_price": (
            benchmark_start_quote.price if benchmark_start_quote else None
        ),
        "benchmark_end_price": (
            benchmark_end_quote.price if benchmark_end_quote else None
        ),
        "benchmark_return_pct": benchmark_return_pct,
        "relative_to_benchmark_pct": (
            return_pct - benchmark_return_pct
            if return_pct is not None and benchmark_return_pct is not None
            else None
        ),
        "matched_benchmark_ending_value": (
            matched_benchmark.ending_value if matched_benchmark else None
        ),
        "matched_benchmark_pnl": (
            matched_benchmark.pnl if matched_benchmark else None
        ),
        "matched_benchmark_return_pct": (
            matched_benchmark.return_pct if matched_benchmark else None
        ),
        "matched_benchmark_money_weighted_return_pct": matched_benchmark_xirr,
        "relative_money_weighted_return_pct": (
            money_weighted_return_pct - matched_benchmark_xirr
            if money_weighted_return_pct is not None
            and matched_benchmark_xirr is not None
            else None
        ),
        "ytd_total_pnl": sum(
            report.ytd_total_pnl for report in reports
            if report.ytd_total_pnl is not None
        ),
        "first_trade_date": first_trade,
        "last_valuation_date": last_date,
    }


def markdown_table(df: pd.DataFrame, max_rows: int | None = None) -> str:
    if df.empty:
        return "_No rows._"
    display = df.copy()
    if max_rows is not None:
        display = display.head(max_rows)
    money_columns = {
        "Avg Buy Cost",
        "Latest Price",
        "Cost Basis",
        "Market Value",
        "Unrealized P&L",
        "Realized P&L",
        "Total P&L",
        "YTD Total P&L",
        "Invested Capital",
        "Drawdown Base",
    }
    for col in display.columns:
        if col in money_columns:
            display[col] = display[col].map(lambda x: money(float(x)) if pd.notna(x) else "n/a")
        elif col == "Return":
            display[col] = display[col].map(lambda x: pct(float(x)) if pd.notna(x) else "n/a")
        elif "Qty" in col:
            display[col] = display[col].map(
                lambda x: f"{float(x):,.2f}".rstrip("0").rstrip(".")
                if pd.notna(x)
                else "n/a"
            )
    return dataframe_to_markdown(display)


def combine_positions(reports: list[AccountReport]) -> pd.DataFrame:
    position_frames = [
        report.positions.assign(Account=report.account)
        for report in reports
        if not report.positions.empty
    ]
    if not position_frames:
        return pd.DataFrame()

    positions = pd.concat(position_frames, ignore_index=True)
    if positions.empty:
        return positions

    grouped = (
        positions.groupby("Symbol", as_index=False)
        .agg(
            {
                "Open Qty": "sum",
                "Cost Basis": "sum",
                "Market Value": "sum",
                "Latest Price": "last",
            }
        )
        .sort_values("Symbol")
    )
    grouped["Avg Buy Cost"] = grouped["Cost Basis"] / grouped["Open Qty"]
    grouped["Unrealized P&L"] = grouped["Market Value"] - grouped["Cost Basis"]
    grouped = grouped[
        [
            "Symbol",
            "Open Qty",
            "Avg Buy Cost",
            "Latest Price",
            "Cost Basis",
            "Market Value",
            "Unrealized P&L",
        ]
    ]
    return grouped


def quotes_to_price_frame(quotes: dict[str, PriceQuote], valuation_date: date) -> pd.DataFrame:
    return pd.DataFrame(
        [{symbol: quote.price for symbol, quote in quotes.items()}],
        index=[pd.Timestamp(valuation_date)],
    )


def valuation_date_from_quotes(
    quotes: dict[str, PriceQuote],
    fallback_date: date,
) -> date:
    quote_dates = []
    for quote in quotes.values():
        if not quote.price_time:
            continue
        try:
            quote_dates.append(pd.to_datetime(quote.price_time).date())
        except (TypeError, ValueError):
            continue
    if not quote_dates:
        return fallback_date
    return min(quote_dates)


def dataframe_to_markdown(df: pd.DataFrame) -> str:
    """Render a small DataFrame as a GitHub-flavored Markdown table."""
    if df.empty:
        return "_No rows._"
    text_df = df.astype(str)
    headers = [str(col) for col in text_df.columns]
    rows = text_df.values.tolist()
    widths = [
        max(len(headers[idx]), *(len(str(row[idx])) for row in rows))
        for idx in range(len(headers))
    ]

    def render_row(values: list[str]) -> str:
        cells = [str(value).ljust(widths[idx]) for idx, value in enumerate(values)]
        return "| " + " | ".join(cells) + " |"

    separator = "| " + " | ".join("-" * width for width in widths) + " |"
    return "\n".join([render_row(headers), separator, *(render_row(row) for row in rows)])


def write_report(
    output_path: Path,
    workbook_path: Path,
    reports: list[AccountReport],
    combined: dict[str, Any],
    run_datetime: datetime,
    quotes: dict[str, PriceQuote],
    nav_chart_path: Path | None,
    annual_risk_free_rate: float,
) -> None:
    lines = [
        "# Account Performance Report",
        "",
        f"- Source workbook: `{workbook_path}`",
        f"- Run date: {run_datetime.strftime('%Y-%m-%d %H:%M:%S')}",
        f"- Valuation date: {combined['last_valuation_date']}",
        "- Method: FIFO realized P&L; buy commissions are included in cost basis; sell commissions reduce proceeds.",
        "- Return definition: total P&L divided by net cash flow when CASH rows exist; otherwise divided by cumulative buy cost including commissions. Simple annualized return uses calendar days from first trade to valuation date; XIRR is preferred when dated CASH flows are available.",
        f"- Benchmark definition: {BENCHMARK_SYMBOL} price return from the closest trading day on or before the first trade date through the valuation price. Simple relative return equals account return minus {BENCHMARK_SYMBOL} return.",
        f"- Cash-flow-matched benchmark: every CASH deposit buys fractional {BENCHMARK_SYMBOL} shares and every withdrawal sells shares at the closest price on or before that cash-flow date. Money-weighted returns use XIRR; relative XIRR equals account XIRR minus matched-{BENCHMARK_SYMBOL} XIRR.",
        "- TWR definition: daily returns geometrically linked after removing CASH deposits and withdrawals, which are assumed to occur before that day's return period. The initial CASH contribution is the first day's return base.",
        f"- Excess-return curve: combined-account cumulative TWR minus cumulative {BENCHMARK_SYMBOL} price return, shown as a percentage on the right axis.",
        f"- Risk definition: maximum drawdown is measured from the TWR high-water mark. Sharpe ratio uses daily TWR returns, {TRADING_DAYS_PER_YEAR} trading days per year, and a {annual_risk_free_rate:.2%} annual risk-free rate.",
        "- YTD Total P&L definition: current total P&L minus total P&L as of the prior December 31.",
        "",
        "## Combined Accounts",
        "",
        summary_block("Combined", combined),
        "",
    ]
    if nav_chart_path is not None:
        lines.extend(
            [
                "## TWR Net Value Curve",
                "",
                f"![TWR net value curve]({nav_chart_path.name})",
                "",
            ]
        )
    lines.extend(
        [
        "### Combined Open Positions",
        "",
        markdown_table(combine_positions(reports)),
        "",
        ]
    )

    for report in reports:
        lines.extend(
            [
                f"## {report.account}",
                "",
                summary_block(report.account, report),
                "",
                "### Open Positions",
                "",
                markdown_table(report.positions),
                "",
            ]
        )

    output_path.write_text("\n".join(lines))


def summary_block(name: str, obj: AccountReport | dict[str, Any]) -> str:
    getter = obj.get if isinstance(obj, dict) else lambda key: getattr(obj, key)
    def text(value: Any) -> Any:
        if value is None or (isinstance(value, float) and math.isnan(value)):
            return "n/a"
        return value

    rows = [
        ("First trade date", text(getter("first_trade_date"))),
        ("Last valuation date", text(getter("last_valuation_date"))),
        ("Invested capital", money(getter("invested_capital"))),
        ("Net cash flow", money(getter("net_cash_flow"))),
        ("Cash balance", money(getter("cash_balance"))),
        ("Ending account value", money(getter("ending_value"))),
        ("Realized P&L", money(getter("realized_pnl"))),
        ("Unrealized P&L", money(getter("unrealized_pnl"))),
        ("Total P&L", money(getter("total_pnl"))),
        ("YTD Total P&L", money(getter("ytd_total_pnl"))),
        ("Return", pct(getter("return_pct"))),
        ("Simple annualized return", pct(getter("annualized_return_pct"))),
        ("Money-weighted return (XIRR)", pct(getter("money_weighted_return_pct"))),
        ("Time-weighted return (TWR)", pct(getter("time_weighted_return_pct"))),
        ("Maximum drawdown", pct(getter("max_drawdown_pct"))),
        ("Sharpe ratio", ratio(getter("sharpe_ratio"))),
        (f"{BENCHMARK_SYMBOL} start price", money(getter("benchmark_start_price"))),
        (f"{BENCHMARK_SYMBOL} end price", money(getter("benchmark_end_price"))),
        (f"{BENCHMARK_SYMBOL} return", pct(getter("benchmark_return_pct"))),
        (
            f"Simple relative to {BENCHMARK_SYMBOL}",
            pct(getter("relative_to_benchmark_pct")),
        ),
        (
            f"Matched {BENCHMARK_SYMBOL} ending value",
            money(getter("matched_benchmark_ending_value")),
        ),
        (f"Matched {BENCHMARK_SYMBOL} P&L", money(getter("matched_benchmark_pnl"))),
        (
            f"Matched {BENCHMARK_SYMBOL} return",
            pct(getter("matched_benchmark_return_pct")),
        ),
        (
            f"Matched {BENCHMARK_SYMBOL} XIRR",
            pct(getter("matched_benchmark_money_weighted_return_pct")),
        ),
        (
            f"Relative XIRR to {BENCHMARK_SYMBOL}",
            pct(getter("relative_money_weighted_return_pct")),
        ),
    ]
    df = pd.DataFrame(rows, columns=["Metric", name])
    return dataframe_to_markdown(df)


def main() -> int:
    args = parse_args()
    if args.risk_free_rate <= -1:
        raise ValueError("--risk-free-rate must be greater than -1.0")
    run_datetime = datetime.now()
    dotenv_path = Path(".env")
    api_key = get_api_key(args.api_key, dotenv_path)

    workbook_path = Path(args.input).expanduser().resolve()
    requested_as_of = (
        datetime.strptime(args.as_of, "%Y-%m-%d").date()
        if args.as_of
        else None
    )

    trades_by_account = read_trades(workbook_path)
    all_trades = pd.concat(trades_by_account.values(), ignore_index=True)
    initial_cutoff_date = requested_as_of or date.today()
    _, _, combined_open_qty, _ = calculate_fifo(all_trades, initial_cutoff_date)
    symbols = sorted(
        symbol for symbol, qty in combined_open_qty.items()
        if abs(qty) > 1e-9 and not is_cash_symbol(symbol)
    )
    quote_symbols = sorted(set(symbols) | {BENCHMARK_SYMBOL})

    if requested_as_of is None:
        quotes = load_latest_prices(
            symbols=quote_symbols,
            api_key=api_key,
            cache_dir=Path(args.cache_dir),
            allow_cache_fallback=args.allow_cache_fallback,
        )
        valuation_date = valuation_date_from_quotes(quotes, date.today())
    else:
        quotes = load_historical_prices(
            symbols=quote_symbols,
            target_date=requested_as_of,
            api_key=api_key,
            cache_dir=Path(args.cache_dir),
            allow_cache_fallback=args.allow_cache_fallback,
        )
        valuation_date = valuation_date_from_quotes(quotes, requested_as_of)
        _, _, final_open_qty, _ = calculate_fifo(all_trades, valuation_date)
        final_symbols = sorted(
            symbol for symbol, qty in final_open_qty.items()
            if abs(qty) > 1e-9 and not is_cash_symbol(symbol)
        )
        if final_symbols != symbols:
            symbols = final_symbols
            quote_symbols = sorted(set(symbols) | {BENCHMARK_SYMBOL})
            quotes = load_historical_prices(
                symbols=quote_symbols,
                target_date=valuation_date,
                api_key=api_key,
                cache_dir=Path(args.cache_dir),
                allow_cache_fallback=args.allow_cache_fallback,
            )
            valuation_date = valuation_date_from_quotes(quotes, valuation_date)

    prices = quotes_to_price_frame(quotes, valuation_date)
    output_name = args.output or f"account_report_{valuation_date:%Y-%m-%d}.md"
    output_path = Path(output_name).expanduser().resolve()
    ytd_start_date = date(valuation_date.year - 1, 12, 31)

    _, _, ytd_start_open_qty, _ = calculate_fifo(all_trades, ytd_start_date)
    ytd_start_symbols = sorted(
        symbol for symbol, qty in ytd_start_open_qty.items()
        if abs(qty) > 1e-9 and not is_cash_symbol(symbol)
    )
    ytd_start_quotes = load_historical_prices(
        symbols=ytd_start_symbols,
        target_date=ytd_start_date,
        api_key=api_key,
        cache_dir=Path(args.cache_dir),
        allow_cache_fallback=args.allow_cache_fallback,
    )
    ytd_start_prices = quotes_to_price_frame(ytd_start_quotes, ytd_start_date)

    first_trade_dates = sorted(
        {
            trades["Date"].min()
            for trades in trades_by_account.values()
            if not trades.empty
        }
    )
    cash_flow_dates = sorted(
        {flow_date for flow_date, _ in external_cash_flows(all_trades, valuation_date)}
    )
    benchmark_quote_dates = sorted(set(first_trade_dates) | set(cash_flow_dates))
    benchmark_end_quote = quotes.get(BENCHMARK_SYMBOL)
    benchmark_historical_quotes = {
        quote_date: (
            benchmark_end_quote
            if quote_date == valuation_date
            else fetch_historical_quote(
                BENCHMARK_SYMBOL,
                quote_date,
                api_key,
                Path(args.cache_dir),
                args.allow_cache_fallback,
            )
        )
        for quote_date in benchmark_quote_dates
    }

    history_trades = all_trades[all_trades["Date"] <= valuation_date]
    history_start_date = history_trades["Date"].min() - timedelta(days=10)
    history_symbols = sorted(
        set(
            history_trades.loc[
                ~history_trades["Symbol"].map(is_cash_symbol),
                "Symbol",
            ]
        )
        | {BENCHMARK_SYMBOL}
    )
    price_history = load_historical_series(
        symbols=history_symbols,
        start_date=history_start_date,
        end_date=valuation_date,
        api_key=api_key,
        cache_dir=Path(args.cache_dir),
        allow_cache_fallback=args.allow_cache_fallback,
    )
    for symbol, quote in quotes.items():
        if symbol in price_history:
            price_history[symbol][valuation_date] = quote.price

    daily_performance_by_account = {
        account: build_daily_performance(
            trades,
            valuation_date,
            price_history,
            args.risk_free_rate,
        )
        for account, trades in trades_by_account.items()
    }
    combined_daily_performance = build_daily_performance(
        all_trades,
        valuation_date,
        price_history,
        args.risk_free_rate,
    )

    missing_symbols = [symbol for symbol in symbols if symbol not in prices.columns]
    if missing_symbols:
        raise RuntimeError(f"Missing latest prices for symbols: {missing_symbols}")

    reports = [
        build_account_report(
            account,
            trades,
            prices,
            quotes,
            ytd_start_prices,
            ytd_start_date,
            benchmark_historical_quotes.get(trades["Date"].min())
            if not trades.empty
            else None,
            benchmark_end_quote,
            benchmark_historical_quotes,
            daily_performance_by_account[account],
        )
        for account, trades in trades_by_account.items()
    ]
    combined_first_trade = min(first_trade_dates) if first_trade_dates else None
    combined = combine_reports(
        reports,
        benchmark_historical_quotes.get(combined_first_trade),
        benchmark_end_quote,
        all_trades,
        benchmark_historical_quotes,
        combined_daily_performance,
    )
    nav_chart_path = output_path.with_name(f"{output_path.stem}_nav.png")
    combined_first_trade = min(first_trade_dates) if first_trade_dates else valuation_date
    benchmark_curve = build_benchmark_curve(
        price_history[BENCHMARK_SYMBOL],
        combined_first_trade,
        valuation_date,
    )
    chart_curves = {
        "Combined": combined_daily_performance.history,
        f"{BENCHMARK_SYMBOL} (price)": benchmark_curve,
    }
    excess_return_curve = build_excess_return_curve(
        combined_daily_performance.history,
        benchmark_curve,
    )
    write_nav_chart(
        nav_chart_path,
        chart_curves,
        excess_return_curve,
    )
    write_report(
        output_path,
        workbook_path,
        reports,
        combined,
        run_datetime,
        quotes,
        nav_chart_path,
        args.risk_free_rate,
    )

    print(f"Wrote report to {output_path}")
    print(f"Accounts: {', '.join(report.account for report in reports)}")
    print(f"Combined total P&L: {money(combined['total_pnl'])}")
    print(f"Combined return: {pct(combined['return_pct'])}")
    print(f"{BENCHMARK_SYMBOL} return: {pct(combined['benchmark_return_pct'])}")
    print(f"Relative to {BENCHMARK_SYMBOL}: {pct(combined['relative_to_benchmark_pct'])}")
    print(f"Account XIRR: {pct(combined['money_weighted_return_pct'])}")
    print(
        f"Matched {BENCHMARK_SYMBOL} XIRR: "
        f"{pct(combined['matched_benchmark_money_weighted_return_pct'])}"
    )
    print(
        f"Relative XIRR to {BENCHMARK_SYMBOL}: "
        f"{pct(combined['relative_money_weighted_return_pct'])}"
    )
    print(f"Combined TWR: {pct(combined['time_weighted_return_pct'])}")
    print(f"Maximum drawdown: {pct(combined['max_drawdown_pct'])}")
    print(f"Sharpe ratio: {ratio(combined['sharpe_ratio'])}")
    print(f"TWR chart: {nav_chart_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
