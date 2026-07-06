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
from collections import defaultdict, deque
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

import pandas as pd


SHEETS = {
    "UOB": "UOB",
    "IB": "IB",
}

REQUIRED_COLUMNS = ["Date", "Symbol", "Price", "Qty", "Comm Fee", "Trade Value"]
FMP_QUOTE_URL = "https://financialmodelingprep.com/stable/quote"
YAHOO_CHART_URL = "https://query1.finance.yahoo.com/v8/finance/chart/{symbol}"


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
    return_pct: float | None
    annualized_return_pct: float | None
    first_trade_date: date | None
    last_valuation_date: date | None


@dataclass
class PriceQuote:
    symbol: str
    price: float
    source: str
    price_time: str | None = None
    fetched_at: str | None = None


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate separate and combined performance reports from the trade workbook."
    )
    parser.add_argument(
        "--input",
        default="/Users/chandlerqian/Downloads/James trade record.xlsx",
        help="Path to the trade record workbook.",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Path for the Markdown report. Defaults to account_report_YYYY-MM-DD.md.",
    )
    parser.add_argument(
        "--as-of",
        default=date.today().isoformat(),
        help="Valuation date in YYYY-MM-DD format. Defaults to today.",
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
        help="Do not use cached fallback prices if live quote refresh fails.",
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
    with urlopen(request, timeout=30) as response:
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


def fetch_latest_quote(
    symbol: str,
    api_key: str | None,
    cache_dir: Path,
    force_refresh: bool,
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
            if cache_file.exists() and not force_refresh:
                cached = PriceQuote(**json.loads(cache_file.read_text()))
                cached.source = f"cached {cached.source}"
                return cached
            raise RuntimeError(f"Could not fetch latest price for {symbol}: {'; '.join(errors)}") from yahoo_exc

    cache_file.write_text(json.dumps(quote.__dict__, indent=2, sort_keys=True))
    return quote


def load_latest_prices(
    symbols: list[str],
    api_key: str | None,
    cache_dir: Path,
    force_refresh: bool,
) -> dict[str, PriceQuote]:
    return {
        symbol: fetch_latest_quote(symbol, api_key, cache_dir, force_refresh)
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
) -> AccountReport:
    first_trade = trades["Date"].min() if not trades.empty else None

    if trades.empty or prices.empty:
        return AccountReport(
            account=account,
            trades=trades,
            positions=pd.DataFrame(),
            realized_pnl=0.0,
            unrealized_pnl=0.0,
            total_pnl=0.0,
            invested_capital=0.0,
            return_pct=None,
            annualized_return_pct=None,
            first_trade_date=None,
            last_valuation_date=None,
        )

    valuation_date = prices.index[0].date()
    last_price_row = prices.iloc[0]
    realized, unrealized, invested, open_qty, open_cost, _ = process_trades_until(
        trades, valuation_date, last_price_row
    )
    total_pnl = realized + unrealized
    return_pct = total_pnl / invested if invested else None
    annualized_return_pct = annualize(return_pct, first_trade, valuation_date)

    positions = build_positions(open_qty, open_cost, last_price_row, quotes)
    return AccountReport(
        account=account,
        trades=trades,
        positions=positions,
        realized_pnl=realized,
        unrealized_pnl=unrealized,
        total_pnl=total_pnl,
        invested_capital=invested,
        return_pct=return_pct,
        annualized_return_pct=annualized_return_pct,
        first_trade_date=first_trade,
        last_valuation_date=valuation_date,
    )


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
        quote = quotes.get(symbol)
        rows.append(
            {
                "Symbol": symbol,
                "Open Qty": qty,
                "Avg Buy Cost": cost_basis / qty if qty else None,
                "Latest Price": latest_price,
                "Cost Basis": cost_basis,
                "Market Value": market_value,
                "Unrealized P&L": market_value - cost_basis,
                "Price Source": quote.source if quote else "n/a",
                "Price Time": quote.price_time if quote and quote.price_time else "n/a",
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


def combine_reports(reports: list[AccountReport]) -> dict[str, Any]:
    first_trade = min(report.first_trade_date for report in reports if report.first_trade_date)
    last_date = max(report.last_valuation_date for report in reports if report.last_valuation_date)
    total_pnl = sum(report.total_pnl for report in reports)
    invested = sum(report.invested_capital for report in reports)
    return_pct = total_pnl / invested if invested else None

    return {
        "realized_pnl": sum(report.realized_pnl for report in reports),
        "unrealized_pnl": sum(report.unrealized_pnl for report in reports),
        "total_pnl": total_pnl,
        "invested_capital": invested,
        "return_pct": return_pct,
        "annualized_return_pct": annualize(return_pct, first_trade, last_date),
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
    positions = pd.concat(
        [
            report.positions.assign(Account=report.account)
            for report in reports
            if not report.positions.empty
        ],
        ignore_index=True,
    )
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
                "Price Source": "last",
                "Price Time": "last",
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
            "Price Source",
            "Price Time",
        ]
    ]
    return grouped


def quotes_to_price_frame(quotes: dict[str, PriceQuote], valuation_date: date) -> pd.DataFrame:
    return pd.DataFrame(
        [{symbol: quote.price for symbol, quote in quotes.items()}],
        index=[pd.Timestamp(valuation_date)],
    )


def price_source_summary(quotes: dict[str, PriceQuote]) -> str:
    if not quotes:
        return "No open positions requiring market prices"
    by_source: dict[str, list[str]] = defaultdict(list)
    for symbol, quote in sorted(quotes.items()):
        by_source[quote.source].append(symbol)
    return "; ".join(
        f"{source}: {', '.join(symbols)}"
        for source, symbols in sorted(by_source.items())
    )


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
) -> None:
    lines = [
        "# Account Performance Report",
        "",
        f"- Source workbook: `{workbook_path}`",
        f"- Run date: {run_datetime.strftime('%Y-%m-%d %H:%M:%S')}",
        f"- Valuation date: {combined['last_valuation_date']}",
        f"- Price source: {price_source_summary(quotes)}",
        "- Method: FIFO realized P&L; buy commissions are included in cost basis; sell commissions reduce proceeds.",
        "- Return definition: total P&L divided by cumulative buy cost including commissions. Annualized return uses calendar days from first trade to valuation date.",
        "",
        "## Combined Accounts",
        "",
        summary_block("Combined", combined),
        "",
        "### Combined Open Positions",
        "",
        markdown_table(combine_positions(reports)),
        "",
    ]

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
    rows = [
        ("First trade date", getter("first_trade_date")),
        ("Last valuation date", getter("last_valuation_date")),
        ("Invested capital", money(getter("invested_capital"))),
        ("Realized P&L", money(getter("realized_pnl"))),
        ("Unrealized P&L", money(getter("unrealized_pnl"))),
        ("Total P&L", money(getter("total_pnl"))),
        ("Return", pct(getter("return_pct"))),
        ("Annualized return", pct(getter("annualized_return_pct"))),
    ]
    df = pd.DataFrame(rows, columns=["Metric", name])
    return dataframe_to_markdown(df)


def main() -> int:
    args = parse_args()
    run_datetime = datetime.now()
    dotenv_path = Path(".env")
    api_key = get_api_key(args.api_key, dotenv_path)

    workbook_path = Path(args.input).expanduser().resolve()
    output_name = args.output or f"account_report_{run_datetime:%Y-%m-%d}.md"
    output_path = Path(output_name).expanduser().resolve()
    as_of = datetime.strptime(args.as_of, "%Y-%m-%d").date()

    trades_by_account = read_trades(workbook_path)
    all_trades = pd.concat(trades_by_account.values(), ignore_index=True)
    _, _, combined_open_qty, _ = calculate_fifo(all_trades, as_of)
    symbols = sorted(
        symbol for symbol, qty in combined_open_qty.items()
        if abs(qty) > 1e-9
    )

    quotes = load_latest_prices(
        symbols=symbols,
        api_key=api_key,
        cache_dir=Path(args.cache_dir),
        force_refresh=args.force_refresh,
    )
    prices = quotes_to_price_frame(quotes, as_of)

    missing_symbols = [symbol for symbol in symbols if symbol not in prices.columns]
    if missing_symbols:
        raise RuntimeError(f"Missing latest prices for symbols: {missing_symbols}")

    reports = [
        build_account_report(account, trades, prices, quotes)
        for account, trades in trades_by_account.items()
    ]
    combined = combine_reports(reports)
    write_report(output_path, workbook_path, reports, combined, run_datetime, quotes)

    print(f"Wrote report to {output_path}")
    print(f"Accounts: {', '.join(report.account for report in reports)}")
    print(f"Combined total P&L: {money(combined['total_pnl'])}")
    print(f"Combined return: {pct(combined['return_pct'])}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
