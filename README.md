# Trade Record Performance

Generate a Markdown performance report from a trade-record Excel workbook.

The script reads two worksheets:

- `UOB`
- `IB`

It calculates separate account reports and a combined report, including:

- realized P&L using FIFO cost basis
- unrealized P&L using latest market prices
- total P&L
- return
- annualized return
- per-account open positions
- combined open positions

Latest prices are fetched from Financial Modeling Prep first. If FMP has no API
key or a quote request fails, the script falls back to Yahoo Finance.

## Setup

Create a local `.env` file from the example:

```bash
cp .env.example .env
```

Then add your FMP key:

```env
FMP_API_KEY=your_fmp_key_here
```

The key is optional because Yahoo Finance is used as a fallback, but FMP is tried
first when a key is available.

## Usage

```bash
python3 generate_account_report.py \
  --input "/path/to/James trade record.xlsx" \
  --as-of 2026-07-02
```

By default, the report is written as:

```text
account_report_YYYY-MM-DD.md
```

You can override the output path:

```bash
python3 generate_account_report.py \
  --input "/path/to/James trade record.xlsx" \
  --output my_report.md
```

## Daily GitHub Actions Run

This repository includes a GitHub Actions workflow at
`.github/workflows/daily-report.yml`.

It runs Monday through Friday at 16:10 America/New_York time and can also be
started manually from the **Actions** tab with **Run workflow**.

Because trade records and generated reports should not be committed to the
repository, the workflow expects the workbook to be available from a private
download URL stored as a GitHub secret.

Configure these repository secrets in GitHub:

| Secret | Required | Purpose |
| --- | --- | --- |
| `TRADE_RECORD_URL` | Yes | Private download URL for the Excel workbook |
| `FMP_API_KEY` | No | FMP API key; Yahoo Finance is used as fallback |

For a Google Sheet, `TRADE_RECORD_URL` can be the normal sheet URL:

```text
https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
```

The workflow converts Google Sheet URLs to `.xlsx` export URLs automatically.
The sheet must be shared so the workflow can download it, for example
**Anyone with the link can view**. If the sheet requires an interactive Google
login, GitHub Actions will download an HTML login page instead of an Excel file.

The workflow downloads the workbook at runtime, generates the report, and saves
the Markdown file as a workflow artifact. It does not commit the workbook or the
generated report back to the repository.

## Workbook Format

Both trade worksheets should contain these columns:

| Column | Meaning |
| --- | --- |
| `Date` | Trade date |
| `Symbol` | Ticker symbol |
| `Price` | Trade price |
| `Qty` | Positive for buy, negative for sell |
| `Comm Fee` | Commission or fee |
| `Trade Value` | Trade value from the source workbook |

Only `UOB` and `IB` are read. Other worksheets are ignored.

## Notes

- `.env`, generated reports, price caches, and local Excel workbooks are ignored
  by git.
- Buy commissions are included in cost basis.
- Sell commissions reduce sale proceeds.
- Realized P&L uses FIFO matching.
- Annualized return uses calendar days from first trade to valuation date.
