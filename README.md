# India ETF Buy-Low Scanner — GitHub + Python

This version replaces the manually-run Google Colab scanner.

## Architecture

GitHub Actions
    ↓
Python scanner
    ↓
Google Sheets
    ├── DASHBOARD
    ├── DAILY_SCAN
    ├── HISTORY
    ├── EXCLUDED
    ├── SETTINGS
    └── NAV

NAV is NOT fetched by Python.

The scanner reads NAV from the existing `NAV` sheet:
- A = Symbol
- F = ISINNumber
- H = NAV

## Spreadsheet

Spreadsheet ID:

`1C0O_uXW2TC44RiLEbilj_zlrS_LcJDEvhD7a_pRnk2M`

## Google service account

1. Create a Google Cloud service account.
2. Create/download its JSON key.
3. Share the spreadsheet with the service-account email as Editor.
4. Add the entire JSON content as a GitHub repository secret named:

`GOOGLE_SERVICE_ACCOUNT_JSON`

Do NOT put the JSON key into the repository.

## GitHub

Create a repository and upload:

- `scanner.py`
- `requirements.txt`
- `.github/workflows/scanner.yml`

Then run the workflow once manually from:
Actions → India ETF Buy-Low Scanner → Run workflow.

After that it is scheduled for 3:35 PM IST Monday-Friday.

GitHub scheduled workflows run automatically on GitHub-hosted runners.

## Settings / allocation

The scanner reads the `SETTINGS` sheet.

Use these columns in row 1:

`Asset Class | Allocation %`

Example:

EQUITY | 60
GOLD | 10
SILVER | 10
LIQUID | 20

The total must equal 100%.

Change these percentages whenever you want. The next scan uses the new allocation.

## Capital

`DASHBOARD!B5` = Total Deployed Capital.

## Priority buy-low rule

The highest-priority setup is:

Price < 20 DMA
AND
20 DMA > 50 DMA > 200 DMA

That setup receives the highest signal tier and an explicit reason:

`BUY-LOW: below 20 DMA while 20>50>200 DMA`

## Speed

The Python version downloads price history in batches rather than calling yfinance separately for every ETF.

## Manual run

You can run:

`python scanner.py`

or use GitHub Actions → Run workflow.

## Important

The GitHub Actions schedule is automated, but GitHub notes that scheduled workflows can occasionally be delayed under high load. The workflow is intentionally scheduled at minute 5 rather than exactly on the hour.
