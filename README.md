# Vendor Credits

Web app: upload a QuickBooks **Balance Sheet** export and an **A/P Aging Summary**
export (`.xlsx`, `.xlsm`, or `.csv`). Download a workbook with three tabs —
**Vendor Credits**, **Other TC**, and **Other AP**.

This is a from-scratch rebuild of the `Vendor_Credits.xlsm` Power Query workbook. The
app does the paste-and-refresh steps for you: it reads the raw QuickBooks reports,
extracts the ticket-credit accounts itself, and produces the same three output tabs.

## What it does

1. **Ticket Credits (from the Balance Sheet).** Every account whose name ends with
   `(TC)` becomes a vendor. ` (TC)` is stripped to get the vendor name; the account's
   balance is the **Ticket Credit Amount**. `Total …` and `(DEP)` lines are ignored.
2. **A/P (from the A/P Aging Summary).** Each vendor's payable is the **sum of the five
   aging buckets** (Current, 1-30, 31-60, 61-90, 91 and over). The report's own Total
   column is ignored and recomputed from the buckets.
3. **Match.** For each ticket-credit vendor, compare its Ticket Credit Amount (TC) to its
   A/P total (AP). The **lower** of the two is the credit that can be applied:
   - `Lower Option` = `TC` if TC < AP, `AP` if AP < TC, else `Equal`.
   - `Difference (AP - TC)` = AP − TC.

### Output tabs

- **Vendor Credits** — vendors that have both a credit and a payable. `Expense Line
  Amount` is the lower of TC / AP (the amount to apply). Columns: Vendor, Payment Date,
  Expense Account, Expense Line Amount, Column1, Lower Option, Ticket Credit Amount,
  AP Aging Amount, Difference (AP - TC).
- **Other TC** — ticket credits with **no payable to offset** (non-negative credit, A/P ≤ 0).
  These are leftover credit balances. Columns: Vendor, Payment Date, Expense Account,
  Ticket Credit Amount.
- **Other AP** — payables for vendors that received **no credit**, broken out by aging
  bucket, with Expense Account set to `Cost of Goods Sold`. Columns: Vendor, Payment Date,
  Expense Account, Expense Line Amount, Column1, Current, 1 - 30, 31 - 60, 61 - 90,
  91 and over.

Each tab carries a live `=SUM(...)` **TOTAL** row on its amount column.

### Matching notes

- Vendor names must match between the two reports (e.g. `Arizona Diamondbacks (TC)` on the
  Balance Sheet matches `Arizona Diamondbacks` on the A/P Aging). A name that exists on only
  one side flows to Other TC or Other AP accordingly.
- **Negative** ticket credits are not applied as a credit; the vendor's payable (if any)
  appears in Other AP.
- The three tabs partition the data: every A/P vendor lands in exactly one of Vendor Credits
  or Other AP.

## Payment date

The **Payment Date** column (and the download filename) use the month-end date. The app
auto-detects it from the `As of …` line in the reports; you can override it with the
optional date field in the UI.

## Input format

The app expects the **raw** QuickBooks exports:

- **Balance Sheet** — an account column containing the `… (TC)` rows and an amount column.
  Amounts may be plain numbers, `=123.45` literal formulas, `$1,234.56`, or `(123)` negatives.
- **A/P Aging Summary** — a Vendor column plus the five aging-bucket columns (matched by
  header name, so column order doesn't matter). A Total column is optional.

## Run locally

```bash
pip install -r requirements.txt
python app.py        # http://localhost:5000
```

## Deploy: GitHub → Railway

1. Push this folder to a GitHub repo.
2. Railway → New Project → Deploy from GitHub repo → pick it.
3. Railway auto-detects Python (Nixpacks) and uses the start command in `railway.json`.
   No env vars needed; `$PORT` is provided automatically.

## Files

| File | Purpose |
|------|---------|
| `app.py` | Flask backend — parsing, reconciliation, workbook builder |
| `index.html` | Single-page upload UI |
| `requirements.txt` | Python dependencies |
| `Procfile` / `railway.json` | Start command for Railway |
