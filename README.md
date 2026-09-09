# Spending Dashboard

A simple local dashboard that visualises your spending from an Excel file you keep
adding to yourself. No cloud, no login, no bank connection — your data never leaves
your computer.

## What you get

- Monthly spend trend, category breakdown, top merchants, and spend-by-account charts
- Headline metrics: total spend, income/refunds, net, average monthly spend
- Filters by date range, category, and account
- A form in the sidebar to add a transaction without opening Excel (optional — you
  can also just edit the spreadsheet directly)
- A filterable transactions table with CSV export

## 1. Requirements

- Python 3.9+ installed
- (Recommended) a virtual environment

## 2. Setup

Unzip the folder, then open a terminal inside it and run:

```bash
# optional but recommended
python3 -m venv venv
source venv/bin/activate      # on Windows: venv\Scripts\activate

pip install -r requirements.txt
```

## 3. Run it

```bash
streamlit run app.py
```

This opens the dashboard in your browser (usually `http://localhost:8501`).

## 4. Add your data

Open `data/transactions.xlsx`. It has one sheet called **Transactions** with these
columns:

| Date | Description | Category | Amount | Account |
|------|-------------|----------|--------|---------|
| 2026-01-15 | Tesco Superstore | Groceries | 42.15 | Current Account |

- **Amount**: enter spending as a **positive** number, income/refunds as **negative**.
  If your bank exports do it the other way round, just tick the "negative = spending"
  checkbox in the dashboard sidebar instead of re-entering everything.
- The highlighted example row is just a formatting guide — edit or delete it.
- A **Read Me** tab inside the workbook repeats these instructions.

Keep adding rows over time (copy in rows from your bank's exported statement, or
type them in manually), save the file, then click **Reload data** in the sidebar
(or just refresh the page) to see the dashboard update.

## Updating from a bank statement export

If your bank lets you export a CSV or Excel file, the easiest workflow is:

1. Export your statement.
2. Copy the Date / Description / Amount columns into the Transactions tab.
3. Fill in a Category and Account for each new row (a quick Find & Replace in Excel
   works well if your bank statement groups payees).

## Notes

- All processing happens locally on your machine — nothing is uploaded anywhere.
- To reset, just replace `data/transactions.xlsx` with a fresh copy of the template
  or clear out the rows below the header.
- Want a different currency symbol? Search `app.py` for `£` and replace it.
