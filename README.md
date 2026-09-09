# Spending Dashboard

A simple local dashboard that visualises your spending from an Excel file you keep
adding to yourself. No cloud, no login, no bank connection — your data never leaves
your computer.

## What you get

Two pages (use the page menu at the top of the sidebar):

**Overview**
- Monthly spend trend, category breakdown, top merchants, and spend-by-account charts
- Headline metrics: total spend, income/refunds, net, average monthly spend
- Filters by date range, category, and account
- A form in the sidebar to add a transaction without opening Excel (optional — you
  can also just edit the spreadsheet directly)
- A filterable transactions table with CSV export

**Budget vs Actual**
- Set a monthly budget per category (on the Budget tab in Excel, or in-app)
- See at a glance which categories are over budget this month, and by how much
- A budget-vs-actual chart per category, colour-coded over/under
- Drill into any category to see exactly which transactions drove an overspend,
  plus how that category has trended over past months against its budget
- Edit the budget directly in the app and save it back to the same Excel file

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

Open `data/transactions.xlsx`. It has three sheets: **Transactions**, **Budget**,
and **Read Me** (which repeats these instructions inside the file itself).

### Transactions

Columns:

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

### Budget

Columns: **Category**, **Monthly Budget**, **Notes** (optional — jot down anything
worth remembering, e.g. "going up after rent renewal in March"). A few starter
rows are included with rough example amounts — edit them, and add or delete rows
so the categories match the ones you actually use in Transactions.

You can also add/edit budget rows directly from the **Budget vs Actual** page in
the app and click "Save budget to Excel" — it writes to the same file, so editing
in Excel or in the app both stay in sync.

### How Budget vs Actual works

- Pick a month in the sidebar. The page totals your actual spend per category
  for that month and compares it to what you budgeted.
- Categories are colour-coded: red = over budget, green = on track, grey = no
  budget set for that category yet.
- Use the "Drill into a category" section to see exactly which transactions
  pushed you over, and how that category has trended over previous months
  against its budget line.

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
