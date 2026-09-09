"""
Personal Spending Dashboard
----------------------------
Reads transactions from an Excel file (data/transactions.xlsx) and
visualises spending by category, over time, and by account.

Run with:  streamlit run app.py
"""

import os
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import load_workbook

DATA_PATH = os.path.join(os.path.dirname(__file__), "data", "transactions.xlsx")
REQUIRED_COLUMNS = ["Date", "Description", "Category", "Amount", "Account"]

st.set_page_config(page_title="Spending Dashboard", page_icon="💷", layout="wide")


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------
def _file_signature(path: str):
    """Used as a cache key so the app auto-refreshes when the Excel file changes."""
    try:
        return os.path.getmtime(path)
    except FileNotFoundError:
        return None


@st.cache_data(show_spinner=False)
def load_data(path: str, _signature) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="Transactions", engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required column(s): {', '.join(missing)}")

    df = df.dropna(subset=["Date", "Amount"]).copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)
    df["Category"] = df["Category"].fillna("Uncategorised").astype(str).str.strip()
    df["Account"] = df["Account"].fillna("Unknown").astype(str).str.strip()
    df["Description"] = df["Description"].fillna("").astype(str)
    df["Month"] = df["Date"].dt.to_period("M").dt.to_timestamp()

    # Convention: spending is a positive number, income/refunds are negative.
    # (If your bank exports the opposite sign convention, flip it in the sidebar.)
    return df.sort_values("Date")


def append_transaction(path: str, row: dict):
    """Append a single row to the Transactions sheet in the Excel file."""
    wb = load_workbook(path)
    ws = wb["Transactions"]
    headers = [cell.value for cell in ws[1]]
    ordered = [row.get(h, "") for h in headers]
    ws.append(ordered)
    wb.save(path)


# ---------------------------------------------------------------------------
# Sidebar controls
# ---------------------------------------------------------------------------
st.sidebar.title("💷 Spending Dashboard")

if not os.path.exists(DATA_PATH):
    st.error(
        f"Couldn't find **{DATA_PATH}**.\n\n"
        "Make sure `transactions.xlsx` is inside the `data/` folder, "
        "or edit `DATA_PATH` in app.py to point at your file."
    )
    st.stop()

flip_sign = st.sidebar.checkbox(
    "My bank exports spending as negative numbers",
    value=False,
    help="Tick this if positive = income and negative = spending in your export.",
)

if st.sidebar.button("🔄 Reload data"):
    st.cache_data.clear()

try:
    data = load_data(DATA_PATH, _file_signature(DATA_PATH))
except Exception as e:
    st.error(f"Couldn't read the data file: {e}")
    st.stop()

if flip_sign:
    data["Amount"] = -data["Amount"]

if data.empty:
    st.warning("No transactions found yet. Add some rows to data/transactions.xlsx and reload.")
    st.stop()

min_date, max_date = data["Date"].min().date(), data["Date"].max().date()
date_range = st.sidebar.date_input(
    "Date range", value=(min_date, max_date), min_value=min_date, max_value=max_date
)
if isinstance(date_range, tuple) and len(date_range) == 2:
    start_date, end_date = date_range
else:
    start_date, end_date = min_date, max_date

all_categories = sorted(data["Category"].unique())
selected_categories = st.sidebar.multiselect("Categories", all_categories, default=all_categories)

all_accounts = sorted(data["Account"].unique())
selected_accounts = st.sidebar.multiselect("Accounts", all_accounts, default=all_accounts)

mask = (
    (data["Date"].dt.date >= start_date)
    & (data["Date"].dt.date <= end_date)
    & (data["Category"].isin(selected_categories))
    & (data["Account"].isin(selected_accounts))
)
filtered = data.loc[mask].copy()

with st.sidebar.expander("➕ Add a transaction"):
    with st.form("add_transaction_form", clear_on_submit=True):
        new_date = st.date_input("Date", value=datetime.today())
        new_desc = st.text_input("Description")
        new_category = st.selectbox("Category", options=all_categories + ["+ New category"])
        if new_category == "+ New category":
            new_category = st.text_input("New category name")
        new_amount = st.number_input("Amount (spend as positive)", step=0.01, format="%.2f")
        new_account = st.selectbox("Account", options=all_accounts + ["+ New account"])
        if new_account == "+ New account":
            new_account = st.text_input("New account name")
        submitted = st.form_submit_button("Add")
        if submitted:
            amount_to_store = -new_amount if flip_sign else new_amount
            append_transaction(
                DATA_PATH,
                {
                    "Date": new_date,
                    "Description": new_desc,
                    "Category": new_category or "Uncategorised",
                    "Amount": amount_to_store,
                    "Account": new_account or "Unknown",
                },
            )
            st.cache_data.clear()
            st.success("Transaction added.")
            st.rerun()


# ---------------------------------------------------------------------------
# Headline metrics
# ---------------------------------------------------------------------------
spend = filtered[filtered["Amount"] > 0]
income = filtered[filtered["Amount"] < 0]

total_spend = spend["Amount"].sum()
total_income = -income["Amount"].sum()
net = total_income - total_spend

months_covered = max(filtered["Month"].nunique(), 1)
avg_monthly_spend = total_spend / months_covered

top_category = (
    spend.groupby("Category")["Amount"].sum().sort_values(ascending=False).index[0]
    if not spend.empty
    else "—"
)

col1, col2, col3, col4 = st.columns(4)
col1.metric("Total spend", f"£{total_spend:,.2f}")
col2.metric("Total income / refunds", f"£{total_income:,.2f}")
col3.metric("Net", f"£{net:,.2f}")
col4.metric("Avg monthly spend", f"£{avg_monthly_spend:,.2f}")

st.caption(f"Showing {len(filtered):,} transactions from {start_date} to {end_date}. Top category: **{top_category}**")

st.divider()

# ---------------------------------------------------------------------------
# Charts
# ---------------------------------------------------------------------------
left, right = st.columns([2, 1])

with left:
    st.subheader("Spending over time")
    monthly = spend.groupby("Month")["Amount"].sum().reset_index()
    fig_trend = px.bar(monthly, x="Month", y="Amount", labels={"Amount": "Spend (£)"})
    fig_trend.update_layout(margin=dict(l=0, r=0, t=10, b=0))
    st.plotly_chart(fig_trend, use_container_width=True)

with right:
    st.subheader("By category")
    by_cat = spend.groupby("Category")["Amount"].sum().sort_values(ascending=False).reset_index()
    fig_cat = px.pie(by_cat, names="Category", values="Amount", hole=0.45)
    fig_cat.update_layout(margin=dict(l=0, r=0, t=10, b=0), showlegend=True)
    st.plotly_chart(fig_cat, use_container_width=True)

st.subheader("Category breakdown by month")
cat_month = spend.groupby(["Month", "Category"])["Amount"].sum().reset_index()
fig_stack = px.bar(
    cat_month, x="Month", y="Amount", color="Category", labels={"Amount": "Spend (£)"}
)
fig_stack.update_layout(margin=dict(l=0, r=0, t=10, b=0), barmode="stack")
st.plotly_chart(fig_stack, use_container_width=True)

col_a, col_b = st.columns(2)

with col_a:
    st.subheader("Top merchants / descriptions")
    top_desc = (
        spend.groupby("Description")["Amount"]
        .sum()
        .sort_values(ascending=False)
        .head(10)
        .reset_index()
    )
    fig_desc = px.bar(top_desc, x="Amount", y="Description", orientation="h", labels={"Amount": "Spend (£)"})
    fig_desc.update_layout(margin=dict(l=0, r=0, t=10, b=0), yaxis=dict(autorange="reversed"))
    st.plotly_chart(fig_desc, use_container_width=True)

with col_b:
    st.subheader("By account")
    by_acct = spend.groupby("Account")["Amount"].sum().sort_values(ascending=False).reset_index()
    fig_acct = px.bar(by_acct, x="Account", y="Amount", labels={"Amount": "Spend (£)"})
    fig_acct.update_layout(margin=dict(l=0, r=0, t=10, b=0))
    st.plotly_chart(fig_acct, use_container_width=True)

st.divider()

st.subheader("Transactions")
st.dataframe(
    filtered.sort_values("Date", ascending=False)[
        ["Date", "Description", "Category", "Amount", "Account"]
    ],
    use_container_width=True,
    hide_index=True,
)

csv = filtered.to_csv(index=False).encode("utf-8")
st.download_button("Download filtered data as CSV", csv, "filtered_transactions.csv", "text/csv")
