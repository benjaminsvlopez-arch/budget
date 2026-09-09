"""
Shared helpers for the Spending Dashboard app.
Both app.py (Overview) and pages/1_Budget_vs_Actual.py import from here so
there's one source of truth for reading/writing the Excel file.
"""

import os

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

DATA_PATH = os.path.join(os.path.dirname(__file__), "data", "transactions.xlsx")

TRANSACTION_COLUMNS = ["Date", "Description", "Category", "Amount", "Account"]
BUDGET_COLUMNS = ["Category", "Monthly Budget", "Notes"]


def file_signature(path: str):
    """Used as a cache key so the app auto-refreshes when the Excel file changes."""
    try:
        return os.path.getmtime(path)
    except FileNotFoundError:
        return None


@st.cache_data(show_spinner=False)
def load_data(path: str, _signature) -> pd.DataFrame:
    """Load and clean the Transactions sheet."""
    df = pd.read_excel(path, sheet_name="Transactions", engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in TRANSACTION_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Transactions sheet is missing column(s): {', '.join(missing)}")

    df = df.dropna(subset=["Date", "Amount"]).copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)
    df["Category"] = df["Category"].fillna("Uncategorised").astype(str).str.strip()
    df["Account"] = df["Account"].fillna("Unknown").astype(str).str.strip()
    df["Description"] = df["Description"].fillna("").astype(str)
    df["Month"] = df["Date"].dt.to_period("M").dt.to_timestamp()

    return df.sort_values("Date")


@st.cache_data(show_spinner=False)
def load_budget(path: str, _signature) -> pd.DataFrame:
    """Load the Budget sheet. Returns an empty frame with the right columns if the sheet is missing."""
    try:
        df = pd.read_excel(path, sheet_name="Budget", engine="openpyxl")
    except ValueError:
        # Sheet doesn't exist yet (older data file) - return an empty budget.
        return pd.DataFrame(columns=BUDGET_COLUMNS)

    df.columns = [str(c).strip() for c in df.columns]
    for col in BUDGET_COLUMNS:
        if col not in df.columns:
            df[col] = "" if col == "Notes" else 0

    df["Category"] = df["Category"].fillna("").astype(str).str.strip()
    df["Monthly Budget"] = pd.to_numeric(df["Monthly Budget"], errors="coerce").fillna(0)
    df["Notes"] = df["Notes"].fillna("").astype(str)
    df = df[df["Category"] != ""]

    return df[BUDGET_COLUMNS].reset_index(drop=True)


def append_transaction(path: str, row: dict):
    """Append a single row to the Transactions sheet, leaving every other sheet untouched."""
    wb = load_workbook(path)
    ws = wb["Transactions"]
    headers = [cell.value for cell in ws[1]]
    ordered = [row.get(h, "") for h in headers]
    ws.append(ordered)
    wb.save(path)


def save_budget(path: str, budget_df: pd.DataFrame):
    """Overwrite the Budget sheet with the given dataframe, leaving every other sheet untouched."""
    clean = budget_df.copy()
    clean["Category"] = clean["Category"].fillna("").astype(str).str.strip()
    clean = clean[clean["Category"] != ""]
    clean["Monthly Budget"] = pd.to_numeric(clean["Monthly Budget"], errors="coerce").fillna(0)
    clean["Notes"] = clean.get("Notes", "").fillna("") if "Notes" in clean else ""
    clean = clean[BUDGET_COLUMNS]

    with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        clean.to_excel(writer, sheet_name="Budget", index=False)

    # Re-apply header styling (ExcelWriter/replace writes a plain sheet).
    wb = load_workbook(path)
    ws = wb["Budget"]
    from openpyxl.styles import Font, PatternFill, Alignment

    for col_idx in range(1, len(BUDGET_COLUMNS) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = Font(name="Arial", bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="2E5266", end_color="2E5266", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 40
    ws.freeze_panes = "A2"
    wb.save(path)
