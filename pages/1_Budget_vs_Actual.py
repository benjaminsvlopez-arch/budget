"""
Budget vs Actual
----------------
Compares actual spend (from the Transactions sheet) against a monthly budget
per category (from the Budget sheet), and helps you drill into *why* and
*where* you went over.
"""

import os

import pandas as pd
import plotly.express as px
import streamlit as st

from common import DATA_PATH, file_signature, load_budget, load_data, save_budget

st.set_page_config(page_title="Spending Dashboard - Budget vs Actual", page_icon="🎯", layout="wide")

st.title("🎯 Budget vs Actual")
st.caption(
    "Set a monthly budget per category on the **Budget** tab of your Excel file "
    "(or edit it right here), then see exactly where you're over or under."
)

if not os.path.exists(DATA_PATH):
    st.error(f"Couldn't find **{DATA_PATH}**. Set it up from the Overview page first.")
    st.stop()

if st.sidebar.button("🔄 Reload data"):
    st.cache_data.clear()

try:
    data = load_data(DATA_PATH, file_signature(DATA_PATH))
    budget = load_budget(DATA_PATH, file_signature(DATA_PATH))
except Exception as e:
    st.error(f"Couldn't read the data file: {e}")
    st.stop()

if data.empty:
    st.warning("No transactions yet — add some on the Overview page or in Excel first.")
    st.stop()

spend_all = data[data["Amount"] > 0].copy()

# ---------------------------------------------------------------------------
# Month picker
# ---------------------------------------------------------------------------
month_options = sorted(spend_all["Month"].unique(), reverse=True)
month_labels = [pd.Timestamp(m).strftime("%B %Y") for m in month_options]

st.sidebar.subheader("Month")
selected_label = st.sidebar.selectbox("Compare budget for", month_labels, index=0)
selected_month = month_options[month_labels.index(selected_label)]

month_spend = spend_all[spend_all["Month"] == selected_month]
actual_by_cat = month_spend.groupby("Category")["Amount"].sum()

budget_indexed = budget.set_index("Category")["Monthly Budget"] if not budget.empty else pd.Series(dtype=float)
notes_indexed = budget.set_index("Category")["Notes"] if not budget.empty else pd.Series(dtype=str)

all_categories = sorted(set(budget_indexed.index) | set(actual_by_cat.index))

if not all_categories:
    st.info("No categories to compare yet.")
    st.stop()

rows = []
for cat in all_categories:
    budgeted = float(budget_indexed.get(cat, 0) or 0)
    actual = float(actual_by_cat.get(cat, 0) or 0)
    remaining = budgeted - actual
    pct_used = (actual / budgeted * 100) if budgeted > 0 else (100.0 if actual == 0 else float("inf"))
    rows.append(
        {
            "Category": cat,
            "Budgeted": budgeted,
            "Actual": actual,
            "Remaining": remaining,
            "% Used": pct_used,
            "Status": "No budget set" if budgeted == 0 else ("Over budget" if remaining < 0 else "On track"),
        }
    )

summary = pd.DataFrame(rows).sort_values("Remaining")

total_budget = summary["Budgeted"].sum()
total_actual = summary["Actual"].sum()
total_remaining = total_budget - total_actual
overspent_categories = summary[summary["Remaining"] < 0]

# ---------------------------------------------------------------------------
# Headline metrics
# ---------------------------------------------------------------------------
col1, col2, col3, col4 = st.columns(4)
col1.metric("Total budget", f"£{total_budget:,.2f}")
col2.metric("Total actual spend", f"£{total_actual:,.2f}")
col3.metric(
    "Overall remaining",
    f"£{total_remaining:,.2f}",
    delta=None if total_budget == 0 else f"{total_remaining/total_budget*100:,.0f}% of budget",
)
col4.metric("Categories over budget", f"{len(overspent_categories)} / {len(summary)}")

st.caption(f"Showing **{selected_label}** — {len(month_spend)} transactions.")

if not overspent_categories.empty:
    worst = overspent_categories.sort_values("Remaining").iloc[0]
    st.warning(
        f"You're over budget in **{len(overspent_categories)}** categor"
        f"{'y' if len(overspent_categories) == 1 else 'ies'} this month. "
        f"Biggest overspend: **{worst['Category']}**, £{abs(worst['Remaining']):,.2f} over."
    )
else:
    st.success("You're within budget in every category this month. 🎉")

st.divider()

# ---------------------------------------------------------------------------
# Budget vs Actual table + chart
# ---------------------------------------------------------------------------
left, right = st.columns([3, 2])

with left:
    st.subheader("By category")

    def highlight_row(row):
        if row["Status"] == "Over budget":
            return ["background-color: #fde2e2"] * len(row)
        if row["Status"] == "No budget set":
            return ["background-color: #f5f5f5"] * len(row)
        return ["background-color: #e3f7e3"] * len(row)

    display_df = summary.copy()
    display_df["Budgeted"] = display_df["Budgeted"].map(lambda x: f"£{x:,.2f}")
    display_df["Actual"] = display_df["Actual"].map(lambda x: f"£{x:,.2f}")
    display_df["Remaining"] = display_df["Remaining"].map(lambda x: f"£{x:,.2f}")
    display_df["% Used"] = display_df["% Used"].map(lambda x: "—" if x == float("inf") else f"{x:,.0f}%")

    styled = summary.style.apply(highlight_row, axis=1)
    st.dataframe(
        styled,
        width='stretch',
        hide_index=True,
        column_config={
            "Budgeted": st.column_config.NumberColumn(format="£%.2f"),
            "Actual": st.column_config.NumberColumn(format="£%.2f"),
            "Remaining": st.column_config.NumberColumn(format="£%.2f"),
            "% Used": st.column_config.NumberColumn(format="%.0f%%"),
        },
    )

with right:
    st.subheader("Budget vs actual")
    chart_df = summary.melt(
        id_vars="Category", value_vars=["Budgeted", "Actual"], var_name="Type", value_name="Amount"
    )
    fig = px.bar(
        chart_df,
        x="Amount",
        y="Category",
        color="Type",
        barmode="group",
        orientation="h",
        color_discrete_map={"Budgeted": "#94a3b8", "Actual": "#2E5266"},
    )
    fig.update_layout(margin=dict(l=0, r=0, t=10, b=0), yaxis=dict(categoryorder="total ascending"))
    st.plotly_chart(fig, width='stretch')

st.divider()

# ---------------------------------------------------------------------------
# Drill down: why and where
# ---------------------------------------------------------------------------
st.subheader("Drill into a category — why and where")

default_idx = 0
if not overspent_categories.empty:
    worst_cat = overspent_categories.sort_values("Remaining").iloc[0]["Category"]
    default_idx = all_categories.index(worst_cat) if worst_cat in all_categories else 0

drill_cat = st.selectbox("Category", sorted(all_categories), index=min(default_idx, len(all_categories) - 1))

cat_month_txns = month_spend[month_spend["Category"] == drill_cat].sort_values("Amount", ascending=False)
cat_note = notes_indexed.get(drill_cat, "")

dcol1, dcol2 = st.columns([3, 2])

with dcol1:
    st.markdown(f"**Top transactions in {drill_cat} this month**")
    if cat_month_txns.empty:
        st.caption("No transactions in this category this month.")
    else:
        st.dataframe(
            cat_month_txns[["Date", "Description", "Amount", "Account"]],
            width='stretch',
            hide_index=True,
            column_config={"Amount": st.column_config.NumberColumn(format="£%.2f")},
        )
    if cat_note:
        st.caption(f"📝 Note on this category: {cat_note}")

with dcol2:
    st.markdown(f"**{drill_cat} — trend over time**")
    trend = (
        spend_all[spend_all["Category"] == drill_cat]
        .groupby("Month")["Amount"]
        .sum()
        .reset_index()
        .sort_values("Month")
    )
    cat_budget_val = float(budget_indexed.get(drill_cat, 0) or 0)
    fig_trend = px.bar(trend, x="Month", y="Amount", labels={"Amount": "Spend (£)"})
    if cat_budget_val > 0:
        fig_trend.add_hline(y=cat_budget_val, line_dash="dash", line_color="#dc2626",
                             annotation_text="Budget", annotation_position="top left")
    fig_trend.update_layout(margin=dict(l=0, r=0, t=10, b=0))
    st.plotly_chart(fig_trend, width='stretch')

st.divider()

# ---------------------------------------------------------------------------
# Edit budget
# ---------------------------------------------------------------------------
st.subheader("Edit your budget")
st.caption(
    "Edit amounts directly here and click Save — or edit the **Budget** tab in "
    "`transactions.xlsx` yourself and reload. Both update the same file."
)

edit_source = budget.copy() if not budget.empty else pd.DataFrame(columns=["Category", "Monthly Budget", "Notes"])
edited = st.data_editor(
    edit_source,
    num_rows="dynamic",
    width='stretch',
    hide_index=True,
    column_config={
        "Monthly Budget": st.column_config.NumberColumn(format="£%.2f", min_value=0),
    },
    key="budget_editor",
)

if st.button("💾 Save budget to Excel"):
    save_budget(DATA_PATH, edited)
    st.cache_data.clear()
    st.success("Budget saved.")
    st.rerun()
