"""Monthly P&L — 2025 actuals + 2026-2030 forecast with QBO actuals overlay."""

import streamlit as st
import pandas as pd
import numpy as np

from dashboard.data_store import DataStore
from dashboard.constants import (
    N, FORECAST_MONTH_LABELS, MONTHS_12, fmt_currency,
)


def show():
    ds = DataStore.get()
    st.header("Monthly P&L")

    tab_2025, tab_forecast = st.tabs(["2025 Actuals", "2026-2030 Forecast"])

    # =====================================================================
    # TAB 1: 2025 Actuals
    # =====================================================================
    with tab_2025:
        st.subheader("2025 Actuals (from QBO)")
        a25 = ds.actuals_2025

        rows_25 = {
            "INCOME": None,
            "  400 Fee Income": a25['fee_income'],
            "  410 Reimbursed Expense": a25['reimbursed_expense_income'],
            "  430 Refunds": a25['refunds'],
            "TOTAL INCOME": [a25['fee_income'][i] + a25['reimbursed_expense_income'][i] + a25['refunds'][i] for i in range(12)],
            "": None,
            "KEY EXPENSES": None,
            "  592 Physician Services": a25['physician_services'],
            "  554 Billing Services": a25['billing_services'],
            "  584 Salaries & Wages": a25['salaries_wages'],
            "  580 Payroll Processing": a25.get('payroll_processing', [0]*12),
        }

        df_25 = pd.DataFrame(
            {label: vals for label, vals in rows_25.items() if vals is not None},
            index=MONTHS_12,
        ).T

        # Add FY total
        df_25["FY 2025"] = df_25.sum(axis=1)

        def _style_25(row):
            styles = []
            is_total = "TOTAL" in str(row.name) or row.name == ""
            for val in row:
                s = "font-weight: bold; " if is_total else ""
                if isinstance(val, (int, float)) and val < 0:
                    s += "color: #e74c3c; "
                styles.append(s)
            return styles

        st.dataframe(
            df_25.style.apply(_style_25, axis=1).format("${:,.0f}"),
            use_container_width=True, height=400,
        )

        # Key metrics
        t = ds.actuals_2025_totals
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("FY 2025 Revenue", fmt_currency(t['total_income']))
        with col2:
            st.metric("FY 2025 Expenses", fmt_currency(t['total_expenses']))
        with col3:
            st.metric("FY 2025 Net Income", fmt_currency(t['net_income']))

    # =====================================================================
    # TAB 2: 2026-2030 Forecast
    # =====================================================================
    with tab_forecast:
        st.subheader("2026-2030 Forecast (Cash Basis)")

        # Run forecast
        forecast = ds.run_forecast()
        pl = forecast['pl']

        # Controls
        col1, col2 = st.columns(2)
        with col1:
            show_months = st.slider("Months to display", 12, 60, 24, key="pl_months")
        with col2:
            show_accrual = st.checkbox("Show accrual reference", value=False)

        labels = FORECAST_MONTH_LABELS[:show_months]
        n_act = ds.n_actuals_2026
        qbo = ds.actuals_2026

        # Build P&L rows
        rows = {}

        # Surgery volumes
        rows["Bobas Volume"] = pl['bobas_volume'][:show_months]
        rows["GAP Volume"] = pl['gap_volume'][:show_months]
        rows["Total Surgeries"] = pl['total_volume'][:show_months]

        if show_accrual:
            rows["---"] = [None] * show_months
            rows["Bobas Earned (Accrual)"] = pl['bobas_earned'][:show_months]
            rows["GAP Earned (Accrual)"] = pl['gap_earned'][:show_months]
            rows["Total Earned (Accrual)"] = pl['total_earned'][:show_months]

        rows[""] = [None] * show_months
        rows["CASH COLLECTED"] = [None] * show_months
        rows["  Bobas Collected"] = pl['bobas_collected'][:show_months]
        rows["  GAP Collected"] = pl['gap_collected'][:show_months]
        rows["  Historical AR"] = [pl['historical_ar_boba'][i] + pl['historical_ar_gap'][i] for i in range(len(pl['historical_ar_boba']))][:show_months]
        rows["TOTAL INCOME"] = pl['total_collected'][:show_months]

        rows[" "] = [None] * show_months
        rows["OVERHEAD"] = [None] * show_months
        rows["  Billing (18%)"] = pl['billing_fees'][:show_months]
        rows["  Payroll (W-2)"] = pl['payroll'][:show_months]
        rows["  Contractors"] = pl['contractors'][:show_months]
        rows["  Operating Expenses"] = pl['total_opex'][:show_months]
        rows["  Expansion Costs"] = pl.get('expansion_total', [0]*60)[:show_months]
        rows["TOTAL OVERHEAD"] = pl['total_overhead'][:show_months]

        rows["  "] = [None] * show_months
        rows["NET EQUITY"] = pl['net_equity'][:show_months]
        rows["  Physician (90%)"] = pl['physician_services'][:show_months]
        rows["NET INCOME (CNS)"] = pl['net_income'][:show_months]

        # Overlay actuals for Jan/Feb
        if n_act > 0:
            for i in range(n_act):
                rows["TOTAL INCOME"][i] = qbo['total_income'][i]
                rows["  Bobas Collected"][i] = 0
                rows["  GAP Collected"][i] = 0
                rows["  Historical AR"][i] = 0
                rows["TOTAL OVERHEAD"][i] = qbo['total_expenses'][i] - qbo['physician_services'][i]
                rows["  Billing (18%)"][i] = 0
                rows["  Payroll (W-2)"][i] = qbo['payroll_expenses'][i]
                rows["  Contractors"][i] = qbo['contracts'][i]
                rows["  Operating Expenses"][i] = (
                    qbo['total_expenses'][i] - qbo['physician_services'][i]
                    - qbo['payroll_expenses'][i] - qbo['contracts'][i])
                rows["  Expansion Costs"][i] = 0
                rows["NET EQUITY"][i] = qbo['net_income'][i] + qbo['physician_services'][i]
                rows["  Physician (90%)"][i] = qbo['physician_services'][i]
                rows["NET INCOME (CNS)"][i] = qbo['net_income'][i]

        # Build DataFrame
        df = pd.DataFrame(
            {label: vals for label, vals in rows.items() if vals is not None and any(v is not None for v in vals)},
            index=labels,
        ).T
        df.index.name = "Line Item"

        # Style
        bold_rows = ["TOTAL INCOME", "TOTAL OVERHEAD", "NET EQUITY", "NET INCOME (CNS)", "Total Surgeries"]

        def _style(row):
            styles = []
            is_bold = row.name in bold_rows
            for val in row:
                s = "font-weight: bold; " if is_bold else ""
                if isinstance(val, (int, float)) and val < 0:
                    s += "color: #e74c3c; "
                styles.append(s)
            return styles

        st.dataframe(
            df.style.apply(_style, axis=1).format(
                lambda x: f"${x:,.0f}" if isinstance(x, (int, float)) else "",
                na_rep="",
            ),
            use_container_width=True, height=700,
        )

        # Year summaries
        st.subheader("Annual Summary")
        metrics = ds.run_dashboard_metrics()
        years = sorted([k for k in metrics.keys() if k != '2025'])

        annual_data = []
        for yr in years:
            m = metrics[yr]
            annual_data.append({
                "Year": yr,
                "Surgeries": m['surgeries_total'],
                "Gross Revenue": m['gross_revenue'],
                "Physician Dist.": m['physician_services'],
                "Savings Deposits": m['savings_deposits'],
                "Locations": m['locations'],
                "Min Cash": m['min_cash'],
            })

        adf = pd.DataFrame(annual_data).set_index("Year")
        st.dataframe(adf.style.format({
            "Surgeries": "{:.0f}",
            "Gross Revenue": "${:,.0f}",
            "Physician Dist.": "${:,.0f}",
            "Savings Deposits": "${:,.0f}",
            "Locations": "{:.0f}",
            "Min Cash": "${:,.0f}",
        }), use_container_width=True)
