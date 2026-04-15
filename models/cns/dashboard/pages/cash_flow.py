"""Cash Flow Forecast — the #1 priority page for CNS."""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from dashboard.data_store import DataStore
from dashboard.constants import (
    N, FORECAST_MONTH_LABELS, month_idx_to_label, fmt_currency, CF_ROWS,
)


def show():
    ds = DataStore.get()
    st.header("Cash Flow Forecast")
    st.caption(f"60-month projection | Actuals: Jan-Feb 2026 | Forecast: Mar 2026 - Dec 2030")

    # Sensitivity
    with st.expander("Sensitivity Adjustments", expanded=False):
        col1, col2, col3 = st.columns(3)
        with col1:
            vol_adj = st.slider("Volume adjustment %", -50, 50, 0, 5, format="%+d%%", key="vol_adj") / 100
        with col2:
            opex_adj = st.slider("OpEx adjustment %", -30, 30, 0, 5, format="%+d%%", key="opex_adj") / 100
        with col3:
            phys_rate = st.slider("Physician rate %", 0, 100,
                                   int(ds.get_assumptions().get('physician_services_rate', 90)),
                                   key="phys_rate_slider")

    # Apply adjustments
    assumptions = ds.get_assumptions()
    if vol_adj != 0:
        assumptions['bobas_volume'] = [max(0, round(v * (1 + vol_adj))) for v in assumptions['bobas_volume']]
        assumptions['gap_volume'] = [max(0, round(v * (1 + vol_adj))) for v in assumptions['gap_volume']]
    if opex_adj != 0:
        for key in ['marketing_monthly', 'contracts_monthly', 'office_software_monthly',
                     'rent_westlake_monthly', 'legal_monthly_recurring']:
            if key in assumptions:
                assumptions[key] = assumptions[key] * (1 + opex_adj)
    assumptions['physician_services_rate'] = phys_rate

    # Run forecast
    from financial_calcs import generate_cash_flow_forecast, generate_monthly_pl_forecast
    pl = generate_monthly_pl_forecast(assumptions)
    cf = generate_cash_flow_forecast(assumptions)

    # Overlay Jan/Feb actuals
    qbo = ds.actuals_2026
    n_act = ds.n_actuals_2026

    # --- KPIs ---
    ending_cash = cf['ending_cash']
    savings = cf['savings_balance']
    physician = cf['physician']

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        # Current cash (end of last actuals month)
        current = ending_cash[n_act - 1] if n_act > 0 else ending_cash[0]
        st.metric("Current Cash", fmt_currency(current))
    with col2:
        total_savings = savings[-1]
        st.metric("Savings (Dec 2030)", fmt_currency(total_savings))
    with col3:
        yr1_available = sum(cf['cash_after_overhead'][:12])
        st.metric("Cash Available (Year 1)", fmt_currency(yr1_available))
    with col4:
        yr1_physician = sum(physician[:12])
        st.metric("Physician Dist. (Year 1)", fmt_currency(yr1_physician))

    # --- Chart: Cash Collected vs Overhead vs Cash Available ---
    labels = FORECAST_MONTH_LABELS
    display_labels = labels[:36]

    # Cash available = cash collected - overhead (what's left to distribute/save/invest)
    cash_available = [cf['cash_in'][i] - cf['cash_overhead'][i] for i in range(36)]

    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=display_labels, y=cf['cash_in'][:36],
        name="Cash Collected", marker_color="#2ecc71", opacity=0.85,
    ))

    fig.add_trace(go.Bar(
        x=display_labels, y=[-v for v in cf['cash_overhead'][:36]],
        name="Overhead (Cash Out)", marker_color="#e74c3c", opacity=0.85,
    ))

    fig.add_trace(go.Scatter(
        x=display_labels, y=cash_available,
        name="Cash Available", line=dict(color="#2c3e50", width=3),
        mode="lines+markers", marker=dict(size=5),
        fill="tozeroy", fillcolor="rgba(44,62,80,0.08)",
    ))

    # Actuals/forecast divider
    if n_act > 0:
        fig.add_vline(x=n_act - 0.5, line_dash="dash", line_color="gray",
                       annotation_text="Forecast →")

    fig.update_layout(
        barmode="relative", height=450, margin=dict(t=30, b=40),
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        hovermode="x unified",
    )
    fig.update_yaxes(title_text="Monthly Cash Flow", tickformat="$,.0f")
    st.plotly_chart(fig, use_container_width=True)

    # --- Detail Table ---
    st.subheader("Cash Flow Detail")
    show_months = st.slider("Months to display", 12, 60, 24, key="cf_months")

    table_data = {
        "Beginning Cash": cf['beginning_cash'][:show_months],
        "Cash Collected": cf['cash_in'][:show_months],
        "Total Overhead": cf['cash_overhead'][:show_months],
        "Cash After Overhead": cf['cash_after_overhead'][:show_months],
        f"Distributable (>{fmt_currency(assumptions.get('minimum_cash_balance', 150000))})": cf['distributable'][:show_months],
        "Physician Services (90%)": cf['physician'][:show_months],
        "Savings Deposit (10%)": cf['savings_deposit'][:show_months],
        "Ending Cash": cf['ending_cash'][:show_months],
        "Savings Balance": cf['savings_balance'][:show_months],
    }

    df = pd.DataFrame(table_data, index=labels[:show_months]).T
    df.index.name = "Line Item"

    summary_rows = ["Cash After Overhead", "Ending Cash", "Savings Balance",
                     f"Distributable (>{fmt_currency(min_cash)})"]

    def _style(row):
        styles = []
        is_bold = any(s in str(row.name) for s in ["Ending", "After Overhead", "Distributable", "Savings Balance"])
        for val in row:
            s = "font-weight: bold; " if is_bold else ""
            if isinstance(val, (int, float)) and val < 0:
                s += "color: #e74c3c; "
            styles.append(s)
        return styles

    st.dataframe(
        df.style.apply(_style, axis=1).format("${:,.0f}"),
        use_container_width=True, height=400,
    )

    # Download
    csv = df.to_csv()
    st.download_button("Download CSV", csv, "cns_cash_flow.csv", "text/csv")
