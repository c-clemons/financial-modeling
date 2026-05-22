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
    qbo_months = ds.actuals_2026.get("months", [])
    n_act_header = ds.n_actuals_2026
    if qbo_months and n_act_header < len(FORECAST_MONTH_LABELS):
        actuals_range = f"{qbo_months[0]} - {qbo_months[-1]}"
        forecast_start = FORECAST_MONTH_LABELS[n_act_header]
        st.caption(
            f"60-month projection | Actuals: {actuals_range} | "
            f"Forecast: {forecast_start} - Dec 2030"
        )
    else:
        st.caption("60-month projection | Forecast: Jan 2026 - Dec 2030")

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

    # Location toggle
    from financial_calcs import generate_cash_flow_forecast, generate_monthly_pl_forecast, generate_pl_by_location
    locations = ds.get_locations()
    selected_view = st.selectbox("View", ["Consolidated"] + locations, key="cf_view")

    # Run forecast
    pl_by_loc = generate_pl_by_location(assumptions)
    cf = generate_cash_flow_forecast(assumptions)

    if selected_view == "Consolidated":
        pl = pl_by_loc['consolidated']
    else:
        pl = pl_by_loc[selected_view]

    # Overlay Jan/Feb actuals
    qbo = ds.actuals_2026
    n_act = ds.n_actuals_2026

    # --- KPIs (always consolidated for cash management) ---
    ending_cash = cf['ending_cash']
    savings = cf['savings_balance']
    physician = cf['physician']

    # For chart: use selected view's revenue/overhead
    chart_collected = pl.get('total_income', pl.get('total_collected', [0]*60))
    chart_overhead = pl.get('total_overhead', [0]*60)

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

    # Cash available = collected - overhead
    cash_available = [chart_collected[i] - chart_overhead[i] for i in range(min(36, len(chart_collected)))]

    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=display_labels, y=chart_collected[:36],
        name="Cash Collected", marker_color="#2ecc71", opacity=0.85,
    ))

    fig.add_trace(go.Bar(
        x=display_labels, y=[-v for v in chart_overhead[:36]],
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
                     "Distributable"]

    # Actual months get an amber tint; forecast months stay default. The
    # range matches what's been committed via the Upload Actuals page —
    # ``n_act`` reflects the P&L upload, and ending_cash for those months
    # is anchored to the balance-sheet upload (see generate_cash_flow_forecast).
    actual_cols_in_view = labels[:min(n_act, show_months)]
    actual_col_set = set(actual_cols_in_view)

    if actual_cols_in_view:
        first_forecast = (labels[n_act]
                          if n_act < len(labels) and n_act < show_months
                          else None)
        if first_forecast:
            st.caption(
                f"📊 **Actuals** (Jan-26 → {actual_cols_in_view[-1]}, "
                f"amber-tinted) from QBO P&L + Balance Sheet uploads. "
                f"**Forecast** ({first_forecast} onward) is model output."
            )
        else:
            st.caption(
                f"📊 **Actuals** (Jan-26 → {actual_cols_in_view[-1]}, "
                "amber-tinted) from QBO P&L + Balance Sheet uploads."
            )

    def _style(row):
        styles = []
        is_bold = any(s in str(row.name) for s in ["Ending", "After Overhead", "Distributable", "Savings Balance"])
        for val in row:
            s = "font-weight: bold; " if is_bold else ""
            if isinstance(val, (int, float)) and val < 0:
                s += "color: #e74c3c; "
            styles.append(s)
        return styles

    def _shade_actuals(col):
        if col.name in actual_col_set:
            return ['background-color: #fff8e1'] * len(col)
        return [''] * len(col)

    styler = df.style.apply(_style, axis=1).apply(_shade_actuals, axis=0).format("${:,.0f}")
    if actual_cols_in_view:
        header_styles = [
            {
                'selector': f'th.col_heading.level0.col{i}',
                'props': [
                    ('background-color', '#fff3c4'),
                    ('font-weight', '600'),
                ],
            }
            for i in range(len(actual_cols_in_view))
        ]
        styler = styler.set_table_styles(header_styles, overwrite=False)

    st.dataframe(styler, use_container_width=True, height=400)

    # Download
    csv = df.to_csv()
    st.download_button("Download CSV", csv, "cns_cash_flow.csv", "text/csv")
