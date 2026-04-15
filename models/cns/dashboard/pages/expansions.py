"""Expansion Planner — multi-location expansion modeling (5 slots)."""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go

from dashboard.data_store import DataStore
from dashboard.constants import FORECAST_MONTH_LABELS, fmt_currency


def show():
    ds = DataStore.get()
    st.header("Expansion Planner")

    expansions = ds.get_expansions()
    exp_detail = ds.run_expansion_detail()

    # --- Summary KPIs ---
    details_list = exp_detail.get('details', [])
    total_ti = sum(sum(d.get('ti', [0]*60)) for d in details_list if d.get('enabled'))
    total_lease_yr1 = sum(sum(d.get('lease', [0]*60)[:12]) for d in details_list if d.get('enabled'))
    total_opex_yr1 = sum(sum(d.get('opex', [0]*60)[:12]) for d in details_list if d.get('enabled'))
    active = sum(1 for d in details_list if d.get('enabled'))

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Active Expansions", active)
    with col2:
        st.metric("Total TI Investment", fmt_currency(total_ti))
    with col3:
        st.metric("Year 1 Lease Cost", fmt_currency(total_lease_yr1))
    with col4:
        st.metric("Year 1 OpEx", fmt_currency(total_opex_yr1))

    st.divider()

    # --- Total Expansion Cost Chart ---
    exp_total = exp_detail.get('total', [0]*60)
    labels = FORECAST_MONTH_LABELS[:36]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=labels, y=exp_total[:36],
        marker_color="#e74c3c", opacity=0.8, name="Monthly Expansion Cost",
    ))
    fig.update_layout(
        height=350, margin=dict(t=10, b=30),
        yaxis_tickformat="$,.0f",
        legend=dict(orientation="h", y=1.1),
    )
    st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # --- Per-Expansion Editors ---
    for idx, exp in enumerate(expansions):
        with st.expander(f"{'🟢' if exp['enabled'] else '⚪'} {exp['name']}", expanded=(idx == 0 and exp['enabled'])):
            col1, col2 = st.columns([1, 3])
            with col1:
                enabled = st.toggle("Enabled", value=exp['enabled'], key=f"exp_en_{idx}")

            if enabled:
                col1, col2, col3 = st.columns(3)
                with col1:
                    name = st.text_input("Location Name", value=exp['name'], key=f"exp_name_{idx}")
                    lease = st.number_input("Lease Monthly", value=int(exp['lease_monthly']),
                                             step=500, key=f"exp_lease_{idx}")
                    lease_start = st.number_input("Lease Start Month (0=Jan-26)", value=int(exp['lease_start_month']),
                                                   min_value=0, max_value=59, key=f"exp_ls_{idx}")
                with col2:
                    ti_share = st.number_input("TI (CNS Share)", value=int(exp['ti_cns_share']),
                                                step=10000, key=f"exp_ti_{idx}")
                    ti_start = st.number_input("TI Start Month", value=int(exp['ti_start_month']),
                                                min_value=0, max_value=59, key=f"exp_tis_{idx}")
                    ti_dur = st.number_input("TI Duration (months)", value=int(exp['ti_duration_months']),
                                              min_value=1, max_value=12, key=f"exp_tid_{idx}")
                with col3:
                    ffe = st.number_input("FF&E Budget", value=int(exp['ffe_budget']),
                                           step=5000, key=f"exp_ffe_{idx}")
                    opex_full = st.number_input("OpEx Monthly (full)", value=int(exp['opex_monthly']),
                                                 step=5000, key=f"exp_opex_{idx}")
                    opex_ramp = st.number_input("OpEx Ramp Monthly", value=int(exp.get('opex_ramp_monthly', 0)),
                                                 step=1000, key=f"exp_ramp_{idx}")
                    ramp_months = st.number_input("Ramp Months", value=int(exp.get('opex_ramp_months', 0)),
                                                   min_value=0, max_value=24, key=f"exp_rm_{idx}")

                # Update expansion
                expansions[idx] = {
                    'name': name, 'enabled': enabled,
                    'lease_start_month': lease_start, 'lease_monthly': lease,
                    'ti_total': exp.get('ti_total', ti_share * 2), 'ti_cns_share': ti_share,
                    'ti_start_month': ti_start, 'ti_duration_months': ti_dur,
                    'ffe_budget': ffe, 'opex_monthly': opex_full,
                    'opex_ramp_monthly': opex_ramp, 'opex_ramp_months': ramp_months,
                }

                # Per-location cost summary
                if idx < len(details_list):
                    detail = details_list[idx]
                    total_cost = sum(detail.get('total', [0]*60))
                    st.caption(f"Total 5-year cost: {fmt_currency(total_cost)}")
            else:
                expansions[idx]['enabled'] = False

    if st.button("Save Expansions", type="primary"):
        ds.set_expansions(expansions)
        st.success("Expansions saved")
        st.rerun()
