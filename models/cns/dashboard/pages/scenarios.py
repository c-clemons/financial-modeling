"""Scenarios — save, load, compare named forecast scenarios."""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import copy

from dashboard.data_store import DataStore
from dashboard.constants import FORECAST_MONTH_LABELS, fmt_currency
from financial_calcs import generate_cash_flow_forecast


def show():
    ds = DataStore.get()
    st.header("Scenarios")

    tab_manage, tab_compare = st.tabs(["Manage", "Compare"])

    with tab_manage:
        st.subheader("Save Current Forecast")
        col1, col2 = st.columns([3, 1])
        with col1:
            name = st.text_input("Scenario Name", placeholder="e.g., Base Case")
        with col2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Save", type="primary") and name:
                clean = name.strip().replace(" ", "_").lower()
                ds.save_scenario(clean)
                st.success(f"Saved: {name}")
                st.rerun()

        st.divider()
        st.subheader("Saved Scenarios")
        scenarios = ds.list_scenarios()
        if not scenarios:
            st.info("No scenarios saved yet.")
        else:
            for sc in scenarios:
                col1, col2, col3 = st.columns([3, 2, 1])
                with col1:
                    st.markdown(f"**{sc['name']}**")
                with col2:
                    st.caption(f"Saved: {sc['saved_at'][:16]}")
                with col3:
                    if st.button("Load", key=f"load_{sc['name']}"):
                        ds.load_scenario(sc['name'])
                        st.success(f"Loaded: {sc['name']}")
                        st.rerun()
                    if st.button("Delete", key=f"del_{sc['name']}"):
                        ds.delete_scenario(sc['name'])
                        st.rerun()

    with tab_compare:
        st.subheader("Compare Scenarios")
        scenarios = ds.list_scenarios()
        if not scenarios:
            st.info("Save at least one scenario to compare.")
            return

        selected = st.multiselect("Select scenarios", [s['name'] for s in scenarios],
                                   default=[s['name'] for s in scenarios[:3]])
        if not selected:
            return

        # Build cash flows
        results = {"Current": generate_cash_flow_forecast(ds.get_assumptions())}
        for sc_name in selected:
            from pathlib import Path
            import json
            sc_path = Path(__file__).parent.parent / "data" / "scenarios" / f"{sc_name}.json"
            if sc_path.exists():
                with open(sc_path) as f:
                    sc_overrides = json.load(f)
                sc_merged = DataStore._deep_merge(ds.defaults, sc_overrides)
                results[sc_name] = generate_cash_flow_forecast(sc_merged)

        labels = FORECAST_MONTH_LABELS[:36]
        colors = ["#2c3e50", "#e74c3c", "#2ecc71", "#3498db", "#f39c12"]

        # Ending Cash
        st.subheader("Ending Cash Balance")
        fig = go.Figure()
        for i, (name, cf) in enumerate(results.items()):
            fig.add_trace(go.Scatter(
                x=labels, y=cf['ending_cash'][:36],
                mode="lines+markers", name=name,
                line=dict(color=colors[i % len(colors)], width=2 if i == 0 else 1.5),
            ))
        fig.add_hline(y=ds.get_assumptions().get('minimum_cash_balance', 150000),
                       line_dash="dash", line_color="red", annotation_text="Minimum")
        fig.update_layout(height=400, yaxis_tickformat="$,.0f",
                           legend=dict(orientation="h", y=1.1), hovermode="x unified")
        st.plotly_chart(fig, use_container_width=True)

        # Metrics table
        st.subheader("Key Metrics (Year 1)")
        metrics = []
        for name, cf in results.items():
            yr1 = slice(0, 12)
            metrics.append({
                "Scenario": name,
                "Total Collected": sum(cf['cash_in'][yr1]),
                "Total Overhead": sum(cf['cash_overhead'][yr1]),
                "Physician Dist.": sum(cf['physician'][yr1]),
                "Savings Deposits": sum(cf['savings_deposit'][yr1]),
                "Ending Cash (Mo 12)": cf['ending_cash'][11],
                "Min Cash": min(cf['ending_cash'][yr1]),
            })
        mdf = pd.DataFrame(metrics).set_index("Scenario")
        st.dataframe(mdf.style.format("${:,.0f}"), use_container_width=True)
