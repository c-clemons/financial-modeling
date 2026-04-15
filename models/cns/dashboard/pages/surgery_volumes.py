"""Surgery Volumes — editable BOBA/GAP forecast with historical context."""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go

from dashboard.data_store import DataStore
from dashboard.constants import (
    N, FORECAST_MONTH_LABELS, SURGERY_TYPES, fmt_currency,
)
from baseline_data import (
    BOBA_VOLUME_2024, BOBA_VOLUME_2025, GAP_VOLUME_2024, GAP_VOLUME_2025,
    MONTHS_12,
)


def show():
    ds = DataStore.get()
    st.header("Surgery Volume Forecast")

    bobas, gap = ds.get_surgery_volumes()
    n_act = ds.n_actuals_2026

    # --- Historical Context ---
    st.subheader("Historical Volumes")
    col1, col2 = st.columns(2)
    with col1:
        st.caption("2024 (Sep-Dec)")
        hist_24 = pd.DataFrame({
            "Month": ["Sep", "Oct", "Nov", "Dec"],
            "BOBA": BOBA_VOLUME_2024,
            "GAP": GAP_VOLUME_2024,
            "Total": [b + g for b, g in zip(BOBA_VOLUME_2024, GAP_VOLUME_2024)],
        }).set_index("Month")
        st.dataframe(hist_24, use_container_width=True)
    with col2:
        st.caption("2025 (Full Year)")
        hist_25 = pd.DataFrame({
            "Month": MONTHS_12,
            "BOBA": BOBA_VOLUME_2025,
            "GAP": GAP_VOLUME_2025,
            "Total": [b + g for b, g in zip(BOBA_VOLUME_2025, GAP_VOLUME_2025)],
        }).set_index("Month")
        st.dataframe(hist_25, use_container_width=True)

    avg_boba_25 = np.mean(BOBA_VOLUME_2025)
    avg_gap_25 = np.mean(GAP_VOLUME_2025)
    st.caption(f"2025 averages: BOBA {avg_boba_25:.1f}/mo, GAP {avg_gap_25:.1f}/mo, Total {avg_boba_25 + avg_gap_25:.1f}/mo")

    st.divider()

    # --- Forecast Chart ---
    st.subheader("2026-2030 Forecast")

    show_months = st.slider("Months to display", 12, 60, 36, key="vol_months")
    labels = FORECAST_MONTH_LABELS[:show_months]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=labels, y=bobas[:show_months], name="BOBA",
        marker_color="#2c3e50",
    ))
    fig.add_trace(go.Bar(
        x=labels, y=gap[:show_months], name="GAP",
        marker_color="#3498db",
    ))
    if n_act > 0:
        fig.add_vline(x=n_act - 0.5, line_dash="dash", line_color="gray",
                       annotation_text="Forecast →")
    fig.update_layout(
        barmode="stack", height=350, margin=dict(t=10, b=30),
        legend=dict(orientation="h", y=1.1),
    )
    st.plotly_chart(fig, use_container_width=True)

    # --- Editable Grid ---
    st.subheader("Edit Forecast Volumes")
    st.caption(f"First {n_act} months are actuals (locked). Edit forecast months below.")

    # Quick-fill
    with st.expander("Quick Fill"):
        col1, col2, col3 = st.columns(3)
        with col1:
            target_boba = st.number_input("BOBA target/month", 0, 20, 8)
        with col2:
            target_gap = st.number_input("GAP target/month", 0, 20, 4)
        with col3:
            ramp_months = st.number_input("Ramp months", 0, 24, 6)
        if st.button("Apply Quick Fill"):
            for i in range(n_act, N):
                months_since_start = i - n_act
                if ramp_months > 0 and months_since_start < ramp_months:
                    pct = (months_since_start + 1) / ramp_months
                    bobas[i] = max(bobas[n_act - 1] if n_act > 0 else 1, round(target_boba * pct))
                    gap[i] = max(gap[n_act - 1] if n_act > 0 else 0, round(target_gap * pct))
                else:
                    bobas[i] = target_boba
                    gap[i] = target_gap
            ds.set_surgery_volumes(bobas, gap)
            st.success("Volumes updated")
            st.rerun()

    # Monthly editor (show 12 at a time)
    cols_per_row = 6
    visible = list(range(show_months))

    changed = False
    new_bobas = list(bobas)
    new_gap = list(gap)

    for row_start in range(0, min(show_months, 24), cols_per_row):
        row_indices = visible[row_start:row_start + cols_per_row]
        cols = st.columns(len(row_indices))
        for ci, idx in enumerate(row_indices):
            with cols[ci]:
                is_actual = idx < n_act
                label = FORECAST_MONTH_LABELS[idx]
                b = st.number_input(f"B {label}", value=int(bobas[idx]), min_value=0,
                                     disabled=is_actual, key=f"b_{idx}")
                g = st.number_input(f"G {label}", value=int(gap[idx]), min_value=0,
                                     disabled=is_actual, key=f"g_{idx}")
                if not is_actual and (b != bobas[idx] or g != gap[idx]):
                    changed = True
                    new_bobas[idx] = b
                    new_gap[idx] = g

    if changed and st.button("Save Volumes", type="primary"):
        ds.set_surgery_volumes(new_bobas, new_gap)
        st.success("Volumes saved")
        st.rerun()
