"""CEO Dashboard — at-a-glance landing page.

KPI tiles, monthly surgery volume by surgery type, and a cash trend chart
(same series as the Cash Flow Forecast page).
"""

from datetime import datetime

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from dashboard.data_store import DataStore
from dashboard.constants import N, FORECAST_MONTH_LABELS, fmt_currency


def show():
    ds = DataStore.get()
    st.header("CEO Dashboard")

    a = ds.get_assumptions()
    forecast = ds.run_forecast()
    cf = forecast['cf']
    pl = forecast['pl']
    n_act = ds.n_actuals_2026
    locations = ds.get_locations()
    volumes_by_loc = ds.get_volumes_by_location()

    # ------------------------------------------------------------------
    # KPI tiles
    # ------------------------------------------------------------------
    ending_cash = cf['ending_cash']
    current_cash = ending_cash[n_act - 1] if n_act > 0 else a.get('starting_cash', 0)
    yr1_collections = sum(pl.get('total_income', pl.get('total_collected', [0] * 60))[:12])
    yr1_physician = sum(cf['physician'][:12])
    yr1_surgeries = sum(a['bobas_volume'][:12]) + sum(a['gap_volume'][:12])
    min_cash = a.get('minimum_cash_balance', 150000)
    runway_pct = (current_cash / min_cash * 100) if min_cash > 0 else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("Current Cash",
                  fmt_currency(current_cash),
                  delta=f"{runway_pct:.0f}% of min target",
                  delta_color="normal" if current_cash >= min_cash else "inverse")
    with c2:
        st.metric("Year-1 Collections", fmt_currency(yr1_collections))
    with c3:
        st.metric("Year-1 Physician Dist.", fmt_currency(yr1_physician))
    with c4:
        st.metric("Year-1 Surgeries", f"{int(yr1_surgeries):,}")

    st.divider()

    # ------------------------------------------------------------------
    # Surgery volume — monthly stacked bar by surgery type
    # ------------------------------------------------------------------
    st.subheader("Monthly Surgery Volume — by Type")

    horizon = st.slider("Months to display", 12, 60, 24, key="ceo_vol_months")
    labels = FORECAST_MONTH_LABELS[:horizon]

    boba_total = a['bobas_volume'][:horizon]
    gap_total = a['gap_volume'][:horizon]

    # Optionally split BOBA/GAP by location (if multiple locations active)
    enabled_locs = [
        loc for loc in locations
        if any(volumes_by_loc.get(loc, {}).get('bobas', [0])[i]
                + volumes_by_loc.get(loc, {}).get('gap', [0])[i] > 0
                for i in range(horizon))
    ]
    show_per_location = (
        len(enabled_locs) > 1
        and st.checkbox("Split by location", value=False, key="ceo_vol_split")
    )

    fig = go.Figure()
    if show_per_location:
        boba_palette = ["#1B2A4A", "#2c3e50", "#34495e", "#5d6d7e", "#85929e"]
        gap_palette = ["#3498db", "#5dade2", "#85c1e9", "#aed6f1", "#d4e6f1"]
        for i, loc in enumerate(enabled_locs):
            fig.add_trace(go.Bar(
                x=labels,
                y=volumes_by_loc[loc]['bobas'][:horizon],
                name=f"BOBA — {loc}",
                marker_color=boba_palette[i % len(boba_palette)],
            ))
        for i, loc in enumerate(enabled_locs):
            fig.add_trace(go.Bar(
                x=labels,
                y=volumes_by_loc[loc]['gap'][:horizon],
                name=f"GAP — {loc}",
                marker_color=gap_palette[i % len(gap_palette)],
            ))
    else:
        fig.add_trace(go.Bar(x=labels, y=boba_total, name="BOBA",
                             marker_color="#1B2A4A"))
        fig.add_trace(go.Bar(x=labels, y=gap_total, name="GAP",
                             marker_color="#3498db"))

    if n_act > 0:
        fig.add_vline(x=n_act - 0.5, line_dash="dash", line_color="gray",
                      annotation_text="Forecast →")

    fig.update_layout(
        barmode="stack", height=380, margin=dict(t=20, b=40),
        legend=dict(orientation="h", y=1.08),
        hovermode="x unified",
    )
    st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # ------------------------------------------------------------------
    # Cash flow trend — same series as Cash Flow Forecast page
    # ------------------------------------------------------------------
    st.subheader("Cash Flow Trend")
    st.caption(
        "Cash collected vs. overhead, with cumulative cash available. "
        "Mirrors the Cash Flow Forecast page."
    )

    chart_collected = pl.get('total_income', pl.get('total_collected', [0] * 60))
    chart_overhead = pl.get('total_overhead', [0] * 60)
    cash_available = [chart_collected[i] - chart_overhead[i] for i in range(horizon)]

    fig2 = go.Figure()
    fig2.add_trace(go.Bar(
        x=labels, y=chart_collected[:horizon],
        name="Cash Collected", marker_color="#2ecc71", opacity=0.85,
    ))
    fig2.add_trace(go.Bar(
        x=labels, y=[-v for v in chart_overhead[:horizon]],
        name="Overhead (Cash Out)", marker_color="#e74c3c", opacity=0.85,
    ))
    fig2.add_trace(go.Scatter(
        x=labels, y=cash_available,
        name="Cash Available (collected − overhead)",
        line=dict(color="#1B2A4A", width=3),
        mode="lines+markers", marker=dict(size=5),
        fill="tozeroy", fillcolor="rgba(27,42,74,0.08)",
    ))

    if n_act > 0:
        fig2.add_vline(x=n_act - 0.5, line_dash="dash", line_color="gray",
                       annotation_text="Forecast →")

    fig2.update_layout(
        barmode="relative", height=420, margin=dict(t=20, b=40),
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        hovermode="x unified",
    )
    fig2.update_yaxes(title_text="Monthly Cash Flow", tickformat="$,.0f")
    st.plotly_chart(fig2, use_container_width=True)

    # Ending-cash trajectory
    fig3 = go.Figure()
    fig3.add_trace(go.Scatter(
        x=FORECAST_MONTH_LABELS[:horizon],
        y=ending_cash[:horizon],
        name="Ending Cash",
        line=dict(color="#1B2A4A", width=3),
        mode="lines+markers",
        fill="tozeroy", fillcolor="rgba(27,42,74,0.10)",
    ))
    fig3.add_hline(y=min_cash, line_dash="dot", line_color="#e74c3c",
                   annotation_text=f"Minimum target {fmt_currency(min_cash)}")
    if n_act > 0:
        fig3.add_vline(x=n_act - 0.5, line_dash="dash", line_color="gray",
                       annotation_text="Forecast →")
    fig3.update_layout(
        height=320, margin=dict(t=20, b=40),
        showlegend=False, hovermode="x unified",
    )
    fig3.update_yaxes(title_text="Cash Balance", tickformat="$,.0f")
    st.subheader("Ending Cash Balance")
    st.plotly_chart(fig3, use_container_width=True)

    # ------------------------------------------------------------------
    # Footer: data freshness
    # ------------------------------------------------------------------
    st.divider()
    last_updated = ds.overrides.get("_last_updated", "—")
    if last_updated and last_updated != "—":
        try:
            dt = datetime.fromisoformat(last_updated)
            last_updated = dt.strftime("%b %d, %Y at %I:%M %p")
        except (ValueError, TypeError):
            pass
    upload_2025 = ds.get_uploaded_actuals_meta(2025).get("uploaded_at", "—")
    upload_2026 = ds.get_uploaded_actuals_meta(2026).get("uploaded_at", "—")

    cols = st.columns(3)
    cols[0].caption(f"**Model last updated:** {last_updated}")
    cols[1].caption(f"**2025 actuals upload:** {upload_2025}")
    cols[2].caption(f"**2026 actuals upload:** {upload_2026}")
