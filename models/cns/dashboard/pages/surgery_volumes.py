"""Surgery Volumes — editable BOBA/GAP forecast with historical context.

Auto-saves on every edit (no Save button). A toggle unlocks historical (2024,
2025) and actual (Jan-Feb 2026) months for cases where the client is still
reconciling their data.
"""

from datetime import datetime

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go

from dashboard.data_store import DataStore
from dashboard.constants import (
    N, FORECAST_MONTH_LABELS, SURGERY_TYPES, fmt_currency,
)
from baseline_data import MONTHS_12

LAST_SAVED_KEY = "vol_last_saved"
UNLOCK_KEY = "vol_unlock_hist"


def _stamp():
    st.session_state[LAST_SAVED_KEY] = datetime.now().strftime("%I:%M:%S %p")


def _last_saved_caption() -> str:
    val = st.session_state.get(LAST_SAVED_KEY)
    return f"💾 Last saved: {val}" if val else "💾 Auto-save active"


def show():
    ds = DataStore.get()
    st.header("Surgery Volume Forecast")

    head_l, head_r = st.columns([3, 2])
    with head_l:
        st.toggle(
            "🔓 Unlock historical & actual months",
            help=("When enabled, you can edit 2024/2025 historical volumes and "
                  "Jan-Feb 2026 actuals. Use this while the client is still "
                  "reconciling data."),
            key=UNLOCK_KEY,
        )
    with head_r:
        st.caption(_last_saved_caption())

    unlock = st.session_state.get(UNLOCK_KEY, False)

    locations = ds.get_locations()
    n_act = ds.n_actuals_2026

    selected_loc = st.selectbox("Location", ["All Locations"] + locations, key="vol_loc")
    volumes_by_loc = ds.get_volumes_by_location()

    if selected_loc == "All Locations":
        bobas = [sum(volumes_by_loc[loc]['bobas'][i] for loc in locations) for i in range(N)]
        gap = [sum(volumes_by_loc[loc]['gap'][i] for loc in locations) for i in range(N)]
    else:
        bobas = list(volumes_by_loc.get(selected_loc, {}).get('bobas', [0] * N))
        gap = list(volumes_by_loc.get(selected_loc, {}).get('gap', [0] * N))

    # ------------------------------------------------------------------
    # Historical volumes (2024 Sep-Dec, 2025 Jan-Dec)
    # ------------------------------------------------------------------
    st.subheader("Historical Volumes")
    boba_24 = ds.get_historical_volumes('boba_2024')
    gap_24 = ds.get_historical_volumes('gap_2024')
    boba_25 = ds.get_historical_volumes('boba_2025')
    gap_25 = ds.get_historical_volumes('gap_2025')

    col1, col2 = st.columns(2)

    with col1:
        st.caption("2024 (Sep-Dec)")
        df_24 = pd.DataFrame({
            "Month": ["Sep", "Oct", "Nov", "Dec"],
            "BOBA": boba_24, "GAP": gap_24,
            "Total": [b + g for b, g in zip(boba_24, gap_24)],
        })
        if unlock:
            edited = st.data_editor(
                df_24, hide_index=True, key="hist24_editor",
                column_config={
                    "Month": st.column_config.TextColumn(disabled=True),
                    "BOBA": st.column_config.NumberColumn(min_value=0, step=1),
                    "GAP": st.column_config.NumberColumn(min_value=0, step=1),
                    "Total": st.column_config.NumberColumn(disabled=True),
                },
                use_container_width=True,
            )
            new_b = [int(x or 0) for x in edited["BOBA"].tolist()]
            new_g = [int(x or 0) for x in edited["GAP"].tolist()]
            if new_b != boba_24 or new_g != gap_24:
                ds.set_historical_volumes(boba_2024=new_b, gap_2024=new_g)
                _stamp()
                st.rerun()
        else:
            st.dataframe(df_24.set_index("Month"), use_container_width=True)

    with col2:
        st.caption("2025 (Full Year)")
        df_25 = pd.DataFrame({
            "Month": MONTHS_12,
            "BOBA": boba_25, "GAP": gap_25,
            "Total": [b + g for b, g in zip(boba_25, gap_25)],
        })
        if unlock:
            edited = st.data_editor(
                df_25, hide_index=True, key="hist25_editor",
                column_config={
                    "Month": st.column_config.TextColumn(disabled=True),
                    "BOBA": st.column_config.NumberColumn(min_value=0, step=1),
                    "GAP": st.column_config.NumberColumn(min_value=0, step=1),
                    "Total": st.column_config.NumberColumn(disabled=True),
                },
                use_container_width=True,
            )
            new_b = [int(x or 0) for x in edited["BOBA"].tolist()]
            new_g = [int(x or 0) for x in edited["GAP"].tolist()]
            if new_b != boba_25 or new_g != gap_25:
                ds.set_historical_volumes(boba_2025=new_b, gap_2025=new_g)
                _stamp()
                st.rerun()
        else:
            st.dataframe(df_25.set_index("Month"), use_container_width=True)

    avg_boba_25 = float(np.mean(boba_25)) if boba_25 else 0
    avg_gap_25 = float(np.mean(gap_25)) if gap_25 else 0
    st.caption(
        f"2025 averages: BOBA {avg_boba_25:.1f}/mo, GAP {avg_gap_25:.1f}/mo, "
        f"Total {avg_boba_25 + avg_gap_25:.1f}/mo"
    )
    if unlock:
        st.info(
            "Note: editing historical volumes updates the displayed history but "
            "does not retroactively recompute AR spillover into the forecast. "
            "Restart the app after large historical edits if AR carry-over "
            "needs to refresh.",
            icon="ℹ️",
        )

    st.divider()

    # ------------------------------------------------------------------
    # 2026-2030 Forecast chart (preview)
    # ------------------------------------------------------------------
    st.subheader("2026-2030 Forecast")
    show_months = st.slider("Months to display", 12, 60, 36, key="vol_months")
    labels = FORECAST_MONTH_LABELS[:show_months]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=labels, y=bobas[:show_months], name="BOBA", marker_color="#2c3e50",
    ))
    fig.add_trace(go.Bar(
        x=labels, y=gap[:show_months], name="GAP", marker_color="#3498db",
    ))
    if n_act > 0:
        fig.add_vline(x=n_act - 0.5, line_dash="dash", line_color="gray",
                      annotation_text="Forecast →")
    fig.update_layout(
        barmode="stack", height=320, margin=dict(t=10, b=30),
        legend=dict(orientation="h", y=1.1),
    )
    st.plotly_chart(fig, use_container_width=True)

    # ------------------------------------------------------------------
    # Quick fill
    # ------------------------------------------------------------------
    with st.expander("Quick Fill"):
        col1, col2, col3 = st.columns(3)
        with col1:
            target_boba = st.number_input("BOBA target/month", 0, 50, 8)
        with col2:
            target_gap = st.number_input("GAP target/month", 0, 50, 4)
        with col3:
            ramp_months = st.number_input("Ramp months", 0, 24, 6)
        if st.button("Apply Quick Fill"):
            ramp_start = n_act
            if selected_loc == "Santa Barbara":
                for exp in ds.get_expansions():
                    if exp.get('name') == 'Santa Barbara' and exp.get('enabled'):
                        ramp_start = max(n_act, exp.get('lease_start_month', 6))
            new_b = list(bobas)
            new_g = list(gap)
            for i in range(ramp_start, N):
                ms = i - ramp_start
                if ramp_months > 0 and ms < ramp_months:
                    pct = (ms + 1) / ramp_months
                    new_b[i] = max(0, round(target_boba * pct))
                    new_g[i] = max(0, round(target_gap * pct))
                else:
                    new_b[i] = target_boba
                    new_g[i] = target_gap
            _save_volumes(ds, selected_loc, locations, new_b, new_g)
            _stamp()
            st.success("Quick fill applied")
            st.rerun()

    # ------------------------------------------------------------------
    # Editable monthly grid (auto-saves on change)
    # ------------------------------------------------------------------
    st.subheader("Edit Monthly Volumes")
    if unlock:
        st.caption("All months editable — including locked actuals. Auto-saves on edit.")
        editor_indices = list(range(N))
    else:
        st.caption(
            f"First {n_act} months locked as actuals. "
            f"Toggle '🔓 Unlock historical & actual months' above to edit them. "
            f"All other edits auto-save."
        )
        if n_act > 0:
            actual_df = pd.DataFrame({
                "Month": FORECAST_MONTH_LABELS[:n_act],
                "BOBA": bobas[:n_act],
                "GAP": gap[:n_act],
                "Total": [bobas[i] + gap[i] for i in range(n_act)],
            }).set_index("Month")
            st.markdown("**Actuals (locked)**")
            st.dataframe(actual_df, use_container_width=True)
        editor_indices = list(range(n_act, N))

    editor_df = pd.DataFrame({
        "Month": [FORECAST_MONTH_LABELS[i] for i in editor_indices],
        "BOBA": [bobas[i] for i in editor_indices],
        "GAP": [gap[i] for i in editor_indices],
        "Total": [bobas[i] + gap[i] for i in editor_indices],
    })

    edited = st.data_editor(
        editor_df,
        hide_index=True,
        key=f"vol_editor_{selected_loc}_{int(unlock)}",
        column_config={
            "Month": st.column_config.TextColumn("Month", disabled=True),
            "BOBA": st.column_config.NumberColumn("BOBA", min_value=0, step=1),
            "GAP": st.column_config.NumberColumn("GAP", min_value=0, step=1),
            "Total": st.column_config.NumberColumn("Total", disabled=True),
        },
        use_container_width=True,
        height=min(38 * (len(editor_indices) + 1) + 10, 600),
    )

    new_bobas = list(bobas)
    new_gap = list(gap)
    for i, fcst_idx in enumerate(editor_indices):
        new_bobas[fcst_idx] = int(edited["BOBA"].iloc[i] or 0)
        new_gap[fcst_idx] = int(edited["GAP"].iloc[i] or 0)

    if new_bobas != list(bobas) or new_gap != list(gap):
        _save_volumes(ds, selected_loc, locations, new_bobas, new_gap)
        _stamp()
        st.rerun()


def _save_volumes(ds: DataStore, selected_loc: str, locations: list, bobas: list, gap: list):
    """Persist edits, handling consolidated vs single-location selection."""
    if selected_loc == "All Locations":
        # Distribute proportionally across enabled locations.
        # Simpler: write to consolidated keys directly (drops per-location detail).
        # To preserve location data, we rewrite Westlake to absorb the diff.
        vbl = ds.get_volumes_by_location()
        cur_b = [sum(vbl[loc]['bobas'][i] for loc in locations) for i in range(N)]
        cur_g = [sum(vbl[loc]['gap'][i] for loc in locations) for i in range(N)]
        delta_b = [bobas[i] - cur_b[i] for i in range(N)]
        delta_g = [gap[i] - cur_g[i] for i in range(N)]
        vbl['Westlake'] = {
            'bobas': [max(0, vbl['Westlake']['bobas'][i] + delta_b[i]) for i in range(N)],
            'gap': [max(0, vbl['Westlake']['gap'][i] + delta_g[i]) for i in range(N)],
        }
        ds.set_volumes_by_location(vbl)
    else:
        vbl = ds.get_volumes_by_location()
        vbl[selected_loc] = {'bobas': list(bobas), 'gap': list(gap)}
        ds.set_volumes_by_location(vbl)
