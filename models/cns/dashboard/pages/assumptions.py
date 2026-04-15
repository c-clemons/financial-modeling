"""Assumptions — collection curves, rates, OpEx, team/payroll."""

import streamlit as st
import pandas as pd
import numpy as np

from dashboard.data_store import DataStore
from dashboard.constants import OPEX_LINE_ITEMS, fmt_currency


def show():
    ds = DataStore.get()
    st.header("Model Assumptions")

    a = ds.get_assumptions()

    tab_rev, tab_team, tab_opex = st.tabs(["Revenue & Collections", "Team & Payroll", "OpEx & Overhead"])

    # =====================================================================
    # TAB 1: Revenue & Collections
    # =====================================================================
    with tab_rev:
        st.subheader("Revenue Per Surgery")
        col1, col2 = st.columns(2)
        with col1:
            boba_rev = st.number_input("BOBA Avg Revenue", value=int(a['avg_revenue_bobas']),
                                        step=5000, key="boba_rev")
        with col2:
            gap_rev = st.number_input("GAP Avg Revenue", value=int(a['avg_revenue_gap']),
                                       step=5000, key="gap_rev")

        st.divider()
        st.subheader("Collection Curves")
        st.caption("Percentage of revenue collected at each month lag (must sum to 100%)")

        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**BOBA Collection Curve** (12 lags)")
            boba_curve = list(a.get('bobas_collection_curve', [0]*12))
            for i in range(12):
                boba_curve[i] = st.number_input(
                    f"M+{i}", value=int(boba_curve[i]), min_value=0, max_value=100,
                    key=f"bc_{i}")
            st.caption(f"Total: {sum(boba_curve)}%")
            if sum(boba_curve) != 100:
                st.warning("Must sum to 100%")

        with col2:
            st.markdown("**GAP Collection Curve** (5 lags)")
            gap_curve = list(a.get('gap_collection_curve', [0]*5))
            for i in range(5):
                gap_curve[i] = st.number_input(
                    f"M+{i}", value=int(gap_curve[i]), min_value=0, max_value=100,
                    key=f"gc_{i}")
            st.caption(f"Total: {sum(gap_curve)}%")
            if sum(gap_curve) != 100:
                st.warning("Must sum to 100%")

        st.divider()
        st.subheader("Fund Flow")
        col1, col2, col3 = st.columns(3)
        with col1:
            billing_rate = st.number_input("Billing Fee Rate %", value=float(a['billing_fee_rate']),
                                            step=1.0, key="billing")
        with col2:
            phys_rate = st.number_input("Physician Services %", value=float(a['physician_services_rate']),
                                         step=5.0, key="phys_rate")
        with col3:
            savings_rate = st.number_input("Savings Rate %", value=float(a['savings_rate']),
                                            step=1.0, key="savings")

        if st.button("Save Revenue Assumptions", type="primary", key="save_rev"):
            ds.set_assumptions_bulk({
                'avg_revenue_bobas': boba_rev,
                'avg_revenue_gap': gap_rev,
                'bobas_collection_curve': boba_curve,
                'gap_collection_curve': gap_curve,
                'billing_fee_rate': billing_rate,
                'physician_services_rate': phys_rate,
                'savings_rate': savings_rate,
            })
            st.success("Saved")
            st.rerun()

    # =====================================================================
    # TAB 2: Team & Payroll
    # =====================================================================
    with tab_team:
        st.subheader("Team Roster")

        roster = ds.get_team_roster()
        qbo = ds.actuals_2026
        jan_payroll = qbo['payroll_expenses'][0] if len(qbo['payroll_expenses']) > 0 else 0
        feb_payroll = qbo['payroll_expenses'][1] if len(qbo['payroll_expenses']) > 1 else 0

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Jan 2026 Actual Payroll", fmt_currency(jan_payroll))
        with col2:
            st.metric("Feb 2026 Actual Payroll", fmt_currency(feb_payroll))
        with col3:
            tax_rate = st.number_input("Payroll Tax Rate %", value=float(a['payroll_tax_rate']),
                                        step=0.5, key="tax_rate")

        st.divider()

        for i, person in enumerate(roster):
            with st.expander(f"{'✓' if person.get('start_month') is not None else '○'} {person.get('notes', f'Slot {i+1}')}"):
                col1, col2, col3 = st.columns(3)
                with col1:
                    salary = st.number_input("Monthly Salary", value=int(person['monthly_salary']),
                                              step=500, key=f"sal_{i}")
                with col2:
                    emp_type = st.selectbox("Type", ["W-2", "Contractor"],
                                            index=0 if person['employment_type'] == 'W-2' else 1,
                                            key=f"type_{i}")
                with col3:
                    start = person.get('start_month')
                    start_val = start if start is not None else -1
                    start_input = st.number_input("Start Month (0=Jan-26, -1=Not hired)",
                                                   value=int(start_val), min_value=-1, max_value=59,
                                                   key=f"start_{i}")

                roster[i]['monthly_salary'] = salary
                roster[i]['employment_type'] = emp_type
                roster[i]['start_month'] = start_input if start_input >= 0 else None

        if st.button("Save Team", type="primary", key="save_team"):
            ds.set_team_roster(roster)
            ds.set_assumption('payroll_tax_rate', tax_rate)
            st.success("Team saved")
            st.rerun()

    # =====================================================================
    # TAB 3: OpEx & Overhead
    # =====================================================================
    with tab_opex:
        st.subheader("Monthly Fixed Costs")

        # QBO context
        jan_total = qbo['total_expenses'][0] - qbo['physician_services'][0] - qbo['payroll_expenses'][0] - qbo['contracts'][0]
        feb_total = qbo['total_expenses'][1] - qbo['physician_services'][1] - qbo['payroll_expenses'][1] - qbo['contracts'][1]

        col1, col2 = st.columns(2)
        with col1:
            st.metric("Jan 2026 Actual OpEx (excl payroll/contracts)", fmt_currency(jan_total))
        with col2:
            st.metric("Feb 2026 Actual OpEx (excl payroll/contracts)", fmt_currency(feb_total))

        st.divider()

        updates = {}
        for key, label in OPEX_LINE_ITEMS:
            col1, col2 = st.columns([3, 1])
            with col1:
                val = st.number_input(label, value=float(a.get(key, 0)), step=100.0, key=f"opex_{key}")
                updates[key] = val

        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            inflation = st.number_input("Annual Inflation %", value=float(a['expense_annual_inflation']),
                                         step=0.5, key="inflation")
            updates['expense_annual_inflation'] = inflation
        with col2:
            min_cash = st.number_input("Minimum Cash Balance", value=int(a['minimum_cash_balance']),
                                        step=10000, key="min_cash")
            updates['minimum_cash_balance'] = min_cash

        if st.button("Save OpEx", type="primary", key="save_opex"):
            ds.set_assumptions_bulk(updates)
            st.success("OpEx saved")
            st.rerun()
