#!/usr/bin/env python3
"""
CNS (California Neurosurgical Specialists) Financial Dashboard

Usage:
    cd /Users/chandlerclemons/financial-modeling/models/cns
    streamlit run dashboard/app.py
"""

import sys
from pathlib import Path

# Add parent directories for imports
CNS_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(CNS_ROOT))
sys.path.insert(0, str(CNS_ROOT / "dashboard"))

import streamlit as st

st.set_page_config(
    page_title="CNS Financial Model",
    layout="wide",
    initial_sidebar_state="expanded",
)

from dashboard.data_store import DataStore

PAGES = {
    "Cash Flow Forecast": "dashboard.pages.cash_flow",
    "Monthly P&L": "dashboard.pages.monthly_pl",
    "Surgery Volumes": "dashboard.pages.surgery_volumes",
    "Assumptions": "dashboard.pages.assumptions",
    "Expansion Planner": "dashboard.pages.expansions",
    "Scenarios": "dashboard.pages.scenarios",
}


def check_password() -> bool:
    if st.session_state.get("authenticated"):
        return True
    try:
        correct = st.secrets["app_password"]
    except (KeyError, FileNotFoundError):
        correct = "cns2026"

    st.title("California Neurosurgical Specialists")
    st.caption("Financial Forecasting Model")
    password = st.text_input("Password", type="password")
    if st.button("Login", type="primary"):
        if password == correct:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Incorrect password")
    return False


def main():
    if not check_password():
        return

    st.markdown("""
    <style>
    div[data-testid="stMetric"] {
        background-color: #f8f9fa;
        padding: 12px 16px;
        border-radius: 8px;
        border: 1px solid #e9ecef;
    }
    </style>
    """, unsafe_allow_html=True)

    if "initialized" not in st.session_state:
        ds = DataStore.get()
        ds.load()
        st.session_state.initialized = True

    ds = DataStore.get()

    # Sidebar
    st.sidebar.title("CNS Financial Model")
    st.sidebar.caption("California Neurosurgical Specialists")
    st.sidebar.divider()

    page = st.sidebar.radio("Navigate", list(PAGES.keys()), label_visibility="collapsed")

    st.sidebar.divider()
    st.sidebar.markdown(f"**Actuals:** Jan-Feb 2026")
    st.sidebar.markdown(f"**Forecast:** Mar 2026 - Dec 2030")

    a = ds.get_assumptions()
    st.sidebar.divider()
    st.sidebar.markdown(f"**Starting Cash:** ${a.get('starting_cash', 0):,.0f}")
    st.sidebar.markdown(f"**Min Cash Target:** ${a.get('minimum_cash_balance', 0):,.0f}")

    # Route
    import importlib
    module = importlib.import_module(PAGES[page])
    module.show()


if __name__ == "__main__":
    main()
