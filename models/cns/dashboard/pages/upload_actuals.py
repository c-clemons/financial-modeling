"""Upload Actuals — drag-and-drop a QBO P&L Excel and commit as actuals.

Handles flexible chart of accounts: any rows whose code/name doesn't match
QBO_ACCOUNTS are flagged for manual mapping before commit.
"""

from __future__ import annotations

import io
from datetime import datetime
from typing import Optional

import pandas as pd
import streamlit as st

from dashboard.data_store import DataStore
from dashboard.qbo_parser import (
    parse_pl_workbook, to_upload_payload, normalize_key,
)
from baseline_data import QBO_ACCOUNTS

PARSED_KEY = "upload_parsed"
MAPPING_KEY = "upload_mapping"  # {label: target_key | "<skip>" | "<new>"}


def _known_keys() -> list[str]:
    seen, out = set(), []
    for meta in QBO_ACCOUNTS.values():
        k = meta["key"]
        if k not in seen:
            seen.add(k)
            out.append(k)
    return sorted(out)


def show():
    ds = DataStore.get()
    st.header("Upload P&L Actuals")
    st.caption(
        "Drag a QBO Profit & Loss Excel export here. The file is parsed, "
        "any new account codes are flagged for review, and once you commit, "
        "the actuals replace the hard-coded baseline for that year — flowing "
        "through the Monthly P&L and Cash Flow Forecast pages."
    )

    # ------------------------------------------------------------------
    # Currently committed uploads
    # ------------------------------------------------------------------
    with st.expander("Currently committed uploads", expanded=False):
        rows = []
        for yr in ("2025", "2026"):
            meta = ds.get_uploaded_actuals_meta(int(yr))
            if meta:
                rows.append({
                    "Year": yr,
                    "Months": ", ".join(meta.get("months", [])),
                    "Source File": meta.get("source_filename", ""),
                    "Uploaded": meta.get("uploaded_at", ""),
                })
        if rows:
            st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
            cols = st.columns(len(rows))
            for i, r in enumerate(rows):
                with cols[i]:
                    if st.button(f"Revert {r['Year']}", key=f"revert_{r['Year']}"):
                        ds.clear_uploaded_actuals(int(r['Year']))
                        st.success(f"Reverted {r['Year']} to baseline")
                        st.rerun()
        else:
            st.caption("No uploads yet — using baseline actuals from `baseline_data.py`.")

    st.divider()

    # ------------------------------------------------------------------
    # File uploader
    # ------------------------------------------------------------------
    uploaded = st.file_uploader(
        "Drop a QBO P&L .xlsx file here",
        type=["xlsx"],
        accept_multiple_files=False,
        help="Standard QuickBooks 'Profit and Loss' export. Cash or Accrual basis.",
    )

    if uploaded is None:
        if PARSED_KEY in st.session_state:
            del st.session_state[PARSED_KEY]
        if MAPPING_KEY in st.session_state:
            del st.session_state[MAPPING_KEY]
        st.info("Upload a file to begin.")
        return

    # Parse on first sight or when the filename changes
    file_bytes = uploaded.getvalue()
    fingerprint = (uploaded.name, len(file_bytes))
    if (PARSED_KEY not in st.session_state
            or st.session_state.get("upload_fingerprint") != fingerprint):
        try:
            parsed = parse_pl_workbook(io.BytesIO(file_bytes),
                                        extras=ds.get_account_mapping_extras())
        except Exception as exc:
            st.error(f"Could not parse this file: {exc}")
            return
        st.session_state[PARSED_KEY] = parsed
        st.session_state[MAPPING_KEY] = {
            u["label"]: u["suggested_key"] for u in parsed["unmapped"]
        }
        st.session_state["upload_fingerprint"] = fingerprint

    parsed = st.session_state[PARSED_KEY]
    mapping = st.session_state[MAPPING_KEY]

    # ------------------------------------------------------------------
    # Summary card
    # ------------------------------------------------------------------
    n = len(parsed["months"])
    st.success(
        f"Parsed **{uploaded.name}** — {n} month(s): "
        f"{parsed['months'][0]} → {parsed['months'][-1]}. "
        f"Year: **{parsed['year']}**. "
        f"{parsed['meta']['parsed_rows']} mapped rows, "
        f"{len(parsed['unmapped'])} unmapped."
    )

    # ------------------------------------------------------------------
    # Reconciliation check: sum of mapped + unmapped lines vs. total_expenses
    # ------------------------------------------------------------------
    sum_lines = [0.0] * n
    for k, vals in parsed["data"].items():
        # Skip totals and revenue keys for expense reconciliation
        if k in ("fee_income", "reimbursed_expense_income", "refunds",
                 "interest_income"):
            continue
        for i, v in enumerate(vals):
            sum_lines[i] += v
    for u in parsed["unmapped"]:
        target = mapping.get(u["label"], "<skip>")
        if target == "<skip>":
            continue
        for i, v in enumerate(u["values"]):
            sum_lines[i] += v

    file_total_exp = parsed["totals"].get("total_expenses", [0.0] * n)
    diffs = [sum_lines[i] - file_total_exp[i] for i in range(n)]

    rec_df = pd.DataFrame({
        "Month": parsed["months"],
        "Sum of line items": sum_lines,
        "File 'Total for Expenses'": file_total_exp,
        "Diff": diffs,
    })

    st.subheader("Reconciliation")
    if all(abs(d) < 1.0 for d in diffs):
        st.success("Line items reconcile to the file's Total for Expenses (within $1).")
    else:
        st.warning(
            "Mismatch detected — review unmapped accounts below before committing. "
            "Most often this is fixed by mapping a new account to an existing key."
        )
    st.dataframe(
        rec_df.style.format({c: "${:,.2f}" for c in rec_df.columns if c != "Month"}),
        hide_index=True, use_container_width=True,
    )

    # ------------------------------------------------------------------
    # Unmapped accounts: let user map each
    # ------------------------------------------------------------------
    if parsed["unmapped"]:
        st.subheader("Unmapped accounts")
        st.caption(
            "These rows didn't match a known account code. Choose how each "
            "should be treated: map to an existing key (recommended for "
            "renamed accounts), keep as a new key, or skip."
        )
        existing = _known_keys()
        for u in parsed["unmapped"]:
            cols = st.columns([3, 2, 3])
            with cols[0]:
                vals_str = ", ".join(f"${v:,.2f}" for v in u["values"])
                st.markdown(f"**{u['label']}**  \n<small>{vals_str}</small>",
                             unsafe_allow_html=True)
            with cols[1]:
                action = st.selectbox(
                    "Action",
                    ["Use suggested new key", "Map to existing", "Skip"],
                    key=f"act_{u['label']}",
                    label_visibility="collapsed",
                )
            with cols[2]:
                if action == "Map to existing":
                    target = st.selectbox(
                        "Target key", existing,
                        key=f"map_{u['label']}",
                        label_visibility="collapsed",
                    )
                    mapping[u["label"]] = target
                elif action == "Skip":
                    mapping[u["label"]] = "<skip>"
                else:
                    new_key = st.text_input(
                        "New key", value=u["suggested_key"],
                        key=f"new_{u['label']}",
                        label_visibility="collapsed",
                    )
                    mapping[u["label"]] = normalize_key(new_key) if new_key else u["suggested_key"]
        st.session_state[MAPPING_KEY] = mapping
    else:
        st.caption("All accounts mapped to known keys ✓")

    # ------------------------------------------------------------------
    # Final preview of what will be saved
    # ------------------------------------------------------------------
    final_data = dict(parsed["data"])
    for u in parsed["unmapped"]:
        target = mapping.get(u["label"], "<skip>")
        if target == "<skip>":
            continue
        cur = final_data.get(target, [0.0] * n)
        final_data[target] = [cur[i] + u["values"][i] for i in range(n)]

    st.subheader("Preview")
    rows = []
    for k in sorted(final_data.keys()):
        rows.append({"Account key": k, **{
            parsed["months"][i]: final_data[k][i] for i in range(n)
        }, "Total": sum(final_data[k])})
    pdf = pd.DataFrame(rows)
    fmt = {c: "${:,.0f}" for c in pdf.columns if c != "Account key"}
    st.dataframe(pdf.style.format(fmt), hide_index=True, use_container_width=True, height=400)

    # ------------------------------------------------------------------
    # Commit button
    # ------------------------------------------------------------------
    st.divider()
    can_commit = all(abs(d) < 1.0 for d in diffs)
    commit_label = (
        f"Commit as {parsed['year']} actuals"
        if can_commit
        else f"Commit as {parsed['year']} actuals (mismatch — review first)"
    )
    col1, col2 = st.columns([1, 3])
    with col1:
        commit = st.button(commit_label, type="primary", use_container_width=True)
    with col2:
        save_mappings = st.checkbox(
            "Remember new mappings for future uploads",
            value=True,
            help="Store the chosen mapping for each unrecognized account label "
                 "so the next time the same label is uploaded, it's matched automatically.",
        )

    if commit:
        # Build payload using post-mapping data
        adjusted_parsed = {**parsed, "data": final_data}
        payload = to_upload_payload(adjusted_parsed, source_filename=uploaded.name)
        ds.set_uploaded_actuals(parsed["year"], payload)

        if save_mappings and parsed["unmapped"]:
            for u in parsed["unmapped"]:
                target = mapping.get(u["label"])
                if target and target != "<skip>":
                    ds.add_account_mapping(u["label"], target)

        st.success(
            f"Committed {n} months of {parsed['year']} actuals from "
            f"**{uploaded.name}**. The Cash Flow Forecast and Monthly P&L "
            "pages now reflect this data."
        )
        for k in (PARSED_KEY, MAPPING_KEY, "upload_fingerprint"):
            st.session_state.pop(k, None)
        st.rerun()
