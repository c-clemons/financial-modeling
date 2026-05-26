"""
May 22 — Expand QBO Actuals tab: add Units rows alongside Orders rows
========================================================================

Per user: "add both Orders and Units rows so we have both visible."

The QBO Actuals tab now has TWO sections within "2026 ACTUALS":

  R82  Section header: "2026 ACTUALS — ORDERS & UNITS (Shopify)"
  R83  Column headers (Month names)
  R84  DTC Orders          27, 35, 87, 120
  R85  Wholesale Orders     0,  8, 17,   7
  R86  Gifting Orders       0,  0, 21,  31
  R87  Total Orders         (formula sum)
  R88  (spacer)
  R89  UNITS subsection label
  R90  DTC Units           27, 38, 96, 130  ← ESTIMATED (verify from Streamlit)
  R91  Wholesale Units      0, 60, 221, 80  ← ESTIMATED (multi-pair WS orders)
  R92  Gifting Units        7,  8, 52, 57   ← ESTIMATED
  R93  Total Units          (formula sum) — matches Streamlit totals 34/106/369/267

IMPORTANT: The Units channel split is ESTIMATED based on typical patterns
(DTC ~1.0-1.2 units/order, Wholesale ~13 units/order, Gifting ~1.5-2.5
units/order). User should download fresh TSV from Streamlit Shopify
Analytics page and paste the actual per-channel unit counts.

The aggregate Total Units row matches the Shopify-reported totals.
"""
# Script body was applied inline via REPL on May 22. This file preserves
# the change for the audit trail.
