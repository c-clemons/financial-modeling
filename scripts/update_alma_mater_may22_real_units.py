"""
May 22 — Replace estimated unit splits with REAL Shopify TSV data
====================================================================

Phase 2 of the units update — user downloaded the actual TSV from
the Streamlit Shopify Analytics page and provided the verified
DTC/WS/Gifting Units breakdown.

CHANGES vs estimates (Jan-Apr 2026):
  DTC Units:       est [27, 38, 96, 130]  → actual [34, 42, 145, 150]
  Wholesale Units: est [0, 60, 221, 80]   → actual [0, 64, 178, 33]
  Gifting Units:   est [7, 8, 52, 57]     → actual [0, 0, 46, 84]

Estimates were systematically off because:
  - DTC orders include some multi-pair purchases (1.1-1.7 units/order)
  - Mar Wholesale was inflated in my estimate (assumed 13 units/order;
    actual was 178/17 = 10.5)
  - Apr Wholesale much lower than I assumed (33 actual vs 80 est)
  - Gifting had ~0 in early months (samples started Mar) — I had
    assumed 1.5-2 units/order from the start

ALSO extended Orders + Units sections through May (5 months of data
shown but Last Actuals Month still = 4 so May is NOT used by the model).
May cells fill = light yellow (FFFFF2CC) to indicate "in progress".

Key observation:
  DTC Units consistently 1.1-1.7× DTC Orders. Current model R66
  pulls QBO Actuals!Cnn84 (Orders) — for inventory consumption
  modeling, should consider switching to Cnn90 (Units) for accuracy.
  Flagged separately to user.
"""
# Inline change via REPL — preserved here for audit trail.
