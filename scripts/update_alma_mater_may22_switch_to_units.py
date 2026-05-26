"""
May 22 — Switch R66 (DTC) and R102 (WS) to pull Units instead of Orders
==========================================================================

Per user: more accurate inventory consumption modeling. DTC Units run
1.1-1.7× DTC Orders due to multi-pair purchases. Inventory consumption
is per UNIT, not per order.

Assumptions R66 (Beta DTC consumption 2026) cols C-F:
  Was:  =IF(month<=$C$47, 'QBO Actuals'!Cnn84, Matt forecast)  [Orders R84]
  Now:  =IF(month<=$C$47, 'QBO Actuals'!Cnn90, Matt forecast)  [Units R90]

Effective Jan-Apr values: 34/42/145/150 (was 27/35/87/120)

Assumptions R102 (Other Wholesale 2026) cols D-G:
  Was:  =IF(...QBO!cnn85, 0)  [WS Orders R85]
  Now:  =IF(...QBO!cnn91, 0)  [WS Units R91]

Effective Jan-Apr values: 0/64/178/33 (was 0/8/17/7)

Net impact: +345 units of inventory consumption in Q1 2026 vs prior model.
2026 ending inventory will be lower after Excel recalc.

May-Dec forecasts UNCHANGED — still uses Matt's original order-count
forecast (which is now technically understated by ~30% since it's
orders, not units). Separate question to address whether to scale
Matt's forecast.

Also pending: Gifting Units (R92) — currently 0/0/46/84 for Jan-Apr,
not yet plumbed into inventory consumption anywhere. ~130 pairs of
gifts went out Q1 2026 that the model doesn't account for. Flagged
separately.
"""
# Inline change via REPL on May 22.
