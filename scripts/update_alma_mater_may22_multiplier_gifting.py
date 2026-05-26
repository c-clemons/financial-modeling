"""
May 22 — Add DTC Units/Order multiplier + include Gifting in DTC consumption
==============================================================================

Per user direction:
  1. Option A on Gifting: treat as DTC consumption (add Gifting Units
     to DTC consumption in R66 for closed months)
  2. Apply multiplier to Matt's order-based forecasts to estimate
     unit consumption

Changes:

R64: New cell "DTC Units/Order Multiplier" = 1.38
  Based on Q1 2026 actuals: 371 units / 269 orders = 1.379
  Yellow editable. User can tune if multiplier shifts over time.

R66 (Beta DTC Units 2026) — rewritten:
  Closed months (Jan-Apr): =IF(month<=$C$47, QBO!Cnn90 + QBO!Cnn92, Matt*$C$64)
    First arg: DTC Units (R90) + Gifting Units (R92) from QBO Actuals
    Fallback: Matt's order forecast × multiplier
  Future months (May-Dec): =Matt_value × $C$64
    e.g., G66 = =194.7*$C$64 → 269 units (was 195 orders)

R77 (Beta DTC Units 2027) — all 12 cols = Matt × $C$64
R88 (Beta DTC Units 2028) — all 12 cols = Matt × $C$64

Impact summary (vs prior state):
  2026 Q1 consumption:  269 orders → 501 units (+232)
                       (Mar/Apr gifting added 130 pairs; multi-pair purchases added 102)
  2026 May-Dec forecast: 1,977 orders → 2,729 units (+752)
  2026 total:           ~2,246 → ~3,230 units (+44%)
  2027 total:           3,610 → 4,982 units (+1,372)
  2028 total:           5,507 → 7,600 units (+2,093)

Inventory adequacy check (rough):
  2026 Beta supply: 2,500 beg + 5,500 POs = 8,000 vs ~4,730 consumption (DTC+WS) → 3,270 ending ✓
  2027 Beta supply: 3,270 + 15,000 POs = 18,270 vs ~8,982 consumption → 9,288 ending ✓
  2028 Beta supply: 9,288 + 15,000 POs = 24,288 vs ~15,600 consumption → 8,688 ending ✓
  No PO scaling needed — current orders cover the higher consumption.

Note: Gifting NOT applied to future months (Matt's forecast has no
gifting). May want to scale based on Q1 trend later (~50-80 units/mo)
once user confirms gifting strategy with Matt.
"""
# Inline change via REPL on May 22.
