"""
May 26 — Parameterized blended Alpha WS ASP + Holiday 2027 sync
=================================================================

Per user direction:
  1. Excel Holiday 2027 (Beta) PO: 3000u → 4000u (fixes 2028 Beta min hitting 0)
  2. Wholesale section: parameterize ASP with Alpha WS Mix % cells

Excel changes:

R96: New row with Alpha WS Mix % parameters
  B96  "Alpha WS Mix % (blends ASP between Beta $144 and Alpha $244)"
  C96  "2026"  D96  5%  (yellow editable)
  E96  "2027"  F96  15%
  G96  "2028"  H96  30%

Wholesale ASP formulas (col Q in each channel row):
  R100-R104 (2026): =$C$8*(1-$D$96)+$C$10*$D$96  → $149 @ 5% mix
  R105-R109 (2027): =$C$8*(1-$F$96)+$C$10*$F$96  → $159 @ 15%
  R110-R114 (2028): =$C$8*(1-$H$96)+$C$10*$H$96  → $174 @ 30%

Beta WS ASP = $C$8 ($144); Alpha WS ASP = $C$10 ($244)
Adjust D96/F96/H96 to sensitivity-test Alpha penetration assumptions.

Resulting WS revenue (per Excel formulas):
  2026: 1,500 GG units × $149 = $223,500 (was $216K @ flat $144)
  2027: 4,000 GG units × $159 = $636,000 (was $576K)
  2028: 8,000 GG units × $174 = $1,392,000 (was $1.152M)

Streamlit BASELINE_WHOLESALE synced:
  - Dropped 2 Alpha entries (Fall 27 Alpha 800u, Fall 28 Alpha 1600u)
  - All remaining 6 entries now product_type='Beta' (matches Excel's
    all-Beta WS inventory assumption)
  - wholesale_price updated to blended values per year ($149/$159/$174)

Inventory health check (passing after Holiday 2027 = 4000u):
  2026 Beta min 1,300u | Dec 4,162u
  2027 Beta min 2,468u | Dec 5,782u
  2028 Beta min   907u | Dec 4,913u
"""
# Inline change via REPL.
