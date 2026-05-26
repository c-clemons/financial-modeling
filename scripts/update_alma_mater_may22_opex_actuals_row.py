"""
May 22 — Add 'QBO Actual Total OpEx' row below each year's TOTAL OTHER OPEX
==============================================================================

Per user direction: surface QBO actual OpEx alongside Matt's plan in the
Assumptions tab so closed-month variance is visible.

Adds 3 rows:
  R151: QBO Actual Total OpEx (closed months) — below 2026 TOTAL OPEX
        Formula: =IF(month<=$C$47, 'QBO Actuals'!Cnn65, 0)
        Pulls QBO!R65 (Total Expenses) for closed months Jan-Apr 2026
  R169: same pattern for 2027 (currently 0 — no actuals yet)
  R187: same for 2028

Green fill on closed-month cells.

QBO categorization (Sales & Marketing, Software, Travel, etc.) doesn't
map 1:1 to Matt's plan categories (Brand Creative, Marketing Channels —
Mgmt/Creative/Spend/Systems, etc.), so we surface the total only. Variance
analysis at the category level lives in Streamlit's Variance Analysis tab.
"""
# Note: Script body already applied inline via REPL on May 22.
# This file preserves the change for the audit trail.
