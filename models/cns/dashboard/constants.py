"""Static lookups and display helpers for the CNS dashboard."""

import sys
from pathlib import Path

# Add parent so we can import baseline_data
sys.path.insert(0, str(Path(__file__).parent.parent))

from baseline_data import FORECAST_MONTH_LABELS, NUM_FORECAST_MONTHS, MONTHS_12, LOCATIONS

N = NUM_FORECAST_MONTHS  # 60

# Month display helpers
def month_idx_to_label(i: int) -> str:
    """Convert 0-based forecast index to label like 'Jan-26'."""
    return FORECAST_MONTH_LABELS[i] if i < len(FORECAST_MONTH_LABELS) else f"M{i}"

def month_idx_to_year(i: int) -> int:
    """Convert 0-based index to calendar year."""
    return 2026 + i // 12

def year_slice(year: int) -> slice:
    """Return slice for a given year within the 60-month array."""
    offset = (year - 2026) * 12
    return slice(offset, offset + 12)

# Overhead expense categories (for display)
OVERHEAD_CATEGORIES = [
    ("billing", "Billing (18% of Collected)"),
    ("payroll", "Payroll (W-2)"),
    ("contractors", "Contractor Costs"),
    ("opex", "Operating Expenses"),
    ("expansion", "Expansion Costs"),
]

# OpEx line items (from Assumptions tab)
OPEX_LINE_ITEMS = [
    ("marketing_monthly", "Advertising & Marketing"),
    ("contracts_monthly", "Contracts & Bonuses"),
    ("general_insurance_monthly", "Insurance (General)"),
    ("malpractice_annual", "Malpractice Insurance (Annual)"),
    ("health_insurance_monthly", "Health Insurance"),
    ("legal_monthly_recurring", "Legal & Accounting"),
    ("mgmt_fee_abc_monthly", "Management Fee (ABC A PC)"),
    ("office_software_monthly", "Office & Software"),
    ("rent_westlake_monthly", "Rent (Westlake)"),
]

# Cash flow line items
CF_ROWS = [
    "Beginning Cash",
    "Cash Collected",
    "Total Overhead",
    "Cash After Overhead",
    "Distributable (above min)",
    "Physician Services (90%)",
    "Savings Deposit (10%)",
    "Ending Cash",
    "Savings Balance",
]

# Surgery types
SURGERY_TYPES = [
    ("bobas", "BOBA", "#2c3e50"),
    ("gap", "GAP", "#3498db"),
]

# Currency format
def fmt_currency(val: float, decimals: int = 0) -> str:
    if val < 0:
        return f"-${abs(val):,.{decimals}f}"
    return f"${val:,.{decimals}f}"
