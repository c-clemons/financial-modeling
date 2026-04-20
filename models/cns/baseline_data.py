"""
CNS (California Neurosurgical Specialists) - Baseline Data
Hard-coded 2025 actuals from QBO and baseline assumptions for 2026-2030 forecast.
QBO Export Date: March 3, 2026 (Accrual Basis)
"""

from datetime import date

# ============================================================
# FORECAST HORIZON
# ============================================================
NUM_FORECAST_MONTHS = 60  # Jan 2026 - Dec 2030

MONTHS_12 = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
             "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

FORECAST_MONTH_LABELS = [f"{m}-{yr}" for yr in range(26, 31) for m in MONTHS_12]


# ============================================================
# HISTORICAL CASE VOLUME (surgery counts by month)
# Now with BOBA/GAP breakdown from client data
# ============================================================

HISTORICAL_CASE_VOLUME = {
    # 2024 (partial - Sep-Dec)
    '2024': {
        'Sep': 5, 'Oct': 8, 'Nov': 14, 'Dec': 10,
    },
    # 2025 (full year)
    '2025': {
        'Jan': 13, 'Feb': 6, 'Mar': 3, 'Apr': 15,
        'May': 11, 'Jun': 11, 'Jul': 7, 'Aug': 9,
        'Sep': 9, 'Oct': 6, 'Nov': 17, 'Dec': 9,
    },
    # 2026 (through Apr - Mar mostly closed, Apr from client schedule)
    '2026': {
        'Jan': 3, 'Feb': 2, 'Mar': 6, 'Apr': 1,
    },
}

# BOBA/GAP breakdown by month (from client GAP/BOBA cases sheet)
# 2024 (Sep-Dec)
BOBA_VOLUME_2024 = [1, 1, 2, 5]       # Sep, Oct, Nov, Dec
GAP_VOLUME_2024 = [0, 2, 1, 0]         # Sep, Oct, Nov, Dec

# 2025 (full year, Jan-Dec)
BOBA_VOLUME_2025 = [5, 1, 1, 3, 0, 3, 1, 1, 3, 3, 1, 4]  # Jan-Dec
GAP_VOLUME_2025 = [0, 1, 0, 0, 0, 1, 2, 2, 0, 0, 4, 1]    # Jan-Dec

# 2026 actuals (Jan-Mar from client data, closed months)
BOBA_2026_ACTUALS = [2, 1, 3]           # Jan, Feb, Mar
GAP_2026_ACTUALS = [0, 1, 1]             # Jan, Feb, Mar

SURGERY_VOLUME_2024 = [5, 8, 14, 10]  # Sep-Dec 2024
SURGERY_VOLUME_2025 = [13, 6, 3, 15, 11, 11, 7, 9, 9, 6, 17, 9]  # Jan-Dec 2025
SURGERY_VOLUME_2026_ACTUALS = [3, 2, 6]  # Jan-Mar 2026

# Derived stats
TOTAL_SURGERIES_2025 = sum(SURGERY_VOLUME_2025)  # 116
AVG_MONTHLY_SURGERIES_2025 = TOTAL_SURGERIES_2025 / 12  # ~9.7
AVG_REVENUE_PER_SURGERY = 2288432.70 / TOTAL_SURGERIES_2025  # ~$19,728

# ============================================================
# SURGERY VOLUME DEFAULTS - BOBAS & GAP (60 months)
# ============================================================

# Westlake volumes: Jan-Mar actuals [2,1,3], Apr-Dec forecast ramp 2→8, then flat 8
_bobas_actuals = [2, 1, 3]  # Jan-Mar 2026 from client data
_bobas_forecast_apr_dec = [2, 3, 4, 4, 5, 5, 6, 7, 8]  # Apr-Dec 2026
_bobas_2027_2030 = [8] * 48
_WESTLAKE_BOBAS = _bobas_actuals + _bobas_forecast_apr_dec + _bobas_2027_2030

_gap_actuals = [0, 1, 1]  # Jan-Mar 2026 from client data
_gap_forecast_apr_dec = [1, 2, 2, 2, 3, 3, 3, 4, 4]  # Apr-Dec 2026
_gap_2027_2030 = [4] * 48
_WESTLAKE_GAP = _gap_actuals + _gap_forecast_apr_dec + _gap_2027_2030

# Consolidated defaults (sum of all locations — backwards compatible)
BOBAS_VOLUME_DEFAULT = _WESTLAKE_BOBAS  # updated below after VOLUMES_BY_LOCATION is defined
GAP_VOLUME_DEFAULT = _WESTLAKE_GAP      # updated below

# ============================================================
# MULTI-LOCATION SUPPORT
# ============================================================

LOCATIONS = ['Westlake', 'Santa Barbara']

# Per-location volumes: Westlake keeps existing forecast, SB ramps after lease start (month 6)
_sb_bobas = [0]*6 + [1, 1, 2, 2, 3, 4] + [4]*48   # starts Jul-26, ramps to 4 over 12mo
_sb_gap = [0]*6 + [0, 1, 1, 1, 1, 2] + [2]*48       # starts Jul-26, ramps to 2

VOLUMES_BY_LOCATION = {
    'Westlake': {
        'bobas': list(BOBAS_VOLUME_DEFAULT),
        'gap': list(GAP_VOLUME_DEFAULT),
    },
    'Santa Barbara': {
        'bobas': _sb_bobas,
        'gap': _sb_gap,
    },
}

# Consolidated volumes (sum across all locations) — backwards compatible
BOBAS_VOLUME_DEFAULT = [sum(loc['bobas'][i] for loc in VOLUMES_BY_LOCATION.values()) for i in range(NUM_FORECAST_MONTHS)]
GAP_VOLUME_DEFAULT = [sum(loc['gap'][i] for loc in VOLUMES_BY_LOCATION.values()) for i in range(NUM_FORECAST_MONTHS)]

# Per-location operating expenses
OPEX_BY_LOCATION = {
    'Westlake': {
        'marketing_monthly': 8000.00,
        'rent_monthly': 6250.00,
        'contracts_monthly': 12000.00,
        'office_software_monthly': 9000.00,
    },
    'Santa Barbara': {
        'marketing_monthly': 3000.00,
        'rent_monthly': 7500.00,      # from expansion config
        'contracts_monthly': 3000.00,
        'office_software_monthly': 2000.00,
    },
}

# Shared overhead (allocated by revenue % across locations)
SHARED_OVERHEAD = {
    'legal_monthly_recurring': 3000.00,
    'malpractice_annual': 7716.00,
    'general_insurance_monthly': 570.00,
    'health_insurance_monthly': 0,
    'mgmt_fee_abc_monthly': 0,
    'bank_fees_monthly': 100.00,
}


# Collection curves (values as percentages, must sum to 100)
# Data-driven from actual paid cases:
#   BOBA: 12 paid cases. Actual $-weighted: M+2=2.5%, M+5=27%, M+6=10%, M+7=32%, M+8=27%, M+10=0.3%
#   Smoothed/normalized to sum to 100:
BOBAS_COLLECTION_CURVE = [0, 0, 3, 0, 0, 27, 10, 32, 23, 0, 5, 0]  # 12 elements
#   GAP: 3 paid cases. $-weighted: M+1=65%, M+4=35%. Extended with small M+2/M+3 buffer:
GAP_COLLECTION_CURVE = [0, 55, 10, 10, 25]  # 5 elements


# ============================================================
# 2025 MONTHLY P&L ACTUALS (from QBO)
# ============================================================

ACTUALS_2025 = {
    'fee_income': [7025.36, 12374.14, 39829.53, 7757.15, 94593.27, 629391.77, 93537.70, 300965.93, 67739.62, 475648.62, 100946.79, 458622.82],
    'reimbursed_expense_income': [53968.59, 29792.92, 52800.00, 87285.79, 14400.00, 0, 62979.68, 12800.00, 20800.00, 0, 0, 0],
    'refunds': [0, 0, 0, 0, 0, 0, -14061.00, 0, 0, 0, 0, 0],
    # Expenses
    'advertising_marketing': [0, 0, 0, 0, 0, 0, 0, 8450.00, 6300.00, 5550.00, 10450.00, 8252.00],
    'bank_fees': [2.05, 2.28, 34.88, 44.00, 1.40, 0.10, 0, 42.59, 26.48, 97.55, 59.18, 86.25],
    'conference': [0, 0, 0, 0, 0, 0, 0, 0, 0, 27.00, 0, 0],
    'contracts': [0, 0, 0, 0, 0, 0, 42044.06, 11430.00, 160.00, 10450.00, 2350.00, 10360.00],
    'contributions': [0, 0, 0, 0, 0, 0, 0, 500.00, 0, 0, 0, 0],
    'dues_subscriptions': [0, 0, 0, 0, 0, 0, 0, 0, 2850.00, 0, 0, 0],
    'insurance': [0, 0, 0, 0, 0, 0, 761.74, 239.34, 238.49, 1705.49, 238.49, 238.49],
    'malpractice_insurance': [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 7713.36, 0],
    'health_insurance': [0, 0, 0, 0, 0, 0, 0, 0, 1112.50, 556.25, 556.25, 556.25],
    'legal_accounting_services': [0, 0, 0, 0, 0, 0, 0, 2400.00, 2400.00, 2560.00, 2432.00, 3840.00],
    'accounting_fees': [0, 0, 0, 0, 0, 0, 0, 3069.50, 0, 7533.50, 0, 0],
    'billing_services': [0, 0, 1405.00, 0, 0, 0, 166412.32, 19632.47, 69456.52, 12525.96, 78475.11, 44769.24],
    'licenses_fees': [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 44.00],
    'mgmt_fee_abc': [0, 0, 0, 21617.16, 0, 0, 6250.00, 12379.62, 0, 0, 0, 12692.13],
    'mgmt_fee_vnsc': [0, 4400.00, 5400.00, 5000.00, 5000.00, 5000.00, 5000.00, 0, 0, 0, 0, 0],
    'meals': [0, 0, 0, 0, 0, 0, 0, 139.56, 57.72, 153.07, 47.22, 82.37],
    'office_expenses': [0, 0, 0, 0, 30.00, 0, 0, 540.00, 128.66, 2722.23, 8248.96, 2051.50],
    'software_apps': [0, 0, 0, 0, 0, 0, 1128.54, 1000.03, 7284.66, 7150.44, 6496.03, 2570.15],
    'payroll_processing': [0, 0, 0, 0, 0, 0, 0, 0, 61.00, 67.00, 67.00, 67.00],
    'salaries_wages': [0, 0, 0, 0, 0, 0, 6800.00, 18000.00, 20392.00, 22732.00, 24640.00, 24029.00],
    'payroll_taxes': [0, 0, 0, 0, 0, 0, 799.00, 1672.20, 1658.05, 1927.94, 1884.96, 2090.22],
    'physician_services': [65000.00, 0, 125000.00, 65000.00, 65000.00, 0, 325187.58, 0, 111283.76, 191548.33, 256444.79, 371208.27],
    'rent_lease': [0, 0, 0, 0, 0, 0, 0, 26108.20, 2500.00, 2500.00, 6240.83, 0],
    'taxes_licenses': [0, 0, 0, 64.00, 130.00, 0, 0, 0, 0, 0, 0, 0],
    'travel': [0, 0, 0, 0, 0, 0, 200.00, 0, 0, 30.99, 0, 0],
    # Other Income/Expenses
    'interest_income': [0, 0, 0, 0, 0, 0, 0.17, 0.28, 0.40, 0.56, 0.73, 0.89],
    'state_taxes': [0, 0, 0, 800.00, 0, 0, 0, 0, 0, 0, 0, 0],
    'penalties': [0, 0, 0, 0, 0, 0, 0, 40.00, 0, 0, 381.99, 0],
    'depreciation': [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 19558.71],
}

ACTUALS_2025_TOTALS = {
    'total_income': 2609198.68,
    'fee_income': 2288432.70,
    'reimbursed_expense_income': 334826.98,
    'refunds': -14061.00,
    'total_expenses': 2413346.26,
    'net_operating_income': 195852.42,
    'net_income': 175074.75,
    'physician_services': 1575672.73,
    'billing_services': 392676.62,
    'total_legal_accounting': 24235.00,
    'total_payroll': 126887.37,
    'total_insurance': 13916.65,
}


# ============================================================
# 2026 QBO ACTUALS (Jan-Feb, Cash Basis)
# Source: P&L 1_2026-2_28_2026.xlsx
# ============================================================

ACTUALS_2026_QBO = {
    'months': ['Jan-26', 'Feb-26'],
    'fee_income': [391983.69, 15487.16],
    'total_income': [391983.69, 15487.16],
    # Expenses
    'advertising_marketing': [8251.00, 15352.00],
    'bank_fees': [570.90, 0],
    'contracts': [600.00, 200.00],
    'insurance': [238.49, 238.49],
    'malpractice_insurance': [7713.36, 0],
    'health_insurance': [556.25, 263.13],
    'supplies_medical': [0, 961.84],
    'legal_fees': [60425.99, 79143.18],
    'licenses_fees': [0, 1206.00],
    'meals': [410.12, 146.43],
    'software_apps': [899.99, 800.35],
    'office_expenses': [1951.33, 1339.77],
    'payroll_expenses': [31710.15, 32007.15],
    'physician_services': [244591.40, 0],
    'rent_lease': [6239.20, 6233.36],
    'repairs_maintenance': [0, 1066.72],
    'travel': [650.00, 0],
    'uniforms': [0, 200.56],
    'total_expenses': [364808.18, 138895.85],
    'net_operating_income': [27175.51, -123408.69],
    'interest_income': [1.00, 0],
    'net_income': [27176.51, -123408.69],
}

# Number of actuals months in 2026 (used by build script to know where forecast starts)
NUM_2026_ACTUALS = len(ACTUALS_2026_QBO['months'])


# ============================================================
# 2025 MONTHLY BALANCE SHEET (from QBO)
# ============================================================

BALANCE_SHEET_2025 = {
    'chase_checking': [4460.47, 42954.08, 3743.73, 6261.51, 45093.38, 669485.05, 156426.24, 369257.72, 217571.30, 403194.59, 45654.38, 150225.89],
    'chase_savings': [0, 0, 0, 0, 0, 0, 36132.12, 36132.40, 48497.66, 69781.37, 98275.97, 123976.86],
    'total_cash': [4460.47, 42954.08, 3743.73, 6261.51, 45093.38, 669485.05, 192558.36, 405390.12, 266068.96, 472975.96, 143930.35, 274202.75],
    'furniture_fixtures': [0, 0, 0, 0, 0, 0, 0, 0, 0, 9944.01, 12366.79, 12366.79],
    'machinery_equipment': [0, 0, 0, 0, 0, 0, 0, 0, 0, 7191.92, 7191.92, 7191.92],
    'accumulated_depreciation': [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -19558.71],
    'total_fixed_assets': [0, 0, 0, 0, 0, 0, 0, 0, 0, 17135.93, 19558.71, 0],
    'total_assets': [4470.47, 42964.08, 3753.73, 6271.51, 45103.38, 669495.05, 192568.36, 405400.12, 266078.96, 490111.89, 163489.06, 274202.75],
    'credit_cards': [0, 0, 0, 0, 0, 0, 0, 4609.06, 2657.72, 20889.22, 45.04, 1694.20],
    'payroll_tax_payable': [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 336.00],
    'due_to_abc_apc': [-4309.45, -3580.62, -3580.62, -3580.62, -3580.62, -3580.62, -3380.62, -3280.62, -3280.62, -3280.62, -3280.62, 149319.78],
    'due_to_vnsc': [65000, 65000, 65000, 65000, 65000, 65000, 0, 0, 0, 0, 0, 0],
    'total_liabilities': [60690.55, 61419.38, 61419.38, 61419.38, 61419.38, 61419.38, -3380.62, 1328.44, -622.90, 17608.60, -3235.58, 151349.98],
    'common_stock': [10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10],
    'retained_earnings': [-52221.98, -52221.98, -52221.98, -52221.98, -52221.98, -52221.98, -52221.98, -52221.98, -52221.98, -52221.98, -52221.98, -52221.98],
    'distributions_benet': [0, 0, 0, 0, 0, 0, 0, 0, 0, -5, -5, -5],
    'distributions_taqi': [0, 0, 0, 0, 0, 0, 0, 0, 0, -5, -5, -5],
    'net_income_cumulative': [-4008.10, 33756.68, -5453.67, -2935.89, 35895.98, 660287.65, 248160.96, 456283.66, 318913.84, 524725.27, 218946.62, 175074.75],
    'total_equity': [-56220.08, -18455.30, -57665.65, -55147.87, -16316.00, 608075.67, 195948.98, 404071.68, 266701.86, 472503.29, 166724.64, 122852.77],
}


# ============================================================
# BASELINE TEAM (payroll paid monthly, end of month)
# ============================================================

# 15 team slots: first 4 are existing, rest are open for new hires.
# start_month: 0-indexed into 60-month forecast (0=Jan-26, 6=Jul-26, etc.)
#              None means slot is empty / not hired yet.
# end_month: None means active through end of forecast.
TEAM_ROSTER = [
    # Westlake staff
    {
        'title': 'Herlyn',
        'monthly_salary': 11000,
        'employment_type': 'W-2',
        'start_month': 0,       # Jan 2026
        'end_month': None,
        'location': 'Westlake',
        'notes': 'W-2 employee, $11K/mo',
    },
    {
        'title': 'Christina',
        'monthly_salary': 10000,
        'employment_type': 'W-2',
        'start_month': 0,       # Jan 2026
        'end_month': None,
        'location': 'Westlake',
        'notes': 'W-2 employee, $10K/mo',
    },
    {
        'title': 'Front Desk (FT)',
        'monthly_salary': 6400,
        'employment_type': 'W-2',
        'start_month': 0,       # Jan 2026
        'end_month': 2,         # Last day March 20, 2026
        'partial_last_month': 20 / 31,
        'location': 'Westlake',
        'notes': 'W-2 employee, 40hr/wk. Last day 3/20/2026.',
    },
    {
        'title': 'Virtual Assistant',
        'monthly_salary': 4000,
        'employment_type': 'Contractor',
        'start_month': 0,
        'end_month': None,
        'location': 'Westlake',
        'notes': 'Via NMed Consulting agency (Philippines) - $16/hour',
    },
    # Santa Barbara staff
    {
        'title': 'SB Office Manager',
        'monthly_salary': 5200,
        'employment_type': 'W-2',
        'start_month': 6,       # Jul 2026 (lease start)
        'end_month': None,
        'location': 'Santa Barbara',
        'notes': 'Santa Barbara office manager',
    },
    # Slots 6-15: open for new hires (location assignable)
    {'title': 'New Hire 6', 'monthly_salary': 5200, 'employment_type': 'W-2', 'start_month': None, 'end_month': None, 'location': 'Westlake', 'notes': ''},
    {'title': 'New Hire 7', 'monthly_salary': 5200, 'employment_type': 'W-2', 'start_month': None, 'end_month': None, 'location': 'Westlake', 'notes': ''},
    {'title': 'New Hire 8', 'monthly_salary': 5200, 'employment_type': 'W-2', 'start_month': None, 'end_month': None, 'location': 'Santa Barbara', 'notes': ''},
    {'title': 'New Hire 9', 'monthly_salary': 5200, 'employment_type': 'W-2', 'start_month': None, 'end_month': None, 'location': 'Santa Barbara', 'notes': ''},
    {'title': 'New Hire 10', 'monthly_salary': 5200, 'employment_type': 'W-2', 'start_month': None, 'end_month': None, 'location': 'Westlake', 'notes': ''},
    {'title': 'New Hire 11', 'monthly_salary': 5200, 'employment_type': 'W-2', 'start_month': None, 'end_month': None, 'location': 'Westlake', 'notes': ''},
    {'title': 'New Hire 12', 'monthly_salary': 5200, 'employment_type': 'W-2', 'start_month': None, 'end_month': None, 'location': 'Santa Barbara', 'notes': ''},
    {'title': 'New Hire 13', 'monthly_salary': 5200, 'employment_type': 'W-2', 'start_month': None, 'end_month': None, 'location': 'Santa Barbara', 'notes': ''},
    {'title': 'New Hire 14', 'monthly_salary': 5200, 'employment_type': 'W-2', 'start_month': None, 'end_month': None, 'location': 'Westlake', 'notes': ''},
    {'title': 'New Hire 15', 'monthly_salary': 5200, 'employment_type': 'W-2', 'start_month': None, 'end_month': None, 'location': 'Westlake', 'notes': ''},
]

PAYROLL_2025_TOTAL = {
    'salaries': 116593.00,
    'payroll_taxes': 10032.37,
    'processing': 262.00,
    'total': 126887.37,
    'dec_monthly_salary': 24029.00,
    'dec_monthly_total': 26186.22,
}


# ============================================================
# BASELINE OPERATING EXPENSES
# ============================================================

BASELINE_OPEX_RECURRING = [
    {'expense_name': 'Advertising & Marketing', 'category': 'Marketing', 'monthly_amount': 8000.00,
     'notes': 'Based on Aug-Dec 2025 average.'},
    {'expense_name': 'Contracts - Bonuses and Reimbursements', 'category': 'Professional Services', 'monthly_amount': 12000.00,
     'notes': 'Based on Jul-Dec 2025 average. Varies month to month.'},
    {'expense_name': 'Insurance - General', 'category': 'Insurance', 'monthly_amount': 570.00,
     'notes': 'General insurance, spread monthly'},
    {'expense_name': 'Insurance - Malpractice', 'category': 'Insurance', 'monthly_amount': 643.00,
     'notes': 'Annual malpractice premium, spread monthly'},
    {'expense_name': 'Insurance - Health', 'category': 'Insurance', 'monthly_amount': 0,
     'notes': 'Employee health insurance, ~$556/mo'},
    {'expense_name': 'Legal & Accounting Services', 'category': 'Legal & Accounting', 'monthly_amount': 3000.00,
     'notes': 'True legal/accounting (accts 550+552). Billing is separate.'},
    {'expense_name': 'Management Fee (ABC A PC)', 'category': 'Management Fees', 'monthly_amount': 0,
     'notes': 'Management fee to ABC A PC. Structure TBD.'},
    {'expense_name': 'Office Expenses & Software', 'category': 'Office & Software', 'monthly_amount': 9000.00,
     'notes': 'Combined office expenses + software.'},
    {'expense_name': 'Rent/Lease - Westlake', 'category': 'Rent', 'monthly_amount': 6250.00,
     'notes': 'Westlake office lease.'},
    {'expense_name': 'NMed Consulting (VA)', 'category': 'Contractors', 'monthly_amount': 0,
     'notes': 'Virtual assistant in Philippines via NMed agency. Now tracked in team roster.'},
    {'expense_name': 'Bank Fees / G&A', 'category': 'G&A', 'monthly_amount': 100.00,
     'notes': 'Bank fees, meals, travel - minimal'},
]

ONETIME_2025 = [
    {'expense_name': 'VNSC Management Fee', 'amount': 29800.00,
     'notes': 'Paid Feb-Jul 2025 only. Ended Jul 2025.'},
    {'expense_name': 'Rent - Initial Setup (Aug deposit/TI)', 'amount': 26108.20,
     'notes': 'Aug 2025 one-time rent charge.'},
    {'expense_name': 'Dues & Subscriptions', 'amount': 2850.00,
     'notes': 'Sep 2025 one-time. May recur annually.'},
]


# ============================================================
# DEFAULT ASSUMPTIONS (2026-2030 Forecast)
# ============================================================

DEFAULT_ASSUMPTIONS = {
    # Revenue - Bobas (data-driven: mean $132K from 15 paid cases, median $93K)
    'bobas_volume': BOBAS_VOLUME_DEFAULT,  # 60-element list
    'avg_revenue_bobas': 132000,
    'bobas_collection_curve': BOBAS_COLLECTION_CURVE,  # percentages

    # Revenue - GAP (data-driven: mean $50K from 8 paid cases, median $47K)
    'gap_volume': GAP_VOLUME_DEFAULT,  # 60-element list
    'avg_revenue_gap': 50000,
    'gap_collection_curve': GAP_COLLECTION_CURVE,  # percentages

    # Reimbursed expenses & refunds (stored as percentages)
    'reimbursed_expense_pct': 0,
    'refund_rate': 0,

    # Fund Flow: Revenue → pay overhead → net equity
    'physician_services_rate': 90,  # percentage
    'savings_rate': 10,  # percentage

    # Billing Services (MD Capital): 18% of cash collected
    'billing_fee_rate': 18,  # percentage

    # Payroll
    'payroll_tax_rate': 8.6,  # percentage
    'salary_annual_increase': 5.0,  # 5% annual raise on salaries

    # Legal & Accounting
    'legal_monthly_recurring': 3000.00,

    # Management Fees
    'mgmt_fee_abc_monthly': 0,

    # Cash
    'starting_cash': 274202.75,
    'minimum_cash_balance': 150000.00,
    'starting_savings': 0,

    # Insurance
    'malpractice_annual': 7716.00,
    'general_insurance_monthly': 570.00,
    'health_insurance_monthly': 0,

    # Inflation (applied to fixed expenses, NOT %-of-revenue items)
    'expense_annual_inflation': 3.0,  # 3% annual increase

    # OpEx (legacy consolidated — used by existing functions for backwards compat)
    'marketing_monthly': 8000.00,
    'rent_westlake_monthly': 6250.00,
    'office_software_monthly': 9000.00,
    'contracts_monthly': 12000.00,
    'nmed_va_monthly': 0,

    # Multi-location support
    'locations': LOCATIONS,
    'volumes_by_location': VOLUMES_BY_LOCATION,
    'opex_by_location': OPEX_BY_LOCATION,
    'shared_overhead': SHARED_OVERHEAD,
    'shared_overhead_allocation': 'revenue_pct',
    'team_roster': TEAM_ROSTER,

    # Per-surgeon compensation at expansion locations
    # Surgeons at new locations get a % of their location's collections
    'surgeon_compensation': {
        'Santa Barbara': {'rate': 70, 'surgeon_name': 'TBD'},
    },

    # Multi-Location Expansions (5 slots)
    'expansions': [
        {
            'name': 'Santa Barbara',
            'enabled': True,
            'lease_start_month': 6,     # Jul 2026 (0-indexed into 60-mo forecast)
            'lease_monthly': 7500,
            'ti_total': 200000,
            'ti_cns_share': 100000,
            'ti_start_month': 4,        # May 2026
            'ti_duration_months': 2,
            'ffe_budget': 15000,
            'opex_monthly': 40000,       # full run-rate after ramp
            'opex_ramp_monthly': 15000,  # first 6 months at reduced rate
            'opex_ramp_months': 6,       # months at ramp rate before full
        },
        {'name': 'Expansion 2', 'enabled': False, 'lease_start_month': 18,
         'lease_monthly': 0, 'ti_total': 0, 'ti_cns_share': 0,
         'ti_start_month': 16, 'ti_duration_months': 2, 'ffe_budget': 0, 'opex_monthly': 0},
        {'name': 'Expansion 3', 'enabled': False, 'lease_start_month': 20,
         'lease_monthly': 0, 'ti_total': 0, 'ti_cns_share': 0,
         'ti_start_month': 18, 'ti_duration_months': 2, 'ffe_budget': 0, 'opex_monthly': 0},
        {'name': 'Expansion 4', 'enabled': False, 'lease_start_month': 22,
         'lease_monthly': 0, 'ti_total': 0, 'ti_cns_share': 0,
         'ti_start_month': 20, 'ti_duration_months': 2, 'ffe_budget': 0, 'opex_monthly': 0},
        {'name': 'Expansion 5', 'enabled': False, 'lease_start_month': 23,
         'lease_monthly': 0, 'ti_total': 0, 'ti_cns_share': 0,
         'ti_start_month': 21, 'ti_duration_months': 2, 'ffe_budget': 0, 'opex_monthly': 0},
    ],
}


# ============================================================
# QBO ACCOUNT MAPPING
# ============================================================

QBO_ACCOUNTS = {
    '400': {'name': 'Fee Income', 'category': 'Revenue', 'key': 'fee_income'},
    '410': {'name': 'Reimbursed Expense Income', 'category': 'Revenue', 'key': 'reimbursed_expense_income'},
    '430': {'name': 'Refunds', 'category': 'Revenue', 'key': 'refunds'},
    '500': {'name': 'Advertising & Marketing', 'category': 'Marketing', 'key': 'advertising_marketing'},
    '505': {'name': 'Bank Fees', 'category': 'G&A', 'key': 'bank_fees'},
    '515': {'name': 'Conference', 'category': 'G&A', 'key': 'conference'},
    '520': {'name': 'Contracts', 'category': 'Professional Services', 'key': 'contracts'},
    '525': {'name': 'Contributions', 'category': 'G&A', 'key': 'contributions'},
    '535': {'name': 'Dues & Subscriptions', 'category': 'G&A', 'key': 'dues_subscriptions'},
    '540': {'name': 'Insurance', 'category': 'Insurance', 'key': 'insurance'},
    '541': {'name': 'Malpractice Insurance', 'category': 'Insurance', 'key': 'malpractice_insurance'},
    '542': {'name': 'Health Insurance', 'category': 'Insurance', 'key': 'health_insurance'},
    '550': {'name': 'Legal & Accounting Services', 'category': 'Legal & Accounting', 'key': 'legal_accounting_services'},
    '552': {'name': 'Accounting Fees', 'category': 'Legal & Accounting', 'key': 'accounting_fees'},
    '554': {'name': 'MD Capital Billing Services', 'category': 'Billing Services', 'key': 'billing_services'},
    '555': {'name': 'Licenses & Fees', 'category': 'G&A', 'key': 'licenses_fees'},
    '556': {'name': 'Management Fee (ABC A PC)', 'category': 'Management Fees', 'key': 'mgmt_fee_abc'},
    '558': {'name': 'Management Fee (VNSC)', 'category': 'Management Fees', 'key': 'mgmt_fee_vnsc'},
    '560': {'name': 'Meals', 'category': 'G&A', 'key': 'meals'},
    '570': {'name': 'Software & Apps', 'category': 'Office & Software', 'key': 'software_apps'},
    '574': {'name': 'Office Expenses', 'category': 'Office & Software', 'key': 'office_expenses'},
    '580': {'name': 'Payroll Processing', 'category': 'Payroll', 'key': 'payroll_processing'},
    '584': {'name': 'Salaries & Wages', 'category': 'Payroll', 'key': 'salaries_wages'},
    '586': {'name': 'Payroll Taxes', 'category': 'Payroll', 'key': 'payroll_taxes'},
    '592': {'name': 'Physician Services (ABC A PC)', 'category': 'Physician Services', 'key': 'physician_services'},
    '600': {'name': 'Rent/Lease', 'category': 'Rent', 'key': 'rent_lease'},
    '605': {'name': 'Taxes & Licenses', 'category': 'G&A', 'key': 'taxes_licenses'},
    '610': {'name': 'Travel', 'category': 'G&A', 'key': 'travel'},
    '440': {'name': 'Interest Income', 'category': 'Other Income', 'key': 'interest_income'},
    '612': {'name': 'State Taxes', 'category': 'Other Expense', 'key': 'state_taxes'},
    '616': {'name': 'Penalties', 'category': 'Other Expense', 'key': 'penalties'},
    '625': {'name': 'Depreciation', 'category': 'Non-cash', 'key': 'depreciation'},
}


# ============================================================
# HISTORICAL AR SPILLOVER
# All past BOBA/GAP surgeries assumed to be in AR (uncollected).
# Apply collection curve from surgery date. Collections falling
# before Jan 2026 are treated as overdue → collected in Jan 2026.
# ============================================================

def _compute_ar_spillover(hist_months_volumes, avg_rev, curve, num_months=60,
                          overdue_spread_months=12):
    """
    hist_months_volumes: list of (month_offset, count) where month_offset is
        relative to Jan 2026 (e.g., Sep 2024 = -16, Dec 2025 = -1).
    avg_rev: average revenue per surgery.
    curve: collection curve percentages (list, sums to 100).
    overdue_spread_months: number of months to spread overdue AR across
        (instead of dumping all into Jan 2026).
    Returns num_months-element list of AR collections landing in forecast.
    """
    spillover = [0.0] * num_months
    overdue_pool = 0.0  # collections that would have landed before Jan 2026
    curve_dec = [p / 100.0 for p in curve]
    for month_offset, count in hist_months_volumes:
        if count == 0:
            continue
        revenue = count * avg_rev
        for lag_idx, pct in enumerate(curve_dec):
            if pct == 0:
                continue
            target = month_offset + lag_idx
            if target >= num_months:
                continue
            if target < 0:
                overdue_pool += revenue * pct  # accumulate overdue
            else:
                spillover[target] += revenue * pct

    # Spread overdue AR evenly across the first N months
    if overdue_pool > 0 and overdue_spread_months > 0:
        monthly_overdue = overdue_pool / overdue_spread_months
        for i in range(min(overdue_spread_months, num_months)):
            spillover[i] += monthly_overdue

    return spillover


# Historical months: label, offset, BOBA still in AR, GAP still in AR
# Data-driven: reconciled against client case sheets (paid/denied excluded)
# BOBA: 22 of 35 total still in AR (13 paid)
# GAP: 9 of 14 total still in AR (3 paid, 3 denied)
HISTORICAL_MONTHS = [
    ("Sep-24", -16, 1, 0),   # BOBA: 1 total, 0 paid → 1 AR. GAP: 0
    ("Oct-24", -15, 0, 2),   # BOBA: 1 paid (UHC $451K). GAP: 2 AR (Kaiser, Medicare)
    ("Nov-24", -14, 1, 1),   # BOBA: 1 paid (Cigna). GAP: 1 AR (Blueshield $3K)
    ("Dec-24", -13, 3, 0),   # BOBA: 5 total, 2 paid → 3 AR. GAP: 0
    ("Jan-25", -12, 3, 0),   # BOBA: 4 total, 1 paid (Cigna $317K) → 3 AR
    ("Feb-25", -11, 0, 0),   # BOBA: 1 paid (UHC $94K). GAP: 1 denied (Medicaid)
    ("Mar-25", -10, 1, 0),   # BOBA: 1 AR (BS of CA, WON $120K pending)
    ("Apr-25",  -9, 2, 0),   # BOBA: 3 total, 1 paid (Anthem $94K) → 2 AR
    ("May-25",  -8, 0, 0),   # No cases
    ("Jun-25",  -7, 2, 0),   # BOBA: 3 total, 1 paid (Aetna $86K) → 2 AR. GAP: 1 denied (BCBS)
    ("Jul-25",  -6, 0, 1),   # BOBA: 1 paid (BCBS $44K). GAP: 3 total, 2 paid → 1 AR
    ("Aug-25",  -5, 0, 0),   # BOBA: 1 paid (Aetna $259K). GAP: 2 total, 1 paid, 1 denied
    ("Sep-25",  -4, 2, 0),   # BOBA: 3 total, 1 paid (Surest $18K) → 2 AR
    ("Oct-25",  -3, 2, 0),   # BOBA: 3 total, 1 paid (Aetna $209K) → 2 AR
    ("Nov-25",  -2, 1, 4),   # BOBA: 1 AR. GAP: 4 AR (UHC, Aetna, Anthem, BS)
    ("Dec-25",  -1, 4, 1),   # BOBA: 4 AR. GAP: 1 AR (UHC)
]

# Build historical month offset lists for AR calculation
_HIST_BOBA = [(h[1], h[2]) for h in HISTORICAL_MONTHS]
_HIST_GAP = [(h[1], h[3]) for h in HISTORICAL_MONTHS]

HISTORICAL_AR_BOBA = _compute_ar_spillover(
    _HIST_BOBA, DEFAULT_ASSUMPTIONS['avg_revenue_bobas'], BOBAS_COLLECTION_CURVE,
)
HISTORICAL_AR_GAP = _compute_ar_spillover(
    _HIST_GAP, DEFAULT_ASSUMPTIONS['avg_revenue_gap'], GAP_COLLECTION_CURVE,
)
HISTORICAL_AR_TOTAL = [HISTORICAL_AR_BOBA[i] + HISTORICAL_AR_GAP[i] for i in range(NUM_FORECAST_MONTHS)]


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def get_actuals_2025():
    return {k: list(v) for k, v in ACTUALS_2025.items()}

def get_balance_sheet_2025():
    return {k: list(v) for k, v in BALANCE_SHEET_2025.items()}

def get_default_assumptions():
    import copy
    return copy.deepcopy(DEFAULT_ASSUMPTIONS)

def get_baseline_opex():
    return [e.copy() for e in BASELINE_OPEX_RECURRING]

def get_team_roster():
    return [m.copy() for m in TEAM_ROSTER]

def get_onetime_2025():
    return [e.copy() for e in ONETIME_2025]

def get_monthly_total_income():
    return [
        ACTUALS_2025['fee_income'][i] +
        ACTUALS_2025['reimbursed_expense_income'][i] +
        ACTUALS_2025['refunds'][i]
        for i in range(12)
    ]

def get_monthly_total_expenses():
    expense_keys = [
        'advertising_marketing', 'bank_fees', 'conference', 'contracts',
        'contributions', 'dues_subscriptions', 'insurance', 'malpractice_insurance',
        'health_insurance', 'legal_accounting_services', 'accounting_fees',
        'billing_services', 'licenses_fees', 'mgmt_fee_abc', 'mgmt_fee_vnsc',
        'meals', 'office_expenses', 'software_apps', 'payroll_processing',
        'salaries_wages', 'payroll_taxes', 'physician_services', 'rent_lease',
        'taxes_licenses', 'travel',
    ]
    return [sum(ACTUALS_2025[k][i] for k in expense_keys) for i in range(12)]
