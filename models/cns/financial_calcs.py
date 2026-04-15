"""
CNS Financial Calculations Module
Revenue forecasting with payment lag, expense projections, cash flow modeling.
60-month forecast horizon: Jan 2026 - Dec 2030.

Fund Flow:
  Revenue → CNS Checking → pay overhead → remaining "net equity":
    physician_services_rate% → ABC A PC (physician services)
    savings_rate% → savings account
"""

from typing import Dict, List, Tuple
from baseline_data import (
    get_default_assumptions, ACTUALS_2025, ACTUALS_2025_TOTALS,
    BALANCE_SHEET_2025, NUM_FORECAST_MONTHS,
    TEAM_ROSTER,
    TOTAL_SURGERIES_2025, SURGERY_VOLUME_2025,
    HISTORICAL_AR_BOBA, HISTORICAL_AR_GAP, HISTORICAL_AR_TOTAL,
)

N = NUM_FORECAST_MONTHS  # 60


def forecast_bobas_revenue(
    volume: List[float],
    avg_rev: float,
    curve: List[float],
) -> Tuple[List[float], List[float]]:
    """
    Forecast Bobas fee income with payment lag.
    volume: N-element list of Bobas surgeries/month.
    avg_rev: avg revenue per Bobas surgery (e.g. 150000).
    curve: up to 12 elements, values as percentages (divide by 100 internally).
    Returns (earned, collected) as N-element lists.
    """
    earned = [s * avg_rev for s in volume]
    curve_dec = [p / 100.0 for p in curve]

    collected = [0.0] * N
    for month_idx in range(N):
        e = earned[month_idx]
        for lag_idx, pct in enumerate(curve_dec):
            collect_month = month_idx + lag_idx
            if collect_month < N:
                collected[collect_month] += e * pct

    return earned, collected


def forecast_gap_revenue(
    volume: List[float],
    avg_rev: float,
    curve: List[float],
) -> Tuple[List[float], List[float]]:
    """
    Forecast GAP fee income with payment lag.
    volume: N-element list of GAP surgeries/month.
    avg_rev: avg revenue per GAP surgery (e.g. 70000).
    curve: up to 3 elements, values as percentages (divide by 100 internally).
    Returns (earned, collected) as N-element lists.
    """
    earned = [s * avg_rev for s in volume]
    curve_dec = [p / 100.0 for p in curve]

    collected = [0.0] * N
    for month_idx in range(N):
        e = earned[month_idx]
        for lag_idx, pct in enumerate(curve_dec):
            collect_month = month_idx + lag_idx
            if collect_month < N:
                collected[collect_month] += e * pct

    return earned, collected


def forecast_payroll(
    team: List[Dict],
    payroll_tax_rate: float = 8.6,
    salary_annual_increase: float = 5.0,
) -> Tuple[List[float], List[float], List[float], List[float]]:
    """
    Calculate payroll from 15-slot team roster. Paid monthly, end of month.
    W-2 employees: salary + taxes + processing.
    Contractors: tracked separately (no taxes).
    payroll_tax_rate: percentage (e.g. 8.6), divided by 100 internally.
    salary_annual_increase: percentage annual raise (e.g. 5.0 = 5%).
    Raises applied at start of each new forecast year (months 12, 24, 36, 48).
    Returns (salaries, taxes, total_w2, contractors) as N-element lists.
    """
    tax_rate = payroll_tax_rate / 100.0
    raise_rate = salary_annual_increase / 100.0
    salaries = [0.0] * N
    contractors = [0.0] * N

    for month in range(N):
        year_idx = month // 12
        escalator = (1 + raise_rate) ** year_idx

        for person in team:
            start = person.get('start_month')
            if start is None:
                continue
            end = person.get('end_month', None)
            if month < start:
                continue
            if end is not None and month > end:
                continue
            base_salary = person['monthly_salary'] * escalator
            if end is not None and month == end:
                partial = person.get('partial_last_month', 1.0)
                base_salary = base_salary * partial

            if person.get('employment_type') == 'Contractor':
                contractors[month] += base_salary
            else:
                salaries[month] += base_salary

    taxes = [s * tax_rate for s in salaries]
    processing = [67.0 if s > 0 else 0.0 for s in salaries]
    total_w2 = [salaries[i] + taxes[i] + processing[i] for i in range(N)]
    return salaries, taxes, total_w2, contractors


def forecast_opex(assumptions: Dict) -> Dict[str, List[float]]:
    """Monthly operating expenses by category. N-element lists.
    Applies annual inflation escalator to fixed costs (not %-of-revenue)."""
    a = assumptions
    inflation_rate = a.get('expense_annual_inflation', 3.0) / 100.0

    def _escalated(base_amount):
        """Return N-element list with annual inflation applied."""
        return [base_amount * (1 + inflation_rate) ** (m // 12) for m in range(N)]

    opex = {}
    opex['advertising_marketing'] = _escalated(a.get('marketing_monthly', 7834))
    opex['contracts'] = _escalated(a.get('contracts_monthly', 12799))

    general = a.get('general_insurance_monthly', 238.49)
    malpractice = a.get('malpractice_annual', 7713.36) / 12
    health = a.get('health_insurance_monthly', 556.25)
    opex['insurance_total'] = _escalated(general + malpractice + health)

    opex['legal_accounting'] = _escalated(a.get('legal_monthly_recurring', 3000))
    opex['mgmt_fee_abc'] = _escalated(a.get('mgmt_fee_abc_monthly', 4411.58))
    opex['office_software'] = _escalated(a.get('office_software_monthly', 5000))
    opex['rent_westlake'] = _escalated(a.get('rent_westlake_monthly', 2500))
    opex['nmed_consulting'] = _escalated(a.get('nmed_va_monthly', 3000))
    opex['general_admin'] = _escalated(100)

    return opex


def forecast_expansion_costs(assumptions: Dict) -> Dict:
    """
    Multi-location expansion costs. Up to 5 expansion slots.
    Applies annual inflation to ongoing lease and opex costs.
    Returns aggregate totals and per-expansion details.
    """
    expansions = assumptions.get('expansions', [])
    inflation_rate = assumptions.get('expense_annual_inflation', 3.0) / 100.0

    agg = {k: [0.0] * N for k in ['ti', 'lease', 'ffe', 'opex', 'total']}
    details = []
    locations_open = [1] * N  # Westlake always open

    for exp in expansions:
        d = {k: [0.0] * N for k in ['ti', 'lease', 'ffe', 'opex', 'total']}
        d['name'] = exp.get('name', 'Unknown')
        d['enabled'] = exp.get('enabled', False)

        if not d['enabled']:
            details.append(d)
            continue

        # TI (one-time capital, no inflation)
        ti_start = exp.get('ti_start_month', 0)
        ti_dur = exp.get('ti_duration_months', 2)
        ti_amount = exp.get('ti_cns_share', 0)
        if ti_dur > 0 and ti_amount > 0:
            per_month = ti_amount / ti_dur
            for m in range(ti_start, min(ti_start + ti_dur, N)):
                d['ti'][m] = per_month

        # Lease (with inflation)
        lease_start = exp.get('lease_start_month', 0)
        lease_amt = exp.get('lease_monthly', 0)
        for m in range(lease_start, N):
            escalator = (1 + inflation_rate) ** (m // 12)
            d['lease'][m] = lease_amt * escalator

        # FF&E (one-time, no inflation)
        ffe = exp.get('ffe_budget', 0)
        if lease_start < N and ffe > 0:
            d['ffe'][lease_start] = ffe

        # Ongoing opex (with inflation)
        sb_opex = exp.get('opex_monthly', 0)
        for m in range(lease_start, N):
            escalator = (1 + inflation_rate) ** (m // 12)
            d['opex'][m] = sb_opex * escalator

        # Location count
        for m in range(lease_start, N):
            locations_open[m] += 1

        d['total'] = [d['ti'][i] + d['lease'][i] + d['ffe'][i] + d['opex'][i] for i in range(N)]

        # Accumulate
        for k in ['ti', 'lease', 'ffe', 'opex', 'total']:
            for i in range(N):
                agg[k][i] += d[k][i]

        details.append(d)

    agg['details'] = details
    agg['locations_open'] = locations_open
    return agg


def generate_monthly_pl_forecast(
    assumptions: Dict = None,
) -> Dict[str, List[float]]:
    """
    Generate 60-month P&L forecast (Jan 2026 - Dec 2030).
    Billing fees on cash basis. Physician services on net equity.
    All rate params stored as percentages; divided by 100 internally.
    """
    if assumptions is None:
        assumptions = get_default_assumptions()
    a = assumptions

    # ---- REVENUE: BOBAS ----
    bobas_vol = a.get('bobas_volume', [8] * N)
    if len(bobas_vol) < N:
        bobas_vol = bobas_vol + [bobas_vol[-1]] * (N - len(bobas_vol))

    bobas_earned, bobas_collected = forecast_bobas_revenue(
        bobas_vol,
        a.get('avg_revenue_bobas', 150000),
        a.get('bobas_collection_curve', [0, 0, 0, 0, 0, 5, 10, 15, 25, 25, 15, 5]),
    )

    # ---- REVENUE: GAP ----
    gap_vol = a.get('gap_volume', [4] * N)
    if len(gap_vol) < N:
        gap_vol = gap_vol + [gap_vol[-1]] * (N - len(gap_vol))

    gap_earned, gap_collected = forecast_gap_revenue(
        gap_vol,
        a.get('avg_revenue_gap', 70000),
        a.get('gap_collection_curve', [0, 90, 10]),
    )

    # ---- COMBINED REVENUE ----
    total_volume = [bobas_vol[i] + gap_vol[i] for i in range(N)]
    total_earned = [bobas_earned[i] + gap_earned[i] for i in range(N)]

    # Add historical AR spillover to collected amounts
    bobas_collected_with_ar = [bobas_collected[i] + HISTORICAL_AR_BOBA[i] for i in range(N)]
    gap_collected_with_ar = [gap_collected[i] + HISTORICAL_AR_GAP[i] for i in range(N)]
    total_collected = [bobas_collected_with_ar[i] + gap_collected_with_ar[i] for i in range(N)]

    # Cash basis: total income = total collected (what actually hits the bank)
    total_income = list(total_collected)

    # ---- BILLING FEES (cash basis) ----
    billing_rate = a.get('billing_fee_rate', 18) / 100.0
    billing_fees = [total_collected[i] * billing_rate for i in range(N)]

    # ---- PAYROLL ----
    _, _, payroll, contractors = forecast_payroll(
        TEAM_ROSTER,
        a.get('payroll_tax_rate', 8.6),
        a.get('salary_annual_increase', 5.0),
    )

    # ---- OPERATING EXPENSES ----
    opex = forecast_opex(a)
    total_opex = [sum(cat[i] for cat in opex.values()) for i in range(N)]

    # ---- EXPANSION COSTS ----
    expansion = forecast_expansion_costs(a)

    # ---- TOTAL OVERHEAD ----
    total_overhead = [
        billing_fees[i] + payroll[i] + contractors[i] + total_opex[i] + expansion['total'][i]
        for i in range(N)
    ]

    # ---- NET EQUITY & FUND FLOW ----
    net_equity = [total_income[i] - total_overhead[i] for i in range(N)]
    physician_rate = a.get('physician_services_rate', 90) / 100.0
    physician_services = [max(0, ne * physician_rate) for ne in net_equity]

    net_income = [net_equity[i] - physician_services[i] for i in range(N)]

    return {
        'bobas_volume': bobas_vol,
        'gap_volume': gap_vol,
        'total_volume': total_volume,
        'bobas_earned': bobas_earned,
        'bobas_collected': bobas_collected,
        'bobas_collected_with_ar': bobas_collected_with_ar,
        'gap_earned': gap_earned,
        'gap_collected': gap_collected,
        'gap_collected_with_ar': gap_collected_with_ar,
        'total_earned': total_earned,
        'total_collected': total_collected,
        'historical_ar_boba': list(HISTORICAL_AR_BOBA),
        'historical_ar_gap': list(HISTORICAL_AR_GAP),
        'total_income': total_income,
        'billing_fees': billing_fees,
        'payroll': payroll,
        'contractors': contractors,
        **{f'opex_{k}': v for k, v in opex.items()},
        'total_opex': total_opex,
        'expansion_ti': expansion['ti'],
        'expansion_lease': expansion['lease'],
        'expansion_ffe': expansion['ffe'],
        'expansion_opex': expansion['opex'],
        'expansion_total': expansion['total'],
        'expansion_details': expansion.get('details', []),
        'locations_open': expansion.get('locations_open', [1] * N),
        'total_overhead': total_overhead,
        'net_equity': net_equity,
        'physician_services': physician_services,
        'net_income': net_income,
    }


def generate_cash_flow_forecast(
    assumptions: Dict = None,
) -> Dict[str, List[float]]:
    """
    60-month cash flow forecast with minimum cash threshold.
    Cash-basis fund flow:
      cash collected → pay overhead → distributable above minimum →
        physician_services_rate% → physician services
        savings_rate% → savings account
    """
    if assumptions is None:
        assumptions = get_default_assumptions()
    a = assumptions

    starting_cash = a.get('starting_cash', 274202.75)
    min_cash = a.get('minimum_cash_balance', 50000.0)
    physician_rate = a.get('physician_services_rate', 90) / 100.0
    savings_rate = a.get('savings_rate', 10) / 100.0
    billing_rate = a.get('billing_fee_rate', 18) / 100.0
    reimb_pct = a.get('reimbursed_expense_pct', 14.6) / 100.0

    # ---- Revenue (collected basis) ----
    bobas_vol = a.get('bobas_volume', [8] * N)
    if len(bobas_vol) < N:
        bobas_vol = bobas_vol + [bobas_vol[-1]] * (N - len(bobas_vol))

    gap_vol = a.get('gap_volume', [4] * N)
    if len(gap_vol) < N:
        gap_vol = gap_vol + [gap_vol[-1]] * (N - len(gap_vol))

    _, bobas_collected = forecast_bobas_revenue(
        bobas_vol,
        a.get('avg_revenue_bobas', 150000),
        a.get('bobas_collection_curve', [0, 0, 0, 0, 0, 5, 10, 15, 25, 25, 15, 5]),
    )
    _, gap_collected = forecast_gap_revenue(
        gap_vol,
        a.get('avg_revenue_gap', 70000),
        a.get('gap_collection_curve', [0, 90, 10]),
    )

    # Add historical AR spillover to collected amounts
    total_collected = [
        bobas_collected[i] + gap_collected[i] + HISTORICAL_AR_BOBA[i] + HISTORICAL_AR_GAP[i]
        for i in range(N)
    ]
    # Cash basis: cash in = total collected
    cash_in = list(total_collected)

    # ---- Cash overhead ----
    billing_fees = [total_collected[i] * billing_rate for i in range(N)]

    _, _, payroll, contractors = forecast_payroll(
        TEAM_ROSTER,
        a.get('payroll_tax_rate', 8.6),
        a.get('salary_annual_increase', 5.0),
    )

    opex = forecast_opex(a)
    total_opex = [sum(cat[i] for cat in opex.values()) for i in range(N)]
    expansion = forecast_expansion_costs(a)

    cash_overhead = [
        billing_fees[i] + payroll[i] + contractors[i] + total_opex[i] + expansion['total'][i]
        for i in range(N)
    ]

    # ---- Month-by-month cash flow with minimum threshold ----
    beginning_cash = [0.0] * N
    cash_after_overhead = [0.0] * N
    distributable = [0.0] * N
    physician = [0.0] * N
    savings_deposit = [0.0] * N
    ending_cash = [0.0] * N
    savings_balance = [0.0] * N

    starting_savings_val = a.get('starting_savings', 0)

    for i in range(N):
        beginning_cash[i] = starting_cash if i == 0 else ending_cash[i - 1]
        cash_after_overhead[i] = beginning_cash[i] + cash_in[i] - cash_overhead[i]
        distributable[i] = max(0, cash_after_overhead[i] - min_cash)
        physician[i] = distributable[i] * physician_rate
        savings_deposit[i] = distributable[i] * savings_rate
        ending_cash[i] = cash_after_overhead[i] - physician[i] - savings_deposit[i]

        prior_savings = starting_savings_val if i == 0 else savings_balance[i - 1]
        savings_balance[i] = prior_savings + savings_deposit[i]

    return {
        'beginning_cash': beginning_cash,
        'cash_in': cash_in,
        'total_collected': total_collected,
        'cash_overhead': cash_overhead,
        'billing_fees': billing_fees,
        'payroll': payroll,
        'total_opex': total_opex,
        'expansion_total': expansion['total'],
        'cash_after_overhead': cash_after_overhead,
        'distributable': distributable,
        'physician': physician,
        'savings_deposit': savings_deposit,
        'ending_cash': ending_cash,
        'savings_balance': savings_balance,
        'starting_cash': starting_cash,
        'minimum_cash_balance': min_cash,
    }


def calculate_dashboard_metrics(assumptions: Dict = None) -> Dict:
    """
    Return year-by-year metrics for the dashboard: 2025 | 2026 | 2027 | 2028 | 2029 | 2030.
    """
    if assumptions is None:
        assumptions = get_default_assumptions()

    pl = generate_monthly_pl_forecast(assumptions)
    cf = generate_cash_flow_forecast(assumptions)

    def _year_slice(data, start, end):
        return sum(data[start:end])

    def _year_min(data, start, end):
        return min(data[start:end])

    def _year_max(data, start, end):
        return max(data[start:end])

    # 2025 from actuals
    yr2025 = {
        'surgeries_bobas': 0,  # not tracked separately in 2025
        'surgeries_gap': 0,
        'surgeries_total': TOTAL_SURGERIES_2025,  # 116
        'gross_revenue': ACTUALS_2025_TOTALS['total_income'],
        'physician_services': ACTUALS_2025_TOTALS['physician_services'],
        'savings_deposits': 0,
        'locations': 1,
        'min_cash': min(BALANCE_SHEET_2025['total_cash']),
        'max_cash': max(BALANCE_SHEET_2025['total_cash']),
        'capex_startup': sum(BALANCE_SHEET_2025['furniture_fixtures'][i] + BALANCE_SHEET_2025['machinery_equipment'][i] for i in range(12) if i == 9) + 26108.20,
        'ending_savings': 0,
    }

    results = {'2025': yr2025}

    # 2026-2030: each year is a 12-month slice
    for yr_idx, year in enumerate(range(2026, 2031)):
        s = yr_idx * 12       # start month index
        e = s + 12             # end month index

        yr_data = {
            'surgeries_bobas': sum(pl['bobas_volume'][s:e]),
            'surgeries_gap': sum(pl['gap_volume'][s:e]),
            'surgeries_total': sum(pl['total_volume'][s:e]),
            'gross_revenue': _year_slice(pl['total_income'], s, e),
            'physician_services': _year_slice(pl['physician_services'], s, e),
            'savings_deposits': _year_slice(cf['savings_deposit'], s, e),
            'locations': max(pl['locations_open'][s:e]),
            'min_cash': _year_min(cf['ending_cash'], s, e),
            'max_cash': _year_max(cf['ending_cash'], s, e),
            'capex_startup': _year_slice(pl['expansion_ti'], s, e) + _year_slice(pl['expansion_ffe'], s, e),
            'ending_savings': cf['savings_balance'][e - 1],
        }
        results[str(year)] = yr_data

    return results


def normalize_2025_pl() -> Dict[str, List[float]]:
    """Normalized 2025 P&L: strips VNSC mgmt fee and Aug rent anomaly."""
    actual_income = [
        ACTUALS_2025['fee_income'][i] +
        ACTUALS_2025['reimbursed_expense_income'][i] +
        ACTUALS_2025['refunds'][i]
        for i in range(12)
    ]

    expense_keys = [
        'advertising_marketing', 'bank_fees', 'conference', 'contracts',
        'contributions', 'dues_subscriptions', 'insurance', 'malpractice_insurance',
        'health_insurance', 'legal_accounting_services', 'accounting_fees',
        'billing_services', 'licenses_fees', 'mgmt_fee_abc', 'mgmt_fee_vnsc',
        'meals', 'office_expenses', 'software_apps', 'payroll_processing',
        'salaries_wages', 'payroll_taxes', 'physician_services', 'rent_lease',
        'taxes_licenses', 'travel',
    ]
    actual_expenses = [sum(ACTUALS_2025[k][i] for k in expense_keys) for i in range(12)]

    adj_expenses = []
    for i in range(12):
        adj = actual_expenses[i]
        adj -= ACTUALS_2025['mgmt_fee_vnsc'][i]
        if i == 7:
            adj -= ACTUALS_2025['rent_lease'][i]
            adj += 2500
        adj_expenses.append(adj)

    return {
        'income': actual_income,
        'expenses_actual': actual_expenses,
        'expenses_normalized': adj_expenses,
        'noi_actual': [actual_income[i] - actual_expenses[i] for i in range(12)],
        'noi_normalized': [actual_income[i] - adj_expenses[i] for i in range(12)],
        'adjustments': {
            'vnsc_removed': sum(ACTUALS_2025['mgmt_fee_vnsc']),
            'rent_anomaly_removed': ACTUALS_2025['rent_lease'][7] - 2500,
        }
    }
