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
            # Convert 1-based month to 0-based array index
            start_idx = start - 1
            end = person.get('end_month', None)
            end_idx = (end - 1) if end is not None else None
            if month < start_idx:
                continue
            if end_idx is not None and month > end_idx:
                continue
            base_salary = person['monthly_salary'] * escalator
            if end_idx is not None and month == end_idx:
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
        # Convert 1-based months to 0-based array indices
        ti_start = exp.get('ti_start_month', 1) - 1
        ti_dur = exp.get('ti_duration_months', 2)
        ti_amount = exp.get('ti_cns_share', 0)
        if ti_dur > 0 and ti_amount > 0:
            per_month = ti_amount / ti_dur
            for m in range(max(0, ti_start), min(ti_start + ti_dur, N)):
                d['ti'][m] = per_month

        # Lease (with inflation)
        lease_start = exp.get('lease_start_month', 1) - 1
        lease_amt = exp.get('lease_monthly', 0)
        for m in range(max(0, lease_start), N):
            escalator = (1 + inflation_rate) ** (m // 12)
            d['lease'][m] = lease_amt * escalator

        # FF&E (one-time, no inflation)
        ffe = exp.get('ffe_budget', 0)
        if 0 <= lease_start < N and ffe > 0:
            d['ffe'][lease_start] = ffe

        # Ongoing opex (with inflation and optional ramp)
        opex_full = exp.get('opex_monthly', 0)
        opex_ramp = exp.get('opex_ramp_monthly', opex_full)
        ramp_months = exp.get('opex_ramp_months', 0)
        for m in range(max(0, lease_start), N):
            escalator = (1 + inflation_rate) ** (m // 12)
            months_since_start = m - lease_start
            if ramp_months > 0 and months_since_start < ramp_months:
                d['opex'][m] = opex_ramp * escalator
            else:
                d['opex'][m] = opex_full * escalator

        # Location count
        for m in range(max(0, lease_start), N):
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


def generate_pl_by_location(assumptions: Dict = None) -> Dict[str, Dict]:
    """
    Per-location P&L with consolidated rollup.

    Returns:
        {
            'Westlake': {per-location P&L dict},
            'Santa Barbara': {per-location P&L dict},
            'consolidated': {aggregated P&L dict},
        }
    """
    from baseline_data import HISTORICAL_AR_BOBA, HISTORICAL_AR_GAP

    a = assumptions or get_default_assumptions()
    locations = a.get('locations', ['Westlake'])
    volumes_by_loc = a.get('volumes_by_location', {})
    opex_by_loc = a.get('opex_by_location', {})
    shared_overhead = a.get('shared_overhead', {})
    team = a.get('team_roster', TEAM_ROSTER)

    avg_rev_bobas = a.get('avg_revenue_bobas', 132000)
    avg_rev_gap = a.get('avg_revenue_gap', 50000)
    bobas_curve = a.get('bobas_collection_curve', [0, 0, 3, 0, 0, 27, 10, 32, 23, 0, 5, 0])
    gap_curve = a.get('gap_collection_curve', [0, 55, 10, 10, 25])
    billing_rate = a.get('billing_fee_rate', 18) / 100
    inflation = a.get('expense_annual_inflation', 3.0) / 100
    payroll_tax_rate = a.get('payroll_tax_rate', 8.6)
    salary_increase = a.get('salary_annual_increase', 5.0)
    surgeon_comp = a.get('surgeon_compensation', {})

    result = {}

    # --- Per-location calculations ---
    for loc in locations:
        loc_vols = volumes_by_loc.get(loc, {'bobas': [0]*N, 'gap': [0]*N})
        bobas_vol = loc_vols.get('bobas', [0]*N)
        gap_vol = loc_vols.get('gap', [0]*N)

        # Revenue
        bobas_earned, bobas_collected = forecast_bobas_revenue(
            bobas_vol, avg_rev_bobas, bobas_curve)
        gap_earned, gap_collected = forecast_gap_revenue(
            gap_vol, avg_rev_gap, gap_curve)

        total_volume = [bobas_vol[i] + gap_vol[i] for i in range(N)]
        total_earned = [bobas_earned[i] + gap_earned[i] for i in range(N)]
        total_collected = [bobas_collected[i] + gap_collected[i] for i in range(N)]

        # Billing
        billing = [total_collected[i] * billing_rate for i in range(N)]

        # Payroll (filtered by location)
        loc_team = [p for p in team if p.get('location') == loc]
        if loc_team:
            _, _, payroll_total, contractors = forecast_payroll(
                loc_team, payroll_tax_rate, salary_increase)
        else:
            payroll_total = [0.0] * N
            contractors = [0.0] * N

        # Direct OpEx (location-specific)
        loc_opex = opex_by_loc.get(loc, {})
        def _escalated(base_val):
            return [base_val * (1 + inflation) ** (i // 12) for i in range(N)]

        direct_opex = [0.0] * N
        opex_detail = {}
        for key, val in loc_opex.items():
            escalated = _escalated(val)
            opex_detail[key] = escalated
            for i in range(N):
                direct_opex[i] += escalated[i]

        # Expansion costs for this location (from expansion config)
        loc_expansion = [0.0] * N
        for exp in a.get('expansions', []):
            if exp.get('name') == loc and exp.get('enabled'):
                exp_result = forecast_expansion_costs({'expansions': [exp],
                                                       'expense_annual_inflation': a.get('expense_annual_inflation', 3.0)})
                loc_expansion = exp_result.get('total', [0.0] * N)

        # Total direct overhead (before shared allocation)
        direct_overhead = [
            billing[i] + payroll_total[i] + contractors[i] + direct_opex[i] + loc_expansion[i]
            for i in range(N)
        ]

        # Surgeon compensation at this location
        loc_surgeon = surgeon_comp.get(loc, {})
        surgeon_rate = loc_surgeon.get('rate', 0) / 100
        surgeon_pay = [total_collected[i] * surgeon_rate for i in range(N)]

        # Contribution (before shared overhead allocation)
        contribution = [total_collected[i] - direct_overhead[i] - surgeon_pay[i] for i in range(N)]

        result[loc] = {
            'bobas_volume': bobas_vol,
            'gap_volume': gap_vol,
            'total_volume': total_volume,
            'bobas_earned': bobas_earned,
            'gap_earned': gap_earned,
            'total_earned': total_earned,
            'bobas_collected': bobas_collected,
            'gap_collected': gap_collected,
            'total_collected': total_collected,
            'billing_fees': billing,
            'payroll': payroll_total,
            'contractors': contractors,
            'direct_opex': direct_opex,
            'opex_detail': opex_detail,
            'expansion_costs': loc_expansion,
            'direct_overhead': direct_overhead,
            'surgeon_compensation': surgeon_pay,
            'surgeon_rate': surgeon_rate * 100,
            'contribution': contribution,
        }

    # --- Shared overhead allocation (by revenue %) ---
    # Calculate each location's share of total collections
    total_all_collected = [0.0] * N
    for loc in locations:
        for i in range(N):
            total_all_collected[i] += result[loc]['total_collected'][i]

    # Shared overhead pool
    shared_pool = [0.0] * N
    malpractice_monthly = shared_overhead.get('malpractice_annual', 0) / 12
    shared_monthly = (
        shared_overhead.get('legal_monthly_recurring', 0)
        + shared_overhead.get('general_insurance_monthly', 0)
        + shared_overhead.get('health_insurance_monthly', 0)
        + shared_overhead.get('mgmt_fee_abc_monthly', 0)
        + shared_overhead.get('bank_fees_monthly', 0)
        + malpractice_monthly
    )
    for i in range(N):
        shared_pool[i] = shared_monthly * (1 + inflation) ** (i // 12)

    # Allocate by revenue %
    for loc in locations:
        allocation = [0.0] * N
        for i in range(N):
            if total_all_collected[i] > 0:
                share = result[loc]['total_collected'][i] / total_all_collected[i]
            else:
                share = 1.0 / len(locations)
            allocation[i] = shared_pool[i] * share
        result[loc]['shared_overhead_allocation'] = allocation
        result[loc]['total_overhead'] = [
            result[loc]['direct_overhead'][i] + allocation[i]
            for i in range(N)
        ]
        # Update contribution to include shared overhead
        result[loc]['contribution'] = [
            result[loc]['total_collected'][i]
            - result[loc]['total_overhead'][i]
            - result[loc]['surgeon_compensation'][i]
            for i in range(N)
        ]

    # --- Consolidated ---
    cons = {k: [0.0] * N for k in [
        'bobas_volume', 'gap_volume', 'total_volume',
        'total_earned', 'total_collected',
        'billing_fees', 'payroll', 'contractors',
        'direct_opex', 'expansion_costs', 'total_overhead',
        'surgeon_compensation', 'contribution',
    ]}
    for loc in locations:
        for key in cons:
            if key in result[loc]:
                for i in range(N):
                    cons[key][i] += result[loc][key][i]

    # Historical AR (consolidated — all historical cases are Westlake)
    historical_ar = [HISTORICAL_AR_BOBA[i] + HISTORICAL_AR_GAP[i] for i in range(N)]
    cons['historical_ar'] = historical_ar
    cons['total_income'] = [cons['total_collected'][i] + historical_ar[i] for i in range(N)]

    # Fund flow at consolidated level
    phys_rate = a.get('physician_services_rate', 90) / 100
    net_equity = [cons['total_income'][i] - cons['total_overhead'][i] - cons['surgeon_compensation'][i] for i in range(N)]
    physician = [max(0, net_equity[i] * phys_rate) for i in range(N)]
    net_income = [net_equity[i] - physician[i] for i in range(N)]

    cons['net_equity'] = net_equity
    cons['physician_services'] = physician
    cons['net_income'] = net_income
    cons['shared_overhead_pool'] = shared_pool

    result['consolidated'] = cons
    return result


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
