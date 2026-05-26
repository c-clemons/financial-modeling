"""
Mighty Pilates Financial Model Builder - FORMULA-DRIVEN

Mirrors the Streamlit dashboard at github.com/c-clemons/mighty-pilates.
All calculated cells use Excel formulas referencing the Assumptions tab.
Yellow cells are the only hard-coded inputs.

Coverage: 2026 (Jan-Apr actuals + May-Dec forecast), 2027, 2028 forecast.
"""

import json
from pathlib import Path
from datetime import datetime

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side,
)
from openpyxl.utils import get_column_letter

# ============================================================
# PATHS
# ============================================================
DASH_DATA = Path("/Users/chandlerclemons/mighty-pilates/dashboard/data")
OUTPUT_PATH = Path(
    "/Users/chandlerclemons/Desktop/Empirica Financial Modeling/Mighty Pilates/"
    "Mighty Pilates Financial Model.xlsx"
)

# ============================================================
# DATA LOADERS
# ============================================================

def load_dashboard_data():
    """Load all source JSONs from the Streamlit dashboard."""
    with open(DASH_DATA / "committed_actuals.json") as f:
        committed = json.load(f)
    with open(DASH_DATA / "baseline.json") as f:
        baseline = json.load(f)
    try:
        with open(DASH_DATA / "user_overrides.json") as f:
            overrides = json.load(f)
    except FileNotFoundError:
        overrides = {}

    # Merge baseline + overrides (overrides win)
    def deep_merge(a, b):
        if isinstance(a, dict) and isinstance(b, dict):
            out = dict(a)
            for k, v in b.items():
                out[k] = deep_merge(a.get(k), v) if k in a else v
            return out
        return b if b is not None else a

    merged = deep_merge(baseline, overrides)

    return {
        "committed": committed,
        "baseline": baseline,
        "overrides": overrides,
        "merged": merged,
    }


# ============================================================
# CONSTANTS
# ============================================================

# 12 active studios (matches Streamlit ACTIVE_STUDIOS)
STUDIOS = [
    ("BK", "Berkeley"),
    ("CC", "Culver City"),
    ("DN", "Danville"),
    ("LF", "Lafayette"),
    ("MR", "Marin"),
    ("OP", "Ocean Park"),
    ("PH", "Presidio Heights"),
    ("RH", "Russian Hill"),
    ("SB", "Santa Barbara"),
    ("SM", "Santa Monica"),
    ("WP", "West Portal"),
    ("WW", "Westwood"),
]

# Development / not-yet-open studios (used in baseline data but excluded from active P&L)
DEV_STUDIOS = [
    ("CDM", "Corona Del Mar"),
    ("PS", "Pasadena"),
]

# OpEx categories (mirrors Streamlit category buckets)
OPEX_CATEGORIES = [
    ("property", "Property Costs", "rent_lease"),
    ("staff", "Staff Costs", "staff"),
    ("utilities", "Utilities", "operating"),
    ("marketing", "Marketing & Promotion", "operating"),
    ("admin", "Administrative & G&A", "operating"),
    ("professional", "Professional Fees", "operating"),
    ("travel", "Travel & Meals", "operating"),
    ("cogs", "Merchant Fees & COGS", "operating"),
    ("startup", "Studio Start Up Costs", "operating"),
    ("taxes", "Taxes", "operating"),
]

# Months: Jan 2026 through Dec 2028 (36 months)
MONTH_LABELS = []
MONTH_KEYS = []
MONTH_DATES = []
for year in (2026, 2027, 2028):
    for mo in range(1, 13):
        MONTH_LABELS.append(f"{['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][mo-1]} {year}")
        MONTH_KEYS.append(f"{year}-{mo:02d}")
        MONTH_DATES.append(datetime(year, mo, 1))

LAST_ACTUALS_MONTH = "2026-04"  # Apr 2026
LAST_ACTUALS_IDX = MONTH_KEYS.index(LAST_ACTUALS_MONTH)  # 3
FIRST_FORECAST_IDX = LAST_ACTUALS_IDX + 1  # 4 (May 2026)

# ============================================================
# STYLES
# ============================================================
NAVY = "1B2A4A"
ACCENT_BLUE = "4472C4"
LIGHT_BLUE = "D6E4F0"
LIGHT_GRAY = "F2F2F2"
MED_GRAY = "D9D9D9"
GREEN_FILL_C = "E2EFDA"
RED_FILL_C = "FCE4EC"
YELLOW_FILL_C = "FFF9E6"
INPUT_FILL_C = "FFFFCC"
ACTUAL_FILL_C = "E8F4E8"  # very light green for actuals columns

title_font = Font(name="Calibri", size=16, bold=True, color=NAVY)
section_font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
subsection_font = Font(name="Calibri", size=11, bold=True, color=NAVY)
header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
data_font = Font(name="Calibri", size=10)
data_bold = Font(name="Calibri", size=10, bold=True)
input_font = Font(name="Calibri", size=10, color="0000CC", bold=False)
metric_value_font = Font(name="Calibri", size=14, bold=True, color=NAVY)
metric_label_font = Font(name="Calibri", size=9, color="666666")

section_fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
header_fill = PatternFill(start_color=ACCENT_BLUE, end_color=ACCENT_BLUE, fill_type="solid")
light_fill = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
alt_fill = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type="solid")
input_fill = PatternFill(start_color=INPUT_FILL_C, end_color=INPUT_FILL_C, fill_type="solid")
actual_fill = PatternFill(start_color=ACTUAL_FILL_C, end_color=ACTUAL_FILL_C, fill_type="solid")
green_fill = PatternFill(start_color=GREEN_FILL_C, end_color=GREEN_FILL_C, fill_type="solid")

thin_border = Border(
    left=Side(style="thin", color=MED_GRAY), right=Side(style="thin", color=MED_GRAY),
    top=Side(style="thin", color=MED_GRAY), bottom=Side(style="thin", color=MED_GRAY),
)
double_bottom = Border(bottom=Side(style="double", color=NAVY))

CURR = '#,##0;[Red](#,##0);"-"'
CURR2 = '#,##0.00'
PCT = '0.0%'
PCT2 = '0.00%'
NUM = '#,##0'
center_align = Alignment(horizontal="center", vertical="center")
right_align = Alignment(horizontal="right", vertical="center")
left_align = Alignment(horizontal="left", vertical="center")


# ============================================================
# HELPERS
# ============================================================

def style_range(ws, row, c1, c2, font=None, fill=None, border=None, alignment=None, number_format=None):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row=row, column=c)
        if font: cell.font = font
        if fill: cell.fill = fill
        if border: cell.border = border
        if alignment: cell.alignment = alignment
        if number_format: cell.number_format = number_format


def section_bar(ws, row, c1, c2, label):
    for c in range(c1, c2 + 1):
        ws.cell(row=row, column=c).fill = section_fill
    cell = ws.cell(row=row, column=c1, value=label)
    cell.font = section_font
    cell.alignment = left_align


def header_row(ws, row, c1, values, fill=None):
    use_fill = fill or header_fill
    for i, v in enumerate(values):
        cell = ws.cell(row=row, column=c1 + i, value=v)
        cell.font = header_font
        cell.fill = use_fill
        cell.alignment = center_align


def input_cell(ws, row, col, value, fmt=CURR):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = input_font
    cell.fill = input_fill
    cell.border = thin_border
    cell.number_format = fmt
    return cell


def formula_cell(ws, row, col, formula, fmt=CURR, bold=False, fill=None):
    cell = ws.cell(row=row, column=col, value=formula)
    cell.font = data_bold if bold else data_font
    if fill:
        cell.fill = fill
    cell.number_format = fmt
    cell.border = thin_border
    return cell


def actual_value_cell(ws, row, col, value, fmt=CURR, bold=False):
    """Hard-coded actuals (closed-month historical data)."""
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = data_bold if bold else data_font
    cell.fill = actual_fill
    cell.border = thin_border
    cell.number_format = fmt
    return cell


def label_cell(ws, row, col, text, bold=False, indent=0):
    cell = ws.cell(row=row, column=col, value=text)
    cell.font = data_bold if bold else data_font
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=indent)
    return cell


# ============================================================
# TAB BUILDERS (stubs - to be implemented)
# ============================================================

def build_cover(wb, data):
    """Cover / instructions tab."""
    ws = wb.active
    ws.title = "Cover"
    ws.sheet_view.showGridLines = False

    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 90

    ws["B2"] = "Mighty Pilates"
    ws["B2"].font = Font(name="Calibri", size=24, bold=True, color=NAVY)
    ws["B3"] = "Financial Model"
    ws["B3"].font = Font(name="Calibri", size=18, color=NAVY)
    ws["B5"] = f"Generated: {datetime.now().strftime('%B %d, %Y')}"
    ws["B5"].font = Font(name="Calibri", size=10, italic=True, color="666666")

    ws["B7"] = "Tab Guide"
    ws["B7"].font = subsection_font

    tabs = [
        ("Summary & Controls", "Top-level KPIs, studio sales & EBITDA tables, charts"),
        ("Assumptions", "All model inputs: OpEx, sales forecast, rev rec curves, loan terms"),
        ("Cash Flow Forecast", "Monthly cash flow from operations + investing + financing"),
        ("P&L Summary", "Consolidated monthly P&L (Revenue → EBITDA)"),
        ("P&L Detail", "Account-level P&L matching accountant chart of accounts"),
        ("All Studios Summary", "Side-by-side studio comparison"),
        ("[Studio] FCST (×12)", "Per-studio P&L forecast (BK, CC, DN, LF, MR, OP, PH, RH, SB, SM, WP, WW)"),
        ("Cash, Debt & Equity", "Balance sheet, loan amortization schedules, equity tracking"),
        ("Sales Forecast", "Studio × month editable sales grid"),
        ("CapEx", "Capital expenditure project schedule"),
        ("QBO Actuals", "Historical accountant-booked P&L, BS, SCF (read-only)"),
    ]
    r = 9
    for tab, desc in tabs:
        ws.cell(row=r, column=2, value=tab).font = data_bold
        ws.cell(row=r, column=3, value=desc).font = data_font
        r += 1

    ws.column_dimensions["C"].width = 70

    ws["B" + str(r + 2)] = "Conventions"
    ws["B" + str(r + 2)].font = subsection_font
    ws.cell(row=r + 4, column=2, value="Yellow cells = inputs (edit these)").fill = input_fill
    ws.cell(row=r + 5, column=2, value="Light green cells = actuals (closed months, do not edit)").fill = actual_fill


def build_assumptions(wb, data):
    """
    Assumptions tab — global model parameters, rev rec curves, loan terms.

    Sections:
      1. Model Control (last actuals month, forecast horizon)
      2. Revenue Recognition Curves (earned/breakage by month lag)
      3. Refund / Discount / Merchant / COGS rates
      4. Annual Escalation Rates (rent vs other OpEx, sales growth)
      5. Loan Terms
      6. Studio List (active studios)
    """
    ws = wb.create_sheet("Assumptions")
    ws.sheet_view.showGridLines = False

    # Column widths
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 14
    ws.column_dimensions["H"].width = 14
    ws.column_dimensions["I"].width = 14
    ws.column_dimensions["J"].width = 18

    # Title
    ws["B2"] = "Assumptions"
    ws["B2"].font = title_font
    ws["B3"] = "Yellow cells are editable inputs. All other tabs reference these values."
    ws["B3"].font = Font(name="Calibri", size=10, italic=True, color="666666")

    row = 5

    # Initialize refs dict for cross-sheet references
    refs = {}

    # ----- Section 1: Model Control -----
    section_bar(ws, row, 2, 9, "1. Model Control")
    row += 1
    label_cell(ws, row, 2, "Last Actuals Month")
    input_cell(ws, row, 3, "Apr 2026", fmt="@")
    refs["last_actuals_row"] = row
    row += 1
    label_cell(ws, row, 2, "Forecast Horizon (months)")
    input_cell(ws, row, 3, 32, fmt=NUM)
    refs["horizon_row"] = row
    row += 1
    label_cell(ws, row, 2, "Forecast End Month")
    input_cell(ws, row, 3, "Dec 2028", fmt="@")
    refs["end_row"] = row
    row += 2

    # ----- Section 2: Revenue Recognition Curves -----
    section_bar(ws, row, 2, 9, "2. Revenue Recognition Curves")
    row += 1
    label_cell(ws, row, 2, "Lag Month (0 = sale month)", bold=True)
    for i in range(7):
        cell = ws.cell(row=row, column=3 + i, value=i)
        cell.font = data_bold
        cell.fill = light_fill
        cell.alignment = center_align
    label_cell(ws, row, 10, "Sum", bold=True)
    ws.cell(row=row, column=10).fill = light_fill
    ws.cell(row=row, column=10).alignment = center_align
    row += 1

    # Earned curve
    label_cell(ws, row, 2, "Earned %")
    curves = data["committed"]["rev_rec_curves"]
    earned = curves["earned"]
    refs["earned_row"] = row
    for i in range(7):
        input_cell(ws, row, 3 + i, float(earned.get(str(i), 0)) / 100.0, fmt=PCT)
    formula_cell(ws, row, 10, f"=SUM(C{row}:I{row})", fmt=PCT, bold=True, fill=light_fill)
    row += 1

    # Breakage curve
    label_cell(ws, row, 2, "Breakage %")
    breakage = curves["breakage"]
    refs["breakage_row"] = row
    for i in range(7):
        input_cell(ws, row, 3 + i, float(breakage.get(str(i), 0)) / 100.0, fmt=PCT)
    formula_cell(ws, row, 10, f"=SUM(C{row}:I{row})", fmt=PCT, bold=True, fill=light_fill)
    row += 1

    # Combined total
    label_cell(ws, row, 2, "Total (Earned + Breakage)", bold=True)
    for i in range(7):
        col = get_column_letter(3 + i)
        formula_cell(ws, row, 3 + i,
                     f"={col}{refs['earned_row']}+{col}{refs['breakage_row']}",
                     fmt=PCT, bold=True, fill=light_fill)
    formula_cell(ws, row, 10, f"=SUM(C{row}:I{row})", fmt=PCT, bold=True, fill=light_fill)
    row += 2

    # ----- Section 3: Rates -----
    section_bar(ws, row, 2, 9, "3. Rates (% of Gross Revenue)")
    row += 1
    rate_specs = [
        ("refund_row", "Refund Rate", float(curves["refund_pct"]) / 100.0,
         "Refunds as % of gross revenue. Negative reduces net revenue."),
        ("discount_row", "Discount Rate", float(curves["discount_pct"]) / 100.0,
         "Discounts as % of gross revenue. Negative reduces net revenue."),
        ("merchant_row", "Merchant Fee %",
         float(data["committed"]["forecast_ratios"]["merchant_fee_pct"]) / 100.0,
         "Payment processor fees as % of revenue."),
        ("cogs_row", "COGS %",
         float(data["committed"]["forecast_ratios"]["cogs_pct"]) / 100.0,
         "Cost of goods sold as % of revenue (retail)."),
    ]
    for key, label, val, note in rate_specs:
        refs[key] = row
        label_cell(ws, row, 2, label)
        input_cell(ws, row, 3, val, fmt=PCT2)
        cell = ws.cell(row=row, column=4, value=note)
        cell.font = Font(name="Calibri", size=9, italic=True, color="666666")
        cell.alignment = left_align
        ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=9)
        row += 1
    row += 1

    # ----- Section 4: Annual Escalation Rates -----
    section_bar(ws, row, 2, 9, "4. Annual Escalation Rates")
    row += 1
    esc_specs = [
        ("rent_esc_row", "Rent / Property Escalation", 0.03,
         "3% annual rent step-up applied to forecast months"),
        ("other_esc_row", "Other OpEx Escalation", 0.04,
         "4% annual escalation for all non-rent OpEx categories"),
        ("sales_growth_row", "Sales YoY Growth", 0.05,
         "5% YoY assumed for 2028 (applied to 2027 seasonal pattern)"),
    ]
    for key, label, val, note in esc_specs:
        refs[key] = row
        label_cell(ws, row, 2, label)
        input_cell(ws, row, 3, val, fmt=PCT)
        cell = ws.cell(row=row, column=4, value=note)
        cell.font = Font(name="Calibri", size=9, italic=True, color="666666")
        cell.alignment = left_align
        ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=9)
        row += 1
    row += 1

    # ----- Section 5: Loan Terms -----
    section_bar(ws, row, 2, 9, "5. Loan Terms (as of Apr 2026)")
    row += 1
    header_row(ws, row, 2, ["Loan", "Original Amt", "Current Balance", "Rate (Annual)",
                             "Avg Monthly Pmt", "Start Date", "Type"])
    row += 1
    loans = data["baseline"].get("loans", [])
    refs["loan_start_row"] = row
    refs["loans"] = []
    for loan in loans:
        loan_row = row
        label_cell(ws, row, 2, loan.get("name", loan.get("id", "?")), bold=True)
        input_cell(ws, row, 3, float(loan.get("original_amount", 0)), fmt=CURR)
        input_cell(ws, row, 4, float(loan.get("current_balance", 0)), fmt=CURR)
        input_cell(ws, row, 5, float(loan.get("rate", 0)), fmt=PCT2)
        input_cell(ws, row, 6, float(loan.get("avg_monthly_payment", 0)), fmt=CURR)
        input_cell(ws, row, 7, str(loan.get("start_date", "")), fmt="@")
        loan_type = "Interest-Only" if loan.get("avg_monthly_payment", 0) == 0 else "Amortizing"
        label_cell(ws, row, 8, loan_type)
        refs["loans"].append({
            "name": loan.get("name", "?"),
            "row": loan_row,
            "is_interest_only": loan.get("avg_monthly_payment", 0) == 0,
        })
        row += 1
    refs["loan_end_row"] = row - 1
    # Totals row
    label_cell(ws, row, 2, "TOTAL DEBT", bold=True)
    formula_cell(ws, row, 4,
                 f"=SUM(D{refs['loan_start_row']}:D{refs['loan_end_row']})",
                 fmt=CURR, bold=True, fill=light_fill)
    formula_cell(ws, row, 6,
                 f"=SUM(F{refs['loan_start_row']}:F{refs['loan_end_row']})",
                 fmt=CURR, bold=True, fill=light_fill)
    refs["loan_total_row"] = row
    row += 2

    # ----- Section 6: Active Studios -----
    section_bar(ws, row, 2, 9, "6. Active Studios")
    row += 1
    header_row(ws, row, 2, ["Code", "Name"])
    row += 1
    for code, name in STUDIOS:
        label_cell(ws, row, 2, code, bold=True)
        label_cell(ws, row, 3, name)
        row += 1

    # Stash refs on workbook for downstream tabs
    if not hasattr(wb, "_mighty_refs"):
        wb._mighty_refs = {}
    wb._mighty_refs["assumptions"] = refs

    return refs


def build_sales_forecast(wb, data):
    """
    Sales Forecast tab: studio × month grid.

    Layout:
      - Col B = Studio code, Col C = Studio name
      - Cols D... = 36 months (Jan 2026 - Dec 2028)
      - First 4 months (Jan-Apr 2026) are actuals (light green, locked)
      - May 2026 onwards are yellow editable forecast cells
      - Bottom row = TOTAL per month (SUM formula)
      - Right side = annual totals + YoY %
    """
    ws = wb.create_sheet("Sales Forecast")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "D4"

    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 22
    for i in range(36):
        ws.column_dimensions[get_column_letter(4 + i)].width = 11

    # Title
    ws["B2"] = "Sales Forecast (Gross)"
    ws["B2"].font = title_font
    ws["B3"] = "Jan-Apr 2026 actuals (locked). May 2026 onwards forecast (editable)."
    ws["B3"].font = Font(name="Calibri", size=10, italic=True, color="666666")

    HEADER_ROW = 5
    # Section bar
    section_bar(ws, HEADER_ROW - 1, 2, 39, "Per-Studio Monthly Sales")

    # Month headers
    label_cell(ws, HEADER_ROW, 2, "Code", bold=True)
    label_cell(ws, HEADER_ROW, 3, "Studio", bold=True)
    ws.cell(row=HEADER_ROW, column=2).fill = header_fill
    ws.cell(row=HEADER_ROW, column=2).font = header_font
    ws.cell(row=HEADER_ROW, column=3).fill = header_fill
    ws.cell(row=HEADER_ROW, column=3).font = header_font
    for i, label in enumerate(MONTH_LABELS):
        cell = ws.cell(row=HEADER_ROW, column=4 + i, value=label)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align

    # Data rows per studio
    csf = data["committed"]["client_sales_forecast"]
    refs = {"studios": {}, "header_row": HEADER_ROW, "first_data_col": 4}

    row = HEADER_ROW + 1
    for code, name in STUDIOS:
        label_cell(ws, row, 2, code, bold=True)
        label_cell(ws, row, 3, name)
        studio_data = csf.get(code, {})
        for i, mk in enumerate(MONTH_KEYS):
            val = float(studio_data.get(mk, 0))
            if i <= LAST_ACTUALS_IDX:
                # Actuals (Jan-Apr 2026)
                actual_value_cell(ws, row, 4 + i, val, fmt=CURR)
            else:
                # Forecast (editable)
                input_cell(ws, row, 4 + i, val, fmt=CURR)
        refs["studios"][code] = row
        row += 1

    # TOTAL row
    TOTAL_ROW = row
    refs["total_row"] = TOTAL_ROW
    label_cell(ws, row, 2, "TOTAL", bold=True)
    label_cell(ws, row, 3, "All Studios", bold=True)
    for c in range(2, 4):
        ws.cell(row=row, column=c).fill = light_fill
    first_studio_row = HEADER_ROW + 1
    for i in range(36):
        col = get_column_letter(4 + i)
        formula_cell(ws, row, 4 + i,
                     f"=SUM({col}{first_studio_row}:{col}{row - 1})",
                     fmt=CURR, bold=True, fill=light_fill)
    row += 2

    # Annual totals section
    section_bar(ws, row, 2, 8, "Annual Summary")
    row += 1
    header_row(ws, row, 2, ["Code", "Studio", "2026", "2027", "2028", "2027 YoY", "2028 YoY"])
    row += 1
    refs["annual_summary_start"] = row
    for code, name in STUDIOS:
        s_row = refs["studios"][code]
        label_cell(ws, row, 2, code, bold=True)
        label_cell(ws, row, 3, name)
        # 2026 = sum of cols D:O (months 1-12)
        formula_cell(ws, row, 4, f"=SUM(D{s_row}:O{s_row})", fmt=CURR)
        # 2027 = cols P:AA (months 13-24)
        formula_cell(ws, row, 5, f"=SUM(P{s_row}:AA{s_row})", fmt=CURR)
        # 2028 = cols AB:AM (months 25-36)
        formula_cell(ws, row, 6, f"=SUM(AB{s_row}:AM{s_row})", fmt=CURR)
        # 2027 YoY % = (2027-2026)/2026
        formula_cell(ws, row, 7,
                     f"=IFERROR((E{row}-D{row})/D{row},0)",
                     fmt=PCT)
        formula_cell(ws, row, 8,
                     f"=IFERROR((F{row}-E{row})/E{row},0)",
                     fmt=PCT)
        row += 1
    # Total annual row
    refs["annual_total_row"] = row
    label_cell(ws, row, 2, "TOTAL", bold=True)
    label_cell(ws, row, 3, "All Studios", bold=True)
    for c in range(2, 4):
        ws.cell(row=row, column=c).fill = light_fill
    formula_cell(ws, row, 4,
                 f"=SUM(D{refs['annual_summary_start']}:D{row - 1})",
                 fmt=CURR, bold=True, fill=light_fill)
    formula_cell(ws, row, 5,
                 f"=SUM(E{refs['annual_summary_start']}:E{row - 1})",
                 fmt=CURR, bold=True, fill=light_fill)
    formula_cell(ws, row, 6,
                 f"=SUM(F{refs['annual_summary_start']}:F{row - 1})",
                 fmt=CURR, bold=True, fill=light_fill)
    formula_cell(ws, row, 7, f"=IFERROR((E{row}-D{row})/D{row},0)",
                 fmt=PCT, bold=True, fill=light_fill)
    formula_cell(ws, row, 8, f"=IFERROR((F{row}-E{row})/E{row},0)",
                 fmt=PCT, bold=True, fill=light_fill)

    wb._mighty_refs["sales_forecast"] = refs
    return refs


# ============================================================
# OPEX DATA AGGREGATION
# ============================================================

# Maps accountant P&L line totals to Cash Flow OpEx categories
PL_TO_CF_CATEGORY = {
    "property": [
        "Total 700000 Property Costs", "Total for 700000 Property Costs",
    ],
    "staff": [
        "Total 602000 Payroll", "Total for 602000 Payroll",
    ],
    "utilities": [
        "Total 616000 Utilities", "Total for 616000 Utilities",
    ],
    "marketing": [
        "Total 601000 Sales & Marketing", "Total for 601000 Sales & Marketing",
    ],
    "admin": [
        "603000 Software & Web Services",
        "608000 Insurance",
        "610000 Office Supplies & General Expense",
        "610100 Furniture & Equipment",
        "611000 Shipping & postage",
        "613000 Bank fees & Service Charges",
        "615000 Parking Lot Rental",
    ],
    "professional": [
        "Total 604000 Professional Fees", "Total for 604000 Professional Fees",
    ],
    "travel": [
        "605000 Travel (Airfare/hotel/ground trans/etc)",
        "606000 Meals", "607000 Entertainment",
    ],
    "finance": [
        "506000 Merchant Account Fees",
        "Total Cost of Goods Sold", "Total for Cost of goods sold",
    ],
    "startup": [
        "630000 Studio Start Up Costs",
        "Total 630000 Studio Start Up Costs", "Total for 630000 Studio Start Up Costs",
    ],
    "taxes": [
        "902000 Taxes Paid", "903000 Property taxes",
        "Total 902000 Taxes Paid", "Total for 902000 Taxes Paid",
    ],
}

# Display order for Cash Flow OpEx rows
CF_OPEX_DISPLAY = [
    ("property", "Property Costs"),
    ("staff", "Staff Costs"),
    ("utilities", "Utilities"),
    ("marketing", "Marketing & Promotion"),
    ("admin", "Administrative & G&A"),
    ("professional", "Professional Fees"),
    ("travel", "Travel & Meals"),
    ("finance", "Merchant Fees & COGS"),
    ("startup", "Studio Start Up Costs"),
    ("taxes", "Taxes"),
]

# Forecast OpEx categories (from override JSON) → CF category
OPEX_OVERRIDE_TO_CF = {
    "property": "property",
    "staff": "staff",
    "utilities": "utilities",
    "marketing": "marketing",
    "admin": "admin",
    "professional_fees": "professional",
    "travel": "travel",
    "finance": "finance",   # merchant+cogs
    "taxes": "taxes",
}


def compute_actuals_opex(data):
    """Return {month_label: {cf_category: value}} for closed months."""
    pl = data["committed"].get("pl", {})
    result = {}
    for month_label, lines in pl.items():
        cats = {cat: 0.0 for cat, _ in CF_OPEX_DISPLAY}
        for cat, patterns in PL_TO_CF_CATEGORY.items():
            for pat in patterns:
                if pat in lines:
                    v = lines[pat]
                    if isinstance(v, (int, float)):
                        cats[cat] += abs(float(v))
        result[month_label] = cats
    return result


def compute_forecast_opex(data):
    """Return {month_key: {cf_category: value}} for forecast months.

    Sums across studios from merged opex_assumptions.
    Finance is added separately at the CF level (computed from revenue × merchant+cogs%).
    """
    merged_opex = data["merged"].get("opex_assumptions", {})
    forecast_keys = MONTH_KEYS[FIRST_FORECAST_IDX:]  # May 2026 onwards
    result = {mk: {cat: 0.0 for cat, _ in CF_OPEX_DISPLAY} for mk in forecast_keys}

    for studio, cats in merged_opex.items():
        if not isinstance(cats, dict):
            continue
        for cat, months in cats.items():
            cf_cat = OPEX_OVERRIDE_TO_CF.get(cat)
            if not cf_cat or not isinstance(months, dict):
                continue
            for mk, v in months.items():
                if mk in result:
                    try:
                        result[mk][cf_cat] += float(v)
                    except (ValueError, TypeError):
                        pass
    return result


def actuals_total_cash_sales(data):
    """Return {month_label: total cash sales} from monthly_sales (the QBO actuals)."""
    ms = data["committed"].get("monthly_sales", {})
    actuals = {}
    for mk in MONTH_KEYS[:FIRST_FORECAST_IDX]:  # Jan-Apr 2026
        actuals[mk] = float(ms.get(mk, 0))
    return actuals


def actuals_other_lines(data):
    """Extract depreciation, interest, other-income lines from accountant P&L."""
    pl = data["committed"].get("pl", {})
    lines = {
        "depreciation": ["810000 Depreciation", "Total 810000 Depreciation"],
        "interest": ["901000 Interest Expense/(Income)"],
        "other_income": ["900000 Other Expense/(Income)"],
    }
    result = {}
    for month_label, plines in pl.items():
        month_dict = {}
        for key, patterns in lines.items():
            for pat in patterns:
                if pat in plines:
                    v = plines[pat]
                    if isinstance(v, (int, float)):
                        month_dict[key] = month_dict.get(key, 0) + float(v)
        result[month_label] = month_dict
    return result


def month_label_to_key(label):
    """Convert 'Jan 2026' → '2026-01'."""
    mo_map = {"Jan":"01","Feb":"02","Mar":"03","Apr":"04","May":"05","Jun":"06",
              "Jul":"07","Aug":"08","Sep":"09","Oct":"10","Nov":"11","Dec":"12"}
    parts = label.split()
    if len(parts) != 2:
        return None
    return f"{parts[1]}-{mo_map.get(parts[0], '00')}"


def build_cash_flow_forecast(wb, data):
    """
    Cash Flow Forecast tab. Mirrors Streamlit Cash Flow page.

    Sections:
      1. Cash Inflows (Total Cash Sales)
      2. Operating Outflows (10 OpEx categories)
      3. Net Cash from Operations
      4. Investing (Equipment, Leasehold, Deposits)
      5. Financing (Loan Proceeds, Loan Repayments, Intercompany)
      6. Net Change in Cash
      7. Beginning / Ending Cash Balance
    """
    ws = wb.create_sheet("Cash Flow Forecast")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "D6"

    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 12
    for i in range(36):
        ws.column_dimensions[get_column_letter(4 + i)].width = 12

    # Title
    ws["B2"] = "Cash Flow Forecast"
    ws["B2"].font = title_font
    ws["B3"] = ("Actuals: Jan-Apr 2026 (light green). Forecast: May 2026 onwards. "
                "Net Change formulas sum operations + investing + financing.")
    ws["B3"].font = Font(name="Calibri", size=10, italic=True, color="666666")

    HEADER_ROW = 5
    section_bar(ws, HEADER_ROW - 1, 2, 39, "Monthly Cash Flow")
    # Month headers
    label_cell(ws, HEADER_ROW, 2, "Line Item", bold=True)
    ws.cell(row=HEADER_ROW, column=2).fill = header_fill
    ws.cell(row=HEADER_ROW, column=2).font = header_font
    label_cell(ws, HEADER_ROW, 3, "Annual Total")
    ws.cell(row=HEADER_ROW, column=3).fill = header_fill
    ws.cell(row=HEADER_ROW, column=3).font = header_font
    ws.cell(row=HEADER_ROW, column=3).alignment = center_align
    for i, label in enumerate(MONTH_LABELS):
        cell = ws.cell(row=HEADER_ROW, column=4 + i, value=label)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align

    sf_refs = wb._mighty_refs["sales_forecast"]
    actuals_opex = compute_actuals_opex(data)
    forecast_opex = compute_forecast_opex(data)
    cash_actuals = actuals_total_cash_sales(data)

    row = HEADER_ROW + 1

    # ---------- INFLOWS ----------
    section_bar(ws, row, 2, 39, "CASH INFLOWS")
    row += 1
    refs = {"cash_sales_row": row}

    label_cell(ws, row, 2, "Total Cash Sales", bold=True)
    # Each month: actuals for closed months, formula =Sales Forecast TOTAL for forecast
    for i, mk in enumerate(MONTH_KEYS):
        col = get_column_letter(4 + i)
        if i <= LAST_ACTUALS_IDX:
            # Hard-coded actual
            actual_value_cell(ws, row, 4 + i, cash_actuals.get(mk, 0), fmt=CURR, bold=True)
        else:
            # Formula: reference Sales Forecast TOTAL row, same month column (cols match: both use col 4+i)
            formula_cell(ws, row, 4 + i,
                         f"='Sales Forecast'!{col}{sf_refs['total_row']}",
                         fmt=CURR, bold=True)
    # Annual total in col C (placeholder — could be SUMIFS)
    # Skip for now; can add later
    row += 2

    # ---------- OPERATING OUTFLOWS ----------
    section_bar(ws, row, 2, 39, "OPERATING OUTFLOWS")
    row += 1
    refs["opex_rows"] = {}
    refs["opex_start_row"] = row
    for cat, label in CF_OPEX_DISPLAY:
        label_cell(ws, row, 2, label)
        for i, mk in enumerate(MONTH_KEYS):
            if i <= LAST_ACTUALS_IDX:
                # Use actuals month label form
                month_label = MONTH_LABELS[i]
                v = actuals_opex.get(month_label, {}).get(cat, 0)
                actual_value_cell(ws, row, 4 + i, v, fmt=CURR)
            else:
                # Forecast: use the precomputed forecast_opex
                # For "finance" specifically, we compute as % of cash sales (matches Streamlit)
                if cat == "finance":
                    col = get_column_letter(4 + i)
                    # =Total Cash Sales × (Merchant+COGS rate)
                    formula_cell(ws, row, 4 + i,
                                 f"={col}{refs['cash_sales_row']}*"
                                 f"(Assumptions!$C${wb._mighty_refs['assumptions']['merchant_row']}"
                                 f"+Assumptions!$C${wb._mighty_refs['assumptions']['cogs_row']})",
                                 fmt=CURR)
                else:
                    v = forecast_opex.get(mk, {}).get(cat, 0)
                    input_cell(ws, row, 4 + i, v, fmt=CURR)
        refs["opex_rows"][cat] = row
        row += 1

    # Total Operating Expenses row
    refs["opex_total_row"] = row
    label_cell(ws, row, 2, "Total Operating Expenses", bold=True)
    for i in range(36):
        col = get_column_letter(4 + i)
        formula_cell(ws, row, 4 + i,
                     f"=SUM({col}{refs['opex_start_row']}:{col}{row - 1})",
                     fmt=CURR, bold=True, fill=light_fill)
    row += 1

    # Net Cash from Operations
    refs["net_ops_row"] = row
    label_cell(ws, row, 2, "Net Cash from Operations", bold=True)
    for i in range(36):
        col = get_column_letter(4 + i)
        formula_cell(ws, row, 4 + i,
                     f"={col}{refs['cash_sales_row']}-{col}{refs['opex_total_row']}",
                     fmt=CURR, bold=True, fill=green_fill)
    row += 2

    # ---------- INVESTING ----------
    section_bar(ws, row, 2, 39, "INVESTING ACTIVITIES")
    row += 1
    refs["investing_start"] = row
    invest_lines = [
        ("equipment", "Equipment & Furniture"),
        ("leasehold", "Leasehold Improvements"),
        ("deposits", "Deposits"),
    ]
    # For actuals, sum from SCF data. For forecast, default to 0 (CapEx handled separately).
    scf = data["committed"].get("scf", {})
    for key, label in invest_lines:
        label_cell(ws, row, 2, label)
        for i, mk in enumerate(MONTH_KEYS):
            if i <= LAST_ACTUALS_IDX:
                month_label = MONTH_LABELS[i]
                v = 0
                # SCF accounts (negative for outflow)
                scf_patterns = {
                    "equipment": ["151000", "152000", "153000", "154000"],
                    "leasehold": ["155"],
                    "deposits": ["171000"],
                }
                scf_month = scf.get(month_label, {})
                for line_label, val in scf_month.items():
                    for pat in scf_patterns.get(key, []):
                        if pat in line_label and isinstance(val, (int, float)):
                            v += float(val)
                            break
                actual_value_cell(ws, row, 4 + i, v, fmt=CURR)
            else:
                input_cell(ws, row, 4 + i, 0, fmt=CURR)
        row += 1
    refs["investing_end"] = row - 1

    # Net Investing
    refs["net_invest_row"] = row
    label_cell(ws, row, 2, "Net Investing", bold=True)
    for i in range(36):
        col = get_column_letter(4 + i)
        formula_cell(ws, row, 4 + i,
                     f"=SUM({col}{refs['investing_start']}:{col}{refs['investing_end']})",
                     fmt=CURR, bold=True, fill=light_fill)
    row += 2

    # ---------- FINANCING ----------
    section_bar(ws, row, 2, 39, "FINANCING ACTIVITIES")
    row += 1
    refs["financing_start"] = row
    finance_lines = [
        ("proceeds", "Loan Proceeds"),
        ("repayments", "Loan Repayments"),
        ("intercompany", "Intercompany / Owner Contributions"),
    ]
    for key, label in finance_lines:
        label_cell(ws, row, 2, label)
        for i, mk in enumerate(MONTH_KEYS):
            if i <= LAST_ACTUALS_IDX:
                month_label = MONTH_LABELS[i]
                v = 0
                scf_month = scf.get(month_label, {})
                if key in ("proceeds", "repayments"):
                    loan_patterns = ["242", "243000", "244000"]
                    for line_label, val in scf_month.items():
                        for pat in loan_patterns:
                            if pat in line_label and isinstance(val, (int, float)):
                                fv = float(val)
                                if key == "proceeds" and fv > 0:
                                    v += fv
                                elif key == "repayments" and fv < 0:
                                    v += fv
                                break
                else:  # intercompany
                    inter_patterns = ["241000", "251000", "Due to", "Opening balance"]
                    for line_label, val in scf_month.items():
                        for pat in inter_patterns:
                            if pat in line_label and isinstance(val, (int, float)):
                                v += float(val)
                                break
                actual_value_cell(ws, row, 4 + i, v, fmt=CURR)
            else:
                input_cell(ws, row, 4 + i, 0, fmt=CURR)
        row += 1
    refs["financing_end"] = row - 1

    # Net Financing
    refs["net_finance_row"] = row
    label_cell(ws, row, 2, "Net Financing", bold=True)
    for i in range(36):
        col = get_column_letter(4 + i)
        formula_cell(ws, row, 4 + i,
                     f"=SUM({col}{refs['financing_start']}:{col}{refs['financing_end']})",
                     fmt=CURR, bold=True, fill=light_fill)
    row += 2

    # ---------- NET CHANGE / CASH BALANCE ----------
    section_bar(ws, row, 2, 39, "NET CHANGE IN CASH")
    row += 1
    refs["net_change_row"] = row
    label_cell(ws, row, 2, "Net Change in Cash", bold=True)
    for i in range(36):
        col = get_column_letter(4 + i)
        formula_cell(ws, row, 4 + i,
                     f"={col}{refs['net_ops_row']}+{col}{refs['net_invest_row']}+"
                     f"{col}{refs['net_finance_row']}",
                     fmt=CURR, bold=True, fill=green_fill)
    row += 1

    # Beginning + Ending Cash Balance
    # Use BS data for actual beginning balance (Jan 2026 opening)
    bs = data["committed"].get("bs", {})
    # Find opening cash from earliest month's "101000 Operating Cash" or similar
    opening_cash = 0
    if "Jan 2026" in bs:
        jan_bs = bs["Jan 2026"]
        for line, val in jan_bs.items():
            if "101000" in line or ("Cash" in line and "operating" in line.lower()):
                if isinstance(val, (int, float)):
                    opening_cash = float(val)
                    break

    refs["beg_cash_row"] = row
    label_cell(ws, row, 2, "Beginning Cash Balance")
    actual_value_cell(ws, row, 4, opening_cash, fmt=CURR, bold=True)
    for i in range(1, 36):
        prev_col = get_column_letter(4 + i - 1)
        # Beginning = prev ending
        formula_cell(ws, row, 4 + i,
                     f"={prev_col}{row + 1}",
                     fmt=CURR, bold=True)
    row += 1

    refs["end_cash_row"] = row
    label_cell(ws, row, 2, "Ending Cash Balance", bold=True)
    for i in range(36):
        col = get_column_letter(4 + i)
        formula_cell(ws, row, 4 + i,
                     f"={col}{refs['beg_cash_row']}+{col}{refs['net_change_row']}",
                     fmt=CURR, bold=True, fill=green_fill)

    wb._mighty_refs["cash_flow"] = refs
    return refs


# ============================================================
# MAIN
# ============================================================

def main():
    print("Loading dashboard data...")
    data = load_dashboard_data()
    print(f"  - committed_actuals: {len(data['committed'])} top-level keys")
    print(f"  - baseline: {len(data['baseline'])} top-level keys")
    print(f"  - merged: ready")

    print("Building workbook...")
    wb = openpyxl.Workbook()
    build_cover(wb, data)
    build_assumptions(wb, data)
    print("  - Assumptions tab built")
    build_sales_forecast(wb, data)
    print("  - Sales Forecast tab built")
    build_cash_flow_forecast(wb, data)
    print("  - Cash Flow Forecast tab built")

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT_PATH)
    print(f"Saved: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
