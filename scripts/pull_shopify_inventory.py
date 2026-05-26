#!/usr/bin/env python3
"""
Pull Shopify inventory snapshot + reconstruct historical month-end balances
============================================================================

Usage:
  python3 scripts/pull_shopify_inventory.py             # current snapshot only
  python3 scripts/pull_shopify_inventory.py --history   # current + back-calculated history
  python3 scripts/pull_shopify_inventory.py --save snapshot_2026-05-22.json

Output is formatted for direct paste into the Excel model's Inventory tab
or QBO Actuals tab.

NOTES on historical reconstruction:
  Shopify's REST API does NOT expose historical month-end inventory.
  This script back-calculates from today's snapshot using order history:
    End_Inv[month N] = End_Inv[today] + Sales[N+1..today] - POs[N+1..today]

  Without PO arrival data (which Shopify doesn't track), the reconstruction
  IGNORES PO arrivals — month-end balances will be UNDERSTATED by however
  much inventory arrived between that month-end and today.

  To get accurate historical balances, also pass --po-data with a JSON
  file containing PO arrival quantities by month, OR cross-reference with
  the Excel model's PO Payment Schedule.

  For PROSPECTIVE tracking: run this script on the 1st of each month and
  save to a dated file. Over time you'll build a real snapshot history.
"""
import sys
import json
import argparse
from pathlib import Path
from datetime import datetime

# Make Streamlit's shopify_client importable
STREAMLIT_DIR = Path(
    "/Users/chandlerclemons/Desktop/Empirica Financial Modeling/Alma Mater/"
    "Streamlit/alma-mater-dashboard-main"
)
sys.path.insert(0, str(STREAMLIT_DIR))

from shopify_client import (
    fetch_inventory_snapshot,
    fetch_products,
    fetch_orders_ytd,
    reconstruct_historical_inventory,
    is_configured,
)


def fmt_snapshot(snap: dict) -> str:
    """Format inventory snapshot for human-readable output."""
    lines = []
    lines.append(f"\n{'='*70}")
    lines.append(f"SHOPIFY INVENTORY SNAPSHOT — as of {snap['as_of']}")
    lines.append(f"{'='*70}")
    lines.append(f"\nBY CATEGORY (paste into Excel Inventory tab):")
    lines.append(f"  Beta:   {snap['by_category']['Beta']:>6,} pairs")
    lines.append(f"  Alpha:  {snap['by_category']['Alpha']:>6,} pairs")
    lines.append(f"  Other:  {snap['by_category']['Other']:>6,} pairs")
    lines.append(f"  TOTAL:  {snap['by_category']['TOTAL']:>6,} pairs")

    lines.append(f"\nBY PRODUCT (top 30 by qty):")
    lines.append(f"  {'Product':<50} {'Cat':<8} {'SKUs':>5} {'Qty':>8}")
    lines.append(f"  {'-'*78}")
    for item in snap['by_product'][:30]:
        title = item['title'][:48]
        lines.append(f"  {title:<50} {item['category']:<8} {item['sku_count']:>5} {item['qty']:>8,}")

    return '\n'.join(lines)


def fmt_historical(history: dict) -> str:
    """Format historical month-end reconstruction."""
    lines = []
    lines.append(f"\n{'='*70}")
    lines.append(f"HISTORICAL MONTH-END INVENTORY (back-calculated)")
    lines.append(f"{'='*70}")
    lines.append(f"\n⚠️  This is back-calculated from today's snapshot using order")
    lines.append(f"    history. PO arrivals were NOT accounted for (no Shopify")
    lines.append(f"    source). Values will be UNDERSTATED for months prior to")
    lines.append(f"    any PO arrivals. Cross-reference with Excel Inventory tab.")
    lines.append(f"\nMonth      Beta     Alpha    Other    TOTAL")
    lines.append(f"-" * 50)
    MONTHS = ['', 'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
    for (y, m) in sorted(history.keys()):
        bal = history[(y, m)]
        total = bal['Beta'] + bal['Alpha'] + bal['Other']
        lines.append(f"{MONTHS[m]} {y}    {bal['Beta']:>6,}    {bal['Alpha']:>6,}    {bal['Other']:>6,}   {total:>6,}")

    return '\n'.join(lines)


def fmt_excel_paste(history: dict, current_snapshot: dict, year: int = 2026) -> str:
    """Format month-end Beta+Alpha balances as a TSV row ready to paste
    into the Excel Inventory tab as 'Ending Inventory (Actual)' rows."""
    lines = []
    MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

    # Build month-by-month for the requested year
    beta_row = ['Beta Ending Inv (Actual)']
    alpha_row = ['Alpha Ending Inv (Actual)']
    for m in range(1, 13):
        if (year, m) in history:
            beta_row.append(str(history[(year, m)]['Beta']))
            alpha_row.append(str(history[(year, m)]['Alpha']))
        else:
            beta_row.append('')
            alpha_row.append('')

    lines.append('\n=== EXCEL PASTE FORMAT (Inventory tab override rows) ===')
    lines.append('Header\t' + '\t'.join(MONTHS))
    lines.append('\t'.join(beta_row))
    lines.append('\t'.join(alpha_row))
    return '\n'.join(lines)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--history', action='store_true',
                        help='Also back-calculate historical month-end balances')
    parser.add_argument('--save', type=str, default=None,
                        help='Save snapshot + history to JSON file')
    parser.add_argument('--year', type=int, default=2026,
                        help='Year to reconstruct (default 2026)')
    parser.add_argument('--po-data', type=str, default=None,
                        help='Optional JSON file with PO arrivals by (year, month)')
    args = parser.parse_args()

    if not is_configured():
        print("❌ Shopify credentials not configured.")
        print("   Set SHOPIFY_STORE + SHOPIFY_ACCESS_TOKEN in")
        print("   ~/Streamlit/alma-mater-dashboard-main/.env")
        sys.exit(1)

    print("Fetching current inventory snapshot from Shopify...")
    snap = fetch_inventory_snapshot()
    print(fmt_snapshot(snap))

    out = {'snapshot': snap}

    if args.history:
        print("\nFetching order history for reconstruction...")
        orders = fetch_orders_ytd(args.year)
        print(f"  {len(orders)} orders pulled for {args.year}")

        po_arrivals = None
        if args.po_data:
            with open(args.po_data) as f:
                # Expected format: {"2026-03": {"Beta": 2000, "Alpha": 1000}, ...}
                raw = json.load(f)
                po_arrivals = {}
                for k, v in raw.items():
                    y_str, m_str = k.split('-')
                    po_arrivals[(int(y_str), int(m_str))] = v

        products = fetch_products()
        history = reconstruct_historical_inventory(
            orders, products,
            current_snapshot=snap,
            po_arrivals=po_arrivals,
        )

        print(fmt_historical(history))
        print(fmt_excel_paste(history, snap, year=args.year))

        # Serialize history with string keys for JSON
        out['history'] = {f"{y}-{m:02d}": v for (y, m), v in history.items()}

    if args.save:
        with open(args.save, 'w') as f:
            json.dump(out, f, indent=2)
        print(f"\n💾 Saved to {args.save}")


if __name__ == "__main__":
    main()
