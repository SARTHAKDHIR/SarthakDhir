"""
generate_sample_data.py
-----------------------
Generates a realistic sample Excel file that works directly with
Demand Allocation Studio v2.0.

Output: sample_data.xlsx  (same folder as this script)

Column layout produced
──────────────────────
  StyleDemand_Class   – demand class label  (A / B / C / D)
  Main_Resource       – primary resource ID
  Main_AVL            – total available units for that resource
  Alt_Resource        – alternate resource ID  (or N/A when none)
  Alt_AVL             – available units for the alt resource    (or N/A)
  Sep_Demand … Mar_Demand  – month-wise demand quantities

Usage
──────
  python generate_sample_data.py

Then open Demand Allocation Studio, load sample_data.xlsx, and map:
  StyleDemand Class Column  →  StyleDemand_Class
  Month 1 Name / Column     →  Sep  /  Sep_Demand
  ...  (up to 7 months)
  Main Resource Column      →  Main_Resource
  Main Availability Column  →  Main_AVL
  Alt Resource Column       →  Alt_Resource
  Alt Availability Column   →  Alt_AVL
"""

import random
import pandas as pd
from pathlib import Path

random.seed(42)

# ── Configuration ─────────────────────────────────────────────────────────────
NUM_ROWS     = 30
MONTHS       = ['Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar']
SDC_CLASSES  = ['A', 'B', 'C', 'D']
OUTPUT_FILE  = Path(__file__).parent / "sample_data.xlsx"

# Resources: 6 main types, 5 alt types
MAIN_RESOURCES = [f"M-{str(i).zfill(3)}" for i in range(1, 7)]   # M-001 … M-006
ALT_RESOURCES  = [f"A-{str(i).zfill(3)}" for i in range(1, 6)]   # A-001 … A-005

# Availability pool for each resource (set once — mirrors how the app uses them)
MAIN_AVL = {r: random.randint(150, 400) for r in MAIN_RESOURCES}
ALT_AVL  = {r: random.randint(80,  250) for r in ALT_RESOURCES}
# ─────────────────────────────────────────────────────────────────────────────


def _demand_for_month(month_index: int) -> int:
    """Return a plausible demand value; early months are lighter."""
    base = random.randint(10, 60)
    ramp = 1.0 + month_index * 0.12          # demand grows slightly each month
    return max(1, int(base * ramp))


def build_rows() -> list[dict]:
    rows = []
    for _ in range(NUM_ROWS):
        main_res = random.choice(MAIN_RESOURCES)

        # ~25 % of rows have no alternate resource
        if random.random() < 0.25:
            alt_res = 'N/A'
            alt_avl = 'N/A'
        else:
            alt_res = random.choice(ALT_RESOURCES)
            alt_avl = ALT_AVL[alt_res]

        row = {
            'StyleDemand_Class': random.choice(SDC_CLASSES),
            'Main_Resource':     main_res,
            'Main_AVL':          MAIN_AVL[main_res],
            'Alt_Resource':      alt_res,
            'Alt_AVL':           alt_avl,
        }

        for i, month in enumerate(MONTHS):
            # ~15 % chance of zero demand for a given month
            row[f'{month}_Demand'] = 0 if random.random() < 0.15 else _demand_for_month(i)

        rows.append(row)
    return rows


def main():
    rows = build_rows()
    df   = pd.DataFrame(rows)

    # Sort by StyleDemand_Class so the output looks tidy
    df.sort_values('StyleDemand_Class', inplace=True, ignore_index=True)

    df.to_excel(OUTPUT_FILE, index=False)
    print(f"Sample data written to: {OUTPUT_FILE}")
    print(f"  Rows      : {len(df)}")
    print(f"  Months    : {MONTHS}")
    print(f"  Resources : {MAIN_RESOURCES}  +  alts {ALT_RESOURCES}")
    print()
    print("Column mapping for Demand Allocation Studio:")
    print("  StyleDemand Class  →  StyleDemand_Class")
    for i, m in enumerate(MONTHS, 1):
        print(f"  Month {i} ({m})        →  {m}_Demand")
    print("  Main Resource      →  Main_Resource")
    print("  Main Availability  →  Main_AVL")
    print("  Alt Resource       →  Alt_Resource")
    print("  Alt Availability   →  Alt_AVL")


if __name__ == "__main__":
    main()
