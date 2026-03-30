# Excel Demand Toolkit

A desktop GUI toolkit for supply chain demand processing — combining vendor data, running VLOOKUPs, and allocating shortages against open POs and PRs.

## Tools Included

**Tool 1 — Vendor Merger**
Merges shortage data with Open PO and Open PR files by item code, consolidating vendor information into a single view.

**Tool 2 — Data Combiner & VLOOKUP**
Combines multiple Excel files and performs VLOOKUP-style joins across datasets using a configurable key column.

**Tool 3 — Shortage Allocation**
Allocates shortage quantities against available PO quantities by item, using a priority-based column selection. Outputs allocated vs. remaining quantities per vendor.

## How to Run

1. Install dependencies:
   ```
   pip install pandas openpyxl
   ```
2. Run the app:
   ```
   python excel_demand_toolkit.py
   ```
3. Use the sample files in this repo to test each tool.

## Input Files

| File | Used In | Description |
|---|---|---|
| `DUMMY_OPEN_PO_OCT.xlsx` | Tool 1, Tool 3 | Open Purchase Orders with vendor & monthly qty |
| `DUMMY_OPEN_PR_OCT.xlsx` | Tool 1 | Open Purchase Requisitions with vendor & monthly qty |
| `DUMMY_Plan_Shortage.xlsx` | Tool 1, Tool 3 | Plan shortage by item for Oct–Dec |
| `vendor_list.xlsx` | Tool 2 | Vendor master for VLOOKUP joins |

## Output Files

| File | Description |
|---|---|
| `shortage_with_vendors.xlsx` | Shortage data merged with vendor details (Tool 1 output) |
| `shortage_with_vendors_with_PO_PR.xlsx` | Above merged with PO & PR quantities (Tool 1 full output) |
| `OUTPUT.xlsx` | Combined & VLOOKUP result (Tool 2 output) |

## Sample Data
- `generate_dummy_data.py` — Script used to generate all dummy input files (10 watch component items, 5 vendors)

## Tech Stack
- Python 3.x
- pandas
- openpyxl
- tkinter (built-in)
