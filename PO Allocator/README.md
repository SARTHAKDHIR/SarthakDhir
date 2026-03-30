# Smart Vendor Order Allocation

A desktop GUI tool for supply chain demand allocation across vendors based on monthly capacity constraints.

## Features
- Load demand and capacity data from Excel (.xlsx)
- Allocate orders to vendors by priority and capacity
- Rollover logic: unallocated quantity from Month N is carried into Month N+1
- Export results to formatted Excel or CSV
- Built with Python + Tkinter (no external GUI framework needed)

## How to Run
1. Install dependencies:
   ```
   pip install pandas openpyxl
   ```
2. Run the app:
   ```
   python Smart_Vendor_Order_Allocation.py
   ```
3. Load `input_dummy.xlsx` as sample input to test.

## Input File Format
Two sheets required:
- **Demand** — Component name, monthly demand columns, and vendor columns
- **Capacity** — Vendor name, priority, and monthly capacity columns

## Sample Data
- `input_dummy.xlsx` — Sample input (bracelet product line)
- `SmartVendorOrderAllocation_Output.xlsx` — Sample output after allocation run
- `create_dummy_data.py` — Script used to generate the dummy input file

## Tech Stack
- Python 3.x
- pandas
- openpyxl
- tkinter (built-in)
