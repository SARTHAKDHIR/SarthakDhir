# Demand Allocation Studio

A desktop GUI tool for multi-month demand allocation across main and alternate resources, with shortage tracking.

## Features
- Load any Excel file with multiple sheets
- Map columns flexibly — StyleDemand class, up to 7 demand months, main & alternate resources
- Month-wise allocation logic: all demand classes compete fairly per month before the next month begins
- Shared resource pool — each resource's availability is counted once and depleted across rows
- Shortage tracking per row per month
- Export results to Excel or CSV

## How to Run

1. Install dependencies:
   ```
   pip install pandas openpyxl
   ```
2. Run the app:
   ```
   python demand_allocation_studio.py
   ```
3. Load `sample_data.xlsx` to test with pre-built sample data.

## Input File Format

One Excel sheet with the following columns (names are flexible — mapped in the UI):

| Column | Description |
|---|---|
| `StyleDemand_Class` | Demand class label (A / B / C / D) |
| `Main_Resource` | Primary resource ID |
| `Main_AVL` | Total available units for the main resource |
| `Alt_Resource` | Alternate resource ID (or `N/A` if none) |
| `Alt_AVL` | Available units for the alternate resource (or `N/A`) |
| `Sep_Demand` … `Mar_Demand` | Monthly demand quantities (up to 7 months) |

## Column Mapping (UI)

When loading `sample_data.xlsx`, map as follows:

```
StyleDemand Class Column  →  StyleDemand_Class
Month 1 Name / Column     →  Sep  /  Sep_Demand
Month 2 Name / Column     →  Oct  /  Oct_Demand
... (up to 7 months)
Main Resource Column      →  Main_Resource
Main Availability Column  →  Main_AVL
Alt Resource Column       →  Alt_Resource
Alt Availability Column   →  Alt_AVL
```

## Sample Data
- `sample_data.xlsx` — 30 rows, 7 months, 6 main resources, 5 alternate resources
- `generate_sample_data.py` — Script used to generate the sample file
- `OUTPUT.xlsx` — Example output after running allocation

## Tech Stack
- Python 3.x
- pandas
- openpyxl
- tkinter (built-in)
