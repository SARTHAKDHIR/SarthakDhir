"""
dumy_data.py — Generate a realistic sample Excel workbook for testing.

Creates supply_chain_sample.xlsx with three sheets:
  - PO_Data         : 60 purchase orders
  - Vendor_Data     : 12 vendors
  - Inventory_Data  : 40 SKUs

Run:
    python dumy_data.py
"""

from __future__ import annotations

import random
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

random.seed(42)

# ─────────────────────────────────────────────────────────────────────────────
# Reference data
# ─────────────────────────────────────────────────────────────────────────────

VENDORS = [
    "AlphaTech Supplies",
    "BlueStar Components",
    "CrestLine Manufacturing",
    "DeltaForge Industries",
    "EagleParts Co.",
    "FineLine Electronics",
    "GreenPath Logistics",
    "HorizonGoods Ltd.",
    "IronClad Materials",
    "JetStream Procurement",
    "KindleMart Corp.",
    "LightSpeed Auto Parts",
]

CATEGORIES = ["Electronics", "Raw Materials", "Mechanical Parts", "Packaging", "Chemicals"]

ITEMS = {
    "Electronics":       ["Microcontroller Unit", "PCB Board", "Power Module", "Sensor Array"],
    "Raw Materials":     ["Steel Coil", "Aluminium Sheet", "Copper Wire", "Carbon Fibre"],
    "Mechanical Parts":  ["Hydraulic Valve", "Gear Assembly", "Bearing Kit", "Shaft Coupling"],
    "Packaging":         ["Corrugated Box", "Foam Insert", "Shrink Wrap", "Pallet Board"],
    "Chemicals":         ["Solvent X-200", "Lubricant Grade A", "Adhesive MX", "Paint Primer"],
}

STATUSES = ["Delivered", "Delivered", "Delivered", "Pending", "Pending", "Cancelled"]
COUNTRIES = ["China", "Germany", "USA", "India", "South Korea", "Taiwan", "Japan"]


def _rdate(start: str, days: int) -> datetime:
    base = datetime.strptime(start, "%Y-%m-%d")
    return base + timedelta(days=random.randint(0, days))


# ─────────────────────────────────────────────────────────────────────────────
# PO_Data
# ─────────────────────────────────────────────────────────────────────────────

def make_po_data(n: int = 60) -> pd.DataFrame:
    rows = []
    for i in range(1, n + 1):
        vendor   = random.choice(VENDORS)
        category = random.choice(CATEGORIES)
        item     = random.choice(ITEMS[category])
        qty      = random.randint(50, 2000)
        price    = round(random.uniform(5.0, 500.0), 2)
        total    = round(qty * price, 2)
        status   = random.choice(STATUSES)
        order_dt = _rdate("2024-01-01", 365)
        exp_del  = order_dt + timedelta(days=random.randint(7, 45))
        delay    = 0
        act_del  = ""

        if status == "Delivered":
            delay   = random.choices([0, 0, 0, random.randint(1, 30)], weights=[6, 6, 4, 3])[0]
            act_del = (exp_del + timedelta(days=delay)).strftime("%Y-%m-%d")

        rows.append({
            "PO_ID":             f"PO-{i:04d}",
            "Vendor_Name":       vendor,
            "Item_Name":         item,
            "Category":          category,
            "Quantity":          qty,
            "Unit_Price":        price,
            "Total_Value":       total,
            "Status":            status,
            "Order_Date":        order_dt.strftime("%Y-%m-%d"),
            "Expected_Delivery": exp_del.strftime("%Y-%m-%d"),
            "Actual_Delivery":   act_del,
            "Delay_Days":        delay,
            "Remarks":           "Urgent" if delay > 14 else "",
        })

    return pd.DataFrame(rows)


# ─────────────────────────────────────────────────────────────────────────────
# Vendor_Data
# ─────────────────────────────────────────────────────────────────────────────

def make_vendor_data() -> pd.DataFrame:
    rows = []
    for vendor in VENDORS:
        reliability = round(random.uniform(3.5, 9.8), 1)
        on_time_pct = round(random.uniform(55.0, 99.0), 1)
        defect_rate = round(random.uniform(0.1, 8.5), 2)

        rows.append({
            "Vendor_Name":           vendor,
            "Country":               random.choice(COUNTRIES),
            "Category":              random.choice(CATEGORIES),
            "Reliability_Score":     reliability,
            "Avg_Lead_Time":         random.randint(5, 40),
            "On_Time_Delivery_Pct":  on_time_pct,
            "Defect_Rate":           defect_rate,
            "Contract_Status":       random.choice(["Active", "Active", "Active", "Under Review", "Expired"]),
            "Partner_Since":         _rdate("2015-01-01", 2000).strftime("%Y-%m-%d"),
            "Contact_Email":         f"procurement@{vendor.lower().replace(' ', '')[:12]}.com",
        })

    return pd.DataFrame(rows)


# ─────────────────────────────────────────────────────────────────────────────
# Inventory_Data
# ─────────────────────────────────────────────────────────────────────────────

def make_inventory_data(n: int = 40) -> pd.DataFrame:
    rows = []
    sku_counter = 1000

    for category, items in ITEMS.items():
        for item in items:
            qty         = random.randint(0, 5000)
            reorder_pt  = random.randint(100, 800)
            reorder_qty = random.randint(200, 1000)
            unit_cost   = round(random.uniform(2.0, 300.0), 2)

            rows.append({
                "SKU":             f"SKU-{sku_counter}",
                "Item_Name":       item,
                "Category":        category,
                "Warehouse":       random.choice(["WH-NORTH", "WH-SOUTH", "WH-EAST", "WH-WEST"]),
                "Qty_On_Hand":     qty,
                "Reorder_Point":   reorder_pt,
                "Reorder_Qty":     reorder_qty,
                "Unit_Cost":       unit_cost,
                "Primary_Supplier":random.choice(VENDORS),
                "Last_Updated":    _rdate("2024-10-01", 180).strftime("%Y-%m-%d"),
            })
            sku_counter += 1

    return pd.DataFrame(rows[:n])


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def generate_sample_excel(output_path: str = "supply_chain_sample.xlsx") -> None:
    po_df  = make_po_data()
    vd_df  = make_vendor_data()
    inv_df = make_inventory_data()

    path = Path(output_path)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        po_df.to_excel(writer,  sheet_name="PO_Data",        index=False)
        vd_df.to_excel(writer,  sheet_name="Vendor_Data",    index=False)
        inv_df.to_excel(writer, sheet_name="Inventory_Data", index=False)

    print(f"Sample workbook written to: {path.resolve()}")
    print(f"   PO_Data:        {len(po_df)} rows")
    print(f"   Vendor_Data:    {len(vd_df)} rows")
    print(f"   Inventory_Data: {len(inv_df)} rows")


if __name__ == "__main__":
    generate_sample_excel()
