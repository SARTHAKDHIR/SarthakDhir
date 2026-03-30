"""
Run this script once to generate dummy Excel files for testing.
It creates the same format as the original data files.
"""
import pandas as pd
import os

output_dir = os.path.dirname(os.path.abspath(__file__))

# ──────────────────────────────────────────────
# SHARED ITEM MASTER (used across all 3 files)
# ──────────────────────────────────────────────
items = {
    "ITM001": "Watch Strap Leather",
    "ITM002": "Watch Glass Crystal",
    "ITM003": "Watch Crown Steel",
    "ITM004": "Watch Battery Lithium",
    "ITM005": "Watch Case Stainless",
    "ITM006": "Watch Dial White",
    "ITM007": "Watch Hand Set",
    "ITM008": "Watch Buckle Gold",
    "ITM009": "Watch Spring Bar",
    "ITM010": "Watch Movement ETA",
}

# ──────────────────────────────────────────────
# VENDOR MASTER
# ──────────────────────────────────────────────
vendors = {
    "V001": "Alpha Supplies Pvt Ltd",
    "V002": "Beta Components Ltd",
    "V003": "Gamma Parts Co",
    "V004": "Delta Traders",
    "V005": "Sigma Industries",
}

# ──────────────────────────────────────────────
# ITEM → VENDOR MAPPING
# Some items have MULTIPLE vendors, some SINGLE
# ──────────────────────────────────────────────
item_vendor_map = {
    "ITM001": ["V001", "V002"],          # Multiple vendors
    "ITM002": ["V001", "V003", "V004"],  # Multiple vendors
    "ITM003": ["V002"],                  # Single vendor
    "ITM004": ["V003", "V005"],          # Multiple vendors
    "ITM005": ["V001"],                  # Single vendor
    "ITM006": ["V002", "V004"],          # Multiple vendors
    "ITM007": ["V003"],                  # Single vendor
    "ITM008": ["V001", "V002", "V005"],  # Multiple vendors
    "ITM009": ["V004"],                  # Single vendor
    "ITM010": ["V002", "V003"],          # Multiple vendors
}

# ──────────────────────────────────────────────
# PO & PR qty per (item, vendor) — each vendor
# gets a distinct split so no two are the same
# ──────────────────────────────────────────────
# Format: "ITEM": [(oct, nov, dec), (oct, nov, dec), ...]
#          one tuple per vendor in item_vendor_map order
po_qty_map = {
    "ITM001": [(300, 200, 100), (150,  80,  50)],
    "ITM002": [(250, 180,  90), (100,  70,  40), ( 80,  50,  30)],
    "ITM003": [(400, 300, 200)],
    "ITM004": [(200, 150,  80), (120,  90,  60)],
    "ITM005": [(350, 280, 180)],
    "ITM006": [(220, 160, 100), (180, 120,  80)],
    "ITM007": [(300, 200, 120)],
    "ITM008": [(200, 150,  80), (120,  90,  50), ( 80,  60,  40)],
    "ITM009": [(250, 180, 100)],
    "ITM010": [(180, 130,  70), (140, 100,  60)],
}

pr_qty_map = {
    "ITM001": [(280, 190,  90), (130,  70,  40)],
    "ITM002": [(230, 160,  80), ( 90,  60,  35), ( 70,  45,  25)],
    "ITM003": [(370, 270, 170)],
    "ITM004": [(180, 130,  70), (100,  75,  50)],
    "ITM005": [(320, 250, 160)],
    "ITM006": [(200, 140,  90), (160, 100,  70)],
    "ITM007": [(270, 180, 100)],
    "ITM008": [(180, 130,  70), (100,  80,  45), ( 70,  50,  30)],
    "ITM009": [(220, 160,  90)],
    "ITM010": [(160, 110,  60), (120,  85,  50)],
}

# ──────────────────────────────────────────────
# 1. OPEN PO FILE
# ──────────────────────────────────────────────
po_rows = []
po_num = 4500
for item_code, vendor_list in item_vendor_map.items():
    item_desc = items[item_code]
    for i, vendor_code in enumerate(vendor_list):
        vendor_name = vendors[vendor_code]
        oct_q, nov_q, dec_q = po_qty_map[item_code][i]
        po_rows.append({
            "PO Number":        f"PO-{po_num}",
            "PO Date":          "01-Oct-2024",
            "Item Code":        item_code,
            "Item Description": item_desc,
            "Vendor Code":      vendor_code,
            "Vendor Name":      vendor_name,
            "Oct Qty":          oct_q,
            "Nov Qty":          nov_q,
            "Dec Qty":          dec_q,
            "UOM":              "PCS",
            "Status":           "Open",
        })
        po_num += 1

df_po = pd.DataFrame(po_rows)
po_path = os.path.join(output_dir, "DUMMY OPEN PO OCT.xlsx")
df_po.to_excel(po_path, index=False, sheet_name="Open PO")
print(f"Created: {po_path}  ({len(df_po)} rows)")

# ──────────────────────────────────────────────
# 2. OPEN PR FILE
# ──────────────────────────────────────────────
pr_rows = []
pr_num = 7100
for item_code, vendor_list in item_vendor_map.items():
    item_desc = items[item_code]
    for i, vendor_code in enumerate(vendor_list):
        vendor_name = vendors[vendor_code]
        oct_q, nov_q, dec_q = pr_qty_map[item_code][i]
        pr_rows.append({
            "PR Number":        f"PR-{pr_num}",
            "PR Date":          "01-Oct-2024",
            "Item Code":        item_code,
            "Item Description": item_desc,
            "Vendor Code":      vendor_code,
            "Vendor Name":      vendor_name,
            "Oct Qty":          oct_q,
            "Nov Qty":          nov_q,
            "Dec Qty":          dec_q,
            "UOM":              "PCS",
            "Status":           "Open",
        })
        pr_num += 1

df_pr = pd.DataFrame(pr_rows)
pr_path = os.path.join(output_dir, "DUMMY OPEN PR OCT.xlsx")
df_pr.to_excel(pr_path, index=False, sheet_name="Open PR")
print(f"Created: {pr_path}  ({len(df_pr)} rows)")

# ──────────────────────────────────────────────
# 3. PLAN SHORTAGE FILE
# All 10 items have a shortage; quantities vary
# ──────────────────────────────────────────────
shortage_rows = []
shortage_qty = {
    "ITM001": (300, 200, 150),
    "ITM002": (500, 350, 200),
    "ITM003": (250,   0, 100),
    "ITM004": (180, 120,  80),
    "ITM005": (400, 300, 200),
    "ITM006": (320, 240, 160),
    "ITM007": (150,  90,  60),
    "ITM008": (600, 450, 300),
    "ITM009": (200, 150, 100),
    "ITM010": (280, 210, 140),
}

for item_code, (oct_s, nov_s, dec_s) in shortage_qty.items():
    item_desc = items[item_code]
    num_vendors = len(item_vendor_map[item_code])
    shortage_rows.append({
        "Item Code":        item_code,
        "Item Description": item_desc,
        "Category":         "Watch Components",
        "Supplier Type":    "Multiple" if num_vendors > 1 else "Single",
        "Oct Shortage":     oct_s,
        "Nov Shortage":     nov_s,
        "Dec Shortage":     dec_s,
        "Total Shortage":   oct_s + nov_s + dec_s,
        "Priority":         "High" if (oct_s + nov_s + dec_s) > 700 else "Medium",
    })

df_shortage = pd.DataFrame(shortage_rows)
shortage_path = os.path.join(output_dir, "DUMMY Plan Shortage.xlsx")
df_shortage.to_excel(shortage_path, index=False, sheet_name="Plan Shortage")
print(f"Created: {shortage_path}  ({len(df_shortage)} rows)")

print("\nAll 3 dummy files generated successfully!")
print("\nSummary:")
print(f"  Items with MULTIPLE vendors : {sum(1 for v in item_vendor_map.values() if len(v) > 1)}")
print(f"  Items with SINGLE vendor    : {sum(1 for v in item_vendor_map.values() if len(v) == 1)}")
print(f"  All 10 items exist in PO, PR, and Shortage files.")
