#!/usr/bin/env python3
"""
NESB STORAGE TANK VISUALIZER - LIVE EXCEL MONITORING SERVER
Reads tank levels, tonnages, capacities, and balances directly from Excel.
Auto-detects file updates when you save changes in Excel!

Requirements:
    pip install flask pandas openpyxl
Run:
    python server.py
Open in Browser:
    http://localhost:5000
"""

import os
import datetime
from flask import Flask, jsonify, request, send_file
import pandas as pd

app = Flask(__name__, static_folder='.', static_url_path='')

# Configuration: Path to your Excel or CSV file
DEFAULT_EXCEL_FILENAME = "NESB_TANK_STOCK.xlsx"
EXCEL_PATH = os.environ.get("TANK_EXCEL_PATH", DEFAULT_EXCEL_FILENAME)

# Initial dataset (all 64 NESB tanks) used if the file does not already exist
INITIAL_TANK_DATA = [
    ("NE 1", "SOAP STOCK ACID OIL", "SSAO/001,002,003", 37.910, 40.000, 2.090, "MS", "NO"),
    ("NE 2", "*NEED CLEAN EX SLB", "", 0.0, 40.000, 40.000, "MS", "NO"),
    ("NE 3", "SOAP STOCK", "*NE PRODUCTION PROGRESS", 27.340, 40.000, 12.660, "MS", "NO"),
    ("NE 6", "LO1025", "BLEND", 21.750, 25.000, 3.250, "AL", "NO"),
    ("NE 7", "LO-LB50", "BLEND", 25.040, 25.000, -0.040, "AL", "NO"),
    ("NE 8 (SUM 2)", "PREPARE DRAIN WATER/OIL", "", 0.0, 24.000, 24.000, "MS", "YES"),
    ("NE 11 (SUM 3)", "PREPARE DRAIN WATER/OIL", "", 0.0, 24.000, 24.000, "MS", "YES"),
    ("NE 14", "MFA", "JUKSAN (BOTTOM OIL) *NEED TRANSFER IBC", 1.880, 35.000, 33.120, "MS", "YES"),
    ("NE 15", "SOAP STOCK ACID OIL", "SSAO/003", 20.030, 35.000, 14.970, "MS", "NO"),
    ("NE 16", "RESIDUE OIL", "BUNGE LIPID", 35.300, 70.000, 34.700, "SS", "NO"),
    ("NE 21", "EFB OIL", "HOOD (MARAN)", 5.080, 17.000, 11.920, "MS", "NO"),
    ("NE 21A", "UCO BLEND", "", 8.450, 17.000, 8.550, "MS", "NO"),
    ("NE 22", "EMPTY", "", 0.0, 8.000, 8.000, "SS", "NO"),
    ("NE 23", "LO-LB50", "BLEND", 17.950, 60.000, 42.050, "SS", "YES"),
    ("NE 24", "EFB OIL", "BIOVISION EFB/168", 35.360, 50.000, 14.640, "SS", "YES"),
    ("NE 25", "EFB OIL", "HOOD (MARAN/PERAK)", 71.270, 80.000, 8.730, "SS", "YES"),
    ("NE 26", "SLB", "RICH OIL", 63.140, 70.000, 6.860, "SS", "YES"),
    ("NE 27", "CRUDE GLYCERINE", "CGLY/001/26-BLD", 20.000, 30.000, 10.000, "SS", "YES"),
    ("NE 28", "EFB OIL", "HOOD (MARAN) EFB/158", 30.000, 30.000, 0.0, "SS", "YES"),
    ("NE 29", "EMPTY", "", 0.0, 30.000, 30.000, "SS", "YES"),
    ("NE 30", "EFB OIL", "HOOD (PERAK)/(MARAN)", 73.470, 85.000, 11.530, "MS", "NO"),
    ("NE 32", "UNDER REPAIR", "", 0.0, 48.000, 48.000, "MS", "YES"),
    ("NE 34", "SLB HIGH MOISTURE", "SLB/097, SLB/113 @ MAO/003", 48.000, 48.000, 0.0, "MS", "YES"),
    ("NE 35", "IFWO (WATER)", "", 20.000, 20.000, 0.0, "SS", "NO"),
    ("NE 36", "SOAP STOCK", "*NE PRODUCTION INPROGRESS", 20.000, 20.000, 0.0, "SS", "YES"),
    ("NE 37", "SLY", "YOGESWARI SLY/048, 049, 050.", 104.710, 105.000, 0.290, "MS", "YES"),
    ("NE 38", "MAO", "PT UNIVERSAL MAO/011", 77.170, 105.000, 27.830, "MS", "YES"),
    ("NE 39", "TANK USED DRAIN/DIRT", "", 0.0, 20.000, 20.000, "SS", "NO"),
    ("NE 40", "SLY", "YOGESWARI SLY/044, 036", 33.720, 50.000, 16.280, "SS", "YES"),
    ("NE 41", "SLY", "YOGESWARI", 4.770, 70.000, 65.230, "MS", "YES"),
    ("NE 42", "SLY", "YOGESWARI SLY/151", 40.320, 70.000, 29.680, "MS", "YES"),
    ("NE 47", "EFB OIL", "HOOD (MARAN)", 25.000, 25.000, 0.0, "MS", "NO"),
    ("NE 48", "EFB OIL", "HOOD (MARAN)", 11.140, 25.000, 13.860, "MS", "NO"),
    ("NE 49", "PKAO", "YOGESWARI PKAO/016", 42.740, 80.000, 37.260, "MS", "NO"),
    ("NE 50", "EFB OIL", "BIOVISION EFB/146", 35.260, 80.000, 44.740, "MS", "NO"),
    ("NE 51", "EFB OIL", "BIOVISION EFB/152,156", 70.920, 80.000, 9.080, "MS", "NO"),
    ("NE 52", "PKAO", "SRI MAJU MANAGEMENT", 84.960, 85.000, 0.040, "MS", "NO"),
    ("NE 53", "PITCH OIL", "PALM OLEO (KLANG)", 59.940, 80.000, 20.060, "MS", "NO"),
    ("NE 54", "EFB OIL", "HOOD BUMI (PERAK/MARAN) EFB/165,155", 71.010, 80.000, 8.990, "MS", "NO"),
    ("NE 55", "EFB OIL", "HOOD BUMI (PERAK) EFB 151/153/164", 64.560, 80.000, 15.440, "MS", "YES"),
    ("NE 56", "LO1025", "BLEND", 20.660, 80.000, 59.340, "MS", "NO"),
    ("NE 57", "EFB BLEND OIL", "*FOR NESTE/BALANCE*", 10.260, 80.000, 69.740, "MS", "NO"),
    ("NE 58", "*CLEAN EX SLY", "", 0.0, 80.000, 80.000, "MS", "NO"),
    ("NE 59", "SBEO", "ECOOILS PGU", 27.680, 85.000, 57.320, "MS", "NO"),
    ("NE 60", "EFB OIL", "HOOD (PERAK)", 34.630, 50.000, 15.370, "SS", "NO"),
    ("NE 61", "EFB OIL", "HOOD (PERAK/MARAN)", 144.340, 160.000, 15.660, "MS", "YES"),
    ("NE 62", "SLB BLEND", "*SLB BLEND INPROGRESS/7MTS/ECO PGU", 64.932, 160.000, 95.068, "MS", "YES"),
    ("NE 63", "OLCP06", "CARGILL P.P *TODAY NEED BLEND 25/09 EMEROL 100", 60.330, 60.000, -0.330, "SS", "NO"),
    ("NE 64", "*CLEAN EX CCCO", "", 0.0, 60.000, 60.000, "SS", "NO"),
    ("NE 65", "PENDING", "WILL FULLY UPDATE INFO BY THIS WEEK", 0.0, 60.000, 60.000, "TBA", "TBA"),
    ("NE 66", "EFB OIL", "BIOVISION EFB/170", 34.400, 60.000, 25.600, "TBA", "TBA"),
    ("NE 67", "SOAP STOCK", "*NE PRODUCTION INPROGRESS", 27.700, 30.000, 2.300, "TBA", "TBA"),
    ("NE 68", "RESIDUE OIL", "BUNGE LIPID", 31.300, 30.000, -1.300, "TBA", "TBA"),
    ("NE 69", "EMPTY", "", 0.0, 20.000, 20.000, "TBA", "TBA"),
    ("NE 70", "EMPTY", "", 0.0, 20.000, 20.000, "TBA", "TBA"),
    ("NE 71", "EFB OIL", "HOOD (PERAK)", 37.280, 50.000, 12.720, "TBA", "TBA"),
    ("NE 72", "EMPTY", "", 0.0, 120.000, 120.000, "SS", "YES"),
    ("NE 73", "EMPTY", "", 0.0, 120.000, 120.000, "SS", "YES"),
    ("NE 74", "EMPTY", "", 0.0, 120.000, 120.000, "SS", "YES"),
    ("NE 75", "EMPTY", "", 0.0, 120.000, 120.000, "SS", "YES"),
    ("NE 76", "EMPTY", "", 0.0, 120.000, 120.000, "SS", "YES"),
    ("NE 77", "MAO", "PT UNIVERSAL MAO/013 *INPROGRESS UNLOAD 25/09-flexi", 95.000, 120.000, 25.000, "SS", "YES"),
    ("NE 78", "EMPTY", "", 0.0, 120.000, 120.000, "SS", "YES"),
    ("SOAP PLANT TANK 4", "MAO", "MITSUI MAO/012", 19.770, 28.000, 8.230, "SS", "YES"),
]

def ensure_excel_file_exists():
    """Generates an initial Excel file with the NESB dataset if not present."""
    if not os.path.exists(EXCEL_PATH):
        print(f"[*] Creating sample Excel workbook: {EXCEL_PATH}")
        df = pd.DataFrame(INITIAL_TANK_DATA, columns=[
            "TANK NO.", "CARGO", "SUPPLIER", "TONNAGE", "TANK CAPACITY", "BALANCE", "TYPE", "HEATING COIL"
        ])
        with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name="TANK_STOCK", index=False, startrow=5)
            ws = writer.sheets["TANK_STOCK"]
            ws["B2"] = "NESB TANK STOCK"
            ws["B4"] = "DATE :25.09.2026 @2PM"

def parse_tank_sheet(filepath):
    """Parses Excel with dynamic header detection (even if data starts on row 6)."""
    if filepath.endswith('.csv'):
        df_raw = pd.read_csv(filepath, header=None)
    else:
        df_raw = pd.read_excel(filepath, header=None)

    # Search first 20 rows for "TANK NO."
    header_row_idx = None
    for r_idx in range(min(20, len(df_raw))):
        row_values = [str(val).upper().strip() for val in df_raw.iloc[r_idx].dropna()]
        if any("TANK NO" in v or "TANK_NO" in v for v in row_values):
            header_row_idx = r_idx
            break

    if header_row_idx is None:
        raise ValueError("Could not locate 'TANK NO.' header row in Excel file.")

    headers = [str(h).strip().upper() for h in df_raw.iloc[header_row_idx]]
    df = df_raw.iloc[header_row_idx + 1:].copy()
    df.columns = headers

    col_tank = next((c for c in headers if "TANK NO" in c or "TANK_NO" in c), None)
    col_cargo = next((c for c in headers if "CARGO" in c or "PRODUCT" in c), None)
    col_supplier = next((c for c in headers if "SUPPLIER" in c or "LOT" in c), None)
    col_tonnage = next((c for c in headers if "TONNAGE" in c or "STOCK" in c or "QTY" in c), None)
    col_capacity = next((c for c in headers if "CAPACITY" in c), None)
    col_balance = next((c for c in headers if "BALANCE" in c or "ULLAGE" in c), None)
    col_type = next((c for c in headers if "TYPE" in c or "MATERIAL" in c), None)
    col_heating = next((c for c in headers if "HEATING" in c or "COIL" in c), None)

    tanks = []
    for _, row in df.iterrows():
        raw_tank_no = str(row.get(col_tank, "")).strip()
        if not raw_tank_no or raw_tank_no == "nan" or "TOTAL" in raw_tank_no.upper():
            continue

        def to_float(val):
            try:
                s = str(val).replace(',', '').strip()
                if s.startswith('-') and len(s) > 1:
                    return -float(s[1:].strip())
                return float(s)
            except:
                return 0.0

        tonnage = to_float(row.get(col_tonnage, 0))
        capacity = to_float(row.get(col_capacity, 0))
        balance = to_float(row.get(col_balance, 0))

        if capacity <= 0:
            continue

        # Helper formulas: TONNAGE / CAPACITY and BALANCE / CAPACITY
        liquid_fill_pct = (tonnage / capacity) * 100.0 if capacity > 0 else 0.0
        empty_space_pct = (balance / capacity) * 100.0 if capacity > 0 else 0.0

        tanks.append({
            "tankNo": raw_tank_no,
            "cargo": str(row.get(col_cargo, "")).strip() if pd.notna(row.get(col_cargo)) else "",
            "supplier": str(row.get(col_supplier, "")).strip() if pd.notna(row.get(col_supplier)) else "",
            "tonnage": round(tonnage, 3),
            "capacity": round(capacity, 3),
            "balance": round(balance, 3),
            "liquidFillPct": round(liquid_fill_pct, 2),
            "emptySpacePct": round(empty_space_pct, 2),
            "type": str(row.get(col_type, "MS")).strip().upper() if pd.notna(row.get(col_type)) else "MS",
            "heatingCoil": "YES" if "YES" in str(row.get(col_heating, "")).upper() else ("TBA" if "TBA" in str(row.get(col_heating, "")).upper() else "NO"),
            "isOverfilled": balance < 0 or tonnage > capacity
        })

    return tanks

@app.route('/')
def serve_index():
    if os.path.exists('index.html'):
        return send_file('index.html')
    return "index.html not found. Place index.html in the same directory.", 404

@app.route('/api/tanks')
def get_tanks():
    ensure_excel_file_exists()
    try:
        tanks = parse_tank_sheet(EXCEL_PATH)
        mtime = os.path.getmtime(EXCEL_PATH)
        mod_time_str = datetime.datetime.fromtimestamp(mtime).strftime("%d/%m/%Y %I:%M:%S %p")

        total_cap = sum(t["capacity"] for t in tanks)
        total_stock = sum(t["tonnage"] for t in tanks)
        total_balance = sum(max(0, t["balance"]) for t in tanks)
        plant_fill = (total_stock / total_cap * 100.0) if total_cap > 0 else 0.0

        return jsonify({
            "success": True,
            "filePath": EXCEL_PATH,
            "lastModified": mod_time_str,
            "tankCount": len(tanks),
            "summary": {
                "totalCapacity": round(total_cap, 3),
                "totalStock": round(total_stock, 3),
                "totalBalance": round(total_balance, 3),
                "plantFillPct": round(plant_fill, 2),
                "heatedCount": len([t for t in tanks if t["heatingCoil"] == "YES"]),
                "overfilledCount": len([t for t in tanks if t["isOverfilled"]]),
                "emptyCount": len([t for t in tanks if t["tonnage"] == 0])
            },
            "tanks": tanks
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/upload', methods=['POST'])
def upload_excel():
    if 'file' not in request.files:
        return jsonify({"success": False, "error": "No file uploaded"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"success": False, "error": "Empty filename"}), 400

    filename = DEFAULT_EXCEL_FILENAME if file.filename.endswith('.xlsx') else file.filename
    file.save(filename)
    return jsonify({"success": True, "message": f"Updated {filename}"})

if __name__ == '__main__':
    ensure_excel_file_exists()
    print("=" * 65)
    print("  NESB STORAGE TANK VISUALIZER - LIVE MONITORING SERVER")
    print(f"  Watching Excel file: {os.path.abspath(EXCEL_PATH)}")
    print("  Open browser: http://localhost:5000")
    print("=" * 65)
    app.run(host='0.0.0.0', port=5000, debug=True)