#!/usr/bin/env python3
"""
NESB STORAGE TANK VISUALIZER - LIVE GOOGLE DRIVE & EXCEL SERVER
Reads tank levels, tonnages, capacities, and balances directly from Google Drive.
Auto-updates every day whenever you edit your file in Google Drive!

Render Start Command: gunicorn app:app
"""

import os
import io
import re
import datetime
from flask import Flask, jsonify, request, send_file
import pandas as pd
import requests

app = Flask(__name__, static_folder='.', static_url_path='')

DEFAULT_EXCEL_FILENAME = "NESB_TANK_STOCK.xlsx"
EXCEL_PATH = os.environ.get("TANK_EXCEL_PATH", DEFAULT_EXCEL_FILENAME)

# Reads Google Drive link from Render environment variable or web UI
GOOGLE_DRIVE_URL = os.environ.get("GOOGLE_DRIVE_URL", os.environ.get("GOOGLE_SHEET_URL", ""))

def get_google_drive_direct_url(url_or_id):
    """
    Transforms any Google Drive or Google Sheets link into a direct download URL.
    """
    raw = str(url_or_id).strip()
    if not raw:
        return None

    # If it is a Google Sheets URL
    sheet_match = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', raw)
    if sheet_match:
        sheet_id = sheet_match.group(1)
        return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"

    # If it is a Google Drive file URL
    file_match = re.search(r'/file/d/([a-zA-Z0-9_-]+)', raw)
    if file_match:
        file_id = file_match.group(1)
        return f"https://drive.google.com/uc?export=download&id={file_id}"

    # Query param ?id=...
    id_param = re.search(r'[?&]id=([a-zA-Z0-9_-]+)', raw)
    if id_param:
        return f"https://drive.google.com/uc?export=download&id={id_param.group(1)}"

    return raw

def fetch_data_from_source():
    """
    Fetches data from Google Drive first. If not configured or offline, falls back to local Excel.
    """
    global GOOGLE_DRIVE_URL
    source_name = "Local Excel File"
    last_mod_str = "N/A"
    df_raw = None

    # 1. Fetch from Google Drive if URL is provided
    if GOOGLE_DRIVE_URL:
        direct_url = get_google_drive_direct_url(GOOGLE_DRIVE_URL)
        try:
            headers = {"User-Agent": "Mozilla/5.0"}
            resp = requests.get(direct_url, headers=headers, timeout=12)
            if resp.status_code == 200 and len(resp.content) > 500:
                excel_bytes = io.BytesIO(resp.content)
                df_raw = pd.read_excel(excel_bytes, header=None)
                source_name = "Google Drive (Live Synced)"
                last_mod_str = datetime.datetime.now().strftime("%d/%m/%Y %I:%M:%S %p")
            else:
                print(f"[Warning] Google Drive returned status {resp.status_code}. Using local file.")
        except Exception as e:
            print(f"[Warning] Failed fetching from Google Drive: {e}. Using local file.")

    # 2. Fallback to local file if Google Drive not set or failed
    if df_raw is None:
        if os.path.exists(EXCEL_PATH):
            df_raw = pd.read_excel(EXCEL_PATH, header=None)
            mtime = os.path.getmtime(EXCEL_PATH)
            last_mod_str = datetime.datetime.fromtimestamp(mtime).strftime("%d/%m/%Y %I:%M:%S %p")
        else:
            raise FileNotFoundError("Neither Google Drive URL nor local Excel file was found.")

    return df_raw, source_name, last_mod_str

def parse_tank_dataframe(df_raw):
    """Dynamically parses tank headers and calculates helper fill & balance formulas."""
    header_row_idx = None
    for r_idx in range(min(25, len(df_raw))):
        row_values = [str(val).upper().strip() for val in df_raw.iloc[r_idx].dropna()]
        if any("TANK NO" in v or "TANK_NO" in v for v in row_values):
            header_row_idx = r_idx
            break

    if header_row_idx is None:
        raise ValueError("Could not find 'TANK NO.' header row in spreadsheet.")

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
    return "index.html not found.", 404

@app.route('/api/tanks')
def get_tanks():
    try:
        df_raw, source_name, last_mod_str = fetch_data_from_source()
        tanks = parse_tank_dataframe(df_raw)

        total_cap = sum(t["capacity"] for t in tanks)
        total_stock = sum(t["tonnage"] for t in tanks)
        total_balance = sum(max(0, t["balance"]) for t in tanks)
        plant_fill = (total_stock / total_cap * 100.0) if total_cap > 0 else 0.0

        return jsonify({
            "success": True,
            "source": source_name,
            "isGoogleDrive": "Google Drive" in source_name,
            "lastModified": last_mod_str,
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

@app.route('/api/config/drive', methods=['GET', 'POST'])
def config_drive():
    global GOOGLE_DRIVE_URL
    if request.method == 'POST':
        data = request.get_json() or {}
        GOOGLE_DRIVE_URL = data.get('url', '').strip()
        return jsonify({"success": True, "message": "Updated Google Drive link!"})
    return jsonify({"googleDriveUrl": GOOGLE_DRIVE_URL})

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
