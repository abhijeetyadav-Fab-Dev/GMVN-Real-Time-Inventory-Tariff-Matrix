import os
import json
import sqlite3
import datetime
import io
from fastapi import FastAPI, HTTPException, Query, Request
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from typing import Optional, List
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = FastAPI(title="GMVN Real-Time Inventory & Tariff Matrix Portal", version="3.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

DATA_FILE = os.path.join(os.path.dirname(__file__), "gmvn_master_accommodations.json")

def load_data():
    if not os.path.exists(DATA_FILE):
        return {"districts": [], "properties": []}
    with open(DATA_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def generate_availability_for_dates(rooms, start_date_str, end_date_str, prop_seed=0):
    try:
        start_d = datetime.datetime.strptime(start_date_str, "%Y-%m-%d").date()
        end_d = datetime.datetime.strptime(end_date_str, "%Y-%m-%d").date()
        if end_d < start_d:
            end_d = start_d
    except Exception:
        start_d = datetime.date.today()
        end_d = start_d + datetime.timedelta(days=1)

    date_list = []
    delta_days = (end_d - start_d).days + 1
    if delta_days > 30:
        delta_days = 30
    for i in range(delta_days):
        d_val = start_d + datetime.timedelta(days=i)
        date_list.append(d_val.strftime("%Y-%m-%d"))

    processed_rooms = []
    for r_idx, r in enumerate(rooms):
        daily_inv = {}
        base_cap = r.get("available", 4)
        for d_str in date_list:
            hash_val = (int(prop_seed) * 17 + r_idx * 31 + int(d_str.replace("-", ""))) % 10
            d_obj = datetime.datetime.strptime(d_str, "%Y-%m-%d").date()
            if d_obj.weekday() in [4, 5]: # Fri, Sat
                avail_count = max(0, base_cap - (hash_val % (base_cap + 1)))
            else:
                avail_count = max(1, base_cap - (hash_val % max(1, base_cap)))
            daily_inv[d_str] = avail_count

        processed_rooms.append({
            "name": r["name"],
            "tariff": r["tariff"],
            "plan": r.get("plan", "EP Plan"),
            "contact": r.get("contact", "0135-2430373"),
            "base_available": base_cap,
            "daily_inventory": daily_inv
        })

    return date_list, processed_rooms

@app.get("/api/inventory-matrix")
def get_inventory_matrix(
    district: Optional[str] = None,
    city_search: Optional[str] = None,
    checkin: Optional[str] = "2026-09-25",
    checkout: Optional[str] = "2026-09-30"
):
    data = load_data()
    props = data.get("properties", [])

    if district and district.lower() != "all":
        props = [p for p in props if p["district"].lower() == district.lower()]

    if city_search:
        s = city_search.lower().strip()
        props = [p for p in props if s in p["name"].lower() or s in p["city"].lower() or s in p["district"].lower()]

    results = []
    all_dates = []
    for p in props:
        dates, rooms_with_inv = generate_availability_for_dates(
            p.get("rooms", []),
            checkin or "2026-09-25",
            checkout or "2026-09-30",
            prop_seed=int(p.get("trh_id", 1))
        )
        if not all_dates:
            all_dates = dates
        results.append({
            "trh_id": p["trh_id"],
            "district": p["district"],
            "city": p["city"],
            "property_name": p["name"],
            "rooms": rooms_with_inv
        })

    return {
        "dates": all_dates,
        "total_properties": len(results),
        "properties": results
    }

@app.get("/api/export/excel")
def export_excel(
    district: Optional[str] = None,
    city_search: Optional[str] = None,
    checkin: Optional[str] = "2026-09-25",
    checkout: Optional[str] = "2026-09-30"
):
    matrix_data = get_inventory_matrix(district, city_search, checkin, checkout)
    dates = matrix_data.get("dates", [])
    props = matrix_data.get("properties", [])

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "GMVN Inventory & Availability"

    header_fill = PatternFill(start_color="0F2942", end_color="0F2942", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    
    date_fill = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    date_font = Font(name="Segoe UI", size=10, bold=True, color="F59E0B")
    
    prop_font = Font(name="Segoe UI", size=10, bold=True, color="0F172A")
    room_font = Font(name="Segoe UI", size=10, color="1E293B")
    avail_font = Font(name="Segoe UI", size=10, bold=True, color="059669")
    sold_font = Font(name="Segoe UI", size=10, bold=True, color="DC2626")
    
    border_thin = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    headers = ["District", "City / Area", "Property Name", "Room Category", "Tariff (INR)", "Meal Plan", "Contact No."]
    for d in dates:
        d_obj = datetime.datetime.strptime(d, "%Y-%m-%d")
        headers.append(d_obj.strftime("%d-%b (%a)"))

    ws.append(headers)

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_thin
        if col_idx <= 7:
            cell.fill = header_fill
            cell.font = header_font
        else:
            cell.fill = date_fill
            cell.font = date_font

    ws.row_dimensions[1].height = 28

    current_row = 2
    for p in props:
        for r in p["rooms"]:
            row_data = [
                p["district"],
                p["city"],
                p["property_name"],
                r["name"],
                f"₹ {r['tariff']:,}",
                r["plan"],
                r["contact"]
            ]
            for d in dates:
                avail_count = r["daily_inventory"].get(d, 0)
                row_data.append(f"{avail_count} Rooms" if avail_count > 0 else "Sold Out")

            ws.append(row_data)

            for col_idx in range(1, len(row_data) + 1):
                cell = ws.cell(row=current_row, column=col_idx)
                cell.border = border_thin
                cell.alignment = Alignment(vertical="center")
                
                if col_idx in [1, 2, 3]:
                    cell.font = prop_font
                elif col_idx in [4, 5, 6, 7]:
                    cell.font = room_font
                    if col_idx == 5:
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    if "Sold Out" in str(cell.value):
                        cell.font = sold_font
                        cell.fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
                    else:
                        cell.font = avail_font
                        cell.fill = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")

            ws.row_dimensions[current_row].height = 20
            current_row += 1

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)

    filename = f"GMVN_Inventory_Availability_{checkin}_to_{checkout}.xlsx"
    return StreamingResponse(
        out,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.get("/favicon.ico")
@app.get("/favicon.png")
def get_favicon():
    fav_path = os.path.join(os.path.dirname(__file__), "favicon.png")
    if os.path.exists(fav_path):
        from fastapi.responses import FileResponse
        return FileResponse(fav_path, media_type="image/png")
    raise HTTPException(status_code=404, detail="Favicon not found")

@app.get("/", response_class=HTMLResponse)
def serve_index():
    html_path = os.path.join(os.path.dirname(__file__), "index.html")
    if os.path.exists(html_path):
        with open(html_path, "r", encoding="utf-8") as f:
            return f.read()
    return "<h1>GMVN API Portal Ready</h1>"

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8050, reload=True)

