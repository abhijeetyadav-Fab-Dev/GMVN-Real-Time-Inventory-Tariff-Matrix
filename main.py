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
        # Honor base_available directly (0 for sold out properties like Auli on early September dates)
        base_cap = r.get("base_available", r.get("available", 0))
        
        for d_str in date_list:
            if "daily_inventory" in r and d_str in r["daily_inventory"]:
                daily_inv[d_str] = r["daily_inventory"][d_str]
            else:
                daily_inv[d_str] = base_cap

        processed_rooms.append({
            "name": r["name"],
            "code": r.get("code", ""),
            "tariff": r["tariff"],
            "plan": r.get("plan", "EP Plan"),
            "contact": r.get("contact", "0135-2430373"),
            "base_available": base_cap,
            "daily_inventory": daily_inv
        })

    return date_list, processed_rooms

import urllib.request
import asyncio
import websockets
from bs4 import BeautifulSoup
import re

# Master in-memory data cache
CACHED_DATA = load_data()
LIVE_SCRAPE_CACHE = {}

async def scrape_gmvn_live_room_tariff(trh_id: str, checkin_str: str, checkout_str: str):
    """
    Connects to Chrome CDP (ws://127.0.0.1:9222) if active, or returns parsed live cache.
    Formats dates to DD-MM-YYYY as expected by GMVN PHP backend.
    """
    try:
        cin_d = datetime.datetime.strptime(checkin_str, "%Y-%m-%d").strftime("%d-%m-%Y")
        cout_d = datetime.datetime.strptime(checkout_str, "%Y-%m-%d").strftime("%d-%m-%Y")
    except Exception:
        cin_d = checkin_str
        cout_d = checkout_str

    cache_key = f"{trh_id}_{cin_d}_{cout_d}"
    if cache_key in LIVE_SCRAPE_CACHE:
        return LIVE_SCRAPE_CACHE[cache_key]

    url = f"https://gmvnonline.com/room-tariff.php?trhID={trh_id}&checkindate={cin_d}&checkoutdate={cout_d}&adults=&child=&log=23f69e30bb19218a"
    try:
        req = urllib.request.Request("http://127.0.0.1:9222/json", headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=1.5) as r:
            tabs = json.loads(r.read())
            ws_url = tabs[0]["webSocketDebuggerUrl"]
        
        async with websockets.connect(ws_url, max_size=10*1024*1024) as ws:
            await ws.send(json.dumps({"id": 1, "method": "Page.navigate", "params": {"url": url}}))
            await ws.recv()
            await asyncio.sleep(2.0)
            await ws.send(json.dumps({"id": 2, "method": "Runtime.evaluate", "params": {"expression": "document.documentElement.outerHTML", "returnByValue": True}}))
            resp = json.loads(await ws.recv())
            html = resp.get("result", {}).get("result", {}).get("value", "")
            
            soup = BeautifulSoup(html, "html.parser")
            rooms = []
            seen = set()
            for box in soup.find_all("div", class_="blog-content"):
                h3 = box.find("h3")
                if not h3 or "tariff" in h3.text.lower():
                    continue
                rname = h3.text.strip()
                if rname in seen:
                    continue
                seen.add(rname)
                
                parent = box.find_parent("div", class_="row") or box.parent
                tariff = 0
                avail = 0
                contact = "9568006683"
                plan = "CP Plan (Bed Tea & Breakfast)"
                if parent:
                    text = parent.text
                    for line in text.splitlines():
                        l = line.strip()
                        if "tariff" in l.lower() or "₹" in l:
                            nums = re.findall(r'\d+', l.replace(",", ""))
                            if nums and tariff == 0: tariff = int(nums[0])
                        if "available room" in l.lower() or "available bed" in l.lower():
                            nums = re.findall(r'\d+', l)
                            if nums: avail = int(nums[0])
                        if "not available" in l.lower():
                            avail = 0
                        if "contact no" in l.lower():
                            m_nums = re.findall(r'[\d\s,]+', l)
                            if m_nums: contact = m_nums[0].strip()
                        if "plan" in l.lower():
                            plan = l
                rooms.append({
                    "name": rname,
                    "tariff": tariff,
                    "available": avail,
                    "contact": contact,
                    "plan": plan
                })
            if rooms:
                LIVE_SCRAPE_CACHE[cache_key] = rooms
                return rooms
    except Exception as e:
        print(f"CDP Scrape fallback for TRH {trh_id}: {e}")

    return None

@app.get("/api/inventory-matrix")
async def get_inventory_matrix(
    district: Optional[str] = None,
    city_search: Optional[str] = None,
    checkin: Optional[str] = "2026-09-25",
    checkout: Optional[str] = "2026-09-30"
):
    global CACHED_DATA
    props = CACHED_DATA.get("properties", [])

    if district and district.lower() != "all":
        props = [p for p in props if p["district"].lower() == district.lower()]

    if city_search:
        s = city_search.lower().strip()
        props = [p for p in props if s in p["name"].lower() or s in p["city"].lower() or s in p["district"].lower()]

    # Calculate date list
    try:
        start_d = datetime.datetime.strptime(checkin, "%Y-%m-%d").date()
        end_d = datetime.datetime.strptime(checkout, "%Y-%m-%d").date()
        if end_d < start_d: end_d = start_d
    except Exception:
        start_d = datetime.date.today()
        end_d = start_d + datetime.timedelta(days=1)
        
    date_list = []
    delta_days = min(31, (end_d - start_d).days + 1)
    for i in range(delta_days):
        date_list.append((start_d + datetime.timedelta(days=i)).strftime("%Y-%m-%d"))
    all_dates = date_list

    results = []
    for p in props:
        trh_id = p.get("trh_id", "")
        # Real-time live scrape for filtered properties
        live_rooms = await scrape_gmvn_live_room_tariff(trh_id, checkin, checkout)
        
        rooms_output = []
        if live_rooms:
            for lr in live_rooms:
                daily_inv = {d_str: lr["available"] for d_str in date_list}
                rooms_output.append({
                    "name": lr["name"],
                    "code": lr["name"][:5].upper().replace(" ", "-"),
                    "tariff": lr["tariff"],
                    "plan": lr["plan"],
                    "contact": lr["contact"],
                    "base_available": lr["available"],
                    "daily_inventory": daily_inv
                })
        else:
            _, rooms_output = generate_availability_for_dates(
                p.get("rooms", []),
                checkin or "2026-09-25",
                checkout or "2026-09-30",
                prop_seed=int(p.get("trh_id", 1))
            )

        results.append({
            "trh_id": p["trh_id"],
            "district": p["district"],
            "city": p["city"],
            "property_name": p["name"],
            "rooms": rooms_output
        })

    return {
        "dates": all_dates,
        "total_properties": len(results),
        "properties": results
    }

@app.post("/api/live-sync")
async def live_sync(
    district: Optional[str] = None,
    city_search: Optional[str] = None,
    checkin: Optional[str] = "2026-09-02",
    checkout: Optional[str] = "2026-09-30"
):
    """
    Directly clears live cache, scrapes all currently filtered GMVN properties, and updates in-memory feed.
    """
    global CACHED_DATA, LIVE_SCRAPE_CACHE
    LIVE_SCRAPE_CACHE.clear()
    
    props = CACHED_DATA.get("properties", [])
    if district and district.lower() != "all":
        props = [p for p in props if p["district"].lower() == district.lower()]
    if city_search:
        s = city_search.lower().strip()
        props = [p for p in props if s in p["name"].lower() or s in p["city"].lower() or s in p["district"].lower()]

    scraped_count = 0
    for p in props:
        trh_id = p.get("trh_id", "")
        rooms = await scrape_gmvn_live_room_tariff(trh_id, checkin, checkout)
        if rooms:
            scraped_count += 1
            # Update memory store
            for r in p.get("rooms", []):
                for lr in rooms:
                    if lr["name"].lower() == r["name"].lower():
                        r["base_available"] = lr["available"]
                        r["tariff"] = lr["tariff"]

    return {
        "success": True,
        "scraped_properties_count": scraped_count,
        "checkin": checkin,
        "checkout": checkout
    }

@app.get("/api/export/excel")
async def export_excel(
    district: Optional[str] = None,
    city_search: Optional[str] = None,
    checkin: Optional[str] = "2026-09-25",
    checkout: Optional[str] = "2026-09-30"
):
    matrix_data = await get_inventory_matrix(district, city_search, checkin, checkout)
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

