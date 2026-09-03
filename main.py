import os
import json
import logging
import datetime
import io
from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from typing import Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import urllib.request
import asyncio
import websockets
from bs4 import BeautifulSoup
import re

# ---------------------------------------------------------------------------
# DEBUG / LOGGING SETUP
# ---------------------------------------------------------------------------
# Set GMVN_DEBUG=0 in your environment to silence the verbose trail.
DEBUG_ENABLED = os.environ.get("GMVN_DEBUG", "1") != "0"

logger = logging.getLogger("gmvn_portal")
logger.setLevel(logging.DEBUG if DEBUG_ENABLED else logging.INFO)
if not logger.handlers:
    _handler = logging.StreamHandler()
    _handler.setFormatter(logging.Formatter(
        "[%(asctime)s] [%(levelname)s] %(message)s", datefmt="%H:%M:%S"
    ))
    logger.addHandler(_handler)


def dbg(msg: str):
    """Shorthand debug logger — every real-vs-fallback decision point logs through here."""
    if DEBUG_ENABLED:
        logger.debug(msg)


app = FastAPI(title="GMVN Real-Time Inventory & Tariff Matrix Portal", version="3.1.0-debug")

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
        dbg(f"DATA_FILE not found at {DATA_FILE} — starting with empty dataset.")
        return {"districts": [], "properties": []}
    with open(DATA_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
        dbg(f"Loaded master dataset: {len(data.get('properties', []))} properties from {DATA_FILE}")
        return data


def generate_availability_for_dates(rooms, start_date_str, end_date_str, prop_seed=0):
    """
    FALLBACK / STATIC generator. This NEVER touches the live website.
    Every value here comes from gmvn_master_accommodations.json (or the
    in-memory CACHED_DATA, which /api/live-sync may have previously updated).
    """
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
            "daily_inventory": daily_inv,
            # --- DEBUG FIELDS ---
            "data_source": "FALLBACK_STATIC",
            "debug_reason": "Served from gmvn_master_accommodations.json / in-memory cache (no live scrape attempted or scrape failed)."
        })

    dbg(f"generate_availability_for_dates(): built {len(processed_rooms)} FALLBACK_STATIC room rows "
        f"for seed={prop_seed}, dates={date_list[0] if date_list else 'NA'}..{date_list[-1] if date_list else 'NA'}")

    return date_list, processed_rooms


# Master in-memory data cache
CACHED_DATA = load_data()
LIVE_SCRAPE_CACHE = {}


# Global Session Cookie Storage & Scraping Config
ACTIVE_SESSION_COOKIES = []
DEFAULT_INTER_REQUEST_DELAY_MS = 1200
DEFAULT_CONCURRENCY_LIMIT = 1
DEFAULT_RATE_LIMIT_RETRIES = 3

@app.post("/api/capture-session-cookies")
async def capture_session_cookies():
    """
    Connects to Chrome CDP, extracts active cf_clearance, PHPSESSID, and security tokens from GMVN.
    """
    global ACTIVE_SESSION_COOKIES
    try:
        req = urllib.request.Request("http://127.0.0.1:9222/json", headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=1.5) as r:
            tabs = json.loads(r.read())
        if not tabs:
            raise Exception("No active Chrome tabs found on 127.0.0.1:9222")
        ws_url = tabs[0]["webSocketDebuggerUrl"]

        async with websockets.connect(ws_url, max_size=10 * 1024 * 1024) as ws:
            await ws.send(json.dumps({"id": 1, "method": "Network.getCookies", "params": {"urls": ["https://gmvnonline.com"]}}))
            resp = json.loads(await ws.recv())
            cookies = resp.get("result", {}).get("cookies", [])
            ACTIVE_SESSION_COOKIES = cookies
            dbg(f"Captured {len(cookies)} live session cookies from Chrome CDP.")
            
            return {
                "success": True,
                "cookie_count": len(cookies),
                "cookies": [{"name": c["name"], "domain": c["domain"], "expires": c.get("expires")} for c in cookies],
                "message": "Live session cookies captured successfully."
            }
    except Exception as e:
        dbg(f"Failed to capture session cookies via CDP: {e}")
        return {
            "success": False,
            "cookie_count": len(ACTIVE_SESSION_COOKIES),
            "error": str(e),
            "message": "Could not connect to Chrome CDP on port 9222."
        }

async def scrape_gmvn_live_room_tariff(
    trh_id: str,
    checkin_str: str,
    checkout_str: str,
    inter_delay_ms: int = DEFAULT_INTER_REQUEST_DELAY_MS,
    retries: int = DEFAULT_RATE_LIMIT_RETRIES
):
    """
    Connects to Chrome CDP (ws://127.0.0.1:9222).
    Extracts dynamic security tokens, navigates with human-like delay, handles Turnstile challenge, and parses live room cards.
    """
    try:
        cin_d = datetime.datetime.strptime(checkin_str, "%Y-%m-%d").strftime("%d-%m-%Y")
        cout_d = datetime.datetime.strptime(checkout_str, "%Y-%m-%d").strftime("%d-%m-%Y")
    except Exception:
        cin_d = checkin_str
        cout_d = checkout_str

    cache_key = f"{trh_id}_{cin_d}_{cout_d}"
    if cache_key in LIVE_SCRAPE_CACHE:
        dbg(f"[TRH {trh_id}] Cache HIT for key '{cache_key}' — returning cached live data.")
        return LIVE_SCRAPE_CACHE[cache_key], "CACHED_LIVE", f"Served from LIVE_SCRAPE_CACHE, key={cache_key}"

    # Human-like inter-request pacing delay
    if inter_delay_ms > 0:
        await asyncio.sleep(inter_delay_ms / 1000.0)

    for attempt in range(retries):
        try:
            req = urllib.request.Request("http://127.0.0.1:9222/json", headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=1.5) as r:
                tabs = json.loads(r.read())
            if not tabs:
                return None, "FALLBACK_STATIC", "CDP endpoint returned no open tabs at 127.0.0.1:9222."
            ws_url = tabs[0]["webSocketDebuggerUrl"]

            async with websockets.connect(ws_url, max_size=10 * 1024 * 1024) as ws:
                # 1. Fetch current dynamic token from active page if present
                await ws.send(json.dumps({
                    "id": 1,
                    "method": "Runtime.evaluate",
                    "params": {"expression": "(document.querySelector('input[name=\"log\"]') || {}).value || '23f69e30bb19218a'", "returnByValue": True}
                }))
                token_resp = json.loads(await ws.recv())
                log_token = token_resp.get("result", {}).get("result", {}).get("value", "23f69e30bb19218a")

                target_url = f"https://gmvnonline.com/room-tariff.php?trhID={trh_id}&checkindate={cin_d}&checkoutdate={cout_d}&adults=&child=&log={log_token}"
                dbg(f"[TRH {trh_id}] Navigating via CDP (Attempt {attempt+1}/{retries}): {target_url}")
                
                await ws.send(json.dumps({"id": 2, "method": "Page.navigate", "params": {"url": target_url}}))
                await ws.recv()

                # Dynamic polling to resolve Cloudflare Turnstile challenge (up to 15s)
                for poll_idx in range(15):
                    await asyncio.sleep(1.0)
                    await ws.send(json.dumps({"id": 100 + poll_idx, "method": "Runtime.evaluate", "params": {"expression": "document.title", "returnByValue": True}}))
                    t_resp = json.loads(await ws.recv())
                    cur_title = t_resp.get("result", {}).get("result", {}).get("value", "")
                    if cur_title and "Just a moment" not in cur_title and "gmvnonline.com" not in cur_title:
                        dbg(f"[TRH {trh_id}] Turnstile challenge resolved at T+{poll_idx+1}s: '{cur_title}'")
                        break

                await ws.send(json.dumps({
                    "id": 3,
                    "method": "Runtime.evaluate",
                    "params": {"expression": "document.documentElement.outerHTML", "returnByValue": True}
                }))
                resp = json.loads(await ws.recv())
                html = resp.get("result", {}).get("result", {}).get("value", "")

                soup = BeautifulSoup(html, "html.parser")
                rooms = []
                seen = set()
                boxes = soup.find_all("div", class_="blog-content")

                for box in boxes:
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
                            if not l: continue
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
                    dbg(f"[TRH {trh_id}] SUCCESS — scraped {len(rooms)} real room rows from live portal.")
                    return rooms, "LIVE_SCRAPE", f"Successfully parsed {len(rooms)} rooms from live HTML."

        except Exception as e:
            dbg(f"[TRH {trh_id}] Attempt {attempt+1} encountered error: {e}")
            await asyncio.sleep(1.0)

    return None, "FALLBACK_STATIC", "HTML fetched but 0 rooms matched or request timed out."

def normalize_room_name(name: str) -> str:
    """Helper to normalize room name strings for robust comparison."""
    if not name:
        return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(name)).lower()

@app.get("/api/inventory-matrix")
async def get_inventory_matrix(
    district: Optional[str] = None,
    city_search: Optional[str] = None,
    checkin: Optional[str] = "2026-09-25",
    checkout: Optional[str] = "2026-09-30"
):
    global CACHED_DATA
    props = CACHED_DATA.get("properties", [])
    dbg(f"=== /api/inventory-matrix called === district={district!r}, city_search={city_search!r}, "
        f"checkin={checkin}, checkout={checkout}")

    if district and district.lower() != "all":
        props = [p for p in props if p["district"].lower() == district.lower()]

    if city_search:
        s = city_search.lower().strip()
        props = [p for p in props if s in p["name"].lower() or s in p["city"].lower() or s in p["district"].lower()]

    dbg(f"Filtered property count: {len(props)}")

    try:
        start_d = datetime.datetime.strptime(checkin, "%Y-%m-%d").date()
        end_d = datetime.datetime.strptime(checkout, "%Y-%m-%d").date()
        if end_d < start_d:
            end_d = start_d
    except Exception:
        start_d = datetime.date.today()
        end_d = start_d + datetime.timedelta(days=1)

    date_list = []
    delta_days = min(31, (end_d - start_d).days + 1)
    for i in range(delta_days):
        date_list.append((start_d + datetime.timedelta(days=i)).strftime("%Y-%m-%d"))
    all_dates = date_list

    results = []
    # Live scraping is only attempted for narrow, specific searches.
    do_live_scrape = bool(city_search and len(props) <= 3)
    dbg(f"do_live_scrape={do_live_scrape} "
        f"(requires: city_search set AND filtered property count <= 3). "
        f"Current: city_search={'set' if city_search else 'NOT set'}, prop_count={len(props)}")

    stats = {"live_scrape": 0, "cached_live": 0, "fallback_static": 0}

    for p in props:
        trh_id = p.get("trh_id", "")
        live_rooms = None
        source_tag = "FALLBACK_STATIC"
        reason = "do_live_scrape=False for this request (broad search / district-only browse)."

        # Format dates to match scrape cache key
        try:
            cin_d = datetime.datetime.strptime(checkin or "2026-09-25", "%Y-%m-%d").strftime("%d-%m-%Y")
            cout_d = datetime.datetime.strptime(checkout or "2026-09-30", "%Y-%m-%d").strftime("%d-%m-%Y")
        except Exception:
            cin_d = checkin or "2026-09-25"
            cout_d = checkout or "2026-09-30"

        cache_key = f"{trh_id}_{cin_d}_{cout_d}"

        # 1. Check if we already have freshly scraped live data in LIVE_SCRAPE_CACHE
        if cache_key in LIVE_SCRAPE_CACHE:
            live_rooms = LIVE_SCRAPE_CACHE[cache_key]
            source_tag = "CACHED_LIVE"
            reason = f"Served from LIVE_SCRAPE_CACHE (key={cache_key}, synced directly from GMVN portal)."
        elif do_live_scrape:
            live_rooms, source_tag, reason = await scrape_gmvn_live_room_tariff(trh_id, checkin, checkout)

        dbg(f"[TRH {trh_id}] property='{p.get('name')}' -> source={source_tag} | reason: {reason}")

        rooms_output = []
        if live_rooms:
            for lr in live_rooms:
                # NOTE (debug flag): this stamps ONE scraped snapshot value across
                # every date in the requested range — it is not a true per-day
                # calendar. Flagged here so it's visible in the API output.
                daily_inv = {d_str: lr["available"] for d_str in date_list}
                rooms_output.append({
                    "name": lr["name"],
                    "code": lr["name"][:5].upper().replace(" ", "-"),
                    "tariff": lr["tariff"],
                    "plan": lr["plan"],
                    "contact": lr["contact"],
                    "base_available": lr["available"],
                    "daily_inventory": daily_inv,
                    "data_source": source_tag,
                    "debug_reason": reason + " [WARNING: single snapshot value repeated across all requested dates, not true daily granularity]"
                })
            if source_tag == "LIVE_SCRAPE":
                stats["live_scrape"] += 1
            elif source_tag == "CACHED_LIVE":
                stats["cached_live"] += 1
        else:
            _, rooms_output = generate_availability_for_dates(
                p.get("rooms", []),
                checkin or "2026-09-25",
                checkout or "2026-09-30",
                prop_seed=int(p.get("trh_id", 1))
            )
            # Attach the specific reason this property fell back (overwrite generic default)
            for r in rooms_output:
                r["debug_reason"] = reason
            stats["fallback_static"] += 1

        results.append({
            "trh_id": p["trh_id"],
            "district": p["district"],
            "city": p["city"],
            "property_name": p["name"],
            "rooms": rooms_output
        })

    dbg(f"=== Request summary: {stats['live_scrape']} live-scraped, {stats['cached_live']} cached-live, "
        f"{stats['fallback_static']} fallback-static (out of {len(props)} properties) ===")

    return {
        "dates": all_dates,
        "total_properties": len(results),
        "properties": results,
        "_debug_stats": {
            "debug_enabled": DEBUG_ENABLED,
            "do_live_scrape_attempted": do_live_scrape,
            "live_scrape_count": stats["live_scrape"],
            "cached_live_count": stats["cached_live"],
            "fallback_static_count": stats["fallback_static"],
            "note": "do_live_scrape is only True when city_search is set AND the filtered result set is <= 3 properties. "
                    "Otherwise every property below is FALLBACK_STATIC regardless of the page title."
        }
    }


@app.post("/api/live-sync")
async def live_sync(
    district: Optional[str] = None,
    city_search: Optional[str] = None,
    checkin: Optional[str] = "2026-09-25",
    checkout: Optional[str] = "2026-09-30",
    inter_delay_ms: Optional[int] = DEFAULT_INTER_REQUEST_DELAY_MS,
    retries: Optional[int] = DEFAULT_RATE_LIMIT_RETRIES,
    concurrency_limit: Optional[int] = DEFAULT_CONCURRENCY_LIMIT
):
    """
    Clears live cache, captures fresh session tokens, scrapes filtered GMVN properties with concurrency & pacing limits, and updates in-memory feed.
    """
    global CACHED_DATA, LIVE_SCRAPE_CACHE
    dbg(f"=== /api/live-sync called === district={district!r}, city_search={city_search!r}, "
        f"checkin={checkin}, checkout={checkout}, inter_delay_ms={inter_delay_ms}, concurrency={concurrency_limit}")
    
    LIVE_SCRAPE_CACHE.clear()
    
    # Pre-flight: Capture live session cookies
    await capture_session_cookies()

    props = CACHED_DATA.get("properties", [])
    if district and district.lower() != "all":
        props = [p for p in props if p["district"].lower() == district.lower()]
    if city_search:
        s = city_search.lower().strip()
        props = [p for p in props if s in p["name"].lower() or s in p["city"].lower() or s in p["district"].lower()]

    dbg(f"live_sync will scrape {len(props)} properties with concurrency={concurrency_limit} and delay={inter_delay_ms}ms")

    semaphore = asyncio.Semaphore(max(1, min(10, concurrency_limit or 2)))
    scraped_count = 0
    failed_count = 0
    per_property_log = []

    async def scrape_property_worker(p):
        nonlocal scraped_count, failed_count
        trh_id = p.get("trh_id", "")
        async with semaphore:
            rooms, source_tag, reason = await scrape_gmvn_live_room_tariff(
                trh_id, checkin, checkout, inter_delay_ms=inter_delay_ms or 1200, retries=retries or 3
            )
            per_property_log.append({
                "trh_id": trh_id, "property_name": p.get("name"), "source": source_tag, "reason": reason
            })
            if rooms:
                scraped_count += 1
                for r in p.get("rooms", []):
                    norm_r = normalize_room_name(r["name"])
                    for lr in rooms:
                        norm_lr = normalize_room_name(lr["name"])
                        if norm_r == norm_lr or norm_r in norm_lr or norm_lr in norm_r:
                            r["base_available"] = lr["available"]
                            r["tariff"] = lr["tariff"]
            else:
                failed_count += 1

    # Run scraping workers
    tasks = [scrape_property_worker(p) for p in props]
    if tasks:
        await asyncio.gather(*tasks)

    # Persist updated master state to disk
    try:
        with open(DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(CACHED_DATA, f, indent=2)
    except Exception as e:
        dbg(f"Failed to persist live-synced master data: {e}")

    dbg(f"=== live-sync summary: {scraped_count} succeeded, {failed_count} failed (out of {len(props)}) ===")

    return {
        "success": True,
        "scraped_properties_count": scraped_count,
        "failed_properties_count": failed_count,
        "checkin": checkin,
        "checkout": checkout,
        "cookies_captured": len(ACTIVE_SESSION_COOKIES),
        "inter_delay_ms": inter_delay_ms,
        "concurrency_limit": concurrency_limit,
        "_debug_per_property": per_property_log
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
    dbg(f"export_excel(): building workbook for {len(props)} properties, "
        f"debug_stats={matrix_data.get('_debug_stats')}")

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
    source_live_font = Font(name="Segoe UI", size=9, bold=True, color="047857")
    source_fallback_font = Font(name="Segoe UI", size=9, bold=True, color="92400E")

    border_thin = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    # NOTE: "Data Source" column added (position 8) so any exported sheet is
    # self-auditing — you can see per-row whether it was LIVE_SCRAPE,
    # CACHED_LIVE, or FALLBACK_STATIC without needing the server logs.
    headers = ["District", "City / Area", "Property Name", "Room Category",
               "Tariff (INR)", "Meal Plan", "Contact No.", "Data Source"]
    for d in dates:
        d_obj = datetime.datetime.strptime(d, "%Y-%m-%d")
        headers.append(d_obj.strftime("%d-%b (%a)"))

    ws.append(headers)

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_thin
        if col_idx <= 8:
            cell.fill = header_fill
            cell.font = header_font
        else:
            cell.fill = date_fill
            cell.font = date_font

    ws.row_dimensions[1].height = 28

    current_row = 2
    for p in props:
        for r in p["rooms"]:
            source_tag = r.get("data_source", "FALLBACK_STATIC")
            source_label = {
                "LIVE_SCRAPE": "LIVE (scraped now)",
                "CACHED_LIVE": "LIVE (cached)",
                "FALLBACK_STATIC": "STATIC fallback"
            }.get(source_tag, source_tag)

            row_data = [
                p["district"],
                p["city"],
                p["property_name"],
                r["name"],
                f"₹ {r['tariff']:,}",
                r["plan"],
                r["contact"],
                source_label
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
                elif col_idx == 8:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    if source_tag == "FALLBACK_STATIC":
                        cell.font = source_fallback_font
                        cell.fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
                    else:
                        cell.font = source_live_font
                        cell.fill = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
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
    dbg(f"export_excel(): workbook built, {current_row - 2} data rows written, file='{filename}'")
    return StreamingResponse(
        out,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


@app.get("/api/debug/status")
async def debug_status():
    """
    Quick health-check endpoint: tells you right now, without touching any
    property, whether the CDP-controlled Chrome bridge is even reachable —
    the single biggest reason live scraping silently falls back.
    """
    cdp_reachable = False
    cdp_error = None
    tab_count = 0
    try:
        req = urllib.request.Request("http://127.0.0.1:9222/json", headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=1.5) as r:
            tabs = json.loads(r.read())
            tab_count = len(tabs)
            cdp_reachable = True
    except Exception as e:
        cdp_error = f"{type(e).__name__}: {e}"

    status = {
        "debug_enabled": DEBUG_ENABLED,
        "cdp_bridge_reachable": cdp_reachable,
        "cdp_open_tab_count": tab_count,
        "cdp_error": cdp_error,
        "live_scrape_cache_size": len(LIVE_SCRAPE_CACHE),
        "live_scrape_cache_keys": list(LIVE_SCRAPE_CACHE.keys()),
        "note": "If cdp_bridge_reachable is false, EVERY request will use FALLBACK_STATIC data "
                "regardless of city_search/property-count filters, because scrape_gmvn_live_room_tariff() "
                "cannot even reach a controllable Chrome tab."
    }
    dbg(f"/api/debug/status -> {status}")
    return status


@app.get("/favicon.ico")
@app.get("/favicon.png")
def get_favicon():
    fav_path = os.path.join(os.path.dirname(__file__), "favicon.png")
    if os.path.exists(fav_path):
        return FileResponse(fav_path, media_type="image/png")
    raise HTTPException(status_code=404, detail="Favicon not found")


@app.get("/gmvn_master_accommodations.json")
def get_master_json():
    json_path = os.path.join(os.path.dirname(__file__), "gmvn_master_accommodations.json")
    if os.path.exists(json_path):
        return FileResponse(json_path, media_type="application/json")
    return CACHED_DATA

@app.get("/", response_class=HTMLResponse)
def serve_index():
    html_path = os.path.join(os.path.dirname(__file__), "index.html")
    if os.path.exists(html_path):
        with open(html_path, "r", encoding="utf-8") as f:
            return f.read()
    return "<h1>GMVN API Portal Ready</h1>"


if __name__ == "__main__":
    import uvicorn
    dbg(f"Starting server. DEBUG_ENABLED={DEBUG_ENABLED}. "
        f"Hit GET /api/debug/status anytime to check if live scraping can even work right now.")
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)