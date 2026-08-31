# GMVN Real-Time Inventory & Tariff Matrix

A high-speed, lightweight web application and API engine for checking live room tariffs, categories, and date-wise availability across all 84 **Garhwal Mandal Vikas Nigam (GMVN)** Tourist Rest Houses (TRHs) in Uttarakhand.

## Features
- **Constant Sticky Control Bar**: Change Arrival & Departure Dates, District filter, and City search anytime while scrolling.
- **Direct Data Hierarchy**: `District >> City / Area >> Property Name >> Room Category >> Tariff >> Date-Wise Availability`.
- **Dynamic Date Grid**: Auto-generates day-by-day availability columns for any selected date range.
- **Native Formatted Excel Export (`.xlsx`)**: One-click download of the active matrix styled with currency formatting, headers, and color-coded availability badges.
- **Zero Image Overhead**: Ultra-fast rendering with zero lag.

## Installation & Running Locally

```bash
# 1. Clone repo
git clone https://github.com/abhijeetyadav-Fab-Dev/GMVN-Real-Time-Inventory-Tariff-Matrix.git
cd GMVN-Real-Time-Inventory-Tariff-Matrix

# 2. Install dependencies
pip install -r requirements.txt

# 3. Launch server
python -m uvicorn main:app --host 0.0.0.0 --port 8050 --reload
```

Access the portal at `http://127.0.0.1:8050/`.
