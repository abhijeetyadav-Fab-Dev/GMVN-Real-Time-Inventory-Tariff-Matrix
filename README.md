# GMVN Real-Time Inventory & Tariff Matrix

A high-speed, production-ready web application and API engine for checking live room tariffs, categories, and date-wise availability across all 84 **Garhwal Mandal Vikas Nigam (GMVN)** Tourist Rest Houses (TRHs) in Uttarakhand.

---

## 🚀 Deployment & Server Hosting Options

### Option 1: Cloud PaaS (Render / Railway / Fly.io / Heroku)
This repository is configured with `render.yaml` and `Procfile` for **1-click zero-config deployment**:
- **Build Command**: `pip install -r requirements.txt`
- **Start Command**: `uvicorn main:app --host 0.0.0.0 --port $PORT --workers 4`

### Option 2: Docker & Docker Compose
```bash
# Build and start container in background
docker compose up -d --build

# View logs
docker compose logs -f
```
The app will be live at `http://localhost:8000/`.

### Option 3: Linux VPS (Ubuntu / Debian / EC2 / DigitalOcean)
```bash
# 1. Clone to /var/www/gmvn-matrix
sudo git clone https://github.com/abhijeetyadav-Fab-Dev/GMVN-Real-Time-Inventory-Tariff-Matrix.git /var/www/gmvn-matrix
cd /var/www/gmvn-matrix

# 2. Set up virtual environment
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# 3. Setup Systemd Service for 24/7 background running
sudo cp gmvn-matrix.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable gmvn-matrix
sudo systemctl start gmvn-matrix

# 4. Setup Nginx Reverse Proxy
sudo cp nginx.conf /etc/nginx/sites-available/gmvn-matrix
sudo ln -s /etc/nginx/sites-available/gmvn-matrix /etc/nginx/sites-enabled/
sudo nginx -t && sudo systemctl reload nginx
```

---

## 💻 Local Development
```bash
pip install -r requirements.txt
python main.py
```
App runs locally at `http://127.0.0.1:8050/`.
