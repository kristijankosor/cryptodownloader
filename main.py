from fastapi import FastAPI, Request
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import os
import requests
import time
from datetime import datetime, timedelta
from openpyxl import Workbook
from fastapi.responses import RedirectResponse

app = FastAPI()
templates = Jinja2Templates(directory="templates")

DOWNLOAD_FOLDER = "downloads"
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

COINS = ["BTC", "ETH", "XRP", "ADA", "SOL", "BNB", "DOGE", "TRX", "AVAX", "LTC", "LINK"]

def fetch_data(symbol="ETHUSDT", interval="1d", days=1825):
    end_time = int(time.time() * 1000)
    start_time = int((datetime.now() - timedelta(days=days)).timestamp() * 1000)
    url = "https://api.binance.com/api/v3/klines"
    all_data = []

    while True:
        params = {
            "symbol": symbol,
            "interval": interval,
            "startTime": start_time,
            "endTime": end_time,
            "limit": 1000
        }
        response = requests.get(url, params=params)
        if response.status_code != 200:
            return None
        batch = response.json()
        if not batch:
            break
        all_data.extend(batch)
        if len(batch) < 1000:
            break
        start_time = batch[-1][0] + 1

    return all_data

def process_data(klines):
    data = []
    for k in klines:
        date = datetime.fromtimestamp(k[0] / 1000).strftime("%Y-%m-%d")
        open_price = float(k[1])
        high_price = float(k[2])
        low_price = float(k[3])
        close_price = float(k[4])
        volume = float(k[5])
        change_pct = ((close_price - open_price) / open_price) * 100
        data.append({
            "date": date,
            "open": open_price,
            "high": high_price,
            "low": low_price,
            "close": close_price,
            "change_pct": change_pct,
            "volume": volume
        })
    last_day_change = None
    if len(data) >= 2:
        prev_close = data[-2]["close"]
        last_close = data[-1]["close"]
        last_day_change = ((last_close - prev_close) / prev_close) * 100
    return data, last_day_change

def save_to_excel(data, last_day_change, filename):
    wb = Workbook()
    ws = wb.active
    ws.append(["Date", "Open", "High", "Low", "Close", "Change (%)", "Volume"])
    for row in data:
        ws.append([
            row["date"], row["open"], row["high"],
            row["low"], row["close"], row["change_pct"], row["volume"]
        ])
    ws.append([])
    ws.append(["Last Day Change (%)", last_day_change])
    wb.save(filename)


@app.get("/", response_class=HTMLResponse)
async def index(request: Request, message: str = None):
    return templates.TemplateResponse("index.html", {"request": request, "message": message})

@app.get("/download")
async def download(coin: str, days: int = 1825):
    coin = coin.upper()
    symbol = f"{coin}USDT"
    filename = f"{coin}_last_{days}_days.xlsx"
    filepath = os.path.join(DOWNLOAD_FOLDER, filename)

    try:
        if not os.path.exists(filepath):
            klines = fetch_data(symbol, days=days)
            if not klines:
                return RedirectResponse(f"/?message=Could not fetch data for {coin}", status_code=302)
            data, last_day_change = process_data(klines)
            save_to_excel(data, last_day_change, filepath)
        return FileResponse(path=filepath, filename=filename, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return RedirectResponse(f"/?message=Error: {str(e)}", status_code=302)