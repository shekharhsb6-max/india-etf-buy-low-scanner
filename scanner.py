"""
INDIA ETF BUY-LOW SCANNER — GITHUB/PYTHON V4

Runs independently of Google Colab.
Google Sheets is used for inputs/outputs.
NAV IS NOT FETCHED BY THIS SCRIPT.

NAV source:
Spreadsheet ID: 1C0O_uXW2TC44RiLEbilj_zlrS_LcJDEvhD7a_pRnk2M
Sheet: NAV
Source columns:
    A = Symbol
    F = ISINNumber
    H = NAV

The scanner reads NAV directly from that source sheet.

Destination spreadsheet:
    1C0O_uXW2TC44RiLEbilj_zlrS_LcJDEvhD7a_pRnk2M

Destination sheets:
    DASHBOARD
    DAILY_SCAN
    HISTORY
    EXCLUDED
    SETTINGS
    NAV

Main speed improvement:
    yfinance downloads the whole ETF universe in batches instead of
    making one network request per ETF.
"""

import os
import json
import re
import time
from datetime import datetime
from zoneinfo import ZoneInfo

import gspread
import numpy as np
import pandas as pd
import yfinance as yf
from google.oauth2.service_account import Credentials


SPREADSHEET_ID = ""
NAV_SOURCE_SPREADSHEET_ID = "1C0O_uXW2TC44RiLEbilj_zlrS_LcJDEvhD7a_pRnk2M"
NAV_SOURCE_SHEET = "NAV"

SHEETS = ["DASHBOARD", "DAILY_SCAN", "HISTORY", "EXCLUDED", "SETTINGS", "NAV"]

MIN_ADTV_CRORE = 1.0
RSI_PERIOD = 14SPREADSHEET_ID = "1iMFuhNvKUpQpQoUMZaMntEhqiQpx9MOM-8NzxGip7I8"
MIN_HISTORY_DAYS = 220
DELIVERY_LOOKBACK = 20

DEBT_KEYWORDS = [
    "BOND","BONDS","DEBT","GILT","G-SEC","GSEC","G SEC",
    "GOVERNMENT SEC","GOVT SEC","TREASURY","T-BILL","TBILL",
    "T BILL","CORPORATE BOND","BANKING PSU DEBT","BANKING & PSU",
    "BANKING AND PSU","SHORT TERM BOND","ULTRA SHORT",
    "TARGET MATURITY","SDL","STATE DEVELOPMENT LOAN"
]
CASH_KEYWORDS = ["LIQUID","LIQUIDCASE","LIQUID ETF","CASH","OVERNIGHT","MONEY MARKET"]
SPECIAL_KEYWORDS = ["INVERSE","2X","3X","LEVERAGE","LEVERAGED"]

DEFAULT_ALLOCATIONS = {
    "EQUITY": 60.0,
    "GOLD": 10.0,
    "SILVER": 10.0,
    "LIQUID": 20.0,
}

SIGNAL_WEIGHTS = {
    "STRONG BUY": 35,
    "BUY": 25,
    "ACCUMULATE": 12,
    "WATCH": 0,
    "AVOID": 0,
}

def credentials():
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if not raw:
        raise RuntimeError(
            "Missing GOOGLE_SERVICE_ACCOUNT_JSON GitHub secret."
        )
    info = json.loads(raw)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    return Credentials.from_service_account_info(info, scopes=scopes)

def connect():
    gc = gspread.authorize(credentials())
    return gc.open_by_key(SPREADSHEET_ID)

def ws(ss, name):
    sh = ss.worksheet(name)
    return sh

def clean_num(v):
    if v is None or str(v).strip() == "":
        return None
    s = str(v).replace(",", "").replace("₹", "").replace("%", "").strip()
    try:
        return float(s)
    except Exception:
        return None

def read_capital(dashboard):
    v = dashboard.acell("B5").value
    n = clean_num(v)
    if not n or n <= 0:
        raise ValueError("DASHBOARD!B5 must contain Total Deployed Capital > 0.")
    return n

def read_nav_source(ss):
    navws = ss.worksheet(NAV_SOURCE_SHEET)
    rows = navws.get_all_values()
    if not rows:
        return {}

    # Source is A=Symbol, F=ISIN, H=NAV.
    nav_map = {}
    for row in rows[1:]:
        symbol = str(row[0]).strip().upper() if len(row) >= 1 else ""
        isin = str(row[5]).strip().upper() if len(row) >= 6 else ""
        nav = clean_num(row[7]) if len(row) >= 8 else None
        if symbol:
            nav_map[symbol] = {"symbol": symbol, "isin": isin, "nav": nav}
    return nav_map

def read_allocations(settings):
    rows = settings.get_all_values()
    if not rows:
        return DEFAULT_ALLOCATIONS.copy()

    headers = [str(x).strip().lower() for x in rows[0]]
    asset_idx = next((i for i,h in enumerate(headers) if h in ("asset class","asset","category")), None)
    alloc_idx = next((i for i,h in enumerate(headers) if h in ("allocation %","allocation","target %","target allocation")), None)

    allocations = {}
    if asset_idx is not None and alloc_idx is not None:
        for row in rows[1:]:
            if len(row) <= max(asset_idx, alloc_idx):
                continue
            asset = str(row[asset_idx]).strip().upper()
            val = clean_num(row[alloc_idx])
            if asset and val is not None:
                allocations[asset] = val

    if not allocations:
        return DEFAULT_ALLOCATIONS.copy()

    total = sum(allocations.values())
    if abs(total - 100.0) > 0.01:
        raise ValueError(
            f"Asset allocation must total 100%. Current total = {total:.2f}%."
        )
    return allocations

def classify_asset(symbol, underlying):
    text = f"{symbol} {underlying}".upper()
    if any(k in text for k in ["GOLD", "GOLDCASE", "GOLDBEES"]):
        return "GOLD"
    if any(k in text for k in ["SILVER", "SILVERBEES"]):
        return "SILVER"
    if any(k in text for k in ["LIQUID", "OVERNIGHT", "MONEY MARKET"]):
        return "LIQUID"
    if any(k in text for k in ["BOND","GILT","G-SEC","GSEC","TREASURY"]):
        return "DEBT"
    return "EQUITY"

def get_nse_etf_universe():
    import requests
    s = requests.Session()
    s.headers.update({
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json,text/plain,*/*",
        "Referer": "https://www.nseindia.com/",
    })
    try:
        s.get("https://www.nseindia.com/", timeout=15)
        r = s.get("https://www.nseindia.com/api/etf", timeout=20)
        r.raise_for_status()
        data = r.json().get("data", [])
        df = pd.DataFrame(data)
        if df.empty:
            raise RuntimeError("NSE returned an empty ETF universe.")
        return df
    except Exception as e:
        raise RuntimeError(f"Unable to obtain NSE ETF universe: {e}")

def find_col(df, names):
    m = {str(c).strip().lower(): c for c in df.columns}
    for n in names:
        if n.lower() in m:
            return m[n.lower()]
    return None

def build_candidates(raw, nav_map):
    symcol = find_col(raw, ["symbol"])
    undercol = find_col(raw, ["underlyingAsset","underlyingAssetName","underlying"])
    if not symcol:
        raise RuntimeError("NSE ETF symbol column not found.")

    out = pd.DataFrame()
    out["Symbol"] = raw[symcol].astype(str).str.strip().str.upper()
    out["Underlying"] = raw[undercol].fillna("").astype(str) if undercol else ""
    out["SEARCH_TEXT"] = (out["Symbol"] + " " + out["Underlying"]).str.upper()
    out["ExclusionReason"] = ""

    out.loc[out.SEARCH_TEXT.apply(lambda x: any(k in x for k in DEBT_KEYWORDS)), "ExclusionReason"] = "DEBT"
    out.loc[(out.ExclusionReason=="") & out.SEARCH_TEXT.apply(lambda x: any(k in x for k in CASH_KEYWORDS)), "ExclusionReason"] = "CASH / LIQUID"
    out.loc[(out.ExclusionReason=="") & out.SEARCH_TEXT.apply(lambda x: any(k in x for k in SPECIAL_KEYWORDS)), "ExclusionReason"] = "INVERSE / LEVERAGED"

    excluded = out[out.ExclusionReason!=""][["Symbol","Underlying","ExclusionReason"]].copy()
    excluded.rename(columns={"ExclusionReason":"Reason"}, inplace=True)

    candidates = out[out.ExclusionReason==""][["Symbol","Underlying"]].drop_duplicates("Symbol").copy()
    candidates["ISIN"] = candidates["Symbol"].map(lambda x: nav_map.get(x,{}).get("isin",""))
    candidates["NAV"] = candidates["Symbol"].map(lambda x: nav_map.get(x,{}).get("nav"))
    return candidates, excluded

def rsi(series, period=14):
    delta = series.diff()
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)
    ag = gain.ewm(alpha=1/period, adjust=False, min_periods=period).mean()
    al = loss.ewm(alpha=1/period, adjust=False, min_periods=period).mean()
    rs = ag / al.replace(0, np.nan)
    return 100 - (100 / (1 + rs))

def download_prices(symbols):
    tickers = [s + ".NS" for s in symbols]
    if not tickers:
        return {}

    raw = yf.download(
        tickers=tickers,
        period="2y",
        interval="1d",
        auto_adjust=False,
        progress=False,
        group_by="ticker",
        threads=True,
    )
    result = {}

    if raw.empty:
        return result

    for symbol in symbols:
        ticker = symbol + ".NS"
        try:
            if isinstance(raw.columns, pd.MultiIndex):
                if ticker not in raw.columns.get_level_values(0):
                    continue
                d = raw[ticker].copy()
            else:
                d = raw.copy()

            if "Close" not in d.columns or "Volume" not in d.columns:
                continue
            d = d.dropna(subset=["Close"])
            if len(d) >= MIN_HISTORY_DAYS:
                result[symbol] = d
        except Exception:
            continue
    return result

def analyse(symbol, underlying, data):
    close = pd.to_numeric(data["Close"], errors="coerce")
    volume = pd.to_numeric(data["Volume"], errors="coerce")
    d = pd.DataFrame({"Close":close, "Volume":volume}).dropna(subset=["Close"])
    if len(d) < MIN_HISTORY_DAYS:
        return None

    for n in (20,50,100,200):
        d[f"DMA{n}"] = d.Close.rolling(n).mean()
    d["RSI14"] = rsi(d.Close, RSI_PERIOD)
    d["ADTV20_Cr"] = (d.Close*d.Volume).rolling(20).mean()/1e7
    d["VolumeRatio"] = d.Volume/d.Volume.rolling(20).mean()

    x = d.iloc[-1]
    vals = [x[c] for c in ["Close","DMA20","DMA50","DMA100","DMA200","RSI14","ADTV20_Cr","VolumeRatio"]]
    if any(pd.isna(v) for v in vals):
        return None

    rising = bool(d["DMA200"].iloc[-1] > d["DMA200"].iloc[-21])
    return {
        "Symbol": symbol, "Underlying": underlying,
        "Price": float(x.Close),
        "DMA20": float(x.DMA20), "DMA50": float(x.DMA50),
        "DMA100": float(x.DMA100), "DMA200": float(x.DMA200),
        "RSI14": float(x.RSI14), "ADTV20_Cr": float(x.ADTV20_Cr),
        "VolumeRatio": float(x.VolumeRatio), "DMA200_Rising": rising,
    }

def liquidity_score(x):
    if x >= 50: return 20
    if x >= 20: return 17
    if x >= 5: return 13
    if x >= 1: return 7
    return 0

def delivery_score(x):
    if pd.isna(x): return 0
    if x >= 70: return 10
    if x >= 60: return 8
    if x >= 50: return 6
    if x >= 40: return 4
    if x >= 30: return 2
    return 0

def trend_score(row):
    # Highest-priority buy-low setup:
    # Price < 20DMA AND 20DMA > 50DMA > 200DMA.
    if row.Price < row.DMA20 and row.DMA20 > row.DMA50 and row.DMA50 > row.DMA200:
        return 25

    score = 0
    if row.Price > row.DMA200: score += 8
    if row.DMA200_Rising: score += 7
    if row.DMA100 > row.DMA200: score += 5
    if row.DMA50 > row.DMA100: score += 3
    if row.Price > row.DMA20: score += 2
    return min(score,25)

def pullback_score(row):
    p,d20,d50,d100,d200 = row.Price,row.DMA20,row.DMA50,row.DMA100,row.DMA200
    score = 0
    if p < d20: score += 5
    if p < d50: score += 7
    for d in (d100,d200):
        dist = abs(p-d)/d*100
        if dist <= 3: score += 8
        elif dist <= 6: score += 5
    return min(score,25)

def rsi_score(x):
    if 40 <= x <= 50: return 10
    if 35 <= x < 40: return 13
    if 30 <= x < 35: return 15
    if x < 30: return 12
    if 50 < x <= 55: return 6
    return 2

def reversal_score(row):
    score = 0
    if 30 <= row.RSI14 <= 45: score += 2
    if row.VolumeRatio >= 1: score += 1
    if row.Price > row.DMA20: score += 2
    return min(score,5)

def generate_reason(row):
    reasons = []
    # Explicitly identify the user's highest-score setup.
    if row.Price < row.DMA20 and row.DMA20 > row.DMA50 and row.DMA50 > row.DMA200:
        reasons.append("BUY-LOW: below 20 DMA while 20>50>200 DMA")
    elif row.Price > row.DMA200 and row.DMA200_Rising:
        reasons.append("Healthy long-term trend")
    elif row.Price > row.DMA200:
        reasons.append("Above 200 DMA")
    else:
        reasons.append("Below 200 DMA")

    if row.Price < row.DMA20: reasons.append("Below 20 DMA")
    if row.Price < row.DMA50: reasons.append("Below 50 DMA")
    if 30 <= row.RSI14 < 35: reasons.append("RSI deeply oversold")
    elif 35 <= row.RSI14 < 40: reasons.append("RSI attractive")
    elif 40 <= row.RSI14 <= 50: reasons.append("RSI accumulation zone")
    elif row.RSI14 < 30: reasons.append("RSI oversold")
    if row.ADTV20_Cr >= 20: reasons.append("Strong liquidity")
    elif row.ADTV20_Cr >= 5: reasons.append("Good liquidity")
    return "; ".join(reasons[:5])

def scan():
    ss = connect()
    dashboard = ws(ss,"DASHBOARD")
    settings = ws(ss,"SETTINGS")
    daily = ws(ss,"DAILY_SCAN")
    history = ws(ss,"HISTORY")
    excluded_ws = ws(ss,"EXCLUDED")
    nav_ws = ws(ss,"NAV")

    capital = read_capital(dashboard)
    allocations = read_allocations(settings)
    nav_map = read_nav_source(ss)

    raw = get_nse_etf_universe()
    candidates, excluded = build_candidates(raw, nav_map)

    price_map = download_prices(candidates.Symbol.tolist())

    results = []
    for _, row in candidates.iterrows():
        d = price_map.get(row.Symbol)
        if d is None:
            continue
        r = analyse(row.Symbol,row.Underlying,d)
        if r:
            r["ISIN"] = row.ISIN
            r["NAV"] = row.NAV
            r["AssetClass"] = classify_asset(row.Symbol,row.Underlying)
            results.append(r)

    df = pd.DataFrame(results)
    if df.empty:
        raise RuntimeError("No ETF technical data was obtained.")

    lowliq = df[df.ADTV20_Cr < MIN_ADTV_CRORE][["Symbol","Underlying"]].copy()
    lowliq["Reason"] = "LOW LIQUIDITY"
    excluded = pd.concat([excluded,lowliq],ignore_index=True)
    df = df[df.ADTV20_Cr >= MIN_ADTV_CRORE].copy()

    # Score
    df["LiquidityScore"] = df.ADTV20_Cr.apply(liquidity_score)
    df["Delivery20Pct"] = np.nan
    df["DeliveryDays"] = 0
    df["DeliveryScore"] = 0
    df["TrendScore"] = df.apply(trend_score,axis=1)
    df["PullbackScore"] = df.apply(pullback_score,axis=1)
    df["RSIScore"] = df.RSI14.apply(rsi_score)
    df["ReversalScore"] = df.apply(reversal_score,axis=1)

    df["TechnicalScore"] = df[["TrendScore","PullbackScore","RSIScore","ReversalScore"]].sum(axis=1)
    df["FinalScore"] = df[["LiquidityScore","DeliveryScore","TechnicalScore"]].sum(axis=1)

    def signal(row):
        # The requested setup receives the highest signal tier.
        special = row.Price < row.DMA20 and row.DMA20 > row.DMA50 and row.DMA50 > row.DMA200
        if special:
            return "STRONG BUY"
        if row.Price < row.DMA200 and not row.DMA200_Rising:
            return "AVOID"
        if row.FinalScore >= 80: return "STRONG BUY"
        if row.FinalScore >= 70: return "BUY"
        if row.FinalScore >= 60: return "ACCUMULATE"
        if row.FinalScore >= 45: return "WATCH"
        return "AVOID"

    df["Signal"] = df.apply(signal,axis=1)
    df["Reason"] = df.apply(generate_reason,axis=1)

    # Asset-class allocation
    class_totals = {}
    for asset, target in allocations.items():
        subset = df[(df.AssetClass==asset) & (df.Signal!="AVOID")].copy()
        weights = subset.Signal.map(SIGNAL_WEIGHTS).fillna(0)
        totalw = weights.sum()
        class_totals[asset] = (subset, target, weights, totalw)

    deploy_pct = []
    amounts = []
    for _, row in df.iterrows():
        item = class_totals.get(row.AssetClass)
        if item and item[3] > 0 and row.Signal != "AVOID":
            _, target, weights, totalw = item
            w = SIGNAL_WEIGHTS.get(row.Signal,0)
            pct = target * w / totalw
        else:
            pct = 0.0
        deploy_pct.append(pct)
        amounts.append(capital*pct/100)

    df["SuggestedDeploy_%"] = deploy_pct
    df["SuggestedAmount"] = amounts

    for n in [20,50,100,200]:
        df[f"Vs{n}DMA_%"] = (df.Price/df[f"DMA{n}"]-1)*100

    df = df.sort_values(["FinalScore","ADTV20_Cr"],ascending=[False,False]).reset_index(drop=True)
    df["Rank"] = df.index+1

    scan_date = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%Y-%m-%d")
    refresh = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%d-%m-%Y %H:%M:%S")

    columns = [
        "ScanDate","Rank","Symbol","ISIN","Underlying","AssetClass","NAV","Price",
        "FinalScore","Signal","Reason","SuggestedDeploy_%","SuggestedAmount",
        "ADTV20_Cr","Delivery20Pct","DeliveryDays","RSI14","DMA20","DMA50","DMA100","DMA200",
        "Vs20DMA_%","Vs50DMA_%","Vs100DMA_%","Vs200DMA_%","VolumeRatio","DMA200_Rising",
        "LiquidityScore","DeliveryScore","TrendScore","PullbackScore","RSIScore","ReversalScore","TechnicalScore"
    ]
    df.insert(0,"ScanDate",scan_date)
    df = df[[c for c in columns if c in df.columns]]

    for c in df.columns:
        if c not in ("ScanDate","Symbol","ISIN","Underlying","AssetClass","Signal","Reason","DMA200_Rising"):
            df[c] = pd.to_numeric(df[c],errors="coerce").round(2)

    write_sheet(daily,df)
    dashboard.update("B6",refresh)
    write_excluded(excluded_ws,excluded,scan_date)
    append_history(history,df)

    # Keep NAV sheet as an imported/source-driven sheet; do not overwrite it.
    # Update only scanner status cells in dashboard.
    dashboard.update("B7", f"ETFs analysed: {len(df)}")
    dashboard.update("B8", f"Strong Buy: {(df.Signal=='STRONG BUY').sum()}")
    dashboard.update("B9", f"Suggested deployment: ₹{df.SuggestedAmount.sum():,.2f}")

    print(f"SCAN COMPLETE: {len(df)} ETFs | {refresh}")
    return df

def write_sheet(sheet, df):
    data = df.replace([np.inf,-np.inf],np.nan).fillna("")
    values = [data.columns.tolist()] + data.values.tolist()
    sheet.clear()
    sheet.update("A1",values,value_input_option="USER_ENTERED")

def write_excluded(sheet, df, scan_date):
    out = df.copy()
    out.insert(0,"ScanDate",scan_date)
    if out.empty:
        out = pd.DataFrame(columns=["ScanDate","Symbol","Underlying","Reason"])
    write_sheet(sheet,out)

def append_history(sheet, df):
    existing = sheet.get_all_values()
    if not existing:
        write_sheet(sheet,df)
        return
    headers = existing[0]
    rows = df.reindex(columns=headers,fill_value="").replace([np.inf,-np.inf],np.nan).fillna("").values.tolist()
    if rows:
        sheet.append_rows(rows,value_input_option="USER_ENTERED")

if __name__ == "__main__":
    scan()
