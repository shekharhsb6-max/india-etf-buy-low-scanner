# ============================================================
# INDIA ETF BUY-LOW SCANNER — GITHUB / PYTHON VERSION
# ============================================================
#
# Runs from GitHub Actions; no Google Colab required.
#
# DESTINATION GOOGLE SHEET:
#   Set GitHub secret GOOGLE_SHEET_ID
#   Default used below: 1iMFuhNvKUpQpQoUMZaMntEhqiQpx9MOM-8NzxGip7I8
#
# NAV SOURCE:
#   Spreadsheet: 1C0O_uXW2TC44RiLEbilj_zlrS_LcJDEvhD7a_pRnk2M
#   Sheet: NAV
#   A = Symbol
#   F = ISINNumber
#   H = NAV
#
# DESTINATION NAV:
#   Sheet NAV
#   A = Symbol
#   B = ISINNumber
#   C = NAV
#
# IMPORTANT:
#   This scanner DOES NOT fetch NAV from AMFI.
#   NAV is read from the source Google Sheet.
#
# ============================================================

import os
import json
import time
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from io import BytesIO

import numpy as np
import pandas as pd
import requests
import yfinance as yf
import gspread
from google.oauth2.service_account import Credentials


# ============================================================
# CONFIGURATION
# ============================================================

DESTINATION_SPREADSHEET_ID = os.getenv(
    "GOOGLE_SHEET_ID",
    "1iMFuhNvKUpQpQoUMZaMntEhqiQpx9MOM-8NzxGip7I8"
)

NAV_SOURCE_SPREADSHEET_ID = (
    "1C0O_uXW2TC44RiLEbilj_zlrS_LcJDEvhD7a_pRnk2M"
)

NAV_SOURCE_SHEET = "NAV"

SHEETS = {
    "NAV": "NAV",
    "DASHBOARD": "DASHBOARD",
    "DAILY_SCAN": "DAILY_SCAN",
    "HISTORY": "HISTORY",
    "EXCLUDED": "EXCLUDED",
    "SETTINGS": "SETTINGS",
}

MIN_ADTV_CRORE = 1.0
DELIVERY_LOOKBACK = 20
RSI_PERIOD = 14
MIN_HISTORY_DAYS = 220
YF_PERIOD = "2y"

NSE_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 Chrome/126.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;"
        "q=0.9,*/*;q=0.8"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.nseindia.com/",
}

DEBT_KEYWORDS = [
    "BOND", "BONDS", "DEBT", "GILT", "G-SEC", "GSEC", "G SEC",
    "GOVERNMENT SEC", "GOVT SEC", "TREASURY", "T-BILL", "TBILL",
    "T BILL", "CORPORATE BOND", "BANKING PSU DEBT",
    "BANKING & PSU", "BANKING AND PSU", "SHORT TERM BOND",
    "ULTRA SHORT", "TARGET MATURITY", "SDL", "STATE DEVELOPMENT LOAN",
]

CASH_KEYWORDS = [
    "LIQUID", "LIQUIDCASE", "LIQUID ETF", "CASH",
    "OVERNIGHT", "MONEY MARKET",
]

SPECIAL_KEYWORDS = [
    "INVERSE", "2X", "3X", "LEVERAGE", "LEVERAGED",
]

# Default allocation. User can change these in SETTINGS.
DEFAULT_ALLOCATIONS = {
    "EQUITY": 60.0,
    "GOLD": 10.0,
    "SILVER": 10.0,
    "INTERNATIONAL": 10.0,
    "CASH": 10.0,
}


# ============================================================
# GOOGLE SHEETS CONNECTION
# ============================================================

def credentials():
    raw = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()

    if not raw:
        raise RuntimeError(
            "Missing GitHub secret GOOGLE_SERVICE_ACCOUNT_JSON."
        )

    try:
        info = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise RuntimeError(
            "GOOGLE_SERVICE_ACCOUNT_JSON is not valid JSON."
        ) from exc

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    return Credentials.from_service_account_info(
        info,
        scopes=scopes,
    )


def connect():
    gc = gspread.authorize(credentials())
    return gc.open_by_key(DESTINATION_SPREADSHEET_ID)


def get_sheet(ss, name):
    ws = ss.worksheet(name)
    return ws


# ============================================================
# SAFE GOOGLE SHEET WRITE
# ============================================================

def write_sheet(ws, dataframe, start_cell="A1", clear=True):
    df = dataframe.copy()

    df = df.replace([np.inf, -np.inf], np.nan)
    df = df.fillna("")

    values = [df.columns.tolist()] + df.astype(object).values.tolist()

    if clear:
        ws.clear()

    # New gspread API order:
    # update(values, range_name, ...)
    ws.update(
        values,
        start_cell,
        value_input_option="USER_ENTERED",
    )


def write_values(ws, range_name, values):
    ws.update(
        values,
        range_name,
        value_input_option="USER_ENTERED",
    )


# ============================================================
# NAV SOURCE — NO AMFI FETCH
# ============================================================

def read_nav_source(gc, spreadsheet):
    """
    Read NAV data from the NAV sheet in the CURRENT spreadsheet.

    NAV sheet:
        A = Symbol
        B = ISIN
        C = NAV

    NAV itself is supplied by IMPORTRANGE in Google Sheets.
    Python does NOT fetch NAV from AMFI or the source spreadsheet.
    """

    print("Reading NAV from current spreadsheet NAV sheet...")

    try:
        nav_ws = spreadsheet.worksheet("NAV")
    except Exception as e:
        raise RuntimeError(
            'NAV sheet was not found in the current spreadsheet.'
        ) from e

    values = nav_ws.get_all_values()

    if not values:
        return {}

    nav_map = {}

    for row in values[1:]:
        if len(row) < 3:
            continue

        symbol = str(row[0]).strip().upper()
        isin = str(row[1]).strip().upper()
        nav_raw = str(row[2]).strip()

        if not symbol:
            continue

        nav = None

        if nav_raw:
            try:
                nav = float(
                    nav_raw
                    .replace(",", "")
                    .replace("₹", "")
                )
            except ValueError:
                nav = None

        nav_map[symbol] = {
            "symbol": symbol,
            "isin": isin,
            "nav": nav
        }

    print(f"NAV records loaded: {len(nav_map)}")

    return nav_map
# ============================================================
# DASHBOARD CAPITAL
# ============================================================

def read_capital(dashboard):
    raw = dashboard.acell("B5").value

    if raw is None or str(raw).strip() == "":
        dashboard.update(
            [[10000]],
            "B5",
            value_input_option="USER_ENTERED",
        )
        return 10000.0

    cleaned = (
        str(raw)
        .replace("₹", "")
        .replace(",", "")
        .replace(" ", "")
        .strip()
    )

    try:
        value = float(cleaned)
    except ValueError as exc:
        raise ValueError(
            f"DASHBOARD!B5 contains invalid capital: {raw}"
        ) from exc

    if value <= 0:
        raise ValueError("DASHBOARD!B5 must be greater than zero.")

    return value


# ============================================================
# SETTINGS / ASSET ALLOCATION
# ============================================================

def normalize_key(value):
    return (
        str(value or "")
        .strip()
        .upper()
        .replace("_", " ")
    )


def read_allocations(settings):
    """
    SETTINGS supported formats:

    Preferred:
      A1 = Asset Class
      B1 = Allocation %
      A2 = EQUITY
      B2 = 60
      ...

    The function also understands labels such as:
      EQUITY ALLOCATION
      GOLD ALLOCATION
    """

    values = settings.get_all_values()

    allocations = {}

    for row in values:
        if len(row) < 2:
            continue

        key = normalize_key(row[0])
        raw = str(row[1]).strip()

        if not key or not raw:
            continue

        key = key.replace(" ALLOCATION", "")

        if key not in {
            "EQUITY",
            "GOLD",
            "SILVER",
            "INTERNATIONAL",
            "CASH",
            "DEBT",
            "OTHER",
        }:
            continue

        try:
            value = float(
                raw.replace("%", "").replace(",", "")
            )
        except ValueError:
            continue

        allocations[key] = value

    # If no allocation rows exist, install defaults.
    if not allocations:
        allocations = DEFAULT_ALLOCATIONS.copy()
        write_allocation_settings(settings, allocations)

    total = sum(allocations.values())

    if total <= 0:
        raise ValueError(
            "Asset allocation total must be greater than 0%."
        )

    # User asked to be able to change allocations. We require 100%.
    if abs(total - 100.0) > 0.01:
        raise ValueError(
            f"Asset allocation must total 100%. Current total = {total:.2f}%."
        )

    return allocations


def write_allocation_settings(settings, allocations):
    rows = [
        ["Asset Class", "Allocation %"],
    ]

    for key, value in allocations.items():
        rows.append([key, value])

    # Only clear the allocation area. Do not overwrite other SETTINGS data.
    clear_rows = max(len(rows), 1)
    settings.batch_clear([f"A1:B{clear_rows}"])

    settings.update(
        rows,
        "A1",
        value_input_option="USER_ENTERED",
    )


def classify_asset(row):
    text = (
        str(row.get("Symbol", ""))
        + " "
        + str(row.get("Underlying", ""))
    ).upper()

    if any(k in text for k in CASH_KEYWORDS):
        return "CASH"

    if any(k in text for k in DEBT_KEYWORDS):
        return "DEBT"

    if any(k in text for k in ["SILVER", "SILV"]):
        return "SILVER"

    if any(k in text for k in [
        "NASDAQ", "S&P", "S&P500", "SP500", "HANG SENG",
        "S&P 500", "US", "AMERICA", "GLOBAL", "WORLD",
        "CHINA", "JAPAN", "EUROPE", "INTERNATIONAL",
    ]):
        return "INTERNATIONAL"

    if any(k in text for k in [
        "GOLD", "GOLDBEES", "GOLD ETF", "GOLDMINES",
    ]):
        return "GOLD"

    return "EQUITY"


def allocation_multiplier(asset_class, allocations):
    """
    Converts the user's asset-class allocation into a deployment
    multiplier. This prevents one asset class from consuming the
    entire deployment simply because it has many ETFs.

    Multiplier = target allocation / number of eligible ETFs in class
    and is applied through class-normalized ranking.
    """
    return allocations.get(asset_class, 0.0)


# ============================================================
# ETF UNIVERSE
# ============================================================

def get_nse_etf_universe():
    session = requests.Session()
    session.headers.update(NSE_HEADERS)

    try:
        session.get(
            "https://www.nseindia.com/",
            timeout=15,
        )
        time.sleep(0.5)

        response = session.get(
            "https://www.nseindia.com/api/etf",
            timeout=30,
        )
        response.raise_for_status()

        data = response.json().get("data", [])
        df = pd.DataFrame(data)

        if df.empty:
            raise RuntimeError(
                "NSE returned an empty ETF universe."
            )

        symbol_col = None
        underlying_col = None

        columns = {
            str(c).strip().lower(): c
            for c in df.columns
        }

        for name in ["symbol"]:
            if name in columns:
                symbol_col = columns[name]
                break

        for name in [
            "underlyingasset",
            "underlyingassetname",
            "underlying",
        ]:
            if name in columns:
                underlying_col = columns[name]
                break

        if symbol_col is None:
            raise RuntimeError(
                "NSE ETF symbol column not found."
            )

        df["Symbol"] = (
            df[symbol_col]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        if underlying_col:
            df["Underlying"] = (
                df[underlying_col]
                .fillna("")
                .astype(str)
            )
        else:
            df["Underlying"] = ""

        df["SEARCH_TEXT"] = (
            df["Symbol"] + " " + df["Underlying"]
        ).str.upper()

        return df[
            ["Symbol", "Underlying", "SEARCH_TEXT"]
        ].drop_duplicates("Symbol")

    except Exception as exc:
        raise RuntimeError(
            f"Unable to obtain NSE ETF universe: {exc}"
        ) from exc


def has_keyword(text, keywords):
    return any(keyword in text for keyword in keywords)


def apply_exclusions(universe):
    universe = universe.copy()
    universe["ExclusionReason"] = ""

    debt = universe["SEARCH_TEXT"].apply(
        lambda x: has_keyword(x, DEBT_KEYWORDS)
    )

    cash = universe["SEARCH_TEXT"].apply(
        lambda x: has_keyword(x, CASH_KEYWORDS)
    )

    special = universe["SEARCH_TEXT"].apply(
        lambda x: has_keyword(x, SPECIAL_KEYWORDS)
    )

    universe.loc[debt, "ExclusionReason"] = "DEBT"
    universe.loc[
        (universe["ExclusionReason"] == "") & cash,
        "ExclusionReason",
    ] = "CASH / LIQUID"
    universe.loc[
        (universe["ExclusionReason"] == "") & special,
        "ExclusionReason",
    ] = "INVERSE / LEVERAGED"

    excluded = universe[
        universe["ExclusionReason"] != ""
    ][
        ["Symbol", "Underlying", "ExclusionReason"]
    ].copy()

    candidates = universe[
        universe["ExclusionReason"] == ""
    ][
        ["Symbol", "Underlying"]
    ].copy()

    return candidates, excluded


# ============================================================
# TECHNICAL INDICATORS
# ============================================================

def calculate_rsi(close, period=14):
    delta = close.diff()
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)

    avg_gain = gain.ewm(
        alpha=1 / period,
        adjust=False,
        min_periods=period,
    ).mean()

    avg_loss = loss.ewm(
        alpha=1 / period,
        adjust=False,
        min_periods=period,
    ).mean()

    rs = avg_gain / avg_loss.replace(0, np.nan)

    return 100 - (100 / (1 + rs))


def flatten_yf_columns(data):
    if isinstance(data.columns, pd.MultiIndex):
        data.columns = data.columns.get_level_values(0)
    return data


def get_price_history(symbol):
    try:
        data = yf.download(
            symbol + ".NS",
            period=YF_PERIOD,
            interval="1d",
            auto_adjust=False,
            progress=False,
            threads=False,
        )

        if data.empty:
            return pd.DataFrame()

        return flatten_yf_columns(data)

    except Exception:
        return pd.DataFrame()


def analyse_etf(symbol, underlying):
    data = get_price_history(symbol)

    if data.empty:
        return None

    required = ["Close", "Volume"]

    if not all(c in data.columns for c in required):
        return None

    data = data.dropna(subset=["Close"]).copy()

    if len(data) < MIN_HISTORY_DAYS:
        return None

    close = pd.to_numeric(
        data["Close"], errors="coerce"
    )

    volume = pd.to_numeric(
        data["Volume"], errors="coerce"
    )

    data["DMA20"] = close.rolling(20).mean()
    data["DMA50"] = close.rolling(50).mean()
    data["DMA100"] = close.rolling(100).mean()
    data["DMA200"] = close.rolling(200).mean()
    data["RSI14"] = calculate_rsi(close, RSI_PERIOD)

    data["TradedValue"] = close * volume

    data["ADTV20_Cr"] = (
        data["TradedValue"].rolling(20).mean() / 1e7
    )

    data["VolumeRatio"] = (
        volume / volume.rolling(20).mean()
    )

    latest = data.iloc[-1]

    dma200_rising = False

    if len(data) >= 221:
        current = data["DMA200"].iloc[-1]
        old = data["DMA200"].iloc[-21]

        if pd.notna(current) and pd.notna(old):
            dma200_rising = bool(current > old)

    fields = [
        latest["Close"],
        latest["DMA20"],
        latest["DMA50"],
        latest["DMA100"],
        latest["DMA200"],
        latest["RSI14"],
        latest["ADTV20_Cr"],
        latest["VolumeRatio"],
    ]

    if any(pd.isna(x) for x in fields):
        return None

    return {
        "Symbol": symbol,
        "Underlying": underlying,
        "Price": float(latest["Close"]),
        "DMA20": float(latest["DMA20"]),
        "DMA50": float(latest["DMA50"]),
        "DMA100": float(latest["DMA100"]),
        "DMA200": float(latest["DMA200"]),
        "RSI14": float(latest["RSI14"]),
        "ADTV20_Cr": float(latest["ADTV20_Cr"]),
        "VolumeRatio": float(latest["VolumeRatio"]),
        "DMA200_Rising": dma200_rising,
    }


# ============================================================
# DELIVERY DATA
# ============================================================

def download_delivery_file(session, date_obj):
    date_str = date_obj.strftime("%d%m%Y")

    url = (
        "https://nsearchives.nseindia.com/products/content/"
        f"sec_bhavdata_full_{date_str}.csv"
    )

    try:
        response = session.get(
            url,
            headers=NSE_HEADERS,
            timeout=15,
        )

        if response.status_code != 200:
            return pd.DataFrame()

        return pd.read_csv(
            BytesIO(response.content)
        )

    except Exception:
        return pd.DataFrame()


def get_delivery_summary(symbols):
    session = requests.Session()
    session.headers.update(NSE_HEADERS)

    try:
        session.get(
            "https://www.nseindia.com/",
            timeout=15,
        )
    except Exception:
        pass

    records = []

    dates = pd.date_range(
        start=datetime.now() - timedelta(days=40),
        end=datetime.now(),
        freq="D",
    )

    for current_date in dates:
        delivery = download_delivery_file(
            session,
            current_date.to_pydatetime(),
        )

        if delivery.empty:
            continue

        delivery.columns = [
            str(c).strip().upper().replace(" ", "_")
            for c in delivery.columns
        ]

        if not all(
            c in delivery.columns
            for c in ["SYMBOL", "DELIV_PER"]
        ):
            continue

        day = delivery[
            ["SYMBOL", "DELIV_PER"]
        ].copy()

        day["SYMBOL"] = (
            day["SYMBOL"]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        day["DELIV_PER"] = pd.to_numeric(
            day["DELIV_PER"],
            errors="coerce",
        )

        day = day[
            day["SYMBOL"].isin(symbols)
        ].copy()

        day["DATE"] = current_date

        records.append(day)

    if not records:
        return pd.DataFrame({
            "Symbol": symbols,
            "Delivery20Pct": np.nan,
            "DeliveryDays": 0,
        })

    df = pd.concat(records, ignore_index=True)

    df = (
        df.sort_values(["SYMBOL", "DATE"])
        .groupby("SYMBOL")
        .tail(DELIVERY_LOOKBACK)
    )

    summary = (
        df.groupby("SYMBOL")
        .agg(
            Delivery20Pct=("DELIV_PER", "mean"),
            DeliveryDays=("DELIV_PER", "count"),
        )
        .reset_index()
        .rename(columns={"SYMBOL": "Symbol"})
    )

    return summary


# ============================================================
# SCORING
# ============================================================

def liquidity_score(x):
    if pd.isna(x):
        return 0
    if x >= 50:
        return 20
    if x >= 20:
        return 17
    if x >= 5:
        return 13
    if x >= 1:
        return 7
    return 0


def delivery_score(x):
    if pd.isna(x):
        return 0
    if x >= 70:
        return 10
    if x >= 60:
        return 8
    if x >= 50:
        return 6
    if x >= 40:
        return 4
    if x >= 30:
        return 2
    return 0


def trend_score(row):
    score = 0

    if row["Price"] > row["DMA200"]:
        score += 8

    if row["DMA200_Rising"]:
        score += 7

    if row["DMA100"] > row["DMA200"]:
        score += 5

    if row["DMA50"] > row["DMA100"]:
        score += 3

    if row["Price"] > row["DMA20"]:
        score += 2

    return min(score, 25)


def pullback_score(row):
    p = row["Price"]
    d20 = row["DMA20"]
    d50 = row["DMA50"]
    d100 = row["DMA100"]
    d200 = row["DMA200"]

    score = 0

    # USER'S PRIORITY RULE:
    # Price < 20DMA AND 20DMA > 50DMA > 200DMA
    # receives the maximum pullback/trend priority.
    if (
        p < d20
        and d20 > d50 > d200
    ):
        score += 25
        return min(score, 25)

    if p < d20:
        score += 5

    if p < d50:
        score += 7

    dist100 = abs(p - d100) / d100 * 100

    if dist100 <= 3:
        score += 8
    elif dist100 <= 6:
        score += 5

    dist200 = abs(p - d200) / d200 * 100

    if dist200 <= 3:
        score += 8
    elif dist200 <= 6:
        score += 5

    return min(score, 25)


def priority_rule_score(row):
    """
    Extra ranking priority requested by the user.

    Maximum score is awarded when:
      Price < 20DMA
      AND 20DMA > 50DMA > 200DMA
    """

    if (
        row["Price"] < row["DMA20"]
        and row["DMA20"] > row["DMA50"] > row["DMA200"]
    ):
        return 25

    return 0


def rsi_score(x):
    if pd.isna(x):
        return 0
    if 40 <= x <= 50:
        return 10
    if 35 <= x < 40:
        return 13
    if 30 <= x < 35:
        return 15
    if x < 30:
        return 12
    if 50 < x <= 55:
        return 6
    return 2


def reversal_score(row):
    score = 0

    if 30 <= row["RSI14"] <= 45:
        score += 2

    if row["VolumeRatio"] >= 1:
        score += 1

    if row["Price"] > row["DMA20"]:
        score += 2

    return min(score, 5)


def determine_signal(row):
    score = row["FinalScore"]

    # User-priority setup must override the normal AVOID rule:
    # Price < 20DMA AND 20DMA > 50DMA > 200DMA
    if (
        row["Price"] < row["DMA20"]
        and row["DMA20"] > row["DMA50"] > row["DMA200"]
    ):
        return "STRONG BUY"

    price = row["Price"]
    dma200 = row["DMA200"]
    rising = row["DMA200_Rising"]

    if price < dma200 and not rising:
        return "AVOID"

    if score >= 80:
        return "STRONG BUY"
    if score >= 70:
        return "BUY"
    if score >= 60:
        return "ACCUMULATE"
    if score >= 45:
        return "WATCH"

    return "AVOID"


# ============================================================
# REASONS
# ============================================================

def generate_reason(row):
    reasons = []

    if (
        row["Price"] < row["DMA20"]
        and row["DMA20"] > row["DMA50"] > row["DMA200"]
    ):
        reasons.append(
            "Priority setup: below 20 DMA while 20>50>200 DMA"
        )
    elif (
        row["Price"] > row["DMA200"]
        and row["DMA200_Rising"]
    ):
        reasons.append("Healthy long-term trend")
    elif row["Price"] > row["DMA200"]:
        reasons.append("Above 200 DMA")
    else:
        reasons.append("Below 200 DMA")

    if row["Price"] < row["DMA20"]:
        reasons.append("Below 20 DMA")

    if row["Price"] < row["DMA50"]:
        reasons.append("Below 50 DMA")

    rsi = row["RSI14"]

    if 30 <= rsi < 35:
        reasons.append("RSI deeply oversold")
    elif 35 <= rsi < 40:
        reasons.append("RSI attractive")
    elif 40 <= rsi <= 50:
        reasons.append("RSI accumulation zone")
    elif rsi < 30:
        reasons.append("RSI oversold")

    if row["ADTV20_Cr"] >= 20:
        reasons.append("Strong liquidity")
    elif row["ADTV20_Cr"] >= 5:
        reasons.append("Good liquidity")

    delivery = row["Delivery20Pct"]

    if pd.notna(delivery):
        if delivery >= 60:
            reasons.append("Strong delivery")
        elif delivery >= 40:
            reasons.append("Healthy delivery")

    return "; ".join(reasons[:5])


# ============================================================
# DEPLOYMENT
# ============================================================

SIGNAL_WEIGHTS = {
    "STRONG BUY": 35,
    "BUY": 25,
    "ACCUMULATE": 12,
    "WATCH": 0,
    "AVOID": 0,
}


def calculate_deployment(df, capital, allocations):
    df = df.copy()

    df["BaseDeployWeight"] = (
        df["Signal"].map(SIGNAL_WEIGHTS).fillna(0)
    )

    df["AssetClass"] = df.apply(
        classify_asset,
        axis=1,
    )

    # A score-priority weight. The user's special setup gets
    # a substantial boost without changing the original score.
    df["PriorityWeight"] = np.where(
        (
            (df["Price"] < df["DMA20"])
            & (df["DMA20"] > df["DMA50"])
            & (df["DMA50"] > df["DMA200"])
        ),
        2.0,
        1.0,
    )

    df["RawWeight"] = (
        df["BaseDeployWeight"]
        * df["PriorityWeight"]
    )

    # Allocate capital separately by asset class.
    # This makes the user's SETTINGS allocation meaningful.
    df["SuggestedDeploy_%"] = 0.0

    for asset_class, target_pct in allocations.items():
        mask = (
            df["AssetClass"].eq(asset_class)
            & df["RawWeight"].gt(0)
        )

        total_class_weight = df.loc[
            mask, "RawWeight"
        ].sum()

        if total_class_weight <= 0:
            continue

        class_pct = (
            target_pct
            * df["RawWeight"]
            / total_class_weight
        )

        df.loc[mask, "SuggestedDeploy_%"] = class_pct

    # If a class has no qualifying signal, its money remains
    # unallocated rather than being silently transferred.
    df["SuggestedAmount"] = (
        capital
        * df["SuggestedDeploy_%"]
        / 100
    )

    return df


# ============================================================
# FINAL SCAN
# ============================================================

def scan():
    print("=" * 65)
    print("INDIA ETF BUY-LOW SCANNER")
    print("=" * 65)

    ss = connect()

    dashboard = get_sheet(
        ss, SHEETS["DASHBOARD"]
    )
    daily = get_sheet(
        ss, SHEETS["DAILY_SCAN"]
    )
    history = get_sheet(
        ss, SHEETS["HISTORY"]
    )
    excluded_ws = get_sheet(
        ss, SHEETS["EXCLUDED"]
    )
    settings = get_sheet(
        ss, SHEETS["SETTINGS"]
    )
    nav_ws = get_sheet(
        ss, SHEETS["NAV"]
    )

    capital = read_capital(dashboard)

    print(f"Deployment capital: ₹{capital:,.2f}")

    allocations = read_allocations(settings)

    print("Asset allocation:")
    for key, value in allocations.items():
        print(f"  {key}: {value:.2f}%")

    # --------------------------------------------------------
    # NAV
    # --------------------------------------------------------

    print("Reading NAV from source sheet...")
    nav_map = read_nav_source(
        gspread.authorize(credentials())
    )

    nav_count = update_destination_nav(
        gspread.authorize(credentials()),
        nav_ws,
        nav_map,
    )

    print(f"NAV records imported: {nav_count}")

    # --------------------------------------------------------
    # ETF UNIVERSE
    # --------------------------------------------------------

    print("Getting NSE ETF universe...")
    universe = get_nse_etf_universe()

    print(f"NSE ETFs found: {len(universe)}")

    candidates, excluded = apply_exclusions(
        universe
    )

    print(f"Candidates after exclusions: {len(candidates)}")

    # --------------------------------------------------------
    # TECHNICAL DATA
    # --------------------------------------------------------

    technical_results = []

    print("Scanning technical data...")

    for index, row in candidates.iterrows():
        result = analyse_etf(
            row["Symbol"],
            row["Underlying"],
        )

        if result is not None:
            technical_results.append(result)

        if index % 25 == 0:
            print(
                f"Technical progress: {index + 1}/"
                f"{len(candidates)}"
            )

    price_df = pd.DataFrame(
        technical_results
    )

    if price_df.empty:
        raise RuntimeError(
            "No ETF technical data was obtained."
        )

    print(
        f"Technical ETFs obtained: {len(price_df)}"
    )

    # --------------------------------------------------------
    # LIQUIDITY
    # --------------------------------------------------------

    low_liquidity = price_df[
        price_df["ADTV20_Cr"] < MIN_ADTV_CRORE
    ].copy()

    if not low_liquidity.empty:
        low_liquidity["Reason"] = "LOW LIQUIDITY"

        excluded = pd.concat(
            [
                excluded,
                low_liquidity[
                    ["Symbol", "Underlying", "Reason"]
                ].rename(
                    columns={
                        "Reason": "ExclusionReason"
                    }
                ),
            ],
            ignore_index=True,
        )

    price_df = price_df[
        price_df["ADTV20_Cr"] >= MIN_ADTV_CRORE
    ].copy()

    # --------------------------------------------------------
    # DELIVERY
    # --------------------------------------------------------

    print("Downloading NSE delivery data...")

    delivery = get_delivery_summary(
        price_df["Symbol"].tolist()
    )

    price_df = price_df.merge(
        delivery,
        on="Symbol",
        how="left",
    )

    if "Delivery20Pct" not in price_df.columns:
        price_df["Delivery20Pct"] = np.nan

    if "DeliveryDays" not in price_df.columns:
        price_df["DeliveryDays"] = 0

    # --------------------------------------------------------
    # SCORES
    # --------------------------------------------------------

    price_df["LiquidityScore"] = (
        price_df["ADTV20_Cr"]
        .apply(liquidity_score)
    )

    price_df["DeliveryScore"] = (
        price_df["Delivery20Pct"]
        .apply(delivery_score)
    )

    price_df["TrendScore"] = (
        price_df.apply(
            trend_score,
            axis=1,
        )
    )

    price_df["PullbackScore"] = (
        price_df.apply(
            pullback_score,
            axis=1,
        )
    )

    price_df["PriorityRuleScore"] = (
        price_df.apply(
            priority_rule_score,
            axis=1,
        )
    )

    price_df["RSIScore"] = (
        price_df["RSI14"]
        .apply(rsi_score)
    )

    price_df["ReversalScore"] = (
        price_df.apply(
            reversal_score,
            axis=1,
        )
    )

    # Original score remains the primary score.
    # PriorityRuleScore is included separately and also added
    # as a ranking boost.
    price_df["TechnicalScore"] = (
        price_df[
            [
                "TrendScore",
                "PullbackScore",
                "RSIScore",
                "ReversalScore",
            ]
        ].sum(axis=1)
    )

    price_df["FinalScore"] = (
        price_df[
            [
                "LiquidityScore",
                "DeliveryScore",
                "TechnicalScore",
            ]
        ].sum(axis=1)
    )

    # The requested setup receives the highest score priority.
    special = price_df["PriorityRuleScore"] > 0

    price_df.loc[
        special,
        "FinalScore"
    ] = np.maximum(
        price_df.loc[special, "FinalScore"],
        95,
    )

    price_df["Signal"] = (
        price_df.apply(
            determine_signal,
            axis=1,
        )
    )

    price_df["Reason"] = (
        price_df.apply(
            generate_reason,
            axis=1,
        )
    )

    # --------------------------------------------------------
    # DEPLOYMENT
    # --------------------------------------------------------

    price_df = calculate_deployment(
        price_df,
        capital,
        allocations,
    )

    # --------------------------------------------------------
    # DMA DISTANCES
    # --------------------------------------------------------

    for n in [20, 50, 100, 200]:
        price_df[f"Vs{n}DMA_%"] = (
            price_df["Price"]
            / price_df[f"DMA{n}"]
            - 1
        ) * 100

    # --------------------------------------------------------
    # RANK
    # --------------------------------------------------------

    price_df = (
        price_df.sort_values(
            [
                "FinalScore",
                "PriorityRuleScore",
                "ADTV20_Cr",
                "Delivery20Pct",
            ],
            ascending=[
                False,
                False,
                False,
                False,
            ],
            na_position="last",
        )
        .reset_index(drop=True)
    )

    price_df["Rank"] = (
        price_df.index + 1
    )

    scan_date = datetime.now(
        ZoneInfo("Asia/Kolkata")
    ).strftime("%Y-%m-%d")

    final_columns = [
        "Rank",
        "Symbol",
        "Underlying",
        "AssetClass",
        "Price",
        "FinalScore",
        "PriorityRuleScore",
        "Signal",
        "Reason",
        "SuggestedDeploy_%",
        "SuggestedAmount",
        "ADTV20_Cr",
        "Delivery20Pct",
        "DeliveryDays",
        "RSI14",
        "DMA20",
        "DMA50",
        "DMA100",
        "DMA200",
        "Vs20DMA_%",
        "Vs50DMA_%",
        "Vs100DMA_%",
        "Vs200DMA_%",
        "VolumeRatio",
        "DMA200_Rising",
        "LiquidityScore",
        "DeliveryScore",
        "TrendScore",
        "PullbackScore",
        "RSIScore",
        "ReversalScore",
        "TechnicalScore",
    ]

    final_df = price_df[
        final_columns
    ].copy()

    final_df.insert(
        0,
        "ScanDate",
        scan_date,
    )

    numeric_columns = [
        "Price",
        "FinalScore",
        "PriorityRuleScore",
        "SuggestedDeploy_%",
        "SuggestedAmount",
        "ADTV20_Cr",
        "Delivery20Pct",
        "RSI14",
        "DMA20",
        "DMA50",
        "DMA100",
        "DMA200",
        "Vs20DMA_%",
        "Vs50DMA_%",
        "Vs100DMA_%",
        "Vs200DMA_%",
        "VolumeRatio",
    ]

    for column in numeric_columns:
        final_df[column] = pd.to_numeric(
            final_df[column],
            errors="coerce",
        ).round(2)

    # --------------------------------------------------------
    # WRITE DAILY SCAN
    # --------------------------------------------------------

    print("Writing DAILY_SCAN...")
    write_sheet(
        daily,
        final_df,
        "A1",
        clear=True,
    )

    # --------------------------------------------------------
    # LAST REFRESH
    # --------------------------------------------------------

    refresh = datetime.now(
        ZoneInfo("Asia/Kolkata")
    ).strftime("%d-%m-%Y %H:%M:%S")

    dashboard.update(
        [[refresh]],
        "B6",
        value_input_option="USER_ENTERED",
    )

    # --------------------------------------------------------
    # HISTORY
    # --------------------------------------------------------

    print("Writing HISTORY...")

    history_values = history.get_all_values()

    if not history_values:
        write_sheet(
            history,
            final_df,
            "A1",
            clear=True,
        )
    else:
        headers = history_values[0]

        aligned = (
            final_df.reindex(
                columns=headers,
                fill_value="",
            )
        )

        rows = (
            aligned.replace(
                [np.inf, -np.inf],
                np.nan,
            )
            .fillna("")
            .astype(object)
            .values
            .tolist()
        )

        if rows:
            history.append_rows(
                rows,
                value_input_option="USER_ENTERED",
            )

    # --------------------------------------------------------
    # EXCLUDED
    # --------------------------------------------------------

    excluded = excluded.copy()

    if not excluded.empty:
        excluded.insert(
            0,
            "ScanDate",
            scan_date,
        )

        excluded = excluded.rename(
            columns={
                "ExclusionReason": "Reason"
            }
        )

    else:
        excluded = pd.DataFrame(
            columns=[
                "ScanDate",
                "Symbol",
                "Underlying",
                "Reason",
            ]
        )

    write_sheet(
        excluded_ws,
        excluded,
        "A1",
        clear=True,
    )

    # --------------------------------------------------------
    # SETTINGS STATUS
    # --------------------------------------------------------

    status_rows = [
        ["Parameter", "Value"],
        [
            "Last Scan",
            refresh,
        ],
        [
            "Total Deployment",
            capital,
        ],
        [
            "Minimum ADTV",
            MIN_ADTV_CRORE,
        ],
        [
            "Delivery Lookback",
            DELIVERY_LOOKBACK,
        ],
        [
            "RSI Period",
            RSI_PERIOD,
        ],
        [
            "Minimum History",
            MIN_HISTORY_DAYS,
        ],
        [
            "Priority Rule",
            "Price < 20DMA AND 20DMA > 50DMA > 200DMA",
        ],
    ]

    for key, value in allocations.items():
        status_rows.append(
            [f"{key} Allocation", value]
        )

    # Keep the user's allocation table at A:B and put
    # scanner status at D:E so the user can edit A:B.
    settings.update(
        status_rows,
        "D1",
        value_input_option="USER_ENTERED",
    )

    # --------------------------------------------------------
    # SUMMARY
    # --------------------------------------------------------

    signal_summary = (
        final_df["Signal"]
        .value_counts()
        .to_dict()
    )

    total_suggested = float(
        final_df["SuggestedAmount"].sum()
    )

    print("")
    print("=" * 65)
    print("SCAN COMPLETED SUCCESSFULLY")
    print("=" * 65)
    print(f"Scan date       : {scan_date}")
    print(f"Deployment      : ₹{capital:,.2f}")
    print(f"ETFs analysed   : {len(final_df)}")
    print(f"Excluded        : {len(excluded)}")
    print("-" * 65)
    print(
        "STRONG BUY      :",
        signal_summary.get("STRONG BUY", 0),
    )
    print(
        "BUY             :",
        signal_summary.get("BUY", 0),
    )
    print(
        "ACCUMULATE      :",
        signal_summary.get("ACCUMULATE", 0),
    )
    print(
        "WATCH           :",
        signal_summary.get("WATCH", 0),
    )
    print(
        "AVOID           :",
        signal_summary.get("AVOID", 0),
    )
    print("-" * 65)
    print(
        f"Suggested total : ₹{total_suggested:,.2f}"
    )
    print(
        f"Unallocated     : "
        f"₹{capital - total_suggested:,.2f}"
    )
    print("=" * 65)
    print(
        "Google Sheet:",
        ss.url,
    )


if __name__ == "__main__":
    scan()
