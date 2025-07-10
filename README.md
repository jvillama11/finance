# finance
Weekly metrics
import yfinance as yf
import pandas as pd
import requests
import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from fredapi import Fred
import os

TODAY = datetime.date.today()
LAST_FRIDAY = TODAY - datetime.timedelta(days=(TODAY.weekday() + 3) % 7 + 2)

FRED_API_KEY = os.getenv("FRED_API_KEY")
fred = Fred(api_key=FRED_API_KEY)

# === EQUITY DATA ===
def get_equity_data():
    tickers = ["^GSPC", "^IXIC", "SOXX", "VNQ", "EEM", "VGK", "XLK", "XLF", "XLE", "XLY", "XLI", "XLB", "XLV", "XLU", "XLC"]
    data = []
    for t in tickers:
        ticker = yf.Ticker(t)
        hist = ticker.history(period="3y")
        if hist.empty or len(hist) < 252:
            continue
        row = {
            "Ticker": t,
            "Current": hist.iloc[-1]["Close"],
            "1W Change %": ((hist.iloc[-1]["Close"] / hist.iloc[-6]["Close"] - 1) * 100),
            "1M Change %": ((hist.iloc[-1]["Close"] / hist.iloc[-21]["Close"] - 1) * 100),
            "1Y Change %": ((hist.iloc[-1]["Close"] / hist.iloc[-252]["Close"] - 1) * 100),
            "3Y Change %": ((hist.iloc[-1]["Close"] / hist.iloc[0]["Close"] - 1) * 100)
        }
        data.append(row)
    return pd.DataFrame(data)

# === FIXED INCOME DATA ===
def get_fixed_income_data():
    series = {
        "2Y UST": "DGS2",
        "10Y UST": "DGS10",
        "30Y UST": "DGS30",
        "Fed Funds 12M Fwd": "EFFR"  # Placeholder; refine using CME data if needed
    }
    dates = [0, 7, 30, 365, 3 * 365]
    labels = ["Current", "1W Ago", "1M Ago", "1Y Ago", "3Y Ago"]
    data = []
    for name, code in series.items():
        row = {"Instrument": name}
        for d, label in zip(dates, labels):
            date = TODAY - datetime.timedelta(days=d)
            try:
                val = fred.get_series(code, date).dropna()
                row[label] = val.iloc[-1]
            except:
                row[label] = None
        data.append(row)
    return pd.DataFrame(data)

# === CURRENCY & COMMODITY DATA ===
def get_currency_commodity_data():
    tickers = {
        "DXY": "DX-Y.NYB",
        "USD/EUR": "EURUSD=X",
        "USD/CHF": "CHFUSD=X",
        "USD/MXN": "MXNUSD=X",
        "USD/JPY": "JPY=X",
        "USD/COP": "COPUSD=X",
        "Gold": "GC=F",
        "Silver": "SI=F",
        "Brent Oil": "BZ=F",
        "Nat Gas": "NG=F"
    }
    data = []
    for name, ticker in tickers.items():
        tk = yf.Ticker(ticker)
        hist = tk.history(period="3y")
        if hist.empty or len(hist) < 252:
            continue
        row = {
            "Asset": name,
            "Current": hist.iloc[-1]["Close"],
            "1M Change %": ((hist.iloc[-1]["Close"] / hist.iloc[-21]["Close"] - 1) * 100),
            "1Y Change %": ((hist.iloc[-1]["Close"] / hist.iloc[-252]["Close"] - 1) * 100),
            "3Y Change %": ((hist.iloc[-1]["Close"] / hist.iloc[0]["Close"] - 1) * 100)
        }
        data.append(row)
    return pd.DataFrame(data)

# === CRYPTO DATA ===
def get_crypto_data():
    ids = ["bitcoin", "ethereum", "solana"]
    data = []
    for cid in ids:
        url = f"https://api.coingecko.com/api/v3/coins/{cid}?localization=false"
        r = requests.get(url)
        if r.status_code != 200:
            continue
        d = r.json()
        market_data = d["market_data"]
        row = {
            "Asset": d["name"],
            "Current Price": market_data["current_price"]["usd"],
            "1W Ago": market_data["price_change_percentage_7d_in_currency"]["usd"],
            "1M Ago": market_data["price_change_percentage_30d_in_currency"]["usd"],
            "1Y Ago": market_data["price_change_percentage_1y_in_currency"]["usd"],
            "Market Cap": market_data["market_cap"]["usd"],
            "Volume (24h)": market_data["total_volume"]["usd"],
            "Fees": d.get("fees") or "N/A",
            "TVL": d.get("total_value_locked") or "N/A"
        }
        data.append(row)
    return pd.DataFrame(data)

# === OPTIONS DATA ===
def get_options_data():
    tickers = ["SPY", "QQQ"]
    data = []
    for t in tickers:
        tk = yf.Ticker(t)
        exps = tk.options[:2]  # Approximate next month & quarter
        for i, exp in enumerate(exps):
            calls = tk.option_chain(exp).calls
            total_oi_value = (calls["openInterest"] * calls["strike"] * 100).sum()
            iv = calls["impliedVolatility"].mean() * 100
            row = {
                "Underlying": t,
                "Cycle": f"Cycle {i+1}",
                "Avg IV %": round(iv, 2),
                "Total OI Value": round(total_oi_value, 2)
            }
            data.append(row)
    return pd.DataFrame(data)

# === ECONOMIC INDICATORS ===
def get_economic_data():
    indicators = {
        "CPI YoY": "CPIAUCNS",
        "PPI YoY": "PPIACO",
        "Unemployment Rate": "UNRATE",
        "Retail Sales YoY": "RSAFS",
        "Government Debt (Total)": "GFDEBTN",
        "Debt to GDP": "GFDGDPA188S",
        "Home Sales (LTM)": "HSN1F",
        "Housing Permits (LTM)": "PERMIT",
        "PMI": "NAPMPI"
    }
    data = []
    for name, code in indicators.items():
        try:
            series = fred.get_series(code).dropna()
            latest_date = series.index[-1].date()
            latest_value = series.iloc[-1]
            data.append({"Indicator": name, "Value": latest_value, "Date": latest_date})
        except:
            data.append({"Indicator": name, "Value": None, "Date": None})
    return pd.DataFrame(data)

# === EXCEL WRITING ===
def write_to_excel(sheets):
    wb = Workbook()
    for sheet_name, df in sheets.items():
        ws = wb.create_sheet(title=sheet_name)
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
    del wb["Sheet"]
    file_name = f"Weekly_Financial_Report_{TODAY}.xlsx"
    wb.save(file_name)
    return file_name

# === MAIN FUNCTION ===
def main():
    equity_df = get_equity_data()
    fixed_income_df = get_fixed_income_data()
    currency_commodity_df = get_currency_commodity_data()
    crypto_df = get_crypto_data()
    options_df = get_options_data()
    economic_df = get_economic_data()
    sheets = {
        "Equities": equity_df,
        "Fixed_Income": fixed_income_df,
        "Currencies_Commodities": currency_commodity_df,
        "Crypto": crypto_df,
        "Options": options_df,
        "Economics": economic_df
    }
    file_path = write_to_excel(sheets)
    print(f"Report generated: {file_path}")

if __name__ == "__main__":
    main()
