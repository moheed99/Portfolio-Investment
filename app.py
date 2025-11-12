# app.py
import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
from datetime import datetime, timedelta, date
from pathlib import Path
import io
import xlsxwriter

st.set_page_config(page_title="Auto Diversified Portfolio", layout="wide")
st.title("Auto Diversified Portfolio — Fixed Pakistani Stocks (Start 1 Sep 2025)")

# -------- USER PARAMETERS --------
TOTAL_INVESTMENT = 10_000_000
START_DATE_FIXED = date(2025, 9, 1)
BASE_TICKERS = ["LUCK", "HBL", "PSO", "ENGRO", "MCB", "OGDC", "FFC"]
SUFFIXES = ["", ".PK", ".PAK", ".KS", ".KSE", ".PA", ".PS"]
HISTORY_PATH = Path("portfolio_history.xlsx")

RUN = st.button("📈 Run / Generate portfolio_history.xlsx (with formulas)")

# ---------- UTILITY FUNCTIONS ----------
def try_find_symbol(base):
    """Try common suffixes and return the first yfinance-valid symbol."""
    for s in SUFFIXES:
        sym = base + s
        try:
            df = yf.download(sym, start="2025-09-01",
                             end=(datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d"),
                             interval="1d", progress=False, auto_adjust=True)
            if not df.empty and 'Close' in df.columns and df['Close'].dropna().shape[0] > 0:
                return sym
        except:
            continue
    return None

def fetch_closes(symbols, start_date, end_date):
    """Fetch daily close prices safely."""
    all_close_list = []
    col_names = []
    for sym in symbols:
        try:
            df = yf.download(sym, start=start_date.strftime("%Y-%m-%d"),
                             end=end_date.strftime("%Y-%m-%d"),
                             interval='1d', progress=False, auto_adjust=True)
            if df.empty:
                s = pd.Series(dtype=float)
            else:
                if 'Close' in df.columns:
                    s = df['Close'].copy()
                else:
                    num_cols = df.select_dtypes(include='number').columns
                    s = df[num_cols[0]].copy() if len(num_cols)>0 else pd.Series(dtype=float)
                s.index = pd.to_datetime(s.index).date
            all_close_list.append(s)
            col_names.append(sym)
        except Exception as e:
            st.warning(f"Failed fetching {sym}: {e}")
            s = pd.Series(dtype=float)
            all_close_list.append(s)
            col_names.append(sym)
    combined = pd.concat(all_close_list, axis=1)
    combined.columns = col_names
    combined = combined.sort_index()
    return combined

def col_letter(idx):
    """Convert 0-based index to Excel column letter."""
    s = ""
    while idx >= 0:
        s = chr(ord('A') + (idx % 26)) + s
        idx = idx // 26 - 1
    return s

# ---------- RUN BUTTON LOGIC ----------
if RUN:
    st.info("Finding valid Yahoo symbols...")
    symbol_map = {}
    failed = []
    for base in BASE_TICKERS:
        sym = try_find_symbol(base)
        if sym:
            symbol_map[base] = sym
        else:
            failed.append(base)
    if failed:
        st.warning(f"Couldn't find market data for: {', '.join(failed)}. Columns will remain empty.")
    st.write("Using symbols:", symbol_map)

    # Determine fetch start
    today = date.today()
    fetch_start = START_DATE_FIXED
    existing_daily = None
    if HISTORY_PATH.exists():
        try:
            prev_daily = pd.read_excel(HISTORY_PATH, sheet_name="Daily Data", index_col=0)
            if not prev_daily.empty:
                prev_dates = pd.to_datetime(prev_daily.index).date
                last_date = prev_dates[-1]
                fetch_start = last_date + timedelta(days=1)
                existing_daily = prev_daily
                st.write(f"Continuing from {fetch_start.isoformat()}.")
        except:
            st.warning("Unable to read existing history. Fetching from 1 Sep 2025.")

    if fetch_start > today:
        st.write("History already up-to-date. Regenerating Excel from existing file.")
        if HISTORY_PATH.exists():
            with open(HISTORY_PATH,"rb") as f:
                st.download_button("💾 Download existing portfolio_history.xlsx", data=f.read(),
                                   file_name="portfolio_history.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.stop()

    symbols_to_fetch = [symbol_map.get(b,b) for b in BASE_TICKERS]
    st.info(f"Fetching daily closes from {fetch_start} to {today} ...")
    closes = fetch_closes(symbols_to_fetch, fetch_start, today)

    # Merge with existing
    if existing_daily is not None:
        prev_close_df = existing_daily[[c for c in existing_daily.columns if "_Close" in c]].copy()
        prev_close_df.columns = [c.replace("_Close","") for c in prev_close_df.columns]
        prev_close_df.index = pd.to_datetime(prev_close_df.index).date
        combined_close = pd.concat([prev_close_df, closes])
        combined_close = combined_close[~combined_close.index.duplicated(keep='last')].sort_index()
        closes = combined_close[symbols_to_fetch].sort_index()

    closes = closes.dropna(how='all')
    if closes.empty:
        st.error("No price data available to write.")
        st.stop()
    st.write(f"Prepared data from {closes.index[0]} to {closes.index[-1]} — rows: {len(closes)}")
    st.dataframe(closes.tail(6))

    # ---------- Portfolio Weights ----------
    vol = closes.pct_change().std().fillna(0)
    vol_adj = vol.replace(0, vol[vol>0].min() if vol[vol>0].any() else 1.0)
    raw_weights = 1.0 / vol_adj
    raw_weights = raw_weights.clip(0.05, 10.0)
    weights = raw_weights / raw_weights.sum()
    latest_prices = closes.ffill().iloc[-1].fillna(0)
    allocation = weights * float(TOTAL_INVESTMENT)
    shares = allocation / latest_prices.replace(0,np.nan)
    shares = shares.fillna(0)

    # ---------- Generate Excel ----------
    out = io.BytesIO()
    workbook = xlsxwriter.Workbook(out, {'in_memory': True})
    ws_hold = workbook.add_worksheet("Holdings")
    ws_daily = workbook.add_worksheet("Daily Data")
    ws_sum = workbook.add_worksheet("Summary")

    # Holdings sheet
    hold_headers = ["Ticker (base)","SymbolUsed","Weight","AllocatedValue","LatestPrice","Shares","MarketValue_formula"]
    for j,h in enumerate(hold_headers):
        ws_hold.write(0,j,h)
    for i, base in enumerate(BASE_TICKERS):
        row = i+1
        sym = symbol_map.get(base, base)
        ws_hold.write(row,0,base)
        ws_hold.write(row,1,sym)
        ws_hold.write_number(row,2,float(weights.get(sym,0.0)))
        ws_hold.write_number(row,3,float(allocation.get(sym,0.0)))
        ws_hold.write_number(row,4,float(latest_prices.get(sym,0.0)))
        ws_hold.write_number(row,5,float(shares.get(sym,0.0)))
        pr_col = col_letter(4)
        sh_col = col_letter(5)
        ws_hold.write_formula(row,6,f"={pr_col}{row+1}*{sh_col}{row+1}")

    # Daily Data sheet
    headers = ["Date"]
    for base in BASE_TICKERS:
        sym = symbol_map.get(base, base)
        headers.append(f"{sym}_Close")
        headers.append(f"{sym}_MktVal")
    headers += ["PortfolioValue","DailyReturn","ProfitLoss"]
    for j,h in enumerate(headers):
        ws_daily.write(0,j,h)

    holdings_shares_cell = {symbol_map.get(b,b):f"Holdings!${col_letter(5)}${i+2}" for i,b in enumerate(BASE_TICKERS)}

    for r, dt in enumerate(closes.index):
        row = r+1
        ws_daily.write_datetime(row,0,datetime.combine(dt,datetime.min.time()))
        for i, base in enumerate(BASE_TICKERS):
            sym = symbol_map.get(base, base)
            close_col = 1+i*2
            mkt_col = close_col+1
            close_val = closes.iloc[r].get(sym,np.nan)
            if pd.isna(close_val):
                ws_daily.write_blank(row, close_col,None)
            else:
                ws_daily.write_number(row,close_col,float(close_val))
            close_cell_ref = f"{col_letter(close_col)}{row+1}"
            shares_cell_ref = holdings_shares_cell.get(sym)
            ws_daily.write_formula(row,mkt_col,f"=IF({close_cell_ref}=\"\",0,{close_cell_ref}*{shares_cell_ref})")
        mkt_cells = [f"{col_letter(1+i*2+1)}{row+1}" for i in range(len(BASE_TICKERS))]
        pv_col_idx = 1+len(BASE_TICKERS)*2
        ws_daily.write_formula(row,pv_col_idx,f"=SUM({','.join(mkt_cells)})")
        pv_cell = f"{col_letter(pv_col_idx)}{row+1}"
        if row==1:
            ws_daily.write_number(row,pv_col_idx+1,0.0)
        else:
            ws_daily.write_formula(row,pv_col_idx+1,f"=IF({col_letter(pv_col_idx)}{row}=0,0,({pv_cell}/{col_letter(pv_col_idx)}{row})-1)")
        ws_daily.write_formula(row,pv_col_idx+2,f"={pv_cell}-Summary!$B$1")

    # Summary sheet
    ws_sum.write(0,0,"Metric")
    ws_sum.write(0,1,"Value")
    ws_sum.write(1,0,"TotalInvestment")
    ws_sum.write_number(1,1,float(TOTAL_INVESTMENT))
    last_pv_cell = f"'Daily Data'!{col_letter(pv_col_idx)}{len(closes)+1}"
    ws_sum.write(2,0,"LatestValue")
    ws_sum.write_formula(2,1,f"={last_pv_cell}")
    ws_sum.write(3,0,"ProfitLoss")
    ws_sum.write_formula(3,1,"=B3-B2")
    ws_sum.write(4,0,"ProfitLossPct")
    ws_sum.write_formula(4,1,"=IF(B2=0,0,B4/B2)")

    ws_daily.freeze_panes(1,1)
    workbook.close()
    out.seek(0)
    with open(HISTORY_PATH,"wb") as f:
        f.write(out.read())
    with open(HISTORY_PATH,"rb") as f:
        st.download_button("💾 Download portfolio_history.xlsx", data=f.read(),
                           file_name="portfolio_history.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.success(f"portfolio_history.xlsx generated successfully!")
