# app.py
import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
from datetime import datetime, timedelta, date
from pathlib import Path
import io
import xlsxwriter

st.set_page_config(page_title="Auto Diversified Portfolio (Excel formulas)", layout="wide")
st.title("Auto Diversified Portfolio — Fixed Basket (Start 1 Sep 2025)")

# -------- USER PARAMETERS (fixed as requested) --------
TOTAL_INVESTMENT = st.number_input("Total investment (PKR)", value=10_000_000, step=10_000)
START_DATE_FIXED = date(2025, 9, 1)

# Fixed diversified basket (you said you won't change — modify here only if needed)
# Common Pakistani large-caps across sectors
BASE_TICKERS = ["LUCK", "HBL", "PSO", "ENGRO", "MCB", "OGDC", "FFC"]

# Tickers suffixes to try on yfinance until a symbol returns data
SUFFIXES = ["", ".PK", ".PAK", ".KS", ".KSE", ".PA", ".PS"]

st.markdown("This app uses a fixed basket (LUCK, HBL, PSO, ENGRO, MCB, OGDC, FFC). It automatically finds workable Yahoo tickers by trying common suffixes if needed.")

RUN = st.button("📈 Run / Generate portfolio_history.xlsx (with Excel formulas)")

HISTORY_PATH = Path("portfolio_history.xlsx")

def try_find_symbol(base):
    """Try common suffixes and return the first yfinance-valid symbol (with at least one price)."""
    for s in SUFFIXES:
        sym = base + s
        try:
            df = yf.download(sym, start="2025-09-01", end=(datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d"), interval="1d", progress=False, auto_adjust=True)
            if not df.empty and 'Close' in df.columns and df['Close'].dropna().shape[0] > 0:
                return sym
        except Exception:
            continue
    return None

def fetch_closes(symbols, start_date, end_date):
    """Fetch daily close prices for a list of symbols. Returns DataFrame indexed by date with columns per symbol (column names = given symbols)."""
    all_close = {}
    for sym in symbols:
        try:
            df = yf.download(sym, start=start_date.strftime("%Y-%m-%d"), end=end_date.strftime("%Y-%m-%d"), interval='1d', progress=False, auto_adjust=True)
            if df.empty:
                all_close[sym] = pd.Series(dtype=float)
            else:
                # If multiindex, use 'Close'
                if isinstance(df.columns, pd.MultiIndex):
                    close = df['Close']
                else:
                    if 'Close' in df.columns:
                        close = df['Close']
                    else:
                        # fallback to first numeric column
                        close = df.iloc[:, 0]
                close.index = pd.to_datetime(close.index).date
                all_close[sym] = close
        except Exception:
            all_close[sym] = pd.Series(dtype=float)
    # Combine into DataFrame
    combined = pd.DataFrame(all_close)
    combined = combined.sort_index()
    return combined

def col_letter(idx):
    s = ""
    while idx >= 0:
        s = chr(ord('A') + (idx % 26)) + s
        idx = idx // 26 - 1
    return s

if RUN:
    st.info("Locating best-available Yahoo symbols for each base ticker...")
    symbol_map = {}
    failed = []
    for base in BASE_TICKERS:
        sym = try_find_symbol(base)
        if sym:
            symbol_map[base] = sym
        else:
            failed.append(base)

    if failed:
        st.warning(f"Couldn't find market data for: {', '.join(failed)}. They will be included but close columns will be blank.")
    st.write("Using symbols:", symbol_map)

    # Determine start date for fetch based on history file
    today = date.today()
    fetch_start = START_DATE_FIXED
    existing_daily = None
    if HISTORY_PATH.exists():
        try:
            prev_daily = pd.read_excel(HISTORY_PATH, sheet_name="Daily Data", index_col=0)
            if not prev_daily.empty:
                # index should be date strings — take last date and continue next day
                prev_dates = pd.to_datetime(prev_daily.index).date
                last_date = prev_dates[-1]
                fetch_start = last_date + timedelta(days=1)
                existing_daily = prev_daily
                st.write(f"Existing history found. Continuing from {fetch_start.isoformat()}.")
        except Exception:
            st.warning("Unable to read existing history. Will fetch from 1 Sep 2025.")

    if fetch_start > today:
        st.write("No new dates to fetch (history already up-to-date). Will still regenerate Excel from existing history.")
        # If history exists, just re-save it with formulas and offer download
        if HISTORY_PATH.exists():
            with open(HISTORY_PATH, "rb") as f:
                data_bytes = f.read()
            st.success("Existing portfolio_history.xlsx ready.")
            st.download_button("💾 Download existing portfolio_history.xlsx", data=data_bytes, file_name="portfolio_history.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("No history exists and no new days to fetch. Nothing to do.")
        st.stop()

    # Build symbol list to fetch (use found symbols; if some bases failed, include their base as column but will be empty)
    symbols_to_fetch = []
    for base in BASE_TICKERS:
        symbols_to_fetch.append(symbol_map.get(base, base))  # if not found, use base as-is (will be blank)

    st.info(f"Fetching daily closes from {fetch_start} to {today} ... This may take a moment.")
    closes = fetch_closes(symbols_to_fetch, fetch_start, today)
    if closes.empty and existing_daily is None:
        st.error("No price data available for the chosen symbols/dates. Check connectivity or ticker mapping.")
        st.stop()

    # If we have existing_daily, we want to combine: existing dates (older) + new closes
    if existing_daily is not None:
        # Attempt to reconstruct existing close-only DataFrame by selecting columns that match our symbols
        # existing_daily columns likely include "<base>_Close" and "<base>_MktVal" etc if generated by this app.
        # We'll try to extract previous close columns heuristically:
        prev_close_df = None
        try:
            # If columns named exactly as symbols, pick them
            intersect_cols = [c for c in existing_daily.columns if c in symbols_to_fetch]
            if len(intersect_cols) == len(symbols_to_fetch):
                prev_close_df = existing_daily[intersect_cols].copy()
                prev_close_df.index = pd.to_datetime(prev_close_df.index).date
            else:
                # Attempt columns that end with "_Close"
                close_cols = [c for c in existing_daily.columns if c.endswith("_Close")]
                if close_cols and len(close_cols) >= 1:
                    # map close_cols to symbols by prefix
                    prev_close_df = existing_daily[close_cols].copy()
                    prev_close_df.columns = [c.replace("_Close", "") for c in close_cols]
                    prev_close_df.index = pd.to_datetime(prev_close_df.index).date
        except Exception:
            prev_close_df = None

        if prev_close_df is not None:
            # Align columns to our symbols (rename prefixes if needed)
            # We'll try to make columns equal to symbols_to_fetch
            # If names are bases, map them to the chosen symbols
            # For simplicity, concat and drop duplicates preferring new data
            combined_close = pd.concat([prev_close_df, closes])
            combined_close = combined_close[~combined_close.index.duplicated(keep='last')].sort_index()
            # Reindex to symbols_to_fetch (if some columns missing, add them)
            for sym in symbols_to_fetch:
                if sym not in combined_close.columns:
                    combined_close[sym] = np.nan
            closes = combined_close[symbols_to_fetch].sort_index()
        else:
            # No usable prev close — keep closes as is
            pass

    # Now closes holds new and (if applicable) old data merged; ensure sorted
    closes = closes.sort_index()
    # Drop rows with all NaN
    closes = closes.dropna(how='all')
    if closes.empty:
        st.error("After processing, there is no close data to write.")
        st.stop()

    st.write(f"Prepared close data from {closes.index[0]} to {closes.index[-1]} — rows: {len(closes)}")
    st.dataframe(closes.tail(6))

    # Determine weights: simple risk-balanced approach -> assign weights inverse to historical volatility (lower vol -> higher weight),
    # but cap extremes and then normalize. This yields "max return / min risk" heuristic without optimization complexity.
    # Compute vol using returns over available close history (use percent daily returns across entire fetched period).
    vol = closes.pct_change().std().fillna(0)
    # avoid zero vol causing inf weight
    vol_adj = vol.replace(0, vol[vol>0].min() if vol[vol>0].any() else 1.0)
    raw_weights = 1.0 / vol_adj  # inverse vol
    # Cap extreme values and normalize
    raw_weights = raw_weights.clip(0.05, 10.0)
    weights = raw_weights / raw_weights.sum()
    # If any symbol had no data, weight will be NaN; replace with equal share among valid symbols
    if weights.isna().any():
        valid = weights.dropna()
        weights = weights.fillna(0)
        if valid.size > 0:
            weights.loc[weights==0] = (1.0 - valid.sum()) / (weights==0).sum() if (weights==0).sum()>0 else 0.0
        else:
            # fallback equal
            weights = pd.Series(1.0/len(symbols_to_fetch), index=symbols_to_fetch)

    # Determine latest available price per symbol (use forward/backfill)
    latest_prices = closes.ffill().iloc[-1]
    # If some latest_prices are NaN, set to a small number to avoid division by zero
    latest_prices = latest_prices.fillna(0.0)

    # Allocation & shares (shares written as values in Holdings)
    allocation = weights * float(TOTAL_INVESTMENT)
    shares = allocation / latest_prices.replace(0, np.nan)
    shares = shares.fillna(0)

    # Build Excel workbook in memory
    out = io.BytesIO()
    workbook = xlsxwriter.Workbook(out, {'in_memory': True})
    # Worksheets
    ws_hold = workbook.add_worksheet("Holdings")
    ws_daily = workbook.add_worksheet("Daily Data")
    ws_sum = workbook.add_worksheet("Summary")

    # -- Write Holdings --
    hold_headers = ["Ticker (base)", "SymbolUsed", "Weight", "AllocatedValue", "LatestPrice", "Shares", "MarketValue_formula"]
    for j, h in enumerate(hold_headers):
        ws_hold.write(0, j, h)
    for i, base in enumerate(BASE_TICKERS):
        row = i + 1
        sym = symbol_map.get(base, base)
        ws_hold.write(row, 0, base)
        ws_hold.write(row, 1, sym)
        ws_hold.write_number(row, 2, float(weights.get(sym, 0.0)))
        ws_hold.write_number(row, 3, float(allocation.get(sym, 0.0)))
        ws_hold.write_number(row, 4, float(latest_prices.get(sym, 0.0)))
        ws_hold.write_number(row, 5, float(shares.get(sym, 0.0)))
        # MarketValue formula = LatestPrice * Shares
        pr_col = col_letter(4)  # LatestPrice col (0-based index 4 -> letter 'E')
        sh_col = col_letter(5)  # Shares col -> 'F'
        mv_formula = f"={pr_col}{row+1}*{sh_col}{row+1}"
        ws_hold.write_formula(row, 6, mv_formula)

    # -- Write Daily Data headers --
    headers = ["Date"]
    for base in BASE_TICKERS:
        sym = symbol_map.get(base, base)
        headers.append(f"{sym}_Close")
        headers.append(f"{sym}_MktVal")
    headers += ["PortfolioValue", "DailyReturn", "ProfitLoss"]
    for j, h in enumerate(headers):
        ws_daily.write(0, j, h)

    # Map holdings shares cell references for formulas (Holdings sheet row numbers)
    holdings_shares_cell = {}
    for i, base in enumerate(BASE_TICKERS):
        row = 1 + i  # zero-based in code -> excel row number row+1
        cell = f"Holdings!${col_letter(5)}${row+1}"  # Shares column is index 5 -> letter
        holdings_shares_cell[ symbol_map.get(base, base) ] = cell

    # Write daily rows
    dates = list(closes.index)
    nrows = len(dates)
    for r, dt in enumerate(dates):
        row = r + 1  # excel row index (0 header)
        # Date
        ws_daily.write_datetime(row, 0, datetime.combine(dt, datetime.min.time()))
        # Per-symbol close and marketvalue formula
        for i, base in enumerate(BASE_TICKERS):
            sym = symbol_map.get(base, base)
            close_col = 1 + i*2
            mkt_col = close_col + 1
            close_val = closes.iloc[r].get(sym, np.nan)
            if pd.isna(close_val):
                ws_daily.write_blank(row, close_col, None)
            else:
                ws_daily.write_number(row, close_col, float(close_val))
            # Market value formula: =IF(<Close_cell>="",0,<Close_cell>*Holdings!$F${row_of_shares})
            close_cell_ref = f"{col_letter(close_col)}{row+1}"
            shares_cell_ref = holdings_shares_cell.get(sym, f"Holdings!${col_letter(5)}$2")
            mkt_formula = f"=IF({close_cell_ref}=\"\",0,{close_cell_ref}*{shares_cell_ref})"
            ws_daily.write_formula(row, mkt_col, mkt_formula)

        # PortfolioValue formula: sum of all MktVal cols in this row
        mkt_cells = []
        for i in range(len(BASE_TICKERS)):
            mkt_idx = 1 + i*2 + 1
            mkt_cells.append(f"{col_letter(mkt_idx)}{row+1}")
        pv_formula = f"=SUM({','.join(mkt_cells)})"
        pv_col_idx = 1 + len(BASE_TICKERS)*2
        ws_daily.write_formula(row, pv_col_idx, pv_formula)

        # DailyReturn formula: if previous pv exists then =(pv / pv_prev) -1 else 0
        pv_cell = f"{col_letter(pv_col_idx)}{row+1}"
        if row == 1:
            ws_daily.write_number(row, pv_col_idx+1, 0.0)
        else:
            pv_prev = f"{col_letter(pv_col_idx)}{row}"
            dr_formula = f"=IF({pv_prev}=0,0,({pv_cell}/{pv_prev})-1)"
            ws_daily.write_formula(row, pv_col_idx+1, dr_formula)

        # ProfitLoss formula: =PortfolioValue - Summary!$B$1 (we will put TotalInvestment at Summary B1)
        profit_formula = f"={pv_cell}-Summary!$B$1"
        ws_daily.write_formula(row, pv_col_idx+2, profit_formula)

    # -- Summary sheet --
    ws_sum.write(0, 0, "Metric")
    ws_sum.write(0, 1, "Value")
    ws_sum.write(1, 0, "TotalInvestment")
    ws_sum.write_number(0, 1, float(TOTAL_INVESTMENT))  # put total investment at B1 so Daily ProfitLoss formula matches
    ws_sum.write(1, 0, "LatestValue")
    # LatestValue -> last PortfolioValue cell in Daily Data
    last_pv_row_excel = 1 + nrows  # last data row number in Excel
    pv_col_idx = 1 + len(BASE_TICKERS)*2
    last_pv_cell = f"'Daily Data'!{col_letter(pv_col_idx)}{last_pv_row_excel}"
    ws_sum.write_formula(1, 1, f"={last_pv_cell}")
    ws_sum.write(2, 0, "ProfitLoss")
    ws_sum.write_formula(2, 1, "=B2-B1")  # LatestValue - TotalInvestment (B2-B1)
    ws_sum.write(3, 0, "ProfitLossPct")
    ws_sum.write_formula(3, 1, "=IF(B1=0,0,B3/B1)")

    # freeze panes for daily data
    ws_daily.freeze_panes(1, 1)

    # Finalize workbook
    workbook.close()
    out.seek(0)

    # Save/overwrite history file
    with open(HISTORY_PATH, "wb") as f:
        f.write(out.read())

    # Offer download
    with open(HISTORY_PATH, "rb") as f:
        data_bytes = f.read()
    st.success(f"portfolio_history.xlsx generated and saved ({HISTORY_PATH.resolve()})")
    st.download_button("💾 Download portfolio_history.xlsx", data=data_bytes, file_name="portfolio_history.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # Quick python-side summary for immediate view
    latest_values = (closes.ffill().iloc[-1] * shares).fillna(0)
    latest_portfolio_value = latest_values.sum()
    profit_loss = latest_portfolio_value - float(TOTAL_INVESTMENT)
    profit_loss_pct = profit_loss / float(TOTAL_INVESTMENT) if TOTAL_INVESTMENT != 0 else 0.0
    st.metric("Latest portfolio value (PY calc)", f"₨ {latest_portfolio_value:,.0f}", delta=f"{profit_loss_pct*100:.2f}%")
    df_hold = pd.DataFrame({
        "BaseTicker": BASE_TICKERS,
        "SymbolUsed": [symbol_map.get(b, b) for b in BASE_TICKERS],
        "Weight": [float(weights.get(symbol_map.get(b, b), 0.0)) for b in BASE_TICKERS],
        "AllocatedValue": [float(allocation.get(symbol_map.get(b, b), 0.0)) for b in BASE_TICKERS],
        "LatestPrice": [float(latest_prices.get(symbol_map.get(b, b), 0.0)) for b in BASE_TICKERS],
        "Shares": [float(shares.get(symbol_map.get(b, b), 0.0)) for b in BASE_TICKERS],
        "MarketValue": [float(latest_prices.get(symbol_map.get(b, b), 0.0) * shares.get(symbol_map.get(b, b), 0.0)) for b in BASE_TICKERS]
    })
    st.dataframe(df_hold.style.format({"AllocatedValue":"{:,}","LatestPrice":"{:.4f}","Shares":"{:.4f}","MarketValue":"{:,}"}))

st.markdown("Notes: Excel contains formulas for per-ticker MarketValue, PortfolioValue (sum), DailyReturn (row-to-row), ProfitLoss and Summary formulas. If a symbol couldn't be located on Yahoo Finance the close column will be blank — you can later edit the generated file or provide corrected mappings in the code.")
