"""
POS Analysis Script
--------------------
Reads raw POS export files (no pivot tables needed), combines them,
and produces: Best Sellers, Rising/Declining, Stopped Selling, Form Factor Trends.

HOW TO USE:
1. Install requirements once:  pip install pandas openpyxl
2. Edit the CONFIG section below (file paths, sheet names, which LV = form factor,
   Sell-In vs Sell-Out).
3. Run:  python pos_analysis.py
4. Output file "POS_Analysis_Output.xlsx" will be created next to this script.
"""

import pandas as pd
import numpy as np
import glob
import os

# ============ CONFIG - EDIT THIS SECTION ============

# List your raw POS files here (xlsx or csv). You can also use glob to grab
# every file in a folder automatically, e.g.:
# file_paths = glob.glob(r"C:\Users\yourname\Documents\POS_files\*.xlsx")
file_paths = [
    r"C:\path\to\2023_pos.xlsx",
    r"C:\path\to\2024_pos.xlsx",
    r"C:\path\to\2025_pos.xlsx",
]

# If your Excel files have multiple sheets and the data is on a specific one,
# set the sheet name here. If it's always the first/only sheet, leave as 0.
SHEET_NAME = 0

# Which column represents "form factor"? Based on your columns, likely LV2 or LV3.
FORM_FACTOR_COL = "LV3"

# Which Sell-In/Out value counts as "actual sales" for this analysis?
# Set to "Out", "In", or None to include both without filtering.
SELL_IN_OUT_FILTER = "Out"

# Output file name
OUTPUT_FILE = "POS_Analysis_Output.xlsx"

# How many recent months count as "recent" for trend / stopped-selling logic
RECENT_MONTHS = 3

# =====================================================


def load_and_combine(file_paths, sheet_name):
    """Read all raw files and stack them into one long dataframe."""
    dfs = []
    for path in file_paths:
        print(f"Reading {path} ...")
        if path.lower().endswith(".csv"):
            df = pd.read_csv(path)
        else:
            df = pd.read_excel(path, sheet_name=sheet_name)
        df["__source_file"] = os.path.basename(path)
        dfs.append(df)
    combined = pd.concat(dfs, ignore_index=True)
    print(f"Combined shape: {combined.shape}")
    return combined


def clean_and_prepare(df):
    """Standardize columns, filter Sell-In/Out, build a Year-Month sort key."""
    df = df.copy()

    # Strip whitespace from column names just in case
    df.columns = [str(c).strip() for c in df.columns]

    # Filter Sell-In/Out if configured
    if SELL_IN_OUT_FILTER is not None and "Sell-In/Out" in df.columns:
        before = len(df)
        df = df[df["Sell-In/Out"].astype(str).str.strip().str.lower()
                 == SELL_IN_OUT_FILTER.lower()]
        print(f"Filtered Sell-In/Out == '{SELL_IN_OUT_FILTER}': {before} -> {len(df)} rows")

    # Make sure QTY is numeric
    df["QTY"] = pd.to_numeric(df["QTY"], errors="coerce").fillna(0)

    # Build a proper Year-Month field. Handles Month as name ("Jan") or number (1).
    month_map = {
        "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
        "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12
    }

    def to_month_num(m):
        if pd.isna(m):
            return np.nan
        m_str = str(m).strip().lower()[:3]
        if m_str in month_map:
            return month_map[m_str]
        try:
            return int(m)
        except ValueError:
            return np.nan

    df["_month_num"] = df["Month"].apply(to_month_num)
    df["_year_num"] = pd.to_numeric(df["Year"], errors="coerce")

    df = df.dropna(subset=["_year_num", "_month_num"])
    df["_year_num"] = df["_year_num"].astype(int)
    df["_month_num"] = df["_month_num"].astype(int)

    df["YearMonth"] = pd.to_datetime(
        dict(year=df["_year_num"], month=df["_month_num"], day=1)
    )

    return df


def build_sku_month_table(df):
    """Aggregate to one row per ITEMCODE x YearMonth."""
    group_cols = ["ITEMCODE", "ITEMNAME", FORM_FACTOR_COL, "YearMonth"]
    group_cols = [c for c in group_cols if c in df.columns]  # safety
    agg = df.groupby(group_cols, as_index=False)["QTY"].sum()
    return agg


def pivot_wide(agg):
    """Turn long SKU x month table into wide SKU-by-month table."""
    wide = agg.pivot_table(
        index=["ITEMCODE", "ITEMNAME", FORM_FACTOR_COL],
        columns="YearMonth",
        values="QTY",
        aggfunc="sum",
        fill_value=0,
    )
    wide = wide.sort_index(axis=1)  # ensure months in chronological order
    wide.columns = [c.strftime("%Y-%m") for c in wide.columns]
    wide = wide.reset_index()
    return wide


def best_sellers(wide, month_cols):
    out = wide.copy()
    out["Total_Units"] = out[month_cols].sum(axis=1)
    out = out.sort_values("Total_Units", ascending=False)
    return out[["ITEMCODE", "ITEMNAME", FORM_FACTOR_COL, "Total_Units"] + month_cols]


def rising_declining(wide, month_cols, recent_n=3):
    out = wide.copy()
    if len(month_cols) < recent_n * 2:
        recent_n = max(1, len(month_cols) // 2)

    recent_cols = month_cols[-recent_n:]
    prior_cols = month_cols[-2 * recent_n:-recent_n]

    out["Recent_Sum"] = out[recent_cols].sum(axis=1)
    out["Prior_Sum"] = out[prior_cols].sum(axis=1)

    def pct_change(row):
        if row["Prior_Sum"] == 0:
            return np.nan if row["Recent_Sum"] == 0 else np.inf
        return (row["Recent_Sum"] - row["Prior_Sum"]) / row["Prior_Sum"]

    out["Pct_Change"] = out.apply(pct_change, axis=1)
    out = out.sort_values("Pct_Change", ascending=False)
    return out[["ITEMCODE", "ITEMNAME", FORM_FACTOR_COL,
                "Prior_Sum", "Recent_Sum", "Pct_Change"]]


def stopped_selling(wide, month_cols, recent_n=2):
    out = wide.copy()
    recent_cols = month_cols[-recent_n:]
    earlier_cols = month_cols[:-recent_n]

    out["Recent_Sum"] = out[recent_cols].sum(axis=1)
    out["Earlier_Sum"] = out[earlier_cols].sum(axis=1) if earlier_cols else 0

    flagged = out[(out["Recent_Sum"] == 0) & (out["Earlier_Sum"] > 0)]
    flagged = flagged.sort_values("Earlier_Sum", ascending=False)
    return flagged[["ITEMCODE", "ITEMNAME", FORM_FACTOR_COL, "Earlier_Sum", "Recent_Sum"]]


def form_factor_trend(agg):
    ff = agg.groupby([FORM_FACTOR_COL, "YearMonth"], as_index=False)["QTY"].sum()
    wide_ff = ff.pivot_table(index=FORM_FACTOR_COL, columns="YearMonth",
                              values="QTY", aggfunc="sum", fill_value=0)
    wide_ff = wide_ff.sort_index(axis=1)
    wide_ff.columns = [c.strftime("%Y-%m") for c in wide_ff.columns]
    wide_ff["Total_Units"] = wide_ff.sum(axis=1)
    wide_ff = wide_ff.sort_values("Total_Units", ascending=False)
    return wide_ff.reset_index()


def main():
    raw = load_and_combine(file_paths, SHEET_NAME)
    clean = clean_and_prepare(raw)
    agg = build_sku_month_table(clean)
    wide = pivot_wide(agg)

    month_cols = [c for c in wide.columns
                  if c not in ["ITEMCODE", "ITEMNAME", FORM_FACTOR_COL]]

    bs = best_sellers(wide, month_cols)
    rd = rising_declining(wide, month_cols, recent_n=RECENT_MONTHS)
    ss = stopped_selling(wide, month_cols, recent_n=2)
    ff = form_factor_trend(agg)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        bs.to_excel(writer, sheet_name="Best Sellers", index=False)
        rd.to_excel(writer, sheet_name="Rising-Declining", index=False)
        ss.to_excel(writer, sheet_name="Stopped Selling", index=False)
        ff.to_excel(writer, sheet_name="Form Factor Trend", index=False)
        wide.to_excel(writer, sheet_name="SKU x Month (raw)", index=False)

    print(f"\nDone! Output written to: {os.path.abspath(OUTPUT_FILE)}")


if __name__ == "__main__":
    main()
