"""
POS Analysis Script (YEARLY VERSION)
-------------------------------------
Reads raw POS export files (no pivot tables needed), combines them,
and produces: Best Sellers, Rising/Declining, Stopped Selling, Form Factor Trends
-- all aggregated by YEAR instead of by month.

HOW TO USE:
1. Install requirements once:  pip install pandas openpyxl
2. Edit the CONFIG section below (file paths, sheet names, which LV = form factor,
   Sell-In vs Sell-Out).
3. Run:  python pos_analysis_yearly.py
4. Output file "POS_Analysis_Output_Yearly.xlsx" will be created next to this script.
"""

import pandas as pd
import numpy as np
import glob
import os

# ============ CONFIG - EDIT THIS SECTION ============

file_paths = [
    r"C:\path\to\2023_pos.xlsx",
    r"C:\path\to\2024_pos.xlsx",
    r"C:\path\to\2025_pos.xlsx",
]

SHEET_NAME = 0

# Which column represents "form factor"? Based on your columns, likely LV2 or LV3.
FORM_FACTOR_COL = "LV3"

# Which Sell-In/Out value counts as "actual sales" for this analysis?
SELL_IN_OUT_FILTER = "Out"

OUTPUT_FILE = "POS_Analysis_Output_Yearly.xlsx"

# How many recent YEARS count as "recent" for stopped-selling logic.
# (e.g. 1 means: flag anything with 0 units in the most recent year
#  but > 0 units in any prior year)
RECENT_YEARS_FOR_STOPPED = 1

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
    """Standardize columns, filter Sell-In/Out, and build a clean Year field."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    if SELL_IN_OUT_FILTER is not None and "Sell-In/Out" in df.columns:
        before = len(df)
        df = df[df["Sell-In/Out"].astype(str).str.strip().str.lower()
                 == SELL_IN_OUT_FILTER.lower()]
        print(f"Filtered Sell-In/Out == '{SELL_IN_OUT_FILTER}': {before} -> {len(df)} rows")

    df["QTY"] = pd.to_numeric(df["QTY"], errors="coerce").fillna(0)

    # We no longer need Month at all for this version - only Year.
    df["_year_num"] = pd.to_numeric(df["Year"], errors="coerce")
    df = df.dropna(subset=["_year_num"])
    df["_year_num"] = df["_year_num"].astype(int)
    df["Year"] = df["_year_num"]  # normalized, clean integer Year column

    return df


def build_sku_year_table(df):
    """Aggregate to one row per ITEMCODE x Year (instead of x Month)."""
    group_cols = ["ITEMCODE", "ITEMNAME", FORM_FACTOR_COL, "Year"]
    group_cols = [c for c in group_cols if c in df.columns]
    agg = df.groupby(group_cols, as_index=False)["QTY"].sum()
    return agg


def pivot_wide(agg):
    """Turn long SKU x year table into wide SKU-by-year table."""
    wide = agg.pivot_table(
        index=["ITEMCODE", "ITEMNAME", FORM_FACTOR_COL],
        columns="Year",
        values="QTY",
        aggfunc="sum",
        fill_value=0,
    )
    wide = wide.sort_index(axis=1)  # ensure years are in chronological order
    wide.columns = [str(int(c)) for c in wide.columns]  # e.g. "2023", "2024"
    wide = wide.reset_index()
    return wide


def best_sellers(wide, year_cols):
    out = wide.copy()
    out["Total_Units"] = out[year_cols].sum(axis=1)
    out = out.sort_values("Total_Units", ascending=False)
    return out[["ITEMCODE", "ITEMNAME", FORM_FACTOR_COL, "Total_Units"] + year_cols]


def rising_declining(wide, year_cols):
    """
    Year-over-year view. For every pair of consecutive years, adds a
    '<year>_vs_<prev_year>_%chg' column, then sorts by the most recent
    year-over-year % change so you can immediately see what's rising
    and what's declining right now.
    """
    out = wide.copy()

    if len(year_cols) < 2:
        raise ValueError("Need at least 2 years of data to compute rising/declining.")

    pct_cols = []
    for i in range(1, len(year_cols)):
        prev_col, cur_col = year_cols[i - 1], year_cols[i]
        pct_col = f"{cur_col}_vs_{prev_col}_%chg"

        def pct_change(row, prev_col=prev_col, cur_col=cur_col):
            prev_val, cur_val = row[prev_col], row[cur_col]
            if prev_val == 0:
                return np.nan if cur_val == 0 else np.inf
            return (cur_val - prev_val) / prev_val

        out[pct_col] = out.apply(pct_change, axis=1)
        pct_cols.append(pct_col)

    latest_pct_col = pct_cols[-1]
    out = out.sort_values(latest_pct_col, ascending=False)

    return out[["ITEMCODE", "ITEMNAME", FORM_FACTOR_COL] + year_cols + pct_cols]


def stopped_selling(wide, year_cols, recent_n=1):
    """Items with sales in an earlier year but 0 in the most recent year(s)."""
    out = wide.copy()
    recent_cols = year_cols[-recent_n:]
    earlier_cols = year_cols[:-recent_n]

    out["Recent_Sum"] = out[recent_cols].sum(axis=1)
    out["Earlier_Sum"] = out[earlier_cols].sum(axis=1) if earlier_cols else 0

    flagged = out[(out["Recent_Sum"] == 0) & (out["Earlier_Sum"] > 0)]
    flagged = flagged.sort_values("Earlier_Sum", ascending=False)
    return flagged[["ITEMCODE", "ITEMNAME", FORM_FACTOR_COL, "Earlier_Sum", "Recent_Sum"]]


def form_factor_trend(agg):
    ff = agg.groupby([FORM_FACTOR_COL, "Year"], as_index=False)["QTY"].sum()
    wide_ff = ff.pivot_table(index=FORM_FACTOR_COL, columns="Year",
                              values="QTY", aggfunc="sum", fill_value=0)
    wide_ff = wide_ff.sort_index(axis=1)
    wide_ff.columns = [str(int(c)) for c in wide_ff.columns]
    wide_ff["Total_Units"] = wide_ff.sum(axis=1)
    wide_ff = wide_ff.sort_values("Total_Units", ascending=False)
    return wide_ff.reset_index()


def main():
    raw = load_and_combine(file_paths, SHEET_NAME)
    clean = clean_and_prepare(raw)
    agg = build_sku_year_table(clean)
    wide = pivot_wide(agg)

    year_cols = [c for c in wide.columns
                 if c not in ["ITEMCODE", "ITEMNAME", FORM_FACTOR_COL]]

    bs = best_sellers(wide, year_cols)
    rd = rising_declining(wide, year_cols)
    ss = stopped_selling(wide, year_cols, recent_n=RECENT_YEARS_FOR_STOPPED)
    ff = form_factor_trend(agg)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        bs.to_excel(writer, sheet_name="Best Sellers", index=False)
        rd.to_excel(writer, sheet_name="Rising-Declining", index=False)
        ss.to_excel(writer, sheet_name="Stopped Selling", index=False)
        ff.to_excel(writer, sheet_name="Form Factor Trend", index=False)
        wide.to_excel(writer, sheet_name="SKU x Year (raw)", index=False)

    print(f"\nDone! Output written to: {os.path.abspath(OUTPUT_FILE)}")


if __name__ == "__main__":
    main()
