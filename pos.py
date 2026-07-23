"""
POS Analysis Script (YEARLY VERSION + PRODUCT LIFECYCLE AWARE)
----------------------------------------------------------------
Reads raw POS export files, combines them, and produces:
  - Best Sellers          (lifetime AND recent-years ranking, one-hit-wonder flag)
  - Rising / Declining    (year-over-year, EXCLUDING new launches & discontinued items)
  - New Product Ramp      (items that just launched, judged on raw unit growth not %)
  - Stopped Selling        (items with prior sales but 0 in the recent year(s))
  - Product Lifecycle      (every item tagged New / Active / Discontinued / Intermittent)
  - Form Factor Trend      (units by form factor, per year)

Why this version is different from the plain yearly script:
Your data spans many years (e.g. 2017-2025), so items get launched and
discontinued constantly. A plain year-over-year %% change is misleading for:
  - NEW items    -> show as "infinite %% growth" (0 -> something), drowning out
                    genuinely important trends among steady sellers.
  - DISCONTINUED items -> just vanish from a %% ranking instead of being called out.
This version classifies every item first, then routes it to the right table.

HOW TO USE:
1. Install requirements once:  pip install pandas openpyxl
2. Edit the CONFIG section below.
3. Run:  python pos_analysis_yearly_v2.py
4. Output file "POS_Analysis_Output_Yearly.xlsx" will be created next to this script.
"""

import pandas as pd
import numpy as np
import os

# ============ CONFIG - EDIT THIS SECTION ============

file_paths = [
    r"C:\path\to\2017_pos.xlsx",
    r"C:\path\to\2018_pos.xlsx",
    # ... one entry per year/file, or use glob to grab a whole folder:
    # glob.glob(r"C:\Users\yourname\Documents\POS_files\*.xlsx")
]

SHEET_NAME = 0

# Which column represents "form factor"? Based on your columns, likely LV2 or LV3.
FORM_FACTOR_COL = "LV3"

# Which Sell-In/Out value counts as "actual sales" for this analysis?
SELL_IN_OUT_FILTER = "Out"

OUTPUT_FILE = "POS_Analysis_Output_Yearly.xlsx"

# An item is "New" if its FIRST year with sales falls within this many
# trailing years of the most recent year in the data (1 = launched this year).
NEW_ITEM_WINDOW_YEARS = 1

# An item is "Discontinued" if it has 0 units in this many trailing years
# but had sales before that.
RECENT_YEARS_FOR_STOPPED = 1

# Window size (in years) for the smoothed multi-year trend comparison,
# e.g. 3 = compare the last 3 years combined vs the 3 years before that.
MULTI_YEAR_WINDOW = 3

# How many recent years count as "recent" for the Best Sellers "recent" ranking.
RECENT_YEARS_FOR_BEST_SELLERS = 2

# Flag an item as a "one-hit wonder" if this fraction (or more) of its
# lifetime volume came from a single year.
ONE_HIT_WONDER_THRESHOLD = 0.70

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

    df["_year_num"] = pd.to_numeric(df["Year"], errors="coerce")
    df = df.dropna(subset=["_year_num"])
    df["_year_num"] = df["_year_num"].astype(int)
    df["Year"] = df["_year_num"]

    return df


def build_sku_year_table(df):
    """Aggregate to one row per ITEMCODE x Year."""
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
    wide = wide.sort_index(axis=1)
    wide.columns = [str(int(c)) for c in wide.columns]
    wide = wide.reset_index()
    return wide


def build_lifecycle_table(wide, year_cols):
    """
    Classify every item as New / Active / Discontinued / Intermittent,
    based on the pattern of zero vs non-zero years.
    """
    years_int = [int(y) for y in year_cols]
    most_recent_year = max(years_int)

    records = []
    for _, row in wide.iterrows():
        series = row[year_cols].astype(float)
        nonzero_years = [int(year_cols[i]) for i, v in enumerate(series) if v > 0]

        if not nonzero_years:
            first_year = last_year = None
            status = "Never Sold"
        else:
            first_year = min(nonzero_years)
            last_year = max(nonzero_years)

            if last_year < most_recent_year:
                # Not selling in the most recent year -> discontinued,
                # regardless of when it launched.
                status = "Discontinued"
            elif first_year > most_recent_year - NEW_ITEM_WINDOW_YEARS:
                # Launched within the "new" window and still selling.
                status = "New"
            else:
                # Selling in the most recent year, launched a while ago.
                # Check for gaps (years within its lifespan with 0 units).
                span_years = [y for y in years_int if first_year <= y <= last_year]
                zero_years_in_span = [y for y in span_years
                                       if row[str(y)] == 0]
                status = "Intermittent" if zero_years_in_span else "Active"

        records.append({
            "ITEMCODE": row["ITEMCODE"],
            "ITEMNAME": row["ITEMNAME"],
            FORM_FACTOR_COL: row[FORM_FACTOR_COL],
            "First_Year_Sold": first_year,
            "Last_Year_Sold": last_year,
            "Status": status,
        })

    return pd.DataFrame(records)


def best_sellers(wide, year_cols, lifecycle):
    out = wide.merge(lifecycle[["ITEMCODE", "Status"]], on="ITEMCODE", how="left")
    out["Total_Units"] = out[year_cols].sum(axis=1)

    recent_cols = year_cols[-RECENT_YEARS_FOR_BEST_SELLERS:]
    out["Recent_Years_Units"] = out[recent_cols].sum(axis=1)

    # One-hit-wonder flag: does a single year account for most of lifetime volume?
    out["Max_Single_Year"] = out[year_cols].max(axis=1)
    out["One_Hit_Wonder"] = np.where(
        out["Total_Units"] > 0,
        (out["Max_Single_Year"] / out["Total_Units"]) >= ONE_HIT_WONDER_THRESHOLD,
        False,
    )

    out = out.sort_values("Recent_Years_Units", ascending=False)
    cols = (["ITEMCODE", "ITEMNAME", FORM_FACTOR_COL, "Status",
             f"Recent_Years_Units", "Total_Units", "One_Hit_Wonder"] + year_cols)
    return out[cols]


def rising_declining(wide, year_cols, lifecycle):
    """
    Year-over-year %% change, restricted to items with status "Active" or
    "Intermittent" (i.e. genuinely comparable both years) so New launches
    and Discontinued items don't distort the ranking with fake +inf%% or
    silently drop out.
    """
    out = wide.merge(lifecycle[["ITEMCODE", "Status"]], on="ITEMCODE", how="left")

    if len(year_cols) < 2:
        raise ValueError("Need at least 2 years of data to compute rising/declining.")

    pct_cols = []
    for i in range(1, len(year_cols)):
        prev_col, cur_col = year_cols[i - 1], year_cols[i]
        pct_col = f"{cur_col}_vs_{prev_col}_%chg"

        def pct_change(row, prev_col=prev_col, cur_col=cur_col):
            prev_val, cur_val = row[prev_col], row[cur_col]
            if prev_val == 0:
                return np.nan  # can't compute a meaningful %% here
            return (cur_val - prev_val) / prev_val

        out[pct_col] = out.apply(pct_change, axis=1)
        pct_cols.append(pct_col)

    # Multi-year smoothed comparison (e.g. last 3 yrs vs prior 3 yrs)
    window = min(MULTI_YEAR_WINDOW, len(year_cols) // 2) if len(year_cols) >= 4 else 1
    recent_block = year_cols[-window:]
    prior_block = year_cols[-2 * window:-window] if len(year_cols) >= 2 * window else []

    out[f"Last_{window}yr_Sum"] = out[recent_block].sum(axis=1)
    out[f"Prior_{window}yr_Sum"] = out[prior_block].sum(axis=1) if prior_block else 0

    def block_pct_change(row):
        prev_val = row[f"Prior_{window}yr_Sum"]
        cur_val = row[f"Last_{window}yr_Sum"]
        if prev_val == 0:
            return np.nan
        return (cur_val - prev_val) / prev_val

    out[f"{window}yr_Block_%chg"] = out.apply(block_pct_change, axis=1)

    # Only Active/Intermittent items belong in the main rising/declining ranking.
    comparable = out[out["Status"].isin(["Active", "Intermittent"])].copy()
    latest_pct_col = pct_cols[-1]
    comparable = comparable.sort_values(latest_pct_col, ascending=False)

    keep_cols = (["ITEMCODE", "ITEMNAME", FORM_FACTOR_COL, "Status"] + year_cols
                 + pct_cols + [f"Prior_{window}yr_Sum", f"Last_{window}yr_Sum",
                                f"{window}yr_Block_%chg"])
    return comparable[keep_cols]


def new_product_ramp(wide, year_cols, lifecycle):
    """
    Items launched recently. Judged by raw unit growth (not %%, since prior
    year is 0 for these by definition) so you can see which launches are
    actually gaining traction.
    """
    out = wide.merge(lifecycle[["ITEMCODE", "Status", "First_Year_Sold"]],
                      on="ITEMCODE", how="left")
    new_items = out[out["Status"] == "New"].copy()

    if new_items.empty:
        return new_items[["ITEMCODE", "ITEMNAME", FORM_FACTOR_COL,
                           "First_Year_Sold"] + year_cols]

    last_col, prev_col = year_cols[-1], year_cols[-2] if len(year_cols) > 1 else None
    new_items["Units_This_Year"] = new_items[last_col]
    if prev_col:
        new_items["Unit_Growth_vs_Prior_Year"] = new_items[last_col] - new_items[prev_col]
    new_items = new_items.sort_values("Units_This_Year", ascending=False)

    cols = ["ITEMCODE", "ITEMNAME", FORM_FACTOR_COL, "First_Year_Sold",
            "Units_This_Year"] + year_cols
    return new_items[cols]


def stopped_selling(wide, year_cols, lifecycle, recent_n=1):
    """Items with sales in an earlier year but 0 in the most recent year(s)."""
    out = wide.merge(lifecycle[["ITEMCODE", "Status", "First_Year_Sold",
                                 "Last_Year_Sold"]], on="ITEMCODE", how="left")
    recent_cols = year_cols[-recent_n:]
    earlier_cols = year_cols[:-recent_n]

    out["Recent_Sum"] = out[recent_cols].sum(axis=1)
    out["Earlier_Sum"] = out[earlier_cols].sum(axis=1) if earlier_cols else 0

    flagged = out[(out["Status"] == "Discontinued")].copy()
    flagged = flagged.sort_values("Earlier_Sum", ascending=False)
    return flagged[["ITEMCODE", "ITEMNAME", FORM_FACTOR_COL, "First_Year_Sold",
                     "Last_Year_Sold", "Earlier_Sum", "Recent_Sum"]]


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

    lifecycle = build_lifecycle_table(wide, year_cols)

    bs = best_sellers(wide, year_cols, lifecycle)
    rd = rising_declining(wide, year_cols, lifecycle)
    npr = new_product_ramp(wide, year_cols, lifecycle)
    ss = stopped_selling(wide, year_cols, lifecycle, recent_n=RECENT_YEARS_FOR_STOPPED)
    ff = form_factor_trend(agg)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        bs.to_excel(writer, sheet_name="Best Sellers", index=False)
        rd.to_excel(writer, sheet_name="Rising-Declining", index=False)
        npr.to_excel(writer, sheet_name="New Product Ramp", index=False)
        ss.to_excel(writer, sheet_name="Stopped Selling", index=False)
        lifecycle.to_excel(writer, sheet_name="Product Lifecycle", index=False)
        ff.to_excel(writer, sheet_name="Form Factor Trend", index=False)
        wide.to_excel(writer, sheet_name="SKU x Year (raw)", index=False)

    print(f"\nDone! Output written to: {os.path.abspath(OUTPUT_FILE)}")
    print(f"Lifecycle breakdown:\n{lifecycle['Status'].value_counts()}")


if __name__ == "__main__":
    main()
