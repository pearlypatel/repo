"""
POS Analysis Script
--------------------
Analyzes POS data and produces:
1. Top 5 Best Sellers
2. Top 5 Rising Products
3. Top 5 Declining Products
4. Top 5 Products That Stopped Selling
5. Top 5 Form Factor Trends

Requirements:
pip install pandas openpyxl

Run:
python pos_analysis.py
"""

import pandas as pd
import numpy as np
import os


# ================= CONFIG =================

file_paths = [
    r"C:\path\to\2023_pos.xlsx",
    r"C:\path\to\2024_pos.xlsx",
    r"C:\path\to\2025_pos.xlsx",
]


SHEET_NAME = 0


# Change this after checking LV1/LV2/LV3/LV4
FORM_FACTOR_COL = "LV3"


# Use "Out" for sell-out analysis
# Use "In" for distributor shipments
# None = include both
SELL_IN_OUT_FILTER = "Out"


OUTPUT_FILE = "POS_Analysis_Output.xlsx"


RECENT_MONTHS = 3


TOP_N = 5


# =========================================


def load_and_combine(file_paths, sheet_name):

    dfs = []

    for path in file_paths:

        print(f"Reading {path}")

        if path.lower().endswith(".csv"):
            df = pd.read_csv(path)

        else:
            df = pd.read_excel(
                path,
                sheet_name=sheet_name
            )

        df["__source_file"] = os.path.basename(path)

        dfs.append(df)


    combined = pd.concat(
        dfs,
        ignore_index=True
    )

    print(
        "Combined rows:",
        len(combined)
    )

    return combined



def clean_and_prepare(df):

    df = df.copy()


    df.columns = [
        str(c).strip()
        for c in df.columns
    ]


    # Filter Sell In / Sell Out

    if SELL_IN_OUT_FILTER and "Sell-In/Out" in df.columns:

        before = len(df)

        df = df[
            df["Sell-In/Out"]
            .astype(str)
            .str.strip()
            .str.lower()
            ==
            SELL_IN_OUT_FILTER.lower()
        ]

        print(
            f"Sell filter: {before} -> {len(df)}"
        )



    df["QTY"] = pd.to_numeric(
        df["QTY"],
        errors="coerce"
    ).fillna(0)



    month_map = {

        "jan":1,
        "feb":2,
        "mar":3,
        "apr":4,
        "may":5,
        "jun":6,
        "jul":7,
        "aug":8,
        "sep":9,
        "oct":10,
        "nov":11,
        "dec":12

    }



    def convert_month(x):

        if pd.isna(x):
            return np.nan


        x = str(x).lower()[:3]


        if x in month_map:
            return month_map[x]


        try:
            return int(x)

        except:
            return np.nan



    df["_month_num"] = df["Month"].apply(
        convert_month
    )


    df["_year_num"] = pd.to_numeric(
        df["Year"],
        errors="coerce"
    )


    df = df.dropna(
        subset=[
            "_year_num",
            "_month_num"
        ]
    )


    df["_year_num"] = (
        df["_year_num"]
        .astype(int)
    )


    df["_month_num"] = (
        df["_month_num"]
        .astype(int)
    )



    df["YearMonth"] = pd.to_datetime(

        dict(
            year=df["_year_num"],
            month=df["_month_num"],
            day=1
        )

    )


    return df




def build_sku_month_table(df):

    cols = [
        "ITEMCODE",
        "ITEMNAME",
        FORM_FACTOR_COL,
        "YearMonth"
    ]


    agg = (
        df.groupby(cols)["QTY"]
        .sum()
        .reset_index()
    )


    return agg




def pivot_wide(agg):

    wide = agg.pivot_table(

        index=[
            "ITEMCODE",
            "ITEMNAME",
            FORM_FACTOR_COL
        ],

        columns="YearMonth",

        values="QTY",

        aggfunc="sum",

        fill_value=0
    )


    wide = wide.sort_index(axis=1)


    wide.columns = [
        x.strftime("%Y-%m")
        for x in wide.columns
    ]


    return wide.reset_index()




def best_sellers(
    wide,
    month_cols,
    top_n
):

    out = wide.copy()


    out["Total_Units"] = (
        out[month_cols]
        .sum(axis=1)
    )


    return (
        out.sort_values(
            "Total_Units",
            ascending=False
        )
        .head(top_n)
        [
            [
                "ITEMCODE",
                "ITEMNAME",
                FORM_FACTOR_COL,
                "Total_Units"
            ]
        ]
    )




def rising_declining(
    wide,
    month_cols,
    recent_n,
    top_n
):

    out = wide.copy()


    recent = month_cols[-recent_n:]

    previous = month_cols[-2*recent_n:-recent_n]


    out["Recent_Sales"] = (
        out[recent]
        .sum(axis=1)
    )


    out["Previous_Sales"] = (
        out[previous]
        .sum(axis=1)
    )


    out = out[
        out["Previous_Sales"] > 0
    ]


    out["Growth_%"] = (

        (
            out["Recent_Sales"]
            -
            out["Previous_Sales"]
        )

        /

        out["Previous_Sales"]

    ) * 100



    columns = [

        "ITEMCODE",
        "ITEMNAME",
        FORM_FACTOR_COL,
        "Previous_Sales",
        "Recent_Sales",
        "Growth_%"

    ]



    rising = (

        out.sort_values(
            "Growth_%",
            ascending=False
        )
        .head(top_n)
        [columns]

    )


    declining = (

        out.sort_values(
            "Growth_%",
            ascending=True
        )
        .head(top_n)
        [columns]

    )


    return rising, declining




def stopped_selling(
    wide,
    month_cols,
    recent_n,
    top_n
):

    out = wide.copy()


    recent = month_cols[-recent_n:]

    previous = month_cols[:-recent_n]


    out["Recent_Sales"] = (
        out[recent]
        .sum(axis=1)
    )


    out["Previous_Sales"] = (
        out[previous]
        .sum(axis=1)
    )



    stopped = out[

        (out["Recent_Sales"] == 0)

        &

        (out["Previous_Sales"] > 0)

    ]



    return (

        stopped.sort_values(
            "Previous_Sales",
            ascending=False
        )
        .head(top_n)

        [

            [
                "ITEMCODE",
                "ITEMNAME",
                FORM_FACTOR_COL,
                "Previous_Sales"
            ]

        ]

    )




def form_factor_trend(
    agg,
    top_n
):

    ff = (

        agg.groupby(
            FORM_FACTOR_COL
        )["QTY"]
        .sum()
        .reset_index()

    )


    return (

        ff.sort_values(
            "QTY",
            ascending=False
        )
        .head(top_n)

    )





def main():

    raw = load_and_combine(
        file_paths,
        SHEET_NAME
    )


    clean = clean_and_prepare(
        raw
    )


    agg = build_sku_month_table(
        clean
    )


    wide = pivot_wide(
        agg
    )


    month_cols = [

        c for c in wide.columns

        if c not in
        [
            "ITEMCODE",
            "ITEMNAME",
            FORM_FACTOR_COL
        ]

    ]



    best = best_sellers(
        wide,
        month_cols,
        TOP_N
    )


    rising, declining = rising_declining(
        wide,
        month_cols,
        RECENT_MONTHS,
        TOP_N
    )


    stopped = stopped_selling(
        wide,
        month_cols,
        RECENT_MONTHS,
        TOP_N
    )


    form_factor = form_factor_trend(
        agg,
        TOP_N
    )



    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl"
    ) as writer:


        best.to_excel(
            writer,
            sheet_name="Top Best Sellers",
            index=False
        )


        rising.to_excel(
            writer,
            sheet_name="Top Rising",
            index=False
        )


        declining.to_excel(
            writer,
            sheet_name="Top Declining",
            index=False
        )


        stopped.to_excel(
            writer,
            sheet_name="Stopped Selling",
            index=False
        )


        form_factor.to_excel(
            writer,
            sheet_name="Form Factor Trend",
            index=False
        )


    print(
        "\nFinished:",
        os.path.abspath(OUTPUT_FILE)
    )



if __name__ == "__main__":
    main()
