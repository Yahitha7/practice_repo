import os
import pandas as pd


# ============================================================
# PATH CONFIGURATION
# ============================================================

BASE_DIR = os.path.dirname(
    os.path.abspath(__file__)
)

HISTORY_DIR = os.path.join(
    BASE_DIR,
    "history"
)

DAILY_HISTORY_FILE = os.path.join(
    HISTORY_DIR,
    "daily_capacity.csv"
)


# ============================================================
# SAVE DAILY HISTORY
# ============================================================

def save_daily_history(df):

    os.makedirs(
        HISTORY_DIR,
        exist_ok=True
    )

    if df is None or df.empty:

        raise ValueError(
            "Cannot save empty capacity data."
        )

    history_df = df.copy()

    # --------------------------------------------------------
    # Required columns from capacity.py
    # --------------------------------------------------------

    required_columns = [
        "Target Storage",
        "Active",
        "Location",
        "Type",
        "Region",
        "Country",
        "City",
        "Total Capacity",
        "Used Capacity",
        "Free Capacity",
        "Used %",
        "Free %",
    ]

    missing_columns = [
        column
        for column in required_columns
        if column not in history_df.columns
    ]

    if missing_columns:

        raise ValueError(
            "Capacity data is missing required columns: "
            + ", ".join(missing_columns)
        )

    # ========================================================
    # TODAY'S DATE
    # ========================================================

    today = pd.Timestamp.now().strftime(
        "%Y-%m-%d"
    )

    month = pd.Timestamp.now().strftime(
        "%Y-%m"
    )

    history_df["Date"] = today
    history_df["Month"] = month

    # ========================================================
    # SELECT COLUMNS
    # ========================================================

    history_df = history_df[
        [
            "Date",
            "Month",
            "Target Storage",
            "Active",
            "Location",
            "Type",
            "Region",
            "Country",
            "City",
            "Total Capacity",
            "Used Capacity",
            "Free Capacity",
            "Used %",
            "Free %",
        ]
    ].copy()

    # ========================================================
    # RENAME FOR HISTORY CSV
    # ========================================================

    history_df = history_df.rename(
        columns={
            "Target Storage": "Cluster",
            "Total Capacity": "Total Capacity (GB)",
            "Used Capacity": "Used Capacity (GB)",
            "Free Capacity": "Free Capacity (GB)",
            "Used %": "Used Capacity (%)",
            "Free %": "Free Capacity (%)",
        }
    )

    # ========================================================
    # NORMALIZE CLUSTER
    # ========================================================

    history_df["Cluster"] = (
        history_df["Cluster"]
        .astype(str)
        .str.strip()
    )

    # ========================================================
    # VALIDATE REGION
    # ========================================================

    invalid_region = (
        history_df["Region"].isna()
        |
        history_df["Region"]
        .astype(str)
        .str.strip()
        .eq("")
        |
        history_df["Region"]
        .astype(str)
        .str.lower()
        .eq("nan")
    )

    if invalid_region.any():

        print()
        print("=" * 70)
        print("ERROR: INVALID REGION IN HISTORY DATA")
        print("=" * 70)

        print(
            history_df.loc[
                invalid_region,
                [
                    "Cluster",
                    "Region",
                    "Country",
                    "City",
                ]
            ].to_string(index=False)
        )

        print("=" * 70)

        raise ValueError(
            "Cannot save history because one or more "
            "clusters have an invalid Region."
        )

    # ========================================================
    # REMOVE INVALID CLUSTERS
    # ========================================================

    history_df = history_df[
        history_df["Cluster"].notna()
        &
        (history_df["Cluster"] != "")
        &
        (history_df["Cluster"] != "nan")
    ]

    # ========================================================
    # ONE RECORD PER CLUSTER FOR TODAY
    # ========================================================

    history_df = (
        history_df
        .drop_duplicates(
            subset=[
                "Date",
                "Cluster",
            ],
            keep="last"
        )
    )

    # ========================================================
    # LOAD EXISTING HISTORY
    # ========================================================

    if os.path.exists(
        DAILY_HISTORY_FILE
    ):

        existing_df = pd.read_csv(
            DAILY_HISTORY_FILE
        )

        if not existing_df.empty:

            combined_df = pd.concat(
                [
                    existing_df,
                    history_df,
                ],
                ignore_index=True
            )

        else:

            combined_df = history_df.copy()

    else:

        combined_df = history_df.copy()

    # ========================================================
    # CLEAN EXISTING HISTORY
    # ========================================================

    # Convert Date to string consistently.

    combined_df["Date"] = (
        pd.to_datetime(
            combined_df["Date"],
            errors="coerce"
        )
        .dt.strftime("%Y-%m-%d")
    )

    # Remove invalid dates.

    combined_df = combined_df[
        combined_df["Date"].notna()
    ]

    # Normalize Cluster.

    combined_df["Cluster"] = (
        combined_df["Cluster"]
        .astype(str)
        .str.strip()
    )

    # ========================================================
    # REPLACE SAME-DAY RECORDS
    #
    # If today's automation runs again, the new data
    # replaces today's previous data.
    # ========================================================

    combined_df = (
        combined_df
        .drop_duplicates(
            subset=[
                "Date",
                "Cluster",
            ],
            keep="last"
        )
    )

    # ========================================================
    # SORT
    # ========================================================

    combined_df = (
        combined_df
        .sort_values(
            by=[
                "Date",
                "Cluster",
            ]
        )
        .reset_index(
            drop=True
        )
    )

    # ========================================================
    # SAVE
    # ========================================================

    combined_df.to_csv(
        DAILY_HISTORY_FILE,
        index=False
    )

    # ========================================================
    # VALIDATION
    # ========================================================

    unique_dates = (
        combined_df["Date"]
        .nunique()
    )

    unique_clusters = (
        combined_df["Cluster"]
        .nunique()
    )

    current_day_rows = len(
        combined_df[
            combined_df["Date"] == today
        ]
    )

    print()
    print("=" * 70)
    print("Historical Data Saved")
    print("=" * 70)

    print(
        "History file       : "
        + DAILY_HISTORY_FILE
    )

    print(
        "Today's snapshot   : "
        + today
    )

    print(
        "Today's rows       : "
        + str(current_day_rows)
    )

    print(
        "Total rows         : "
        + str(len(combined_df))
    )

    print(
        "Unique dates       : "
        + str(unique_dates)
    )

    print(
        "Unique clusters    : "
        + str(unique_clusters)
    )

    print("=" * 70)

    # ========================================================
    # REGION SUMMARY
    # ========================================================

    print()
    print("Current snapshot by region:")

    current_snapshot = combined_df[
        combined_df["Date"] == today
    ]

    print(
        current_snapshot[
            "Region"
        ]
        .value_counts(
            dropna=False
        )
        .to_string()
    )

    return DAILY_HISTORY_FILE
