import pandas as pd


# ============================================================
# CONFIGURATION
# ============================================================

GROWTH_WINDOW_DAYS = 14


# ============================================================
# PREPARE HISTORY
# ============================================================

def prepare_history(history_df):

    df = history_df.copy()

    if df.empty:
        return df

    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(
            df["Date"],
            errors="coerce"
        )

    numeric_columns = [
        "Total Capacity (GB)",
        "Used Capacity (GB)",
        "Free Capacity (GB)",
        "Used Capacity (%)",
        "Free Capacity (%)"
    ]

    for column in numeric_columns:

        if column in df.columns:

            df[column] = pd.to_numeric(
                df[column],
                errors="coerce"
            )

    if "Cluster" in df.columns:

        df["Cluster"] = (
            df["Cluster"]
            .astype(str)
            .str.strip()
        )

    df = (
        df
        .dropna(subset=["Date"])
        .sort_values(
            ["Cluster", "Date"]
        )
        .reset_index(drop=True)
    )

    return df


# ============================================================
# THRESHOLD STATUS
# ============================================================

def flag_threshold_status(
    used_percentage,
    thresholds
):

    if used_percentage is None:
        return "Unknown"

    try:
        value = float(
            used_percentage
        )
    except (
        TypeError,
        ValueError
    ):
        return "Unknown"

    # Support both 0.80 and 80
    if value > 1:
        value = value / 100.0

    warning = thresholds.get(
        "warning"
    )

    critical = thresholds.get(
        "critical"
    )

    immediate = thresholds.get(
        "immediate"
    )

    if (
        immediate is not None
        and value >= immediate
    ):
        return "Immediate Action"

    if (
        critical is not None
        and value >= critical
    ):
        return "Critical"

    if (
        warning is not None
        and value >= warning
    ):
        return "Warning"

    return "OK"


# ============================================================
# CALCULATE GROWTH
# ============================================================

def calculate_growth(
    history_df,
    cluster
):

    df = prepare_history(
        history_df
    )

    empty_result = {
        "Cluster": cluster,
        "Latest Date": None,
        "Daily Growth (GB)": None,
        "Weekly Growth (GB)": None,
        "Monthly Growth (GB)": None,
        "Average Daily Consumption (GB)": None,
        "Latest Used Capacity (GB)": None,
        "Latest Free Capacity (GB)": None,
        "Latest Total Capacity (GB)": None,
        "Used Capacity (%)": None
    }

    if df.empty:
        return empty_result

    cluster_df = df[
        df["Cluster"] == str(cluster).strip()
    ].copy()

    if cluster_df.empty:
        return empty_result

    cluster_df = (
        cluster_df
        .sort_values("Date")
        .drop_duplicates(
            subset=["Date"],
            keep="last"
        )
        .reset_index(drop=True)
    )

    latest = cluster_df.iloc[-1]

    latest_date = pd.Timestamp(
        latest["Date"]
    )

    latest_used = float(
        latest["Used Capacity (GB)"]
    )

    latest_free = float(
        latest["Free Capacity (GB)"]
    )

    latest_total = float(
        latest["Total Capacity (GB)"]
    )

    latest_used_pct = float(
        latest["Used Capacity (%)"]
    )

    # --------------------------------------------------------
    # Daily growth
    # --------------------------------------------------------

    daily_growth = None

    if len(cluster_df) >= 2:

        previous = cluster_df.iloc[-2]

        daily_growth = (
            latest_used
            -
            float(
                previous["Used Capacity (GB)"]
            )
        )

    # --------------------------------------------------------
    # Weekly growth
    # --------------------------------------------------------

    weekly_growth = None

    if len(cluster_df) >= 2:

        target_date = (
            latest_date
            - pd.Timedelta(days=7)
        )

        previous_week = cluster_df[
            cluster_df["Date"] <= target_date
        ]

        if not previous_week.empty:

            previous_row = (
                previous_week.iloc[-1]
            )

            weekly_growth = (
                latest_used
                -
                float(
                    previous_row[
                        "Used Capacity (GB)"
                    ]
                )
            )

    # --------------------------------------------------------
    # Monthly growth
    # --------------------------------------------------------

    monthly_growth = None

    if len(cluster_df) >= 2:

        target_date = (
            latest_date
            - pd.DateOffset(months=1)
        )

        previous_month = cluster_df[
            cluster_df["Date"] <= target_date
        ]

        if not previous_month.empty:

            previous_row = (
                previous_month.iloc[-1]
            )

            monthly_growth = (
                latest_used
                -
                float(
                    previous_row[
                        "Used Capacity (GB)"
                    ]
                )
            )

    # --------------------------------------------------------
    # Average daily consumption
    # --------------------------------------------------------

    average_daily = None

    window_start = (
        latest_date
        - pd.Timedelta(
            days=GROWTH_WINDOW_DAYS
        )
    )

    recent = cluster_df[
        cluster_df["Date"] >= window_start
    ].copy()

    if len(recent) >= 2:

        recent["Daily Delta"] = (
            recent[
                "Used Capacity (GB)"
            ]
            .diff()
        )

        deltas = (
            recent["Daily Delta"]
            .dropna()
        )

        if not deltas.empty:
            average_daily = float(
                deltas.mean()
            )

    return {
        "Cluster": cluster,
        "Latest Date": latest_date,
        "Daily Growth (GB)": daily_growth,
        "Weekly Growth (GB)": weekly_growth,
        "Monthly Growth (GB)": monthly_growth,
        "Average Daily Consumption (GB)": average_daily,
        "Latest Used Capacity (GB)": latest_used,
        "Latest Free Capacity (GB)": latest_free,
        "Latest Total Capacity (GB)": latest_total,
        "Used Capacity (%)": latest_used_pct
    }


# ============================================================
# FORECAST CAPACITY
# ============================================================

def forecast_capacity(
    history_df,
    cluster
):

    result = calculate_growth(
        history_df,
        cluster
    )

    average_daily = result.get(
        "Average Daily Consumption (GB)"
    )

    free_capacity = result.get(
        "Latest Free Capacity (GB)"
    )

    latest_date = result.get(
        "Latest Date"
    )

    # --------------------------------------------------------
    # Insufficient history
    # --------------------------------------------------------

    if average_daily is None:

        result[
            "Estimated Days Until Exhaustion"
        ] = None

        result[
            "Estimated Exhaustion Date"
        ] = "Insufficient History"

        return result

    # --------------------------------------------------------
    # No positive growth
    # --------------------------------------------------------

    if average_daily <= 0:

        result[
            "Estimated Days Until Exhaustion"
        ] = None

        result[
            "Estimated Exhaustion Date"
        ] = "Not Growing"

        return result

    # --------------------------------------------------------
    # Invalid free capacity
    # --------------------------------------------------------

    if free_capacity is None:

        result[
            "Estimated Days Until Exhaustion"
        ] = None

        result[
            "Estimated Exhaustion Date"
        ] = "Unknown"

        return result

    # --------------------------------------------------------
    # Already exhausted
    # --------------------------------------------------------

    if free_capacity <= 0:

        result[
            "Estimated Days Until Exhaustion"
        ] = 0

        result[
            "Estimated Exhaustion Date"
        ] = latest_date.strftime(
            "%Y-%m-%d"
        )

        return result

    # --------------------------------------------------------
    # Forecast
    # --------------------------------------------------------

    estimated_days = (
        free_capacity
        /
        average_daily
    )

    exhaustion_date = (
        latest_date
        +
        pd.Timedelta(
            days=estimated_days
        )
    )

    result[
        "Estimated Days Until Exhaustion"
    ] = float(
        estimated_days
    )

    result[
        "Estimated Exhaustion Date"
    ] = exhaustion_date.strftime(
        "%Y-%m-%d"
    )

    return result


# ============================================================
# FORECAST TABLE
# ============================================================

def build_forecast_table(
    history_df
):

    df = prepare_history(
        history_df
    )

    if df.empty:
        return pd.DataFrame()

    clusters = (
        df["Cluster"]
        .dropna()
        .unique()
    )

    results = []

    for cluster in clusters:

        results.append(
            forecast_capacity(
                df,
                cluster
            )
        )

    return pd.DataFrame(
        results
    )


# ============================================================
# DIRECT TEST
# ============================================================

if __name__ == "__main__":

    print(
        "growth.py loaded successfully."
    )
