import os
from datetime import datetime

import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import (
    Font,
    PatternFill,
    Border,
    Side,
    Alignment
)
from openpyxl.worksheet.table import Table, TableStyleInfo

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

OUTPUT_DIR = os.path.join(
    BASE_DIR,
    "output"
)

# ============================================================
# MONTHLY CONFIGURATION
# ============================================================

NUMBER_OF_MONTHS = 3

TH_LOW = 0.80
TH_MEDIUM = 0.90
TH_HIGH = 0.95

# ============================================================
# EXCEL STYLES
# ============================================================

HEADER_FILL = PatternFill(
    fill_type="solid",
    fgColor="1F4E78"
)

HEADER_FONT = Font(
    color="FFFFFF",
    bold=True
)

TITLE_FONT = Font(
    bold=True,
    size=16
)

THIN_SIDE = Side(
    style="thin",
    color="D9E1F2"
)

BORDER = Border(
    left=THIN_SIDE,
    right=THIN_SIDE,
    top=THIN_SIDE,
    bottom=THIN_SIDE
)

# ============================================================
# LOAD HISTORY
# ============================================================

def load_history():

    if not os.path.exists(
        DAILY_HISTORY_FILE
    ):
        raise FileNotFoundError(
            "Historical file not found: "
            + DAILY_HISTORY_FILE
        )

    print(
        "Historical file used : "
        + DAILY_HISTORY_FILE
    )

    # IMPORTANT:
    # keep_default_na=False prevents Pandas from converting
    # the valid region name "NA" into NaN.
    df = pd.read_csv(
        DAILY_HISTORY_FILE,
        keep_default_na=False
    )

    if df.empty:
        raise ValueError(
            "Historical file is empty: "
            + DAILY_HISTORY_FILE
        )

    required_columns = [
        "Date",
        "Cluster",
        "Active",
        "Location",
        "Type",
        "Region",
        "Country",
        "City",
        "Total Capacity (GB)",
        "Used Capacity (GB)",
        "Free Capacity (GB)",
        "Used Capacity (%)",
        "Free Capacity (%)"
    ]

    missing_columns = [
        column
        for column in required_columns
        if column not in df.columns
    ]

    if missing_columns:
        raise ValueError(
            "Historical file is missing columns: "
            + ", ".join(missing_columns)
        )

    # --------------------------------------------------------
    # Date
    # --------------------------------------------------------

    df["Date"] = pd.to_datetime(
        df["Date"],
        errors="coerce"
    )

    df = df.dropna(
        subset=["Date"]
    )

    # --------------------------------------------------------
    # Numeric columns
    # --------------------------------------------------------

    numeric_columns = [
        "Total Capacity (GB)",
        "Used Capacity (GB)",
        "Free Capacity (GB)",
        "Used Capacity (%)",
        "Free Capacity (%)"
    ]

    for column in numeric_columns:

        df[column] = pd.to_numeric(
            df[column],
            errors="coerce"
        )

    # --------------------------------------------------------
    # Normalize text
    # --------------------------------------------------------

    text_columns = [
        "Cluster",
        "Active",
        "Location",
        "Type",
        "Region",
        "Country",
        "City"
    ]

    for column in text_columns:

        df[column] = (
            df[column]
            .astype(str)
            .str.strip()
        )

    # --------------------------------------------------------
    # Restore Region from cluster mapping if required
    # --------------------------------------------------------

    try:

        from cluster_mapping import CLUSTER_MAPPING

        for index in df.index:

            cluster = df.at[
                index,
                "Cluster"
            ]

            region = df.at[
                index,
                "Region"
            ]

            if (
                region == ""
                or str(region).lower() == "nan"
            ):

                mapping = CLUSTER_MAPPING.get(
                    cluster
                )

                if mapping:

                    df.at[
                        index,
                        "Region"
                    ] = mapping.get(
                        "Region",
                        ""
                    )

                    if (
                        df.at[
                            index,
                            "Country"
                        ] == ""
                    ):

                        df.at[
                            index,
                            "Country"
                        ] = mapping.get(
                            "Country",
                            ""
                        )

                    if (
                        df.at[
                            index,
                            "City"
                        ] == ""
                    ):

                        df.at[
                            index,
                            "City"
                        ] = mapping.get(
                            "City",
                            ""
                        )

    except ImportError:

        pass

    # --------------------------------------------------------
    # Remove invalid clusters
    # --------------------------------------------------------

    df = df[
        (df["Cluster"] != "")
        &
        (df["Cluster"].str.lower() != "nan")
    ]

    # --------------------------------------------------------
    # Sort
    # --------------------------------------------------------

    df = df.sort_values(
        by=[
            "Date",
            "Cluster"
        ]
    ).reset_index(
        drop=True
    )

    print(
        "Historical rows loaded : "
        + str(len(df))
    )

    print(
        "Unique dates           : "
        + str(df["Date"].nunique())
    )

    print(
        "Unique clusters        : "
        + str(df["Cluster"].nunique())
    )

    return df


# ============================================================
# GET MONTHS
# ============================================================

def get_latest_months(df):

    if df.empty:
        return []

    # Convert each date into calendar month.
    df = df.copy()

    df["Month_Period"] = (
        df["Date"]
        .dt.to_period("M")
    )

    available_months = sorted(
        df["Month_Period"].unique()
    )

    selected_months = (
        available_months[
            -NUMBER_OF_MONTHS:
        ]
    )

    return selected_months


# ============================================================
# BUILD MONTHLY TABLE
# ============================================================

def build_monthly_table(
    history,
    selected_months
):

    if history.empty:
        return pd.DataFrame()

    monthly_records = []

    for month_period in selected_months:

        month_data = history[
            history["Date"].dt.to_period("M")
            == month_period
        ].copy()

        if month_data.empty:
            continue

        # ----------------------------------------------------
        # Take the latest snapshot available in that month
        # for every cluster.
        # ----------------------------------------------------

        month_data = (
            month_data
            .sort_values("Date")
            .groupby(
                "Cluster",
                as_index=False
            )
            .tail(1)
        )

        month_data = month_data[
            [
                "Cluster",
                "Region",
                "Country",
                "City",
                "Used Capacity (%)"
            ]
        ].copy()

        month_data["Month_Period"] = (
            month_period
        )

        monthly_records.append(
            month_data
        )

    if not monthly_records:
        return pd.DataFrame()

    monthly_data = pd.concat(
        monthly_records,
        ignore_index=True
    )

    # --------------------------------------------------------
    # Pivot months into columns
    # --------------------------------------------------------

    pivot = monthly_data.pivot_table(
        index=[
            "Cluster",
            "Region",
            "Country",
            "City"
        ],
        columns="Month_Period",
        values="Used Capacity (%)",
        aggfunc="last"
    ).reset_index()

    pivot.columns.name = None

    # --------------------------------------------------------
    # Rename month columns
    # Example:
    # 2026-08 -> Aug-26
    # --------------------------------------------------------

    rename_map = {}

    for month_period in selected_months:

        if month_period in pivot.columns:

            rename_map[
                month_period
            ] = pd.Timestamp(
                month_period.start_time
            ).strftime(
                "%b-%y"
            )

    pivot = pivot.rename(
        columns=rename_map
    )

    # --------------------------------------------------------
    # Month column names
    # --------------------------------------------------------

    month_columns = []

    for month_period in selected_months:

        month_name = pd.Timestamp(
            month_period.start_time
        ).strftime(
            "%b-%y"
        )

        if month_name in pivot.columns:

            month_columns.append(
                month_name
            )

    # --------------------------------------------------------
    # Final columns
    # --------------------------------------------------------

    result_columns = [
        "Cluster",
        "Region",
        "Country",
        "City"
    ]

    result_columns.extend(
        month_columns
    )

    pivot = pivot[
        result_columns
    ]

    # --------------------------------------------------------
    # Sort by city
    # --------------------------------------------------------

    pivot = pivot.sort_values(
        by=[
            "City",
            "Cluster"
        ]
    ).reset_index(
        drop=True
    )

    return pivot


# ============================================================
# CREATE REGIONAL SHEET
# ============================================================

def create_region_sheet(
    wb,
    region_data,
    region_name,
    month_columns
):

    ws = wb.create_sheet(
        title=region_name
    )

    # --------------------------------------------------------
    # Title
    # --------------------------------------------------------

    ws["A1"] = (
        "Rubrik Monthly Capacity Growth"
    )

    ws["A1"].font = TITLE_FONT

    # --------------------------------------------------------
    # Header
    # --------------------------------------------------------

    start_row = 3

    headers = [
        "City Location",
        "TH %-80",
        "TH %-90",
        "TH %-95"
    ]

    headers.extend(
        month_columns
    )

    for column_number, header in enumerate(
        headers,
        start=1
    ):

        cell = ws.cell(
            row=start_row,
            column=column_number,
            value=header
        )

        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = BORDER

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    # --------------------------------------------------------
    # Data
    # --------------------------------------------------------

    for row_offset, row_data in enumerate(
        region_data.itertuples(
            index=False,
            name=None
        ),
        start=1
    ):

        row_number = (
            start_row
            + row_offset
        )

        # Tuple structure:
        #
        # 0 Cluster
        # 1 Region
        # 2 Country
        # 3 City
        # 4+ monthly values

        city = row_data[3]

        values = [
            city,
            TH_LOW,
            TH_MEDIUM,
            TH_HIGH
        ]

        for index in range(
            4,
            len(row_data)
        ):

            values.append(
                row_data[index]
            )

        for column_number, value in enumerate(
            values,
            start=1
        ):

            cell = ws.cell(
                row=row_number,
                column=column_number,
                value=value
            )

            cell.border = BORDER

            cell.alignment = Alignment(
                vertical="center"
            )

            # Threshold percentages
            if column_number in [
                2,
                3,
                4
            ]:

                cell.number_format = "0%"

            # Monthly capacity percentages
            elif column_number >= 5:

                cell.number_format = "0%"

    # --------------------------------------------------------
    # Column widths
    # --------------------------------------------------------

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14

    for column_number in range(
        5,
        ws.max_column + 1
    ):

        ws.column_dimensions[
            ws.cell(
                row=1,
                column=column_number
            ).column_letter
        ].width = 14

    ws.freeze_panes = "A4"

    return ws


# ============================================================
# CREATE RAW DATA SHEET
# ============================================================

def create_raw_data_sheet(
    wb,
    history
):

    ws = wb.create_sheet(
        title="Raw Data"
    )

    ws["A1"] = (
        "Rubrik Monthly Capacity Raw History"
    )

    ws["A1"].font = TITLE_FONT

    # --------------------------------------------------------
    # Raw columns
    # --------------------------------------------------------

    columns = [
        "Date",
        "Cluster",
        "Region",
        "Country",
        "City",
        "Total Capacity (GB)",
        "Used Capacity (GB)",
        "Free Capacity (GB)",
        "Used Capacity (%)",
        "Free Capacity (%)",
        "Active",
        "Location",
        "Type"
    ]

    start_row = 3

    for column_number, header in enumerate(
        columns,
        start=1
    ):

        cell = ws.cell(
            row=start_row,
            column=column_number,
            value=header
        )

        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = BORDER

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

    # --------------------------------------------------------
    # Raw data rows
    # --------------------------------------------------------

    for row_offset, row_data in enumerate(
        history.itertuples(
            index=False,
            name=None
        ),
        start=1
    ):

        row_number = (
            start_row
            + row_offset
        )

        record = history.iloc[
            row_offset - 1
        ]

        values = [
            record["Date"],
            record["Cluster"],
            record["Region"],
            record["Country"],
            record["City"],
            record["Total Capacity (GB)"],
            record["Used Capacity (GB)"],
            record["Free Capacity (GB)"],
            record["Used Capacity (%)"],
            record["Free Capacity (%)"],
            record["Active"],
            record["Location"],
            record["Type"]
        ]

        for column_number, value in enumerate(
            values,
            start=1
        ):

            # Excel does not accept pandas Period
            # or some pandas-specific objects.
            if isinstance(
                value,
                pd.Period
            ):

                value = str(value)

            cell = ws.cell(
                row=row_number,
                column=column_number,
                value=value
            )

            cell.border = BORDER

            cell.alignment = Alignment(
                vertical="center"
            )

            if column_number in [
                9,
                10
            ]:

                cell.number_format = "0.00%"

            elif column_number in [
                6,
                7,
                8
            ]:

                cell.number_format = "#,##0.00"

    # --------------------------------------------------------
    # Widths
    # --------------------------------------------------------

    widths = {
        "A": 14,
        "B": 16,
        "C": 12,
        "D": 18,
        "E": 20,
        "F": 20,
        "G": 20,
        "H": 20,
        "I": 18,
        "J": 18,
        "K": 12,
        "L": 35,
        "M": 15
    }

    for column, width in widths.items():

        ws.column_dimensions[
            column
        ].width = width

    ws.freeze_panes = "A4"

    return ws


# ============================================================
# GENERATE MONTHLY REPORT
# ============================================================

def generate_monthly_report():

    print()
    print("=" * 70)
    print("Rubrik Monthly Capacity Report")
    print("=" * 70)

    # --------------------------------------------------------
    # Load history
    # --------------------------------------------------------

    history = load_history()

    # --------------------------------------------------------
    # Get latest months
    # --------------------------------------------------------

    selected_months = get_latest_months(
        history
    )

    print()
    print(
        "Months available : "
        + str(len(selected_months))
    )

    for month in selected_months:

        print(
            "  "
            + pd.Timestamp(
                month.start_time
            ).strftime(
                "%B %Y"
            )
        )

    # --------------------------------------------------------
    # Build monthly table
    # --------------------------------------------------------

    monthly_table = build_monthly_table(
        history,
        selected_months
    )

    print()

    print(
        "Clusters included : "
        + str(len(monthly_table))
    )

    if monthly_table.empty:

        raise ValueError(
            "No monthly capacity data available."
        )

    # --------------------------------------------------------
    # Month columns
    # --------------------------------------------------------

    month_columns = []

    for month in selected_months:

        month_name = pd.Timestamp(
            month.start_time
        ).strftime(
            "%b-%y"
        )

        if month_name in monthly_table.columns:

            month_columns.append(
                month_name
            )

    print(
        "Monthly columns    : "
        + ", ".join(month_columns)
    )

    # --------------------------------------------------------
    # Create output directory
    # --------------------------------------------------------

    os.makedirs(
        OUTPUT_DIR,
        exist_ok=True
    )

    # --------------------------------------------------------
    # Report filename
    # --------------------------------------------------------

    today = datetime.now()

    report_name = today.strftime(
        "Backup_Capacity_Monthly_%d_%b_%Y.xlsx"
    )

    report_path = os.path.join(
        OUTPUT_DIR,
        report_name
    )

    # --------------------------------------------------------
    # Create workbook
    # --------------------------------------------------------

    wb = Workbook()

    default_sheet = wb.active

    wb.remove(
        default_sheet
    )

    # --------------------------------------------------------
    # Regional sheets
    # --------------------------------------------------------

    regions = [
        "NA",
        "NA_DD",
        "LATAM",
        "EU",
        "APAC",
        "CHINA"
    ]

    for region in regions:

        if region == "CHINA":

            region_data = monthly_table[
                monthly_table["Country"]
                .astype(str)
                .str.upper()
                .eq("CHINA")
            ]

        elif region == "NA_DD":

            # Required sheet but currently no
            # clusters are assigned to NA_DD.
            region_data = monthly_table.iloc[
                0:0
            ]

        else:

            region_data = monthly_table[
                monthly_table["Region"]
                .astype(str)
                .str.upper()
                .eq(region)
            ]

        create_region_sheet(
            wb,
            region_data,
            region,
            month_columns
        )

    # --------------------------------------------------------
    # Raw Data
    # --------------------------------------------------------

    create_raw_data_sheet(
        wb,
        history
    )

    # --------------------------------------------------------
    # Sheet order
    # --------------------------------------------------------

    desired_order = [
        "NA",
        "NA_DD",
        "LATAM",
        "EU",
        "APAC",
        "CHINA",
        "Raw Data"
    ]

    wb._sheets = [
        wb[name]
        for name in desired_order
        if name in wb.sheetnames
    ]

    # --------------------------------------------------------
    # Save
    # --------------------------------------------------------

    wb.save(
        report_path
    )

    # --------------------------------------------------------
    # Verify
    # --------------------------------------------------------

    file_exists = os.path.exists(
        report_path
    )

    print()
    print("=" * 70)
    print("Monthly Excel Report Generated")
    print("=" * 70)

    print(
        "Report Location : "
        + report_path
    )

    print(
        "Absolute Path   : "
        + os.path.abspath(
            report_path
        )
    )

    print(
        "Clusters        : "
        + str(len(monthly_table))
    )

    print(
        "Months Included : "
        + str(len(selected_months))
    )

    print(
        "File Exists     : "
        + str(file_exists)
    )

    print(
        "Sheets          : "
        + ", ".join(
            wb.sheetnames
        )
    )

    print("=" * 70)

    return report_path


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":

    report_file = (
        generate_monthly_report()
    )

    print()
    print(
        "Sending Monthly Capacity Report Email..."
    )

    from email_report import send_email

    send_email(
        report_file
    )
