import os
from datetime import datetime

import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
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

HISTORY_FILE = os.path.join(
    HISTORY_DIR,
    "daily_capacity.csv"
)

OUTPUT_DIR = os.path.join(
    BASE_DIR,
    "output"
)


# ============================================================
# WEEKLY CONFIGURATION
# ============================================================

NUMBER_OF_WEEKS = 4

# Backup team thresholds
TH_LOW = 0.30
TH_HIGH = 0.85


# ============================================================
# EXCEL STYLES
# ============================================================

HEADER_FILL = PatternFill(
    fill_type="solid",
    fgColor="D9EAF7"
)

HEADER_FONT = Font(
    bold=True,
    color="000000"
)

TITLE_FONT = Font(
    bold=True,
    size=14
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

    if not os.path.exists(HISTORY_FILE):

        raise FileNotFoundError(
            "Historical file not found: "
            + HISTORY_FILE
        )

    print(
        "Historical file used : "
        + HISTORY_FILE
    )

    df = pd.read_csv(
        HISTORY_FILE,
        keep_default_na=False
    )

    if df.empty:

        raise ValueError(
            "Historical file is empty: "
            + HISTORY_FILE
        )

    # --------------------------------------------------------
    # Required columns
    # --------------------------------------------------------

    required_columns = [
        "Date",
        "Cluster",
        "Region",
        "Country",
        "City",
        "Used Capacity (%)"
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
    # Used Capacity %
    # --------------------------------------------------------

    df["Used Capacity (%)"] = pd.to_numeric(
        df["Used Capacity (%)"],
        errors="coerce"
    )

    # --------------------------------------------------------
    # Normalize text fields
    # --------------------------------------------------------

    for column in [
        "Cluster",
        "Region",
        "Country",
        "City"
    ]:

        df[column] = (
            df[column]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    # --------------------------------------------------------
    # Remove invalid clusters
    # --------------------------------------------------------

    df = df[
        (df["Cluster"] != "")
        &
        (df["Cluster"].str.lower() != "nan")
    ]

    # --------------------------------------------------------
    # One record per cluster per day
    # --------------------------------------------------------

    df = df.drop_duplicates(
        subset=[
            "Date",
            "Cluster"
        ],
        keep="last"
    )

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
# BUILD WEEK INFORMATION
# ============================================================

def get_weekly_snapshots(df):

    data = df.copy()

    # --------------------------------------------------------
    # Calculate week ending Sunday.
    #
    # Monday = 0
    # Sunday = 6
    # --------------------------------------------------------

    data["Week_End"] = (
        data["Date"]
        + pd.to_timedelta(
            6 - data["Date"].dt.weekday,
            unit="D"
        )
    )

    data["Week_End"] = (
        pd.to_datetime(
            data["Week_End"]
        )
        .dt.normalize()
    )

    # --------------------------------------------------------
    # Available weeks
    # --------------------------------------------------------

    available_weeks = sorted(
        data["Week_End"]
        .dropna()
        .unique()
    )

    selected_weeks = available_weeks[
        -NUMBER_OF_WEEKS:
    ]

    # Convert to Timestamp explicitly
    selected_weeks = [
        pd.Timestamp(
            week
        ).normalize()
        for week in selected_weeks
    ]

    return (
        data,
        selected_weeks
    )


# ============================================================
# BUILD WEEKLY TABLE
# ============================================================

def build_weekly_table(
    data,
    selected_weeks
):

    if len(selected_weeks) == 0:

        return pd.DataFrame()

    # --------------------------------------------------------
    # Normalize selected weeks
    # --------------------------------------------------------

    selected_weeks = [
        pd.Timestamp(
            week
        ).normalize()
        for week in selected_weeks
    ]

    weekly_records = []

    # --------------------------------------------------------
    # Process each week
    # --------------------------------------------------------

    for week_end in selected_weeks:

        # Convert comparison column explicitly
        week_dates = (
            pd.to_datetime(
                data["Week_End"]
            )
            .dt.normalize()
        )

        week_data = data[
            week_dates == week_end
        ].copy()

        if week_data.empty:
            continue

        # ----------------------------------------------------
        # Latest available snapshot for every cluster
        # in that week
        # ----------------------------------------------------

        week_data = (
            week_data
            .sort_values(
                "Date"
            )
            .groupby(
                "Cluster",
                as_index=False
            )
            .tail(1)
        )

        week_data = week_data[
            [
                "Cluster",
                "Region",
                "Country",
                "City",
                "Used Capacity (%)"
            ]
        ].copy()

        week_data["Week_End"] = week_end

        weekly_records.append(
            week_data
        )

    if not weekly_records:

        return pd.DataFrame()

    # --------------------------------------------------------
    # Combine all weekly records
    # --------------------------------------------------------

    weekly_data = pd.concat(
        weekly_records,
        ignore_index=True
    )

    # --------------------------------------------------------
    # City Location
    # --------------------------------------------------------

    weekly_data["City Location"] = (
        weekly_data["City"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # --------------------------------------------------------
    # Pivot weekly percentages
    # --------------------------------------------------------

    pivot = weekly_data.pivot_table(
        index=[
            "Cluster",
            "Region",
            "Country",
            "City Location"
        ],
        columns="Week_End",
        values="Used Capacity (%)",
        aggfunc="last"
    ).reset_index()

    pivot.columns.name = None

    # --------------------------------------------------------
    # Detect and rename datetime columns
    #
    # This avoids the previous Period/Timestamp comparison
    # problem.
    # --------------------------------------------------------

    week_column_map = {}

    for column in list(
        pivot.columns
    ):

        try:

            column_date = pd.Timestamp(
                column
            ).normalize()

        except Exception:

            continue

        for week in selected_weeks:

            if column_date == week:

                formatted_name = (
                    week.strftime(
                        "%d-%b-%y"
                    )
                )

                week_column_map[
                    column
                ] = formatted_name

                break

    pivot = pivot.rename(
        columns=week_column_map
    )

    # --------------------------------------------------------
    # Determine actual week columns
    # --------------------------------------------------------

    week_columns = []

    for week in selected_weeks:

        formatted_name = (
            week.strftime(
                "%d-%b-%y"
            )
        )

        if formatted_name in pivot.columns:

            week_columns.append(
                formatted_name
            )

    # --------------------------------------------------------
    # Final columns
    #
    # Cluster/Region/Country are retained internally so
    # regional sheets can be created correctly.
    # --------------------------------------------------------

    result_columns = [
        "Cluster",
        "Region",
        "Country",
        "City Location"
    ]

    result_columns.extend(
        week_columns
    )

    pivot = pivot[
        result_columns
    ]

    # --------------------------------------------------------
    # Sort
    # --------------------------------------------------------

    pivot = pivot.sort_values(
        by=[
            "City Location",
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
    week_columns,
    table_number
):

    ws = wb.create_sheet(
        title=region_name
    )

    # --------------------------------------------------------
    # Title
    # --------------------------------------------------------

    ws["A1"] = (
        "Rubrik Weekly Capacity Growth"
    )

    ws["A1"].font = TITLE_FONT

    # --------------------------------------------------------
    # Headers
    # --------------------------------------------------------

    start_row = 3

    headers = [
        "City Location",
        "TH %-Low",
        "TH %-High"
    ]

    headers.extend(
        week_columns
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
            vertical="center",
            wrap_text=True
        )

    # --------------------------------------------------------
    # Data rows
    # --------------------------------------------------------

    row_number = start_row + 1

    for _, row in region_data.iterrows():

        values = [
            row["City Location"],
            TH_LOW,
            TH_HIGH
        ]

        for week_column in week_columns:

            value = row.get(
                week_column,
                None
            )

            values.append(
                value
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
                horizontal="center",
                vertical="center"
            )

            # ------------------------------------------------
            # Threshold + capacity columns are percentages
            # ------------------------------------------------

            if column_number >= 2:

                cell.number_format = "0%"

        row_number += 1

    # --------------------------------------------------------
    # Excel table
    # --------------------------------------------------------

    if row_number > start_row + 1:

        last_column = len(headers)

        last_column_letter = (
            ws.cell(
                row=start_row,
                column=last_column
            ).column_letter
        )

        table_ref = (
            "A"
            + str(start_row)
            + ":"
            + last_column_letter
            + str(row_number - 1)
        )

        table = Table(
            displayName=(
                "WeeklyTable"
                + str(table_number)
            ),
            ref=table_ref
        )

        table_style = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )

        table.tableStyleInfo = table_style

        ws.add_table(
            table
        )

    # --------------------------------------------------------
    # Column widths
    # --------------------------------------------------------

    ws.column_dimensions[
        "A"
    ].width = 25

    ws.column_dimensions[
        "B"
    ].width = 14

    ws.column_dimensions[
        "C"
    ].width = 14

    for column_number in range(
        4,
        len(headers) + 1
    ):

        column_letter = (
            ws.cell(
                row=start_row,
                column=column_number
            ).column_letter
        )

        ws.column_dimensions[
            column_letter
        ].width = 15

    # --------------------------------------------------------
    # Freeze
    # --------------------------------------------------------

    ws.freeze_panes = "D4"

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

    raw_data = history.copy()

    # --------------------------------------------------------
    # Excel-safe date values
    # --------------------------------------------------------

    raw_data["Date"] = (
        pd.to_datetime(
            raw_data["Date"]
        )
        .dt.strftime(
            "%Y-%m-%d"
        )
    )

    raw_data["Week_End"] = (
        pd.to_datetime(
            raw_data["Week_End"]
        )
        .dt.strftime(
            "%Y-%m-%d"
        )
    )

    # --------------------------------------------------------
    # Headers
    # --------------------------------------------------------

    headers = list(
        raw_data.columns
    )

    for column_number, header in enumerate(
        headers,
        start=1
    ):

        cell = ws.cell(
            row=1,
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

    for row_number, row in enumerate(
        raw_data.itertuples(
            index=False,
            name=None
        ),
        start=2
    ):

        for column_number, value in enumerate(
            row,
            start=1
        ):

            cell = ws.cell(
                row=row_number,
                column=column_number,
                value=value
            )

            cell.border = BORDER

    # --------------------------------------------------------
    # Widths
    # --------------------------------------------------------

    for column_number in range(
        1,
        ws.max_column + 1
    ):

        column_letter = (
            ws.cell(
                row=1,
                column=column_number
            ).column_letter
        )

        ws.column_dimensions[
            column_letter
        ].width = 18

    ws.freeze_panes = "A2"

    return ws


# ============================================================
# GENERATE WEEKLY REPORT
# ============================================================

def generate_weekly_report():

    print()
    print("=" * 70)
    print("Rubrik Weekly Capacity Report")
    print("=" * 70)

    # --------------------------------------------------------
    # Load history
    # --------------------------------------------------------

    history = load_history()

    # --------------------------------------------------------
    # Calculate weekly buckets
    # --------------------------------------------------------

    data, selected_weeks = (
        get_weekly_snapshots(
            history
        )
    )

    print()

    print(
        "Weeks included : "
        + str(
            len(selected_weeks)
        )
    )

    for week in selected_weeks:

        print(
            "  "
            + week.strftime(
                "%d %b %Y"
            )
        )

    # --------------------------------------------------------
    # Build weekly table
    # --------------------------------------------------------

    weekly_table = build_weekly_table(
        data,
        selected_weeks
    )

    if weekly_table.empty:

        raise ValueError(
            "No weekly capacity data available."
        )

    print()

    print(
        "Clusters included : "
        + str(
            len(weekly_table)
        )
    )

    # --------------------------------------------------------
    # Determine week columns
    # --------------------------------------------------------

    week_columns = []

    for week in selected_weeks:

        formatted_name = (
            week.strftime(
                "%d-%b-%y"
            )
        )

        if formatted_name in weekly_table.columns:

            week_columns.append(
                formatted_name
            )

    print(
        "Weekly columns     : "
        + ", ".join(
            week_columns
        )
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
        "Backup_Capacity_Weekly_%d_%b_%Y.xlsx"
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

    table_number = 1

    for region in regions:

        if region == "CHINA":

            region_data = weekly_table[
                weekly_table[
                    "Country"
                ]
                .astype(str)
                .str.upper()
                .eq("CHINA")
            ]

        elif region == "NA_DD":

            region_data = weekly_table.iloc[
                0:0
            ]

        else:

            region_data = weekly_table[
                weekly_table[
                    "Region"
                ]
                .astype(str)
                .str.upper()
                .eq(region)
            ]

        create_region_sheet(
            wb,
            region_data,
            region,
            week_columns,
            table_number
        )

        table_number += 1

    # --------------------------------------------------------
    # Raw Data
    # --------------------------------------------------------

    create_raw_data_sheet(
        wb,
        data
    )

    # --------------------------------------------------------
    # Final sheet order
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
    print("Weekly Excel Report Generated")
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
        + str(
            len(weekly_table)
        )
    )

    print(
        "Weeks Included  : "
        + str(
            len(selected_weeks)
        )
    )

    print(
        "Weekly Columns  : "
        + str(
            len(week_columns)
        )
    )

    print(
        "File Exists     : "
        + str(
            file_exists
        )
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
        generate_weekly_report()
    )

    print()

    print(
        "Sending Weekly Capacity Report Email..."
    )

    from email_report import send_email

    send_email(
        report_file
    )
