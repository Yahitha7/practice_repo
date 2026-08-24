import os
from datetime import datetime

import pandas as pd
import xlsxwriter


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = os.path.dirname(
    os.path.abspath(__file__)
)

HISTORY_FILE = os.path.join(
    BASE_DIR,
    "history",
    "daily_capacity.csv"
)

OUTPUT_DIR = os.path.join(
    BASE_DIR,
    "output"
)

NUMBER_OF_WEEKS = 4

TH_LOW = 0.30
TH_HIGH = 0.85

REGIONS = [
    "NA",
    "NA_DD",
    "LATAM",
    "EU",
    "APAC",
    "CHINA"
]


# ============================================================
# LOAD HISTORY
# ============================================================

def load_history():

    if not os.path.exists(HISTORY_FILE):

        raise FileNotFoundError(
            "Historical file not found: "
            + HISTORY_FILE
        )

    print()
    print(
        "Historical file used : "
        + HISTORY_FILE
    )

    # IMPORTANT:
    # keep_default_na=False prevents pandas from
    # converting the Region value "NA" into NaN.
    df = pd.read_csv(
        HISTORY_FILE,
        keep_default_na=False
    )

    if df.empty:

        raise ValueError(
            "Historical file is empty: "
            + HISTORY_FILE
        )

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
    # Text cleanup
    # --------------------------------------------------------

    for column in [
        "Cluster",
        "Region",
        "Country",
        "City"
    ]:

        df[column] = (
            df[column]
            .astype(str)
            .str.strip()
        )

    # --------------------------------------------------------
    # Fix any old NA values that may already exist
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
                not region
                or region.lower() == "nan"
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

                    if not df.at[
                        index,
                        "Country"
                    ]:

                        df.at[
                            index,
                            "Country"
                        ] = mapping.get(
                            "Country",
                            ""
                        )

                    if not df.at[
                        index,
                        "City"
                    ]:

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
    # Used Capacity
    # --------------------------------------------------------

    df["Used Capacity (%)"] = pd.to_numeric(
        df["Used Capacity (%)"],
        errors="coerce"
    )

    df = df.dropna(
        subset=["Used Capacity (%)"]
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

    print()
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

    print()

    print("Regions:")

    print(
        df["Region"].value_counts(
            dropna=False
        )
    )

    return df


# ============================================================
# BUILD WEEKLY SNAPSHOTS
# ============================================================

def build_weekly_data(df):

    working = df.copy()

    # --------------------------------------------------------
    # Create a Monday-based week
    #
    # Every daily record is assigned to its Monday-Sunday
    # reporting week.
    # --------------------------------------------------------

    working["Week_Start"] = (
        working["Date"]
        - pd.to_timedelta(
            working["Date"].dt.weekday,
            unit="D"
        )
    )

    working["Week_End"] = (
        working["Week_Start"]
        + pd.Timedelta(
            days=6
        )
    )

    # --------------------------------------------------------
    # Latest available snapshot for each cluster in each week
    # --------------------------------------------------------

    working = working.sort_values(
        by=[
            "Week_Start",
            "Cluster",
            "Date"
        ]
    )

    weekly = (
        working
        .groupby(
            [
                "Week_Start",
                "Cluster"
            ],
            as_index=False
        )
        .tail(1)
        .copy()
    )

    # --------------------------------------------------------
    # Keep only latest NUMBER_OF_WEEKS
    # --------------------------------------------------------

    available_weeks = sorted(
        weekly["Week_Start"].unique()
    )

    selected_weeks = available_weeks[
        -NUMBER_OF_WEEKS:
    ]

    weekly = weekly[
        weekly["Week_Start"].isin(
            selected_weeks
        )
    ].copy()

    # --------------------------------------------------------
    # Use the actual latest snapshot date as the report column
    #
    # This avoids creating fake dates when history is incomplete.
    # --------------------------------------------------------

    weekly["Report Date"] = weekly[
        "Date"
    ]

    # --------------------------------------------------------
    # Format date column
    # --------------------------------------------------------

    weekly["Week Label"] = weekly[
        "Report Date"
    ].apply(
        format_excel_date
    )

    # --------------------------------------------------------
    # Keep required columns
    # --------------------------------------------------------

    weekly = weekly[
        [
            "Cluster",
            "Region",
            "Country",
            "City",
            "Week_Start",
            "Report Date",
            "Week Label",
            "Used Capacity (%)"
        ]
    ]

    return weekly, selected_weeks


# ============================================================
# DATE FORMAT
# ============================================================

def format_excel_date(value):

    timestamp = pd.Timestamp(
        value
    )

    # Linux-safe replacement for %-d
    day = str(
        timestamp.day
    )

    month = timestamp.strftime(
        "%b"
    )

    year = timestamp.strftime(
        "%y"
    )

    return (
        day
        + "-"
        + month
        + "-"
        + year
    )


# ============================================================
# BUILD REGIONAL TABLE
# ============================================================

def build_region_table(
    weekly,
    region
):

    region_upper = region.upper()

    # --------------------------------------------------------
    # CHINA gets its own sheet based on Country
    # --------------------------------------------------------

    if region_upper == "CHINA":

        region_data = weekly[
            weekly["Country"]
            .str.upper()
            .eq("CHINA")
        ].copy()

    # --------------------------------------------------------
    # NA_DD remains as an empty required sheet
    # --------------------------------------------------------

    elif region_upper == "NA_DD":

        region_data = weekly.iloc[
            0:0
        ].copy()

    # --------------------------------------------------------
    # Normal regional sheets
    # --------------------------------------------------------

    else:

        region_data = weekly[
            weekly["Region"]
            .str.upper()
            .eq(region_upper)
        ].copy()

    if region_data.empty:

        return pd.DataFrame()

    # --------------------------------------------------------
    # Pivot dates into columns
    # --------------------------------------------------------

    pivot = region_data.pivot_table(
        index=[
            "Cluster",
            "Region",
            "Country",
            "City"
        ],
        columns="Report Date",
        values="Used Capacity (%)",
        aggfunc="last"
    ).reset_index()

    pivot.columns.name = None

    # --------------------------------------------------------
    # Sort cities
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
# CREATE EXCEL FORMAT OBJECTS
# ============================================================

def create_formats(workbook):

    formats = {}

    formats["title"] = workbook.add_format({
        "bold": True,
        "font_size": 14
    })

    formats["blue_header"] = workbook.add_format({
        "bold": True,
        "bg_color": "#D9EAF7",
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })

    formats["blue_cell"] = workbook.add_format({
        "bg_color": "#D9EAF7",
        "border": 1
    })

    formats["normal"] = workbook.add_format({
        "border": 1
    })

    formats["percent"] = workbook.add_format({
        "border": 1,
        "num_format": "0%"
    })

    formats["percent_bold"] = workbook.add_format({
        "border": 1,
        "bold": True,
        "num_format": "0%"
    })

    formats["date_percent"] = workbook.add_format({
        "border": 1,
        "num_format": "0%"
    })

    formats["raw_header"] = workbook.add_format({
        "bold": True,
        "bg_color": "#1F4E78",
        "font_color": "white",
        "border": 1,
        "align": "center"
    })

    formats["raw_normal"] = workbook.add_format({
        "border": 1
    })

    formats["raw_percent"] = workbook.add_format({
        "border": 1,
        "num_format": "0.00%"
    })

    return formats


# ============================================================
# CREATE REGIONAL SHEET
# ============================================================

def create_region_sheet(
    workbook,
    formats,
    region,
    region_table,
    week_labels
):

    ws = workbook.add_worksheet(
        region
    )

    # --------------------------------------------------------
    # Top filters
    # --------------------------------------------------------

    ws.write(
        "A1",
        "Region",
        formats["blue_cell"]
    )

    ws.write(
        "B1",
        region,
        formats["blue_cell"]
    )

    ws.write(
        "A2",
        "Type (Avamar/Netbackup/Rubrik)",
        formats["blue_cell"]
    )

    ws.write(
        "B2",
        "(All)",
        formats["blue_cell"]
    )

    # --------------------------------------------------------
    # Main heading
    # --------------------------------------------------------

    ws.write(
        "A5",
        "Sum of Used Capacity (%)",
        formats["title"]
    )

    # --------------------------------------------------------
    # Month heading
    # --------------------------------------------------------

    if week_labels:

        ws.write(
            "D5",
            "Month",
            formats["title"]
        )

    # --------------------------------------------------------
    # Main table headers
    # --------------------------------------------------------

    headers = [
        "City Location",
        "TH %-Low",
        "TH %-High"
    ]

    headers.extend(
        week_labels
    )

    header_row = 5

    for column, header in enumerate(
        headers
    ):

        ws.write(
            header_row,
            column,
            header,
            formats["blue_header"]
        )

    # --------------------------------------------------------
    # Main table data
    # --------------------------------------------------------

    if not region_table.empty:

        for row_number in range(
            len(region_table)
        ):

            excel_row = (
                header_row
                + 1
                + row_number
            )

            city = region_table.iloc[
                row_number
            ]["City"]

            if not city:

                city = region_table.iloc[
                    row_number
                ]["Cluster"]

            ws.write(
                excel_row,
                0,
                city,
                formats["normal"]
            )

            ws.write(
                excel_row,
                1,
                TH_LOW,
                formats["percent_bold"]
            )

            ws.write(
                excel_row,
                2,
                TH_HIGH,
                formats["percent_bold"]
            )

            for week_index, week in enumerate(
                week_labels
            ):

                # Find the actual column in the pivot
                date_column = None

                for column in region_table.columns:

                    if isinstance(
                        column,
                        pd.Timestamp
                    ):

                        if format_excel_date(
                            column
                        ) == week:

                            date_column = column
                            break

                    elif hasattr(
                        column,
                        "strftime"
                    ):

                        if format_excel_date(
                            column
                        ) == week:

                            date_column = column
                            break

                value = None

                if date_column is not None:

                    value = region_table.iloc[
                        row_number
                    ][date_column]

                if pd.notna(value):

                    ws.write(
                        excel_row,
                        3 + week_index,
                        float(value),
                        formats["percent"]
                    )

                else:

                    ws.write_blank(
                        excel_row,
                        3 + week_index,
                        None,
                        formats["percent"]
                    )

    # --------------------------------------------------------
    # Right-side chart source table
    # --------------------------------------------------------

    right_col = 9

    for column, header in enumerate(
        headers,
        start=right_col
    ):

        ws.write(
            header_row,
            column,
            header,
            formats["blue_header"]
        )

    if not region_table.empty:

        for row_number in range(
            len(region_table)
        ):

            excel_row = (
                header_row
                + 1
                + row_number
            )

            city = region_table.iloc[
                row_number
            ]["City"]

            if not city:

                city = region_table.iloc[
                    row_number
                ]["Cluster"]

            ws.write(
                excel_row,
                right_col,
                city,
                formats["normal"]
            )

            ws.write(
                excel_row,
                right_col + 1,
                TH_LOW,
                formats["percent_bold"]
            )

            ws.write(
                excel_row,
                right_col + 2,
                TH_HIGH,
                formats["percent_bold"]
            )

            for week_index, week in enumerate(
                week_labels
            ):

                date_column = None

                for column in region_table.columns:

                    if hasattr(
                        column,
                        "strftime"
                    ):

                        if format_excel_date(
                            column
                        ) == week:

                            date_column = column
                            break

                value = None

                if date_column is not None:

                    value = region_table.iloc[
                        row_number
                    ][date_column]

                if pd.notna(value):

                    ws.write(
                        excel_row,
                        right_col
                        + 3
                        + week_index,
                        float(value),
                        formats["percent"]
                    )

                else:

                    ws.write_blank(
                        excel_row,
                        right_col
                        + 3
                        + week_index,
                        None,
                        formats["percent"]
                    )

    # --------------------------------------------------------
    # Column widths
    # --------------------------------------------------------

    ws.set_column(
        "A:A",
        18
    )

    ws.set_column(
        "B:G",
        12
    )

    ws.set_column(
        "J:J",
        18
    )

    ws.set_column(
        "K:P",
        12
    )

    # --------------------------------------------------------
    # Freeze panes
    # --------------------------------------------------------

    ws.freeze_panes(
        6,
        0
    )

    # --------------------------------------------------------
    # Chart
    # --------------------------------------------------------

    if region_table.empty:

        return ws

    number_of_cities = len(
        region_table
    )

    if number_of_cities == 0:

        return ws

    chart = workbook.add_chart({
        "type": "column"
    })

    # --------------------------------------------------------
    # Bar colors
    # --------------------------------------------------------

    chart_colors = [
        "#5B9BD5",
        "#9DC3E6",
        "#BDD7EE",
        "#DDEBF7"
    ]

    # --------------------------------------------------------
    # Weekly capacity series
    # --------------------------------------------------------

    for index, week in enumerate(
        week_labels
    ):

        chart.add_series({

            "name": [
                region,
                header_row,
                right_col + 3 + index
            ],

            "categories": [
                region,
                header_row + 1,
                right_col,
                header_row + number_of_cities,
                right_col
            ],

            "values": [
                region,
                header_row + 1,
                right_col + 3 + index,
                header_row + number_of_cities,
                right_col + 3 + index
            ],

            "fill": {
                "color": chart_colors[
                    index
                    % len(chart_colors)
                ]
            },

            "border": {
                "color": chart_colors[
                    index
                    % len(chart_colors)
                ]
            },

            "data_labels": {
                "value": True,
                "num_format": "0%"
            }
        })

    # --------------------------------------------------------
    # Low threshold line
    # --------------------------------------------------------

    line_chart = workbook.add_chart({
        "type": "line"
    })

    line_chart.add_series({

        "name": [
            region,
            header_row,
            right_col + 1
        ],

        "categories": [
            region,
            header_row + 1,
            right_col,
            header_row + number_of_cities,
            right_col
        ],

        "values": [
            region,
            header_row + 1,
            right_col + 1,
            header_row + number_of_cities,
            right_col + 1
        ],

        "line": {
            "color": "#00B050",
            "width": 2.25
        },

        "marker": {
            "type": "none"
        }
    })

    # --------------------------------------------------------
    # High threshold line
    # --------------------------------------------------------

    line_chart.add_series({

        "name": [
            region,
            header_row,
            right_col + 2
        ],

        "categories": [
            region,
            header_row + 1,
            right_col,
            header_row + number_of_cities,
            right_col
        ],

        "values": [
            region,
            header_row + 1,
            right_col + 2,
            header_row + number_of_cities,
            right_col + 2
        ],

        "line": {
            "color": "#FF0000",
            "width": 2.25
        },

        "marker": {
            "type": "none"
        }
    })

    # --------------------------------------------------------
    # Combine charts
    # --------------------------------------------------------

    chart.combine(
        line_chart
    )

    chart.set_title({
        "name": ""
    })

    chart.set_legend({
        "position": "top"
    })

    chart.set_y_axis({

        "num_format": "0%",

        "min": 0,

        "max": 1,

        "major_unit": 0.25,

        "major_gridlines": {
            "visible": True,
            "line": {
                "color": "#D9D9D9"
            }
        }
    })

    # IMPORTANT:
    # City names appear underneath the bars.

    chart.set_x_axis({

        "name": "",

        "label_position": "low",

        "num_font": {
            "size": 8
        }
    })

    chart.set_plotarea({

        "border": {
            "color": "#FFFFFF"
        },

        "fill": {
            "color": "#FFFFFF"
        }
    })

    chart.set_chartarea({

        "border": {
            "color": "#D9D9D9"
        },

        "fill": {
            "color": "#FFFFFF"
        }
    })

    # --------------------------------------------------------
    # Insert chart
    # --------------------------------------------------------

    ws.insert_chart(
        "E14",
        chart,
        {
            "x_scale": 1.45,
            "y_scale": 1.35
        }
    )

    return ws


# ============================================================
# RAW DATA SHEET
# ============================================================

def create_raw_data_sheet(
    workbook,
    formats,
    history
):

    ws = workbook.add_worksheet(
        "Raw Data"
    )

    columns = [
        "Date",
        "Cluster",
        "Region",
        "Country",
        "City",
        "Used Capacity (%)"
    ]

    for column_number, header in enumerate(
        columns
    ):

        ws.write(
            0,
            column_number,
            header,
            formats["raw_header"]
        )

    for row_number in range(
        len(history)
    ):

        row = history.iloc[
            row_number
        ]

        ws.write(
            row_number + 1,
            0,
            row["Date"].strftime(
                "%Y-%m-%d"
            ),
            formats["raw_normal"]
        )

        ws.write(
            row_number + 1,
            1,
            row["Cluster"],
            formats["raw_normal"]
        )

        ws.write(
            row_number + 1,
            2,
            row["Region"],
            formats["raw_normal"]
        )

        ws.write(
            row_number + 1,
            3,
            row["Country"],
            formats["raw_normal"]
        )

        ws.write(
            row_number + 1,
            4,
            row["City"],
            formats["raw_normal"]
        )

        ws.write(
            row_number + 1,
            5,
            float(
                row["Used Capacity (%)"]
            ),
            formats["raw_percent"]
        )

    ws.set_column(
        "A:A",
        14
    )

    ws.set_column(
        "B:B",
        16
    )

    ws.set_column(
        "C:C",
        12
    )

    ws.set_column(
        "D:D",
        16
    )

    ws.set_column(
        "E:E",
        20
    )

    ws.set_column(
        "F:F",
        20
    )

    ws.freeze_panes(
        1,
        0
    )


# ============================================================
# GENERATE WEEKLY REPORT
# ============================================================

def generate_weekly_report():

    print()
    print("=" * 70)
    print("Rubrik Weekly Capacity Growth Report")
    print("=" * 70)

    # --------------------------------------------------------
    # Load history
    # --------------------------------------------------------

    history = load_history()

    # --------------------------------------------------------
    # Build weekly data
    # --------------------------------------------------------

    weekly, selected_weeks = build_weekly_data(
        history
    )

    if weekly.empty:

        raise ValueError(
            "No weekly capacity data available."
        )

    # --------------------------------------------------------
    # Week labels
    # --------------------------------------------------------

    week_labels = []

    for week in selected_weeks:

        week_data = weekly[
            weekly["Week_Start"].eq(
                week
            )
        ]

        if week_data.empty:
            continue

        # Use latest actual snapshot date
        latest_date = week_data[
            "Report Date"
        ].max()

        label = format_excel_date(
            latest_date
        )

        if label not in week_labels:

            week_labels.append(
                label
            )

    print()
    print(
        "Weeks included : "
        + str(len(week_labels))
    )

    for week in week_labels:

        print(
            "  "
            + week
        )

    print()

    print(
        "Clusters included : "
        + str(
            history["Cluster"].nunique()
        )
    )

    print(
        "Weekly columns     : "
        + ", ".join(
            week_labels
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
    # Filename
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

    workbook = xlsxwriter.Workbook(
        report_path
    )

    formats = create_formats(
        workbook
    )

    # --------------------------------------------------------
    # Create regional sheets
    # --------------------------------------------------------

    for region in REGIONS:

        region_table = build_region_table(
            weekly,
            region
        )

        create_region_sheet(
            workbook,
            formats,
            region,
            region_table,
            week_labels
        )

    # --------------------------------------------------------
    # Raw Data
    # --------------------------------------------------------

    create_raw_data_sheet(
        workbook,
        formats,
        history
    )

    # --------------------------------------------------------
    # Close workbook
    # --------------------------------------------------------

    workbook.close()

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
            history["Cluster"].nunique()
        )
    )

    print(
        "Weeks Included  : "
        + str(
            len(week_labels)
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
            REGIONS
            + ["Raw Data"]
        )
    )

    print("=" * 70)

    return report_path


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":

    report_file = generate_weekly_report()

    print()
    print(
        "Sending Weekly Capacity Report Email..."
    )

    from email_report import send_email

    send_email(
        report_file
    )
