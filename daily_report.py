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
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import (
    Table,
    TableStyleInfo
)
from openpyxl.chart import (
    BarChart,
    LineChart,
    Reference
)

from growth import (
    build_forecast_table,
    flag_threshold_status
)


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = os.path.dirname(
    os.path.abspath(__file__)
)

OUTPUT_DIR = os.path.join(
    BASE_DIR,
    "output"
)

DAILY_LOW = 0.80
DAILY_MEDIUM = 0.90
DAILY_HIGH = 0.95


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

TITLE_FILL = PatternFill(
    fill_type="solid",
    fgColor="17365D"
)

TITLE_FONT = Font(
    color="FFFFFF",
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
# REPORT COLUMNS
# ============================================================

REPORT_COLUMNS = [
    "Month",
    "Active",
    "Location",
    "Type",
    "Target Storage",
    "Region",
    "Country",
    "City",
    "Total Capacity (GB)",
    "Used Capacity (GB)",
    "Free Capacity (GB)",
    "Used Capacity (%)",
    "Free Capacity (%)",
    "TH %-80",
    "TH %-90",
    "TH %-95",
    "Threshold Status",
    "Daily Growth (GB)",
    "Average Daily Consumption (GB)",
    "Estimated Days Until Exhaustion",
    "Estimated Exhaustion Date"
]


# ============================================================
# PREPARE DATAFRAME
# ============================================================

def prepare_dataframe(
    df,
    forecast_df=None
):

    data = df.copy()

    # --------------------------------------------------------
    # Normalize capacity column names
    # --------------------------------------------------------

    rename_map = {
        "Total Capacity":
            "Total Capacity (GB)",

        "Used Capacity":
            "Used Capacity (GB)",

        "Free Capacity":
            "Free Capacity (GB)",

        "Used %":
            "Used Capacity (%)",

        "Free %":
            "Free Capacity (%)"
    }

    data = data.rename(
        columns=rename_map
    )

    # --------------------------------------------------------
    # Month
    # --------------------------------------------------------

    if "Month" not in data.columns:

        data["Month"] = (
            datetime.now().strftime(
                "%Y-%m-%d"
            )
        )

    # --------------------------------------------------------
    # Thresholds
    # --------------------------------------------------------

    data["TH %-80"] = DAILY_LOW
    data["TH %-90"] = DAILY_MEDIUM
    data["TH %-95"] = DAILY_HIGH

    # --------------------------------------------------------
    # Default growth fields
    # --------------------------------------------------------

    data["Threshold Status"] = "Unknown"

    data["Daily Growth (GB)"] = None

    data[
        "Average Daily Consumption (GB)"
    ] = None

    data[
        "Estimated Days Until Exhaustion"
    ] = None

    data[
        "Estimated Exhaustion Date"
    ] = "Insufficient History"

    # --------------------------------------------------------
    # Merge forecast information
    #
    # history uses Cluster
    # current data uses Target Storage
    # --------------------------------------------------------

    if (
        forecast_df is not None
        and not forecast_df.empty
        and "Cluster" in forecast_df.columns
        and "Target Storage" in data.columns
    ):

        forecast_columns = [
            "Cluster",
            "Daily Growth (GB)",
            "Average Daily Consumption (GB)",
            "Estimated Days Until Exhaustion",
            "Estimated Exhaustion Date"
        ]

        available_columns = [
            column
            for column in forecast_columns
            if column in forecast_df.columns
        ]

        forecast_merge = forecast_df[
            available_columns
        ].copy()

        forecast_merge = forecast_merge.rename(
            columns={
                "Cluster":
                    "Target Storage"
            }
        )

        data = data.merge(
            forecast_merge,
            on="Target Storage",
            how="left",
            suffixes=(
                "",
                "_forecast"
            )
        )

        for column in [
            "Daily Growth (GB)",
            "Average Daily Consumption (GB)",
            "Estimated Days Until Exhaustion",
            "Estimated Exhaustion Date"
        ]:

            forecast_column = (
                f"{column}_forecast"
            )

            if forecast_column in data.columns:

                data[column] = (
                    data[forecast_column]
                    .combine_first(
                        data[column]
                    )
                )

                data.drop(
                    columns=[
                        forecast_column
                    ],
                    inplace=True
                )

    # --------------------------------------------------------
    # Threshold status
    # --------------------------------------------------------

    if "Used Capacity (%)" in data.columns:

        data["Threshold Status"] = (
            data["Used Capacity (%)"]
            .apply(
                lambda value:
                    flag_threshold_status(
                        value,
                        {
                            "warning":
                                DAILY_LOW,

                            "critical":
                                DAILY_MEDIUM,

                            "immediate":
                                DAILY_HIGH
                        }
                    )
            )
        )

    # --------------------------------------------------------
    # Ensure all columns exist
    # --------------------------------------------------------

    for column in REPORT_COLUMNS:

        if column not in data.columns:

            data[column] = ""

    # --------------------------------------------------------
    # Keep required columns
    # --------------------------------------------------------

    data = data[
        REPORT_COLUMNS
    ]

    return data


# ============================================================
# FORMAT WORKSHEET
# ============================================================

def format_worksheet(
    ws,
    df
):

    # --------------------------------------------------------
    # Header
    # --------------------------------------------------------

    for column_number, header in enumerate(
        df.columns,
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
            vertical="center",
            wrap_text=True
        )

    # --------------------------------------------------------
    # Body
    # --------------------------------------------------------

    for row in ws.iter_rows(
        min_row=2,
        max_row=ws.max_row
    ):

        for cell in row:

            cell.border = BORDER

            cell.alignment = Alignment(
                vertical="center"
            )

    # --------------------------------------------------------
    # Number formats
    # --------------------------------------------------------

    for column_number, header in enumerate(
        df.columns,
        start=1
    ):

        for row_number in range(
            2,
            ws.max_row + 1
        ):

            cell = ws.cell(
                row=row_number,
                column=column_number
            )

            if (
                "Capacity (%)" in header
                or header.startswith("TH %-")
            ):

                cell.number_format = "0.00%"

            elif (
                "Capacity (GB)" in header
                or "Growth (GB)" in header
                or "Consumption (GB)" in header
            ):

                cell.number_format = "#,##0.00"

            elif (
                "Days Until Exhaustion"
                in header
            ):

                cell.number_format = "0.0"

    # --------------------------------------------------------
    # Column widths
    # --------------------------------------------------------

    for column_number, column_name in enumerate(
        df.columns,
        start=1
    ):

        max_length = len(
            str(column_name)
        )

        for row_number in range(
            2,
            ws.max_row + 1
        ):

            value = ws.cell(
                row=row_number,
                column=column_number
            ).value

            if value is not None:

                max_length = max(
                    max_length,
                    len(str(value))
                )

        ws.column_dimensions[
            get_column_letter(
                column_number
            )
        ].width = min(
            max(max_length + 2, 12),
            40
        )

    ws.freeze_panes = "A2"


# ============================================================
# ADD TABLE
# ============================================================

def add_table(
    ws,
    table_name
):

    if ws.max_row < 2:
        return

    table = Table(
        displayName=table_name,
        ref=ws.dimensions
    )

    style = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    )

    table.tableStyleInfo = style

    ws.add_table(
        table
    )


# ============================================================
# CREATE REGION SHEET
# ============================================================

def create_region_sheet(
    wb,
    df,
    sheet_name
):

    ws = wb.create_sheet(
        sheet_name
    )

    if df.empty:

        ws["A1"] = (
            f"No capacity data available for "
            f"{sheet_name}"
        )

        ws["A1"].font = Font(
            bold=True
        )

        ws.column_dimensions[
            "A"
        ].width = 50

        return ws

    # --------------------------------------------------------
    # Write headers
    # --------------------------------------------------------

    for column_number, header in enumerate(
        df.columns,
        start=1
    ):

        ws.cell(
            row=1,
            column=column_number,
            value=header
        )

    # --------------------------------------------------------
    # Write data
    # --------------------------------------------------------

    for row_number, row_data in enumerate(
        df.itertuples(
            index=False,
            name=None
        ),
        start=2
    ):

        for column_number, value in enumerate(
            row_data,
            start=1
        ):

            ws.cell(
                row=row_number,
                column=column_number,
                value=value
            )

    format_worksheet(
        ws,
        df
    )

    safe_name = (
        sheet_name
        .replace(" ", "")
        .replace("-", "")
    )

    add_table(
        ws,
        f"Tbl{safe_name}"
    )

    return ws


# ============================================================
# CREATE RAW DATA SHEET
# ============================================================

def create_raw_data_sheet(
    wb,
    df
):

    ws = wb.create_sheet(
        "Raw Data"
    )

    if df.empty:

        ws["A1"] = (
            "No raw data available."
        )

        return ws

    for column_number, header in enumerate(
        df.columns,
        start=1
    ):

        ws.cell(
            row=1,
            column=column_number,
            value=header
        )

    for row_number, row_data in enumerate(
        df.itertuples(
            index=False,
            name=None
        ),
        start=2
    ):

        for column_number, value in enumerate(
            row_data,
            start=1
        ):

            ws.cell(
                row=row_number,
                column=column_number,
                value=value
            )

    format_worksheet(
        ws,
        df
    )

    add_table(
        ws,
        "TblRawData"
    )

    return ws


# ============================================================
# CREATE GLOBAL VIEW
# ============================================================

def create_global_view(
    wb,
    df,
    history_df=None
):

    ws = wb.create_sheet(
        "Global View"
    )

    # --------------------------------------------------------
    # Title
    # --------------------------------------------------------

    ws["A1"] = (
        "Rubrik Daily Capacity Report"
    )

    ws["A1"].font = Font(
        bold=True,
        size=16
    )

    # --------------------------------------------------------
    # Report date
    # --------------------------------------------------------

    ws["A3"] = "Report Date"

    if not df.empty:

        report_date = df.iloc[0][
            "Month"
        ]

    else:

        report_date = (
            datetime.now().strftime(
                "%Y-%m-%d"
            )
        )

    ws["B3"] = report_date

    # --------------------------------------------------------
    # Cluster count
    # --------------------------------------------------------

    ws["A4"] = "Clusters Included"
    ws["B4"] = len(df)

    # --------------------------------------------------------
    # Regional summary
    # --------------------------------------------------------

    if df.empty:

        ws["A6"] = (
            "No capacity data available."
        )

        return ws

    summary = (
        df.groupby("Region")
        .agg(
            Clusters=(
                "Target Storage",
                "count"
            ),
            Total_Capacity_GB=(
                "Total Capacity (GB)",
                "sum"
            ),
            Used_Capacity_GB=(
                "Used Capacity (GB)",
                "sum"
            ),
            Free_Capacity_GB=(
                "Free Capacity (GB)",
                "sum"
            )
        )
        .reset_index()
    )

    # --------------------------------------------------------
    # Regional percentages
    # --------------------------------------------------------

    summary[
        "Used Capacity (%)"
    ] = (
        summary["Used_Capacity_GB"]
        /
        summary["Total_Capacity_GB"]
    )

    summary[
        "Free Capacity (%)"
    ] = (
        summary["Free_Capacity_GB"]
        /
        summary["Total_Capacity_GB"]
    )

    # --------------------------------------------------------
    # Headers
    # --------------------------------------------------------

    start_row = 7

    headers = [
        "Region",
        "Clusters",
        "Total Capacity (GB)",
        "Used Capacity (GB)",
        "Free Capacity (GB)",
        "Used Capacity (%)",
        "Free Capacity (%)"
    ]

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
    # Summary data
    # --------------------------------------------------------

    for row_offset, row_data in enumerate(
        summary.itertuples(
            index=False,
            name=None
        ),
        start=1
    ):

        row_number = (
            start_row
            + row_offset
        )

        values = [
            row_data[0],
            row_data[1],
            row_data[2],
            row_data[3],
            row_data[4],
            row_data[5],
            row_data[6]
        ]

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

    # --------------------------------------------------------
    # Number formats
    # --------------------------------------------------------

    for row_number in range(
        start_row + 1,
        ws.max_row + 1
    ):

        for column_number in [
            3,
            4,
            5
        ]:

            ws.cell(
                row=row_number,
                column=column_number
            ).number_format = (
                "#,##0.00"
            )

        for column_number in [
            6,
            7
        ]:

            ws.cell(
                row=row_number,
                column=column_number
            ).number_format = (
                "0.00%"
            )

    # --------------------------------------------------------
    # Regional Used Capacity chart
    # --------------------------------------------------------

    if len(summary) > 0:

        chart = BarChart()

        chart.type = "bar"
        chart.style = 10

        chart.title = (
            "Regional Used Capacity"
        )

        chart.y_axis.title = "Region"

        chart.x_axis.title = (
            "Used Capacity (GB)"
        )

        chart_data = Reference(
            ws,
            min_col=4,
            min_row=start_row,
            max_row=(
                start_row
                + len(summary)
            )
        )

        categories = Reference(
            ws,
            min_col=1,
            min_row=start_row + 1,
            max_row=(
                start_row
                + len(summary)
            )
        )

        chart.add_data(
            chart_data,
            titles_from_data=True
        )

        chart.set_categories(
            categories
        )

        chart.height = 8
        chart.width = 15

        ws.add_chart(
            chart,
            "I7"
        )

    # --------------------------------------------------------
    # Forecast summary
    # --------------------------------------------------------

    if (
        history_df is not None
        and not history_df.empty
    ):

        forecast_df = build_forecast_table(
            history_df
        )

        if not forecast_df.empty:

            forecast_start = 22

            ws.cell(
                row=forecast_start,
                column=1,
                value="Forecast Summary"
            )

            ws.cell(
                row=forecast_start,
                column=1
            ).font = Font(
                bold=True,
                size=14
            )

            forecast_headers = [
                "Cluster",
                "Used Capacity (GB)",
                "Free Capacity (GB)",
                "Average Daily Consumption (GB)",
                "Estimated Days Until Exhaustion",
                "Estimated Exhaustion Date",
                "Threshold Status"
            ]

            header_row = (
                forecast_start + 1
            )

            for column_number, header in enumerate(
                forecast_headers,
                start=1
            ):

                cell = ws.cell(
                    row=header_row,
                    column=column_number,
                    value=header
                )

                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.border = BORDER

            # ------------------------------------------------
            # Match current data for status
            # ------------------------------------------------

            status_map = {}

            for _, row in df.iterrows():

                status_map[
                    str(
                        row["Target Storage"]
                    )
                ] = row[
                    "Threshold Status"
                ]

            for row_offset, (_, row) in enumerate(
                forecast_df.iterrows(),
                start=1
            ):

                row_number = (
                    header_row
                    + row_offset
                )

                cluster = str(
                    row["Cluster"]
                )

                values = [
                    cluster,
                    row.get(
                        "Latest Used Capacity (GB)"
                    ),
                    row.get(
                        "Latest Free Capacity (GB)"
                    ),
                    row.get(
                        "Average Daily Consumption (GB)"
                    ),
                    row.get(
                        "Estimated Days Until Exhaustion"
                    ),
                    row.get(
                        "Estimated Exhaustion Date"
                    ),
                    status_map.get(
                        cluster,
                        "Unknown"
                    )
                ]

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

            # ------------------------------------------------
            # Forecast number formats
            # ------------------------------------------------

            for row_number in range(
                header_row + 1,
                ws.max_row + 1
            ):

                ws.cell(
                    row=row_number,
                    column=2
                ).number_format = (
                    "#,##0.00"
                )

                ws.cell(
                    row=row_number,
                    column=3
                ).number_format = (
                    "#,##0.00"
                )

                ws.cell(
                    row=row_number,
                    column=4
                ).number_format = (
                    "#,##0.00"
                )

                ws.cell(
                    row=row_number,
                    column=5
                ).number_format = (
                    "0.0"
                )

    # --------------------------------------------------------
    # Historical growth trend chart
    #
    # Only created when at least two dates exist.
    # --------------------------------------------------------

    if (
        history_df is not None
        and not history_df.empty
        and "Date" in history_df.columns
    ):

        history = history_df.copy()

        history["Date"] = pd.to_datetime(
            history["Date"],
            errors="coerce"
        )

        history = history.dropna(
            subset=["Date"]
        )

        if history["Date"].nunique() >= 2:

            trend = (
                history
                .groupby("Date")[
                    "Used Capacity (GB)"
                ]
                .sum()
                .reset_index()
                .sort_values("Date")
            )

            trend_start = 22

            # Put trend data to the right side.
            trend_col = 9

            ws.cell(
                row=trend_start,
                column=trend_col,
                value="Date"
            )

            ws.cell(
                row=trend_start,
                column=trend_col + 1,
                value="Total Used Capacity (GB)"
            )

            for cell in ws[
                trend_start
            ][
                trend_col - 1:
                trend_col + 1
            ]:

                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.border = BORDER

            for offset, (_, row) in enumerate(
                trend.iterrows(),
                start=1
            ):

                row_number = (
                    trend_start
                    + offset
                )

                ws.cell(
                    row=row_number,
                    column=trend_col,
                    value=row["Date"]
                )

                ws.cell(
                    row=row_number,
                    column=trend_col + 1,
                    value=row[
                        "Used Capacity (GB)"
                    ]
                )

            chart = LineChart()

            chart.title = (
                "Capacity Growth Trend"
            )

            chart.y_axis.title = (
                "Used Capacity (GB)"
            )

            chart.x_axis.title = (
                "Date"
            )

            chart.height = 8
            chart.width = 16

            chart_data = Reference(
                ws,
                min_col=trend_col + 1,
                min_row=trend_start,
                max_row=(
                    trend_start
                    + len(trend)
                )
            )

            chart_categories = Reference(
                ws,
                min_col=trend_col,
                min_row=trend_start + 1,
                max_row=(
                    trend_start
                    + len(trend)
                )
            )

            chart.add_data(
                chart_data,
                titles_from_data=True
            )

            chart.set_categories(
                chart_categories
            )

            ws.add_chart(
                chart,
                "I40"
            )

    # --------------------------------------------------------
    # Widths
    # --------------------------------------------------------

    widths = {
        "A": 18,
        "B": 15,
        "C": 22,
        "D": 22,
        "E": 22,
        "F": 20,
        "G": 20,
        "I": 18,
        "J": 22,
        "K": 18
    }

    for column, width in widths.items():

        ws.column_dimensions[
            column
        ].width = width

    ws.freeze_panes = "A7"

    return ws


# ============================================================
# GENERATE DAILY EXCEL REPORT
# ============================================================

def generate_daily_report(
    df,
    history_df=None
):

    # --------------------------------------------------------
    # Create output directory
    # --------------------------------------------------------

    os.makedirs(
        OUTPUT_DIR,
        exist_ok=True
    )

    # --------------------------------------------------------
    # Build growth / forecast table
    # --------------------------------------------------------

    forecast_df = pd.DataFrame()

    if (
        history_df is not None
        and not history_df.empty
    ):

        forecast_df = build_forecast_table(
            history_df
        )

    # --------------------------------------------------------
    # Prepare current data
    # --------------------------------------------------------

    data = prepare_dataframe(
        df,
        forecast_df
    )

    # --------------------------------------------------------
    # File name
    # --------------------------------------------------------

    today = datetime.now()

    file_name = today.strftime(
        "Backup_Capacity_Daily_%d_%b_%Y.xlsx"
    )

    report_path = os.path.join(
        OUTPUT_DIR,
        file_name
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
    # Global View
    # --------------------------------------------------------

    create_global_view(
        wb,
        data,
        history_df
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

            region_df = data[
                data["Country"]
                .astype(str)
                .str.upper()
                .eq("CHINA")
            ]

        elif region == "NA_DD":

            region_df = data.iloc[0:0]

        else:

            region_df = data[
                data["Region"]
                .astype(str)
                .str.upper()
                .eq(region)
            ]

        create_region_sheet(
            wb,
            region_df,
            region
        )

    # --------------------------------------------------------
    # Raw Data
    # --------------------------------------------------------

    create_raw_data_sheet(
        wb,
        data
    )

    # --------------------------------------------------------
    # Workbook order
    # --------------------------------------------------------

    desired_order = [
        "Global View",
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
    print("Daily Excel Report Generated")
    print("=" * 70)

    print(
        f"Report Location : {report_path}"
    )

    print(
        f"Absolute Path   : "
        f"{os.path.abspath(report_path)}"
    )

    print(
        f"Clusters        : {len(data)}"
    )

    print(
        f"File Exists     : {file_exists}"
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
# DIRECT EXECUTION
# ============================================================

if __name__ == "__main__":

    print(
        "daily_report.py loaded successfully."
    )
