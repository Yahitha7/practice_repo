import pandas as pd

from auth import get_token
from capacity import get_capacity, build_dataframe
from history import save_daily_history
from daily_report import generate_daily_report
from email_report import send_email


def main():

    print("=" * 70)
    print("Rubrik Daily Capacity Report")
    print("=" * 70)

    # --------------------------------------------------------
    # Authentication
    # --------------------------------------------------------

    print()
    print("Authenticating...")

    token = get_token()

    print("Authentication Successful")

    # --------------------------------------------------------
    # Collect Rubrik capacity
    # --------------------------------------------------------

    print()
    print("Collecting Capacity Information...")

    data = get_capacity(
        token
    )

    print(
        f"Retrieved {len(data)} clusters."
    )

    # --------------------------------------------------------
    # Build DataFrame
    # --------------------------------------------------------

    print()
    print("Building Report...")

    df = build_dataframe(
        data
    )

    print(
        f"Clusters included in report: "
        f"{len(df)}"
    )

    # --------------------------------------------------------
    # Validate
    # --------------------------------------------------------

    if len(df) == 0:

        raise ValueError(
            "No clusters were included in the report."
        )

    # --------------------------------------------------------
    # Save historical snapshot
    # --------------------------------------------------------

    print()
    print("Saving Historical Data...")

    history_file = save_daily_history(
        df
    )

    print(
        f"Historical data saved to: "
        f"{history_file}"
    )

    # --------------------------------------------------------
    # Load historical data
    # --------------------------------------------------------

    print()
    print("Loading Historical Data...")

    history_df = pd.read_csv(
        history_file
    )

    print(
        f"Historical rows loaded: "
        f"{len(history_df)}"
    )

    if "Date" in history_df.columns:

        print(
            f"Historical dates: "
            f"{history_df['Date'].nunique()}"
        )

    if "Cluster" in history_df.columns:

        print(
            f"Historical clusters: "
            f"{history_df['Cluster'].nunique()}"
        )

    # --------------------------------------------------------
    # Generate Daily Excel
    # --------------------------------------------------------

    print()
    print("Generating Daily Excel Report...")

    report_file = generate_daily_report(
        df,
        history_df
    )

    print(
        f"Daily report generated: "
        f"{report_file}"
    )

    # --------------------------------------------------------
    # Send Email
    # --------------------------------------------------------

    print()
    print("Sending Email...")

    send_email(
        report_file
    )

    # --------------------------------------------------------
    # Final status
    # --------------------------------------------------------

    print()
    print("=" * 70)
    print(
        "Rubrik Daily Capacity Report "
        "Completed Successfully"
    )
    print("=" * 70)

    print(
        f"Report  : {report_file}"
    )

    print(
        f"History : {history_file}"
    )

    print("=" * 70)


if __name__ == "__main__":
    main()
