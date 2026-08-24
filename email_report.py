import os
import smtplib

from email.message import EmailMessage
from datetime import datetime


def send_email(report_file):

    smtp_server = "mail.mattel.com"
    smtp_port = 25

    sender = "svc_iAutomate_Prod@Mattel.com"

    recipients = [
        "yahitha.kasireddy@mattel.com"
    ]

    today = datetime.now().strftime(
        "%d %b %Y"
    )

    # --------------------------------------------------------
    # Determine report type from filename
    # --------------------------------------------------------

    file_name = os.path.basename(
        report_file
    )

    if "Daily" in file_name:

        report_type = "Daily"

    elif "Weekly" in file_name:

        report_type = "Weekly"

    elif "Monthly" in file_name:

        report_type = "Monthly"

    else:

        report_type = "Capacity"

    # --------------------------------------------------------
    # Email
    # --------------------------------------------------------

    msg = EmailMessage()

    msg["Subject"] = (
        f"Rubrik {report_type} Capacity Report - "
        f"{today}"
    )

    msg["From"] = sender

    msg["To"] = ", ".join(
        recipients
    )

    msg.set_content(
        f"""
Hi Team,

Please find attached the latest Rubrik {report_type} Capacity Report generated on {today}.

Report Details

Report Type : {report_type}
Report Name : {file_name}
Generated On: {today}

Regards,
Automation Team
"""
    )

    # --------------------------------------------------------
    # Verify report exists
    # --------------------------------------------------------

    if not os.path.isfile(
        report_file
    ):

        raise FileNotFoundError(
            f"Report file not found: {report_file}"
        )

    # --------------------------------------------------------
    # Attach Excel
    # --------------------------------------------------------

    with open(
        report_file,
        "rb"
    ) as attachment:

        msg.add_attachment(
            attachment.read(),
            maintype="application",
            subtype=(
                "vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            ),
            filename=file_name
        )

    # --------------------------------------------------------
    # SMTP
    # --------------------------------------------------------

    print()
    print("=" * 70)
    print("Sending Email...")
    print("=" * 70)

    print(
        f"From       : {sender}"
    )

    print(
        f"To         : {', '.join(recipients)}"
    )

    print(
        f"Report Type: {report_type}"
    )

    print(
        f"Attachment : {file_name}"
    )

    print("=" * 70)

    with smtplib.SMTP(
        smtp_server,
        smtp_port
    ) as smtp:

        smtp.ehlo()

        smtp.send_message(
            msg,
            from_addr=sender,
            to_addrs=recipients
        )

    print("=" * 70)
    print("Email Sent Successfully")
    print("=" * 70)
