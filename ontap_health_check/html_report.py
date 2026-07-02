from datetime import datetime


def generate_html_report(results):

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>ONTAP Health Report</title>

        <style>

            body {{
                font-family: Arial;
                margin:40px;
                background:#f4f4f4;
            }}

            h1 {{
                color:#005cb9;
            }}

            h2 {{
                color:#333333;
                border-bottom:2px solid #005cb9;
            }}

            pre {{
                background:white;
                border:1px solid #cccccc;
                padding:10px;
                overflow:auto;
            }}

        </style>

    </head>

    <body>

    <h1>ONTAP Health Report</h1>

    <p><b>Generated :</b> {datetime.now()}</p>

    """

    for controller, commands in results.items():

        html += f"<h2>Cluster : {controller}</h2>"

        for command, output in commands.items():

            html += f"""
            <h3>{command}</h3>

            <pre>{output}</pre>
            """

    html += """
    </body>
    </html>
    """

    with open("reports/ontap_health_report.html", "w") as file:
        file.write(html)

    print("\nHTML report generated successfully.")
