from config import load_credentials
from config import load_controllers
from config import load_commands
from ssh_connection import connect_ontap
from command_executor import execute_command
from html_report import generate_html_report


def main():
    """
    Main function to perform ONTAP Health Check
    """

    # Load configuration
    credentials = load_credentials()
    controllers = load_controllers()
    commands = load_commands()

    username = credentials["username"]
    password = credentials["password"]

    # Store results from all clusters
    all_results = {}

    # Loop through each ONTAP Cluster
    for controller in controllers:

        print("=" * 60)
        print(f"Connecting to ONTAP Cluster : {controller}")
        print("=" * 60)

        try:
            # Connect to ONTAP
            ssh = connect_ontap(controller, username, password)

            print("✅ Connected Successfully\n")

            # Store command outputs for one cluster
            cluster_results = {}

            # Execute each command
            for command in commands:

                print("-" * 60)
                print(f"Executing Command : {command}")
                print("-" * 60)

                output = execute_command(ssh, command)

                # Save output
                cluster_results[command] = output

                # Print output to terminal
                print(output)
                print()

            # Save all command outputs for this controller
            all_results[controller] = cluster_results

            ssh.close()
            print("SSH Connection Closed.\n")

        except Exception as e:

            print(f"❌ Connection Failed to {controller}")
            print(f"Reason : {e}")

            all_results[controller] = {
                "Connection Status": f"FAILED - {e}"
            }

    # Print Summary
    print("=" * 60)
    print("ONTAP Health Check Summary")
    print("=" * 60)

    for controller, results in all_results.items():
        print(f"Cluster : {controller}")
        print(f"Commands Executed : {len(results)}")
        print("-" * 60)

    # Generate HTML Report
    generate_html_report(all_results)

    print("\n✅ HTML Report Generated Successfully")
    print("Location : reports/ontap_health_report.html")

    return all_results


if __name__ == "__main__":
    main()
