from config import load_credentials
from config import load_controllers
from config import load_commands
from ssh_connection import connect_ontap
from command_executor import execute_command


def main():

    # Load configuration files
    credentials = load_credentials()
    controllers = load_controllers()
    commands = load_commands()

    username = credentials["username"]
    password = credentials["password"]

    # Dictionary to store results from all ONTAP clusters
    all_results = {}

    # Connect to each ONTAP controller
    for controller in controllers:

        print("=" * 60)
        print(f"Connecting to ONTAP Cluster : {controller}")
        print("=" * 60)

        try:
            ssh = connect_ontap(controller, username, password)

            print("✅ Connected Successfully\n")

            # Dictionary to store one controller's command outputs
            cluster_results = {}

            # Execute all commands
            for command in commands:

                print("-" * 60)
                print(f"Executing Command : {command}")
                print("-" * 60)

                output = execute_command(ssh, command)

                # Store command output
                cluster_results[command] = output

                # Display output on terminal
                print(output)
                print()

            # Save this controller's results
            all_results[controller] = cluster_results

            ssh.close()
            print("SSH Connection Closed.\n")

        except Exception as e:
            print(f"❌ Connection Failed: {e}")

    # Display summary
    print("\n" + "=" * 60)
    print("ONTAP Health Check Summary")
    print("=" * 60)

    for controller, results in all_results.items():
        print(f"Cluster : {controller}")
        print(f"Commands Executed : {len(results)}")
        print("-" * 60)

    # Return results for HTML report generation
    return all_results


if __name__ == "__main__":
    results = main()
