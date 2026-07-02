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

    # Connect to each ONTAP controller
    for controller in controllers:

        print("=" * 60)
        print(f"Connecting to ONTAP Cluster : {controller}")
        print("=" * 60)

        try:
            ssh = connect_ontap(controller, username, password)

            print("✅ Connected Successfully\n")

            # Execute all commands from commands.txt
            for command in commands:

                print("-" * 60)
                print(f"Executing Command : {command}")
                print("-" * 60)

                output = execute_command(ssh, command)

                print(output)
                print()

            ssh.close()
            print("SSH Connection Closed.\n")

        except Exception as e:
            print(f"❌ Connection Failed: {e}")


if __name__ == "__main__":
    main()
