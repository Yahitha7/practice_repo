from config import load_credentials
from config import load_controllers
from ssh_connection import connect_ontap

credentials = load_credentials()
controllers = load_controllers()

username = credentials["username"]
password = credentials["password"]

for controller in controllers:

    print(f"Connecting to {controller}...")

    ssh = connect_ontap(
        controller,
        username,
        password
    )

    print("Connection Successful!")

    ssh.close()
