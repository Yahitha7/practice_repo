from config import load_credentials
from config import load_controllers
from config import load_commands

credentials = load_credentials()
controllers = load_controllers()
commands = load_commands()

print("Username :", credentials["username"])

print("\nControllers:")
for controller in controllers:
    print(controller)

print("\nCommands:")
for command in commands:
    print(command)
