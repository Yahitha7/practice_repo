import json
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CREDENTIALS_FILE = os.path.join(BASE_DIR, "configs", "credentials.json")
CONTROLLERS_FILE = os.path.join(BASE_DIR, "inventory", "controllers.txt")
COMMANDS_FILE = os.path.join(BASE_DIR, "commands", "commands.txt")


def load_credentials():
    with open(CREDENTIALS_FILE, "r") as file:
        data = json.load(file)
    return data


def load_controllers():
    with open(CONTROLLERS_FILE, "r") as file:
        controllers = []
        for line in file:
            line = line.strip()
            if line:
                controllers.append(line)
    return controllers


def load_commands():
    with open(COMMANDS_FILE, "r") as file:
        commands = []
        for line in file:
            line = line.strip()
            if line:
                commands.append(line)
    return commands
