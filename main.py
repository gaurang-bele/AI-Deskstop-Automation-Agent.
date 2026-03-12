import time
from agent import process_command

COMMAND_FILE = "commands.txt"
RESPONSE_FILE = "response.txt"

last_command = ""

while True:

    with open(COMMAND_FILE, "r") as f:
        command = f.read().strip()

    if command != "" and command != last_command:

        print("New command:", command)

        result = process_command(command)

        print("Result:", result)

        with open(RESPONSE_FILE, "w") as f:
            f.write(result)

        last_command = command

    time.sleep(2)