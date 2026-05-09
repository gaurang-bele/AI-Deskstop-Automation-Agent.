# main.py — Alternate entry point (polling-based, no watchdog)
# Use this if watchdog causes issues. Normally run agent.py instead.
import time
from agent import on_command

COMMAND_FILE = "commands.txt"

last_command = ""

print("[MAIN] Polling commands.txt every 2s. Edit and save to run a command.")

while True:
    try:
        with open(COMMAND_FILE, "r") as f:
            command = f.read().strip()
        if command and command != last_command:
            print(f"[MAIN] New command: {command}")
            on_command(command)
            last_command = command
    except Exception as e:
        print(f"[MAIN] Error: {e}")
    time.sleep(2)