# ============================================================
#  watcher.py — Watches commands.txt for changes
# ============================================================
#
#  WHY THIS FILE EXISTS:
#    Instead of you running the agent manually every time,
#    this file "listens" to commands.txt in the background.
#    The moment you save new text in Notepad → it fires.
#
#  HOW IT WORKS (simple analogy):
#    Think of it like a guard standing outside a door.
#    The moment someone opens the door (saves the file),
#    the guard wakes up and calls for action.
# ============================================================


# ── IMPORTS ──────────────────────────────────────────────────

import time
import os
# time → we use time.sleep() to add small delays so the file
# is fully written before we try to read it

from watchdog.observers import Observer
# Observer → this is the "guard" that runs in a background thread
# and monitors the folder for any file system changes

from watchdog.events import FileSystemEventHandler
# FileSystemEventHandler → a base class we inherit from.
# It has methods like on_modified(), on_created() that trigger
# when files change. We override on_modified() to do OUR logic.


# ── THE WATCHER CLASS ─────────────────────────────────────────

class CommandWatcher(FileSystemEventHandler):
    # WHY A CLASS:
    #   watchdog requires us to subclass FileSystemEventHandler.
    #   We extend it by adding our own callback logic.
    #   A "callback" = a function we pass in to be called later.

    def __init__(self, callback):
        # __init__ runs when we create a CommandWatcher object.
        # callback → the function in agent.py we want to call
        #            when a new command is detected
        self.callback = callback

        # last_signature -> (file mtime, content) of last processed save.
        # WHY: watchdog can fire multiple times for one save.
        # We want exactly one trigger per save, even if command text is unchanged.
        self.last_signature = None

    def on_modified(self, event):
        # on_modified() is called by watchdog automatically
        # every time ANY file in the watched folder is modified.
        # event.src_path → the full path of the file that changed

        # WHY THIS CHECK:
        #   We only care about commands.txt, not other files.
        #   endswith() checks if the path ends with our filename.
        if not event.src_path.endswith("commands.txt"):
            return  # ignore all other file changes

        try:
            # Small delay → sometimes watchdog fires before the OS
            # has finished writing the file. 0.2s gives it time.
            time.sleep(0.2)

            # Open and read the file content
            # "r" = read mode (not write)
            with open("commands.txt", "r", encoding="utf-8") as f:
                content = f.read().strip()
            # .strip() removes leading/trailing whitespace/newlines
            # so "  open excel  \n" becomes "open excel"

            signature = (os.path.getmtime("commands.txt"), content)

            # Process once per save. Allow same command text on later saves.
            if content and signature != self.last_signature:
                self.last_signature = signature
                print(f"\n[WATCHER] New command detected: '{content}'")
                self.callback(content)        # fire the agent!

        except Exception as e:
            # If the file is locked or unreadable, don't crash —
            # just print the error and keep watching
            print(f"[WATCHER ERROR] {e}")


# ── START FUNCTION ────────────────────────────────────────────

def start_watching(callback):
    # This function creates and starts the file watcher.
    # We call it from agent.py, passing our on_command function.

    # Create our custom event handler with the callback attached
    event_handler = CommandWatcher(callback)

    # Create the Observer — this is the background thread
    observer = Observer()

    # Tell the observer: watch the current folder "."
    # recursive=False → don't watch subfolders, just root
    observer.schedule(event_handler, path=".", recursive=False)

    # Start the observer thread (runs in background, non-blocking)
    observer.start()

    print("[WATCHER] Started. Watching commands.txt for changes...")
    print("[WATCHER] Open commands.txt in Notepad, type a command, and save.\n")

    # Return the observer so agent.py can stop it later (on Ctrl+C)
    return observer
