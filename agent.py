# ============================================================
#  agent.py — The Main Entry Point (Run This File)
# ============================================================
#
#  WHY THIS FILE EXISTS:
#    This is the "conductor" that ties every other file together.
#    It starts the file watcher, receives commands, calls the
#    LLM for a plan, runs the executor, and writes responses.
#
#  HOW TO RUN:
#    python agent.py
#
#  WHAT HAPPENS AFTER RUNNING:
#    1. Agent starts and prints a welcome message
#    2. It watches commands.txt in the background
#    3. You open commands.txt in Notepad
#    4. You type a command and save (Ctrl+S)
#    5. The agent wakes up, plans, executes, responds
#    6. You open response.txt to see what happened
#    7. Repeat from step 3 for a new command
#
#  TO STOP THE AGENT:
#    Press Ctrl+C in the terminal
#    OR move your mouse to the top-left corner of the screen
# ============================================================


# ── IMPORTS ──────────────────────────────────────────────────

import time
# time → used in the main loop (time.sleep keeps it alive)

import logging
# logging → we set up a log file here so ALL modules write
# to the same agent.log file

import os
import re
from urllib.parse import quote_plus
# os → used to create initial files and folders on startup

# Import our own modules (the other files we built)
from watcher   import start_watching   # the file watcher
from llm       import get_action_plan, get_excel_action_plan, generate_professional_email, verify_step  # NVIDIA brain
from executor  import execute_action, get_clipboard_text   # action runner
from excel_automation import (
    EXCEL_ACTIONS,
    is_excel_command,
    parse_excel_actions_from_command,
    validate_excel_actions,
)
from email_automation import (
    build_email_actions,
    is_email_command,
    parse_email_actions_from_command,
    parse_email_request,
    validate_email_actions,
    wants_model_polish,
)
from responder import write_response, write_thinking, write_error, write_research_report
# Each import pulls the function we defined in those files


GENERIC_ACTIONS = {
    "open_app",
    "open_file",
    "open_url",
    "click",
    "double_click",
    "right_click",
    "click_text",
    "type_text",
    "paste_text",
    "press_key",
    "scroll",
    "wait",
    "screenshot",
    "drag",
    "move_mouse",
    "delete_rows",
}


def _get_excel_mode() -> str:
    """Choose Excel routing strategy: ui, hybrid, or structured."""
    mode = os.getenv("AGENT_EXCEL_MODE", "hybrid").strip().lower()
    if mode not in {"ui", "hybrid", "structured"}:
        return "hybrid"
    return mode


def _unsupported_actions(actions: list[dict]) -> list[str]:
    supported = GENERIC_ACTIONS.union(EXCEL_ACTIONS)

    unknown = []
    for action in actions or []:
        name = action.get("action") if isinstance(action, dict) else None
        if not name or name not in supported:
            unknown.append(str(name))
    return unknown


def _is_web_research_command(command: str) -> bool:
    lower = command.lower()
    research_terms = (
        "fetch information",
        "find information",
        "research",
        "search for",
        "look up",
        "lookup",
    )
    return any(term in lower for term in research_terms) and (
        "google" in lower or "about" in lower or "related to" in lower
    )


def _parse_research_request(command: str) -> tuple[str, list[str], str]:
    lower = command.lower()

    output_file = "response.txt"
    file_match = re.search(r"\b(?:in|into|to)\s+([a-zA-Z0-9_.-]+\.txt)\b", command, re.IGNORECASE)
    if file_match:
        output_file = file_match.group(1)

    fields = []
    if "price" in lower or "cost" in lower:
        fields.append("price")
    if "density" in lower:
        fields.append("density")
    if "use" in lower or "used for" in lower or "application" in lower:
        fields.append("use")
    if not fields:
        fields = ["price", "density", "use"]

    query = command
    for pattern in (r"related to\s+(.+)", r"about\s+(.+)", r"for\s+(.+)"):
        match = re.search(pattern, command, re.IGNORECASE)
        if match:
            query = match.group(1)
            break
    query = re.sub(r"\s+and\s+(?:save|write|paste|put)\b.*$", "", query, flags=re.IGNORECASE).strip(" .")
    query = re.sub(r"\s+in(?:to)?\s+[a-zA-Z0-9_.-]+\.txt$", "", query, flags=re.IGNORECASE).strip(" .")
    if not query:
        query = command.strip()

    return query, fields, output_file


def _extract_findings_from_clipboard(text: str, fields: list[str]) -> dict[str, list[str]]:
    snippets = re.split(r"(?<=[.!?])\s+", text or "")
    findings = {field: [] for field in fields}

    for snippet in snippets:
        s = snippet.strip()
        if len(s) < 30 or len(s) > 240:
            continue
        lower = s.lower()

        if "price" in fields and (
            "price" in lower
            or "unit price" in lower
            or re.search(r"(\$|usd|eur|inr|₹|€)\s?\d+(?:[.,]\d+)?", s, re.IGNORECASE)
        ):
            if len(findings["price"]) < 5:
                findings["price"].append(s)

        if "density" in fields and "density" in lower:
            if len(findings["density"]) < 5:
                findings["density"].append(s)

        if "use" in fields and any(token in lower for token in ("use", "used for", "application", "suitable for")):
            if len(findings["use"]) < 5:
                findings["use"].append(s)

    return {k: v for k, v in findings.items() if v}


# ── LOGGING SETUP ─────────────────────────────────────────────

# Configure the logging system ONCE here — all modules use it.
# filename="agent.log" → write logs to this file
# level=INFO          → log INFO, WARNING, ERROR (not DEBUG spam)
# format             → every line: "2024-03-18 14:32:01 [INFO] message"
logging.basicConfig(
    filename="agent.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)


# ── STARTUP INITIALIZATION ────────────────────────────────────

def initialize():
    # WHY THIS FUNCTION:
    #   Creates required files and folders when the agent
    #   starts for the first time, so there are no errors.

    # Create screenshots folder for verification screenshots
    os.makedirs("screenshots", exist_ok=True)

    # Create commands.txt if it doesn't exist
    # This is the file you'll type commands into
    if not os.path.exists("commands.txt"):
        with open("commands.txt", "w", encoding="utf-8") as f:
            f.write("")  # create empty file
        print("[AGENT] Created commands.txt — open this in Notepad!")

    # Create an initial response.txt so it always exists
    if not os.path.exists("response.txt"):
        with open("response.txt", "w", encoding="utf-8") as f:
            f.write("Agent is ready. Type a command in commands.txt and save.\n")

    print("[AGENT] Initialization complete.")


# ── CORE FUNCTION: Handle a New Command ──────────────────────

def on_command(command: str):
    # WHY THIS FUNCTION:
    #   This is the callback we pass to start_watching().
    #   Every time you save a new command in commands.txt,
    #   watchdog calls this function with the command text.
    #
    # PARAMETER:
    #   command → the text you typed in commands.txt
    #
    # FLOW:
    #   1. Write "Thinking..." to response.txt immediately
    #   2. Ask NVIDIA for an action plan
    #   3. Execute each action one by one
    #   4. Track results
    #   5. Write final report to response.txt

    print("\n" + "─" * 50)
    print(f"[AGENT] 📥 New command received: '{command}'")
    print("─" * 50)
    logging.info(f"Command received: {command}")

    # ── STEP 1: Tell user we're thinking ──────────────────────
    # We write "thinking" status IMMEDIATELY so if you check
    # response.txt right away, you see progress, not old result
    write_thinking(command)

    try:
        excel_mode = _get_excel_mode()
        excel_request = is_excel_command(command)
        email_request = is_email_command(command)
        research_request = _is_web_research_command(command)
        actions = None

        # ── STEP 2: Build Action Plan ───────────────────────
        if research_request:
            print("[AGENT] 🌐 Web research command detected - using UI automation flow...")
            query, fields, output_file = _parse_research_request(command)
            google_query = quote_plus(f"{query} {' '.join(fields)}")
            search_url = f"google.com/search?q={google_query}"

            # UI-only flow: open browser, search, open result pages, copy text.
            actions = [
                {"action": "open_app", "app": "chrome"},
                {"action": "wait", "seconds": 2},
                {"action": "press_key", "keys": ["ctrl", "l"]},
                {"action": "type_text", "text": search_url},
                {"action": "press_key", "keys": ["enter"]},
                {"action": "wait", "seconds": 3},
                # Open first 3 likely result links in sequence using keyboard navigation.
                {"action": "press_key", "keys": ["tab"]},
                {"action": "press_key", "keys": ["tab"]},
                {"action": "press_key", "keys": ["tab"]},
                {"action": "press_key", "keys": ["enter"]},
                {"action": "wait", "seconds": 3},
                {"action": "press_key", "keys": ["ctrl", "a"]},
                {"action": "press_key", "keys": ["ctrl", "c"]},
            ]

            results = []
            for i, action in enumerate(actions, 1):
                success, note = execute_action(action)
                results.append(
                    {
                        "step": i,
                        "action": action.get("action", "unknown"),
                        "params": action,
                        "success": success,
                        "note": note,
                    }
                )
                if not success:
                    break

            clipboard_text = get_clipboard_text()
            findings = _extract_findings_from_clipboard(clipboard_text, fields)
            write_research_report(
                command=command,
                query=query,
                fields=fields,
                findings=findings,
                sources=[],
                output_file=output_file,
                error=None if findings else "Could not extract requested details from copied page text.",
            )
            return

        # Email commands use a deterministic Gmail parser to support
        # both structured and natural-language requests.
        elif email_request:
            print("[AGENT] 📧 Email command detected - building Gmail action plan...")
            if wants_model_polish(command):
                details = parse_email_request(command)
                raw_intent = details.get("body") or details.get("raw_text") or command
                recipient_email = (details.get("to") or [None])[0]
                try:
                    polished = generate_professional_email(raw_intent, recipient_email=recipient_email)
                    details["subject"] = polished.get("subject", details.get("subject", "Quick update"))
                    details["body"] = polished.get("body", details.get("body", ""))
                    print("[AGENT] ✨ Email draft polished by model.")
                except Exception as polish_error:
                    print(f"[AGENT] ⚠ Could not polish email with model: {polish_error}")
                actions = build_email_actions(details)
            else:
                actions = parse_email_actions_from_command(command)
            valid, note = validate_email_actions(actions)
            if not valid:
                raise Exception(f"Email action plan invalid: {note}")

        # Default is hybrid: try generic desktop planner first,
        # then fallback to structured Excel only when needed.
        elif excel_request and excel_mode == "structured":
            print("[AGENT] 📊 Excel mode=structured - using Excel parser/DSL...")
            actions = parse_excel_actions_from_command(command)
            valid, note = validate_excel_actions(actions)
            if not valid:
                print(f"[AGENT] Excel parser not specific enough ({note}) - asking NVIDIA for structured Excel actions...")
                actions = get_excel_action_plan(command)
                valid, note = validate_excel_actions(actions)
                if not valid:
                    raise Exception(f"Excel action plan invalid: {note}")
        else:
            print("[AGENT] 🧠 Asking NVIDIA for generic action plan...")
            actions = get_action_plan(command)

            unknown = _unsupported_actions(actions)
            if unknown and excel_request and excel_mode == "hybrid":
                print(f"[AGENT] ⚠ Planner returned unsupported action(s) {unknown} - switching to structured Excel fallback...")
                actions = parse_excel_actions_from_command(command)
                valid, note = validate_excel_actions(actions)
                if not valid:
                    actions = get_excel_action_plan(command)
                    valid, note = validate_excel_actions(actions)
                    if not valid:
                        raise Exception(f"Excel action plan invalid: {note}")

        # Print the full plan so developer can see it in terminal
        print(f"[AGENT] 📋 Planned {len(actions)} actions:")
        for i, a in enumerate(actions, 1):
            print(f"         Step {i}: {a}")

        # ── STEP 3: Execute Each Action ───────────────────────
        results = []  # we'll collect results here

        for i, action in enumerate(actions, 1):
            print(f"\n[AGENT] ▶ Running Step {i}/{len(actions)}: {action.get('action')}")

            # execute_action() returns (success: bool, note: str)
            success, note = execute_action(action)

            # Build a result record for this step
            result = {
                "step":    i,
                "action":  action.get("action", "unknown"),
                "params":  action,        # store the full action dict
                "success": success,
                "note":    note
            }

            # If this was a screenshot action, also ask NVIDIA to
            # verify whether the task looks correct on screen
            if action.get("action") == "screenshot" and success:
                label = action.get("label", "step")
                print(f"[AGENT] 🔍 Verifying step with NVIDIA vision...")
                verified = verify_step(command, label)
                result["verified"] = verified
                result["note"]     = note + (" | Verified ✓" if verified else " | Verify failed ✗")

            results.append(result)

            # ── STOP EARLY IF CRITICAL STEP FAILS ─────────────
            # If a step failed, we stop. No point continuing because
            # next steps likely depend on the previous one succeeding.
            # e.g. if "open Excel" failed, no point clicking inside it
            if not success:
                print(f"[AGENT] ✗ Step {i} failed — stopping execution")
                logging.warning(f"Stopped at step {i}: {note}")
                break

            print(f"[AGENT] ✓ Step {i} done: {note}")

        # ── STEP 4: Write Final Response ──────────────────────
        # Pass command, full plan, and execution results to responder
        write_response(command, actions, results)

    except Exception as e:
        # ── CATCH ANY UNEXPECTED ERROR ────────────────────────
        # This catches: bad API key, network errors, JSON parse
        # failures, etc. We write a clear error to response.txt
        error_msg = str(e)
        print(f"[AGENT] ❌ Error: {error_msg}")
        logging.error(f"Agent error for command '{command}': {error_msg}", exc_info=True)
        # exc_info=True → also logs the full stack trace to agent.log
        write_error(command, error_msg)


# ── ENTRY POINT ───────────────────────────────────────────────

# WHY "if __name__ == '__main__'":
#   This block ONLY runs when you directly run: python agent.py
#   If another file imports agent.py, this block is skipped.
#   It's a Python best practice for the main script.

if __name__ == "__main__":

    # ── WELCOME BANNER ────────────────────────────────────────
    print("\n" + "═" * 50)
    print("   🤖  DESKTOP AI AGENT  v1.0")
    print("   Powered by Google NVIDIA")
    print("═" * 50)
    print()

    # ── INITIALIZE ────────────────────────────────────────────
    initialize()
    print()

    # ── INSTRUCTIONS ──────────────────────────────────────────
    print("📌 HOW TO USE:")
    print("   1. Open  commands.txt  in Notepad")
    print("   2. Type a command  (e.g. 'open calculator and calculate 5+5')")
    print("   3. Save the file  (Ctrl+S)")
    print("   4. Watch the agent work on your screen")
    print("   5. Open  response.txt  to see the result")
    print()
    print("🛑 TO STOP:  Press Ctrl+C  OR  move mouse to TOP-LEFT corner")
    print()

    # ── START FILE WATCHER ────────────────────────────────────
    # start_watching(on_command) → passes our function as callback
    # watchdog will call on_command(text) every time you save
    observer = start_watching(on_command)

    # ── MAIN LOOP ─────────────────────────────────────────────
    # WHY THIS LOOP:
    #   The watcher runs in a background thread.
    #   If we don't keep the main thread alive, the program exits.
    #   time.sleep(1) keeps it alive while using almost no CPU.
    #   KeyboardInterrupt fires when you press Ctrl+C.
    try:
        while True:
            time.sleep(1)   # keep the program alive, wait for commands

    except KeyboardInterrupt:
        # Gracefully shut down the watcher thread
        print("\n\n[AGENT] Ctrl+C detected — shutting down...")
        observer.stop()    # tell the observer thread to stop
        observer.join()    # wait for it to fully finish
        print("[AGENT] Goodbye! ✌️")
