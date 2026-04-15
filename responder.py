# ============================================================
#  responder.py — Writes the Success/Failure Report
# ============================================================
#
#  WHY THIS FILE EXISTS:
#    After the agent runs all actions, you need to KNOW
#    what happened. Did it succeed? Which step failed?
#    This file writes a clear report to response.txt
#    that you can read in Notepad immediately.
#
#  WHAT IT WRITES:
#    - Overall status: SUCCESS / PARTIAL / FAILURE
#    - How many steps passed out of total
#    - Per-step breakdown with ✓ or ✗
#    - Timestamp
#    - Full JSON log (for developers who want raw data)
#
#  STATUS MEANINGS:
#    SUCCESS  → all steps completed with no errors
#    PARTIAL  → some steps worked, some failed
#    FAILURE  → no steps worked at all
# ============================================================


# ── IMPORTS ──────────────────────────────────────────────────

import json
# json → used to serialize the full report as JSON at the bottom
# json.dumps() converts a Python dict → JSON string

import datetime
# datetime → used to get the current timestamp
# datetime.datetime.now().isoformat() → "2024-03-18T14:32:01.123456"


def write_research_report(
    command: str,
    query: str,
    fields: list[str],
    findings: dict[str, list[str]],
    sources: list[dict[str, str]],
    output_file: str = "response.txt",
    error: str | None = None,
):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    status_line = "✅ RESEARCH COMPLETE" if findings else "⚠️ RESEARCH INCOMPLETE"

    with open(output_file, "w", encoding="utf-8") as f:
        f.write("=" * 60 + "\n")
        f.write("  WEB RESEARCH REPORT\n")
        f.write("=" * 60 + "\n\n")
        f.write(f"  Status  : {status_line}\n")
        f.write(f"  Time    : {timestamp}\n")
        f.write(f"  Command : {command}\n")
        f.write(f"  Query   : {query}\n")
        f.write(f"  Fields  : {', '.join(fields)}\n\n")

        if findings:
            for field in fields:
                field_hits = findings.get(field, [])
                if not field_hits:
                    continue
                f.write(f"[{field.upper()}]\n")
                for idx, hit in enumerate(field_hits, 1):
                    f.write(f"{idx}. {hit}\n")
                f.write("\n")
        else:
            f.write("No relevant facts were extracted.\n")
            if error:
                f.write(f"Reason: {error}\n")
            f.write("\n")

        if sources:
            f.write("Sources checked:\n")
            for idx, source in enumerate(sources, 1):
                f.write(f"{idx}. {source.get('title', 'Untitled')} - {source.get('url', '')}\n")
            f.write("\n")
        else:
            f.write("Sources checked: UI automation mode (page text copied from browser).\n\n")

    print(f"[RESPONSE] {status_line} — report written to {output_file}")


# ── MAIN FUNCTION ─────────────────────────────────────────────

def write_response(command: str, actions: list, results: list):
    # WHY THIS FUNCTION:
    #   Takes all execution results, calculates pass/fail,
    #   and writes a human-readable + machine-readable report.
    #
    # PARAMETERS:
    #   command  → the original command string from commands.txt
    #   actions  → the full list of actions Gemini planned
    #   results  → list of result dicts from executor.py
    #              each result looks like:
    #              {
    #                "step": 1,
    #                "action": "open_app",
    #                "success": True,
    #                "note": "Opened excel.exe"
    #              }

    # ── CALCULATE STATISTICS ──────────────────────────────────

    total  = len(results)
    # Count how many results have success=True
    passed = sum(1 for r in results if r["success"])
    failed = total - passed

    # ── DETERMINE OVERALL STATUS ──────────────────────────────

    # All passed → SUCCESS
    # Some passed → PARTIAL (something went wrong mid-way)
    # None passed → FAILURE
    if failed == 0:
        overall = "✅ SUCCESS"
        status_symbol = "✅"
    elif passed > 0:
        overall = "⚠️  PARTIAL"
        status_symbol = "⚠️ "
    else:
        overall = "❌ FAILURE"
        status_symbol = "❌"

    # ── GET TIMESTAMP ─────────────────────────────────────────
    # isoformat() → standard readable format: 2024-03-18T14:32:01
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # ── BUILD THE FULL DATA DICT ───────────────────────────────
    # We store everything in a dict so we can:
    #   1. Write a readable text report
    #   2. Dump it as JSON at the bottom for tools/scripts
    report = {
        "timestamp":      timestamp,
        "command":        command,
        "overall_status": overall,
        "summary":        f"{passed}/{total} steps completed",
        "steps":          results
    }

    # ── WRITE TO response.txt ─────────────────────────────────
    # "w" = write mode → overwrites the file every time
    # This means response.txt always shows the LATEST result
    with open("response.txt", "w", encoding="utf-8") as f:

        # Header bar
        f.write("=" * 50 + "\n")
        f.write("       DESKTOP AI AGENT — RESPONSE REPORT\n")
        f.write("=" * 50 + "\n\n")

        # Summary section
        f.write(f"  Status  :  {overall}\n")
        f.write(f"  Steps   :  {passed}/{total} succeeded\n")
        f.write(f"  Time    :  {timestamp}\n")
        f.write(f"  Command :  {command}\n")
        f.write("\n")

        # Per-step breakdown
        f.write("-" * 50 + "\n")
        f.write("  STEP-BY-STEP BREAKDOWN\n")
        f.write("-" * 50 + "\n")

        for r in results:
            # Pick ✓ for success, ✗ for failure
            icon = "✓" if r["success"] else "✗"

            # Format: [✓] Step 1: open_app → Opened excel.exe
            f.write(f"  [{icon}] Step {r['step']}: {r['action']}\n")
            f.write(f"       └─ {r['note']}\n")

        # If some steps weren't reached (agent stopped early)
        steps_done = len(results)
        total_planned = len(actions)
        if steps_done < total_planned:
            skipped = total_planned - steps_done
            f.write(f"\n  ⚠  {skipped} step(s) were skipped due to earlier failure.\n")

        f.write("\n")

        # Separator before raw JSON
        f.write("=" * 50 + "\n")
        f.write("  RAW JSON LOG (for developers)\n")
        f.write("=" * 50 + "\n")

        # json.dumps with indent=2 → pretty-printed JSON
        # ensure_ascii=False → allows emoji in the output
        f.write(json.dumps(report, indent=2, ensure_ascii=False))
        f.write("\n")

    # Also print to terminal so developer can see it live
    print(f"\n[RESPONSE] {status_symbol} {overall} — {passed}/{total} steps completed")
    print(f"[RESPONSE] Full report written to response.txt\n")


# ── HELPER: Write "Thinking" Status ───────────────────────────

def write_thinking(command: str):
    # WHY THIS FUNCTION:
    #   While the agent is processing (before it finishes),
    #   we write a "THINKING" status to response.txt.
    #   This way, if you open response.txt while it's running,
    #   you see something instead of the old result.

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open("response.txt", "w", encoding="utf-8") as f:
        f.write("=" * 50 + "\n")
        f.write("  🤔 AGENT IS THINKING...\n")
        f.write("=" * 50 + "\n\n")
        f.write(f"  Time    : {timestamp}\n")
        f.write(f"  Command : {command}\n\n")
        f.write("  Gemini is analyzing your screen and planning actions.\n")
        f.write("  Please wait...\n")


# ── HELPER: Write Error Status ────────────────────────────────

def write_error(command: str, error: str):
    # WHY: if something goes catastrophically wrong (network error,
    # bad API key, etc.), we write a clear error to response.txt
    # instead of leaving the old result or crashing silently.

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open("response.txt", "w", encoding="utf-8") as f:
        f.write("=" * 50 + "\n")
        f.write("  ❌ AGENT ERROR\n")
        f.write("=" * 50 + "\n\n")
        f.write(f"  Time    : {timestamp}\n")
        f.write(f"  Command : {command}\n\n")
        f.write(f"  Error   : {error}\n\n")
        f.write("  TROUBLESHOOTING:\n")
        f.write("  - Check your GEMINI_API_KEY in .env\n")
        f.write("  - Check agent.log for full traceback\n")
        f.write("  - Make sure you have internet connection\n")
