# ============================================================
#  responder.py — Writes Success/Failure Reports
#
#  CHANGES IN THIS VERSION:
#    1. write_response() now also writes structured JSON status
#       that matches the Desktop Automation AI Agent spec format:
#       { goal, plan, actions[], validation, status, error }
#    2. Part code JSON results handled cleanly
#    3. write_research_report() unchanged (backward compat)
# ============================================================

import json
import datetime


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
    status_line = "RESEARCH COMPLETE" if findings else "RESEARCH INCOMPLETE"

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


def write_response(command: str, actions: list, results: list, llm_usage: dict | None = None):
    total = len(results)
    def _step_passed(result: dict) -> bool:
        if not result.get("success"):
            return False
        if result.get("action") == "screenshot" and result.get("verified") is False:
            return False
        return True

    passed = sum(1 for r in results if _step_passed(r))
    failed = total - passed

    if failed == 0:
        overall = "SUCCESS"
        status_symbol = "OK"
        agent_status = "done"
    elif passed > 0:
        overall = "PARTIAL"
        status_symbol = "WARN"
        agent_status = "continue"
    else:
        overall = "FAILURE"
        status_symbol = "FAIL"
        agent_status = "failed"

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    usage = {
        "requests": int((llm_usage or {}).get("requests", 0)),
        "prompt_tokens": int((llm_usage or {}).get("prompt_tokens", 0)),
        "completion_tokens": int((llm_usage or {}).get("completion_tokens", 0)),
        "total_tokens": int((llm_usage or {}).get("total_tokens", 0)),
    }

    # Build structured report matching Desktop Automation AI Agent spec
    report = {
        "timestamp": timestamp,
        "command": command,
        "goal": command,
        "overall_status": overall,
        "agent_status": agent_status,   # "done | continue | failed"
        "summary": f"{passed}/{total} steps completed",
        "steps": results,
        "validation": {
            "expected_result": "All planned actions completed successfully",
            "method": "visual",
            "actual_result": f"{passed}/{total} steps passed",
        },
        "llm_usage": usage,
        "error": None if agent_status == "done" else f"{failed} step(s) failed",
    }

    with open("response.txt", "w", encoding="utf-8") as f:
        f.write("=" * 50 + "\n")
        f.write("       DESKTOP AI AGENT — RESPONSE REPORT\n")
        f.write("=" * 50 + "\n\n")

        f.write(f"  Status  :  [{status_symbol}] {overall}\n")
        f.write(f"  Steps   :  {passed}/{total} succeeded\n")
        f.write(f"  Time    :  {timestamp}\n")
        f.write(f"  Command :  {command}\n")
        if usage["requests"] > 0:
            f.write(
                f"  LLM     :  {usage['total_tokens']} tokens "
                f"({usage['prompt_tokens']} prompt + {usage['completion_tokens']} completion) "
                f"across {usage['requests']} call(s)\n"
            )
        f.write("\n")

        f.write("-" * 50 + "\n")
        f.write("  STEP-BY-STEP BREAKDOWN\n")
        f.write("-" * 50 + "\n")

        for r in results:
            icon = "+" if _step_passed(r) else "X"
            f.write(f"  [{icon}] Step {r['step']}: {r['action']}\n")
            f.write(f"       |_ {r['note']}\n")

        steps_done = len(results)
        total_planned = len(actions)
        if steps_done < total_planned:
            skipped = total_planned - steps_done
            f.write(f"\n  WARN: {skipped} step(s) skipped due to earlier failure.\n")

        f.write("\n")
        f.write("=" * 50 + "\n")
        f.write("  RAW JSON LOG\n")
        f.write("=" * 50 + "\n")
        f.write(json.dumps(report, indent=2, ensure_ascii=False))
        f.write("\n")

    print(f"\n[RESPONSE] [{status_symbol}] {overall} — {passed}/{total} steps completed")
    print(f"[RESPONSE] Report written to response.txt\n")


def write_thinking(command: str):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open("response.txt", "w", encoding="utf-8") as f:
        f.write("=" * 50 + "\n")
        f.write("  AGENT IS THINKING...\n")
        f.write("=" * 50 + "\n\n")
        f.write(f"  Time    : {timestamp}\n")
        f.write(f"  Command : {command}\n\n")
        f.write("  The model is analyzing your screen and planning actions.\n")
        f.write("  Please wait...\n")


def write_error(command: str, error: str):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open("response.txt", "w", encoding="utf-8") as f:
        f.write("=" * 50 + "\n")
        f.write("  AGENT ERROR\n")
        f.write("=" * 50 + "\n\n")
        f.write(f"  Time    : {timestamp}\n")
        f.write(f"  Command : {command}\n\n")
        f.write(f"  Error   : {error}\n\n")
        f.write("  TROUBLESHOOTING:\n")
        f.write("  - Check OPENROUTER_API_KEY or NVIDIA_API_KEY in .env\n")
        f.write("  - Check agent.log for full traceback\n")
        f.write("  - Make sure you have internet connection\n")

        # Also write as JSON for programmatic consumers
        error_json = {
            "timestamp": timestamp,
            "command": command,
            "agent_status": "failed",
            "error": error,
        }
        f.write("\n")
        f.write(json.dumps(error_json, indent=2, ensure_ascii=False))
        f.write("\n")