# ============================================================
#  agent.py — Main Orchestrator
#
#  FIXES IN THIS VERSION:
#    1. TRUE AGENTIC LOOP — after each action, current screenshot
#       + last command + full action_history sent back to LLM.
#       LLM decides next action. Loop runs until status=done/failed
#       or MAX_LOOP_ITERATIONS reached.
#    2. PART CODES — always saved to part_code_results.txt (not response.txt)
#    3. LLM LOG — 401 errors surface with clear fix instructions
#    4. LOOP PREVENTION — same screenshot label 2x stops execution
#    5. VERIFY FUNCTION — each screenshot step verified; failure
#       increments a counter; 3 consecutive failures = abort
#    6. EMAIL SEND — fixed: uses ctrl+enter, no Tesseract needed
# ============================================================

import time
import logging
import os
import re
import json
import uuid
from urllib.parse import quote_plus

from watcher import start_watching
from llm import (
    get_action_plan,
    get_excel_action_plan,
    generate_professional_email,
    get_usage_totals,
    set_command_context,
    clear_command_context,
    log_command_event,
    verify_step,
)
from executor import execute_action, get_clipboard_text, set_log_context
from excel_automation import (
    EXCEL_ACTIONS,
    is_excel_command,
    parse_excel_actions_from_command,
    validate_excel_actions,
)
from email_automation import (
    build_email_actions,
    is_email_command,
    parse_email_request,
    validate_email_actions,
    wants_model_polish,
)
from responder import write_response, write_thinking, write_error, write_research_report


GENERIC_ACTIONS = {
    "open_app", "open_file", "open_url", "click", "double_click", "right_click",
    "click_text", "type_text", "paste_text", "press_key", "scroll", "wait",
    "screenshot", "drag", "move_mouse", "delete_rows",
}

# Maximum iterations for the agentic loop to prevent runaway execution
MAX_LOOP_ITERATIONS = 20

# ── Logging ───────────────────────────────────────────────────

logging.basicConfig(
    filename="agent.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)


# ── Excel mode ────────────────────────────────────────────────

def _get_excel_mode() -> str:
    mode = os.getenv("AGENT_EXCEL_MODE", "hybrid").strip().lower()
    return mode if mode in {"ui", "hybrid", "structured"} else "hybrid"


def _unsupported_actions(actions: list[dict]) -> list[str]:
    supported = GENERIC_ACTIONS.union(EXCEL_ACTIONS)
    return [
        str(action.get("action"))
        for action in (actions or [])
        if isinstance(action, dict) and action.get("action") not in supported
    ]


# ── Part code detection ───────────────────────────────────────

_PART_CODE_RE = re.compile(
    r"\b[A-Z0-9]{2,}[A-Z][0-9]{2,}[A-Z0-9]*\b"
    r"|\b[A-Z]{2,}\d{3,}[A-Z0-9]*\b",
    re.IGNORECASE,
)

def _looks_like_part_code(token: str) -> bool:
    if len(token) < 5:
        return False
    has_letter = any(c.isalpha() for c in token)
    has_digit = any(c.isdigit() for c in token)
    return has_letter and has_digit


def _extract_part_codes(command: str) -> list[str]:
    tokens = re.findall(r"[A-Z0-9]{5,}", command, re.IGNORECASE)
    unique_codes: list[str] = []
    seen: set[str] = set()
    for token in tokens:
        normalized = token.upper()
        if not _looks_like_part_code(normalized) or normalized in seen:
            continue
        seen.add(normalized)
        unique_codes.append(normalized)
    return unique_codes


_KNOWN_COMPONENT_MAKERS = {
    "avx", "kyocera", "murata", "tdk", "yageo", "vishay",
    "samsung", "kemet", "nichicon", "panasonic", "walsin", "taiyo yuden",
}


# ── Command classification ────────────────────────────────────

def _is_part_code_command(command: str) -> bool:
    lower = command.lower()
    part_triggers = (
        "part code", "part number", "component", "datasheet",
        "manufacturer", "mouser", "digikey", "digi-key", "octopart",
        "specification", "product details", "get me details",
    )
    has_trigger = any(t in lower for t in part_triggers)
    has_part_code = bool(_extract_part_codes(command))
    if has_part_code and (has_trigger or "search" in lower or "find" in lower or "get" in lower):
        return True
    return False


def _is_web_research_command(command: str) -> bool:
    lower = command.lower()
    if is_excel_command(command) or is_email_command(command):
        return False
    if _is_part_code_command(command):
        return True

    research_terms = (
        "fetch information", "find information", "collect information",
        "research", "search for", "search ", "look up", "lookup",
        "get me details", "get details", "find details",
        "product details", "open google", "google search",
    )
    if any(term in lower for term in research_terms):
        return True
    if lower.strip().endswith("search"):
        return True

    detail_terms = ("price", "density", "use", "uses", "application", "applications", "information")
    return ("google" in lower or "search" in lower) and any(t in lower for t in detail_terms)


def _parse_research_request(command: str) -> tuple[str, list[str], str]:
    """Returns (query, fields, output_file)."""
    lower = command.lower()

    # Part code commands ALWAYS save to part_code_results.txt
    if _is_part_code_command(command):
        output_file = "part_code_results.txt"
        fields = ["manufacturer", "description", "specifications", "datasheet"]
    else:
        # Check for explicit file target in command
        file_match = re.search(r"\b(?:in|into|to)\s+([a-zA-Z0-9_.-]+\.txt)\b", command, re.IGNORECASE)
        output_file = file_match.group(1) if file_match else "response.txt"
        fields = []
        if "price" in lower or "cost" in lower:
            fields.append("price")
        if "density" in lower:
            fields.append("density")
        if "use" in lower or "used for" in lower or "application" in lower:
            fields.append("use")
        if not fields:
            fields = ["price", "density", "use"]

    # Extract query subject
    query = command
    for pattern in (
        r"\brelated to\s+(.+)", r"\babout\s+(.+)", r"\bon\s+(.+)",
        r"\bregarding\s+(.+)", r"\bfor\s+(.+)", r"\bof\s+(.+)",
    ):
        match = re.search(pattern, command, re.IGNORECASE)
        if match:
            query = match.group(1)
            break

    query = re.sub(r"\s+and\s+(?:save|write|paste|put)\b.*$", "", query, flags=re.IGNORECASE).strip(" .")
    query = re.sub(r"\s+in(?:to)?\s+[a-zA-Z0-9_.-]+\.txt$", "", query, flags=re.IGNORECASE).strip(" .")

    part_codes = _extract_part_codes(command)
    if part_codes and _is_part_code_command(command):
        query = part_codes[0]

    if not query:
        query = command.strip()

    return query, fields, output_file


def _split_research_and_email_command(command: str) -> tuple[str | None, str | None]:
    if not command or not is_email_command(command):
        return None, None

    email_match = re.search(r"\b(?:mail|email|gmail|compose)\b", command, re.IGNORECASE)
    if not email_match:
        email_match = re.search(r"\bsend\s+(?:an\s+)?email\b", command, re.IGNORECASE)
    if not email_match:
        return None, None

    research_fragment = command[:email_match.start()].strip()
    email_fragment = command[email_match.start():].strip()
    if not research_fragment or not email_fragment:
        return None, None

    research_fragment = re.sub(
        r"(?:,|\s)+(?:after\s+that|afterwards|then|and\s+then|and)\s*$",
        "", research_fragment, flags=re.IGNORECASE,
    ).strip(" ,.")

    if research_fragment.lower() in {
        "open google", "open browser", "open chrome",
        "open edge", "open gmail",
    }:
        return None, None

    if _is_part_code_command(research_fragment) or _is_web_research_command(research_fragment):
        return research_fragment, email_fragment
    return None, None


def _build_research_email_body(report_file: str, max_chars: int = 2500) -> str:
    try:
        with open(report_file, "r", encoding="utf-8") as file_obj:
            content = file_obj.read().strip()
    except OSError:
        return ""

    if len(content) <= max_chars:
        return content

    truncated = content[:max_chars].rstrip()
    return f"{truncated}\n\n[Results truncated to fit email body.]"


def _build_email_intent_for_model(command: str, details: dict) -> str:
    body = str(details.get("body", "") or "").strip()
    subject = str(details.get("subject", "") or "").strip()
    raw_text = str(details.get("raw_text", "") or "").strip()

    noisy_tokens = ("mail", "email", "subject", "body", "by yourself", "@")
    body_lower = body.lower()
    noisy_body = (
        not body
        or any(token in body_lower for token in noisy_tokens)
        or "and main" in body_lower
        or body_lower.startswith("and ")
        or len(body.split()) < 5
    )
    if body and not noisy_body:
        return body
    if subject and subject.lower() != "quick update":
        return f"Write an engaging professional email about: {subject}"
    return raw_text or command


def _run_research_flow(command: str, default_output_file: str | None = None) -> str:
    """Execute web/part research flow and write findings report."""
    query, fields, output_file = _parse_research_request(command)

    # For non-part-code research, allow caller to override the output file
    if default_output_file and not _is_part_code_command(command) and output_file == "response.txt":
        output_file = default_output_file

    part_code_request = _is_part_code_command(command)
    part_codes = _extract_part_codes(command) if part_code_request else []

    if part_code_request:
        print(f"[AGENT] Part code search detected: {part_codes} -> saving to {output_file}")
        findings_by_code: dict[str, dict[str, list[str]]] = {}
        search_codes = part_codes or ([query] if query else [])
        for idx, code in enumerate(search_codes, 1):
            print(f"[AGENT] Searching part code {idx}/{len(search_codes)}: {code}")
            actions = _build_part_code_search_actions(code)
            _execute_actions_until_failure(actions)

            clipboard_text = get_clipboard_text()
            if not _is_relevant_part_clipboard(clipboard_text, [code]):
                print(f"[AGENT] Clipboard text does not match requested part code ({code}); skipping noisy extraction.")
                clipboard_text = ""
            findings_by_code[code] = _extract_findings_from_clipboard(clipboard_text, fields)

        _save_part_code_result(search_codes, findings_by_code, output_file)
    else:
        print("[AGENT] Web research command detected...")
        google_query = quote_plus(f"{query} {' '.join(fields)}")
        search_url = f"https://google.com/search?q={google_query}"
        actions = [
            {"action": "open_app", "app": "chrome", "url": search_url},
            {"action": "wait", "seconds": 4},
            {"action": "press_key", "keys": ["ctrl", "a"]},
            {"action": "press_key", "keys": ["ctrl", "c"]},
        ]
        _execute_actions_until_failure(actions)
        clipboard_text = get_clipboard_text()
        findings = _extract_findings_from_clipboard(clipboard_text, fields)
        write_research_report(
            command=command,
            query=query,
            fields=fields,
            findings=findings,
            sources=[],
            output_file=output_file,
            error=None if findings else "Could not extract details from page text.",
        )
    return output_file


def _execute_actions_until_failure(actions: list[dict]) -> list[dict]:
    results = []
    for i, action in enumerate(actions, 1):
        success, note = execute_action(action)
        results.append({
            "step": i,
            "action": action.get("action", "unknown"),
            "params": action,
            "success": success,
            "note": note,
        })
        if not success:
            break
    return results


def _extract_findings_from_clipboard(text: str, fields: list[str]) -> dict[str, list[str]]:
    snippets = re.split(r"(?<=[.!?])\s+|\n+", text or "")
    findings = {field: [] for field in fields}

    def _add(field: str, value: str, limit: int) -> None:
        if field not in findings:
            return
        cleaned = re.sub(r"\s+", " ", (value or "").strip())
        if not cleaned or cleaned in findings[field] or len(findings[field]) >= limit:
            return
        findings[field].append(cleaned)

    maker_re = re.compile(r"\b(?:" + "|".join(re.escape(m) for m in sorted(_KNOWN_COMPONENT_MAKERS, key=len, reverse=True)) + r")\b", re.IGNORECASE)
    spec_re = re.compile(r"\b\d+(?:\.\d+)?\s*(?:v|kv|mv|ma|a|ua|pf|nf|uf|ohm|kohm|mohm|w|mw|mhz|ghz|%|ppm|c|°c|tol(?:erance)?)\b", re.IGNORECASE)
    datasheet_url_re = re.compile(r"https?://\S*(?:datasheet|\.pdf)\S*", re.IGNORECASE)
    component_desc_re = re.compile(r"\b(?:capacitor|resistor|inductor|transistor|diode|mosfet|sensor|module|chip|ic)\b", re.IGNORECASE)

    for snippet in snippets:
        s = snippet.strip()
        if len(s) < 20 or len(s) > 320:
            continue
        lower = s.lower()

        if "price" in fields and ("price" in lower or re.search(r"(\$|usd|eur|inr|₹|€)\s?\d+", s, re.IGNORECASE)):
            _add("price", s, 5)
        if "density" in fields and re.search(r"\bdensity\b", lower):
            _add("density", s, 5)
        if "use" in fields and re.search(r"\b(use|used\s+for|application|applications|suitable\s+for)\b", lower):
            _add("use", s, 5)
        if "manufacturer" in fields:
            explicit = re.search(r"\bmanufacturer\b\s*[:\-]?\s*([A-Za-z0-9&., +\-/]{2,80})", s, re.IGNORECASE)
            if explicit:
                _add("manufacturer", explicit.group(1).strip(), 3)
            elif maker_re.search(s):
                _add("manufacturer", maker_re.search(s).group(0), 3)
        if "description" in fields and component_desc_re.search(s):
            _add("description", s, 3)
        if "specifications" in fields and (spec_re.search(s) or re.search(r"\b(tolerance|temperature|voltage|rating|size|package|dielectric)\b", lower)):
            _add("specifications", s, 5)
        if "datasheet" in fields:
            ds_match = datasheet_url_re.search(s)
            if ds_match:
                _add("datasheet", ds_match.group(0), 3)
            elif re.search(r"\bdatasheet\b", lower) and re.search(r"https?://", s, re.IGNORECASE):
                _add("datasheet", s, 3)

    return {k: v for k, v in findings.items() if v}


def _is_relevant_part_clipboard(text: str, part_codes: list[str]) -> bool:
    haystack = (text or "").upper()
    if any(code in haystack for code in (part_codes or [])):
        return True
    return bool(re.search(r"\b(datasheet|mouser|digikey|octopart|alldatasheet|manufacturer)\b", text or "", re.IGNORECASE))


def _build_part_code_search_actions(query: str) -> list[dict]:
    google_query = quote_plus(f"{query} datasheet specifications manufacturer")
    search_url = f"https://google.com/search?q={google_query}"
    return [
        {"action": "open_app", "app": "chrome", "url": search_url},
        {"action": "wait", "seconds": 4},
        {"action": "screenshot", "label": "search_results"},
        {"action": "press_key", "keys": ["ctrl", "a"]},
        {"action": "press_key", "keys": ["ctrl", "c"]},
        {"action": "wait", "seconds": 1},
    ]


def _save_part_code_result(part_codes: list[str], findings: dict, output_file: str) -> None:
    """Write structured part code JSON result to part_code_results.txt."""
    results = []
    for code in part_codes:
        code_findings = findings.get(code, {}) if isinstance(findings, dict) else {}

        raw_manufacturer = (code_findings.get("manufacturer") or ["Unknown"])[0]
        manufacturer = re.sub(r"(?i)^manufacturer\s*[:\-]?\s*", "", raw_manufacturer).strip() or "Unknown"
        description_hits = code_findings.get("description") or []
        description = description_hits[0] if description_hits else ""
        datasheet_hits = code_findings.get("datasheet") or []
        datasheet = datasheet_hits[0] if datasheet_hits else "Not found"

        result = {
            "part_code": code,
            "manufacturer": manufacturer,
            "description": description,
            "specs": {"raw_specs": code_findings.get("specifications") or []},
            "datasheet": datasheet,
        }
        results.append(result)

    with open(output_file, "w", encoding="utf-8") as f:
        f.write("=" * 60 + "\n")
        f.write("  PART CODE SEARCH RESULTS\n")
        f.write("=" * 60 + "\n\n")
        for r in results:
            f.write(f"Part Code    : {r['part_code']}\n")
            f.write(f"Manufacturer : {r['manufacturer']}\n")
            f.write(f"Description  : {r['description'] or 'N/A'}\n")
            f.write(f"Datasheet    : {r['datasheet']}\n")
            if r["specs"]["raw_specs"]:
                f.write("Specifications:\n")
                for spec in r["specs"]["raw_specs"]:
                    f.write(f"  - {spec}\n")
            f.write("\n")
        f.write("\nRAW JSON:\n")
        f.write(json.dumps(results, indent=2, ensure_ascii=False))
        f.write("\n")

    print(f"[AGENT] Part code results saved to {output_file}")


def _usage_delta(before: dict[str, int], after: dict[str, int]) -> dict[str, int]:
    keys = ("requests", "prompt_tokens", "completion_tokens", "total_tokens")
    return {k: max(0, int(after.get(k, 0)) - int(before.get(k, 0))) for k in keys}


# ── Initialization ────────────────────────────────────────────

def initialize():
    os.makedirs("screenshots", exist_ok=True)
    if not os.path.exists("commands.txt"):
        with open("commands.txt", "w", encoding="utf-8") as f:
            f.write("")
        print("[AGENT] Created commands.txt — open this in Notepad!")
    if not os.path.exists("response.txt"):
        with open("response.txt", "w", encoding="utf-8") as f:
            f.write("Agent is ready. Type a command in commands.txt and save.\n")
    print("[AGENT] Initialization complete.")


# ── Agentic Loop ──────────────────────────────────────────────

def _run_agentic_loop(command: str) -> tuple[list[dict], list[dict]]:
    """
    Core agentic loop:
    - Sends command + current screenshot + full action_history to LLM
    - Executes returned actions one by one
    - After each action, updates history and re-queries LLM
    - Loop stops when: LLM returns status=done, status=failed,
      no more actions, MAX_LOOP_ITERATIONS reached,
      or same screenshot label repeats 2+ times (loop detection)

    Returns (all_planned_actions, all_results)
    """
    action_history: list[dict] = []
    all_results: list[dict] = []
    all_planned_actions: list[dict] = []
    consecutive_verify_failures = 0
    MAX_CONSECUTIVE_FAILURES = 3
    step_counter = 0

    print(f"[AGENT] Starting agentic loop (max {MAX_LOOP_ITERATIONS} iterations)")

    for iteration in range(MAX_LOOP_ITERATIONS):
        print(f"\n[AGENT] -- Loop iteration {iteration + 1}/{MAX_LOOP_ITERATIONS} --")
        print(f"[AGENT] History so far: {len(action_history)} steps")

        # Ask LLM what to do next given current screen + history
        try:
            next_actions = get_action_plan(command, action_history=action_history)
        except Exception as llm_err:
            err_str = str(llm_err)
            if "401" in err_str or "Unauthorized" in err_str:
                print("[AGENT] LLM API key rejected (401). Check OPENROUTER_API_KEY in .env")
                print("[AGENT] Get a valid key at https://openrouter.ai/keys")
            raise

        if not next_actions:
            print("[AGENT] LLM returned no more actions — task complete.")
            break

        print(f"[AGENT] LLM returned {len(next_actions)} action(s) for this iteration")
        all_planned_actions.extend(next_actions)

        # Execute each action in the returned batch
        for action in next_actions:
            step_counter += 1
            act_type = action.get("action", "unknown")
            print(f"\n[AGENT] Step {step_counter}: {act_type} | {action}")

            success, note = execute_action(action)

            result = {
                "step": step_counter,
                "action": act_type,
                "params": action,
                "success": success,
                "note": note,
            }

            # ── Screenshot: verify + loop detection ──────────
            if act_type == "screenshot" and success:
                label = action.get("label", "step")

                # Loop detection: same label 2+ times = stuck
                prev_labels = [
                    r.get("params", {}).get("label")
                    for r in all_results
                    if r.get("action") == "screenshot"
                ]
                if prev_labels.count(label) >= 2:
                    print(f"[AGENT] Loop detected — label '{label}' repeated 3x. Stopping.")
                    result["note"] = note + " | LOOP DETECTED — stopped"
                    result["verified"] = False
                    all_results.append(result)
                    action_history.append(result)
                    return all_planned_actions, all_results

                # Verify the step visually
                print(f"[AGENT] Verifying step with model vision...")
                verified = verify_step(command, label)
                result["verified"] = verified

                if verified is True:
                    result["note"] = note + " | Verified ✓"
                    consecutive_verify_failures = 0
                elif verified is False:
                    result["note"] = note + " | Verify failed ✗"
                    consecutive_verify_failures += 1
                    print(f"[AGENT] Consecutive verify failures: {consecutive_verify_failures}/{MAX_CONSECUTIVE_FAILURES}")
                    if consecutive_verify_failures >= MAX_CONSECUTIVE_FAILURES:
                        print("[AGENT] Too many consecutive verify failures — aborting loop.")
                        all_results.append(result)
                        action_history.append(result)
                        return all_planned_actions, all_results
                else:
                    result["note"] = note + " | Verification skipped"

            all_results.append(result)
            action_history.append(result)

            if not success:
                print(f"[AGENT] Step failed — retrying once...")
                # Single retry for transient failures
                if act_type in {"press_key", "type_text", "paste_text", "click"}:
                    time.sleep(1)
                    success2, note2 = execute_action(action)
                    if success2:
                        all_results[-1]["success"] = True
                        all_results[-1]["note"] = note2 + " (retried)"
                        action_history[-1] = all_results[-1]
                        print("[AGENT] Retry succeeded.")
                        continue
                # Stop this iteration batch on failure
                print(f"[AGENT] Step failed permanently — sending failure to LLM for recovery")
                break  # LLM will see the failure in next iteration and recover

    else:
        print(f"[AGENT] Reached MAX_LOOP_ITERATIONS ({MAX_LOOP_ITERATIONS}) — stopping.")

    return all_planned_actions, all_results


# ── Core command handler ──────────────────────────────────────

def on_command(command: str):
    command_id = uuid.uuid4().hex[:10]
    set_command_context(command_id, command)
    set_log_context(command_id)
    log_command_event("command_start", {})

    print("\n" + "-" * 50)
    print(f"[AGENT][cmd:{command_id}] New command: '{command}'")
    print("-" * 50)
    logging.info(f"[cmd:{command_id}] Command received: {command}")

    write_thinking(command)

    try:
        route = "unknown"
        usage_before = get_usage_totals()
        command_for_routing = command
        research_email_body = ""

        research_fragment, email_fragment = _split_research_and_email_command(command)
        if research_fragment and email_fragment:
            print("[AGENT] Mixed command detected: research + email")
            details_file = _run_research_flow(
                research_fragment,
                default_output_file="part_code_results.txt",
            )
            print(f"[AGENT] Research details saved to: {details_file}")
            research_email_body = _build_research_email_body(details_file)
            command_for_routing = email_fragment

        excel_mode = _get_excel_mode()
        excel_request = is_excel_command(command_for_routing)
        email_request = is_email_command(command_for_routing)
        research_request = _is_web_research_command(command_for_routing)

        # ── RESEARCH / PART CODE PATH ─────────────────────────
        if research_request:
            route = "research"
            log_command_event("command_route", {"route": route, "routed_command": command_for_routing})
            _run_research_flow(command_for_routing)

        # ── EMAIL PATH ────────────────────────────────────────
        elif email_request:
            route = "email"
            log_command_event("command_route", {"route": route, "routed_command": command_for_routing})
            print("[AGENT] Email command detected...")
            details = parse_email_request(command_for_routing)
            if research_email_body and not (details.get("body") or "").strip():
                details["body"] = research_email_body
                if (details.get("subject") or "").strip().lower() == "quick update":
                    details["subject"] = "Part code search results"

            if wants_model_polish(command_for_routing):
                raw_intent = _build_email_intent_for_model(command_for_routing, details)
                recipient_email = (details.get("to") or [None])[0]
                try:
                    polished = generate_professional_email(raw_intent, recipient_email=recipient_email)
                    details["subject"] = polished.get("subject", details.get("subject", "Quick update"))
                    details["body"] = polished.get("body", details.get("body", ""))
                    print("[AGENT] Email polished by model.")
                except Exception as polish_error:
                    raise RuntimeError(
                        "Email drafting requested but model draft failed. "
                        f"Fix LLM credentials and retry. Root error: {polish_error}"
                    ) from polish_error
            actions = build_email_actions(details)
            valid, note = validate_email_actions(actions)
            if not valid:
                raise Exception(f"Email action plan invalid: {note}")

            # Email uses simple linear execution (not agentic loop)
            results = _execute_actions_until_failure(actions)
            usage_after = get_usage_totals()
            usage_delta = _usage_delta(usage_before, usage_after)
            write_response(command, actions, results, llm_usage=usage_delta)

        # ── EXCEL PATH ────────────────────────────────────────
        elif excel_request:
            route = "excel"
            log_command_event("command_route", {"route": route, "routed_command": command_for_routing})
            print(f"[AGENT] Excel command (mode={excel_mode})...")
            actions = parse_excel_actions_from_command(command_for_routing)
            valid, note = validate_excel_actions(actions)

            if valid:
                print(f"[AGENT] Using {len(actions)} structured Excel actions.")
            elif excel_mode in {"structured", "hybrid"}:
                print(f"[AGENT] Parser insufficient ({note}) — asking LLM...")
                actions = get_excel_action_plan(command_for_routing)
                valid, note = validate_excel_actions(actions)
                if not valid:
                    if excel_mode == "hybrid":
                        print("[AGENT] Structured Excel plan invalid — falling back to generic...")
                        actions = get_action_plan(command_for_routing)
                        unknown = _unsupported_actions(actions)
                        if unknown:
                            raise Exception(f"Excel plan invalid and generic fallback has unsupported actions: {unknown}")
                    else:
                        raise Exception(f"Excel action plan invalid: {note}")
            else:
                actions = get_action_plan(command_for_routing)
                unknown = _unsupported_actions(actions)
                if unknown:
                    raise Exception(f"Generic Excel plan has unsupported actions: {unknown}")

            results = _execute_actions_until_failure(actions)
            usage_after = get_usage_totals()
            usage_delta = _usage_delta(usage_before, usage_after)
            write_response(command, actions, results, llm_usage=usage_delta)

        # ── GENERIC PATH: TRUE AGENTIC LOOP ───────────────────
        else:
            route = "generic"
            log_command_event("command_route", {"route": route, "routed_command": command_for_routing})
            print("[AGENT] Generic command — starting agentic loop with memory...")
            all_planned_actions, all_results = _run_agentic_loop(command_for_routing)

            usage_after = get_usage_totals()
            usage_delta = _usage_delta(usage_before, usage_after)
            write_response(command, all_planned_actions, all_results, llm_usage=usage_delta)

        log_command_event(
            "command_completed",
            {
                "route": route,
                "status": "success",
            },
        )

    except Exception as e:
        error_msg = str(e)
        print(f"[AGENT] Error: {error_msg}")
        logging.error(f"[cmd:{command_id}] Agent error for '{command}': {error_msg}", exc_info=True)
        log_command_event(
            "command_completed",
            {
                "status": "failed",
                "error": error_msg,
            },
        )
        write_error(command, error_msg)
    finally:
        set_log_context(None)
        clear_command_context()


# ── process_command (for main.py compatibility) ───────────────

def process_command(command: str) -> str:
    on_command(command)
    try:
        with open("response.txt", "r", encoding="utf-8") as f:
            return f.read()
    except Exception:
        return "Done."


# ── Entry point ───────────────────────────────────────────────

if __name__ == "__main__":
    print("\n" + "=" * 50)
    print("   DESKTOP AI AGENT  v3.0")
    print("   Agentic Loop with Memory")
    print("=" * 50 + "\n")

    initialize()

    print("HOW TO USE:")
    print("  1. Open commands.txt in Notepad")
    print("  2. Type a command and save (Ctrl+S)")
    print("  3. Watch the agent work")
    print("  4. Open response.txt for results")
    print("  5. Part code results -> part_code_results.txt")
    print("\nSTOP: Ctrl+C or move mouse to top-left corner\n")

    observer = start_watching(on_command)

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n[AGENT] Shutting down...")
        observer.stop()
        observer.join()
        print("[AGENT] Goodbye!")
