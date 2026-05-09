# ============================================================
#  llm.py — The Brain: Talks to NVIDIA / OpenRouter API
#
#  FIXES IN THIS VERSION:
#    1. get_action_plan() always sends current screenshot +
#       command + full action_history (true memory per call)
#    2. 401 errors → clear fix instructions printed + logged
#    3. LLM log records every call with full context
#    4. verify_step() logs success/failure to llm_steps.log
#    5. Loop-aware: returns status=done to stop agentic loop
# ============================================================

import os
import json
import re
import base64
import time
import requests
from datetime import datetime
from pathlib import Path
from PIL import ImageGrab, Image
from dotenv import load_dotenv
from excel_automation import validate_excel_actions

load_dotenv()

NVIDIA_INVOKE_URL = "https://integrate.api.nvidia.com/v1/chat/completions"
OPENROUTER_INVOKE_URL = "https://openrouter.ai/api/v1/chat/completions"

NVIDIA_API_KEY = ""
_openai_like_key = ""
OPENROUTER_API_KEY = ""
LLM_PROVIDER = "auto"


def _resolve_provider() -> tuple[str, str, str]:
    if LLM_PROVIDER == "openrouter":
        if not OPENROUTER_API_KEY:
            raise RuntimeError("LLM_PROVIDER=openrouter but OPENROUTER_API_KEY is missing.")
        return "openrouter", OPENROUTER_INVOKE_URL, OPENROUTER_API_KEY
    if LLM_PROVIDER == "nvidia":
        if not NVIDIA_API_KEY:
            raise RuntimeError("LLM_PROVIDER=nvidia but NVIDIA_API_KEY is missing.")
        return "nvidia", NVIDIA_INVOKE_URL, NVIDIA_API_KEY
    if OPENROUTER_API_KEY:
        return "openrouter", OPENROUTER_INVOKE_URL, OPENROUTER_API_KEY
    if NVIDIA_API_KEY:
        return "nvidia", NVIDIA_INVOKE_URL, NVIDIA_API_KEY
    return "none", "", ""


PROVIDER = "none"
INVOKE_URL = ""
ACTIVE_API_KEY = ""
MODEL = ""
_PROJECT_ROOT = Path(__file__).resolve().parent
_configured_log_file = os.getenv("LLM_LOG_FILE", "llm_steps.log").strip() or "llm_steps.log"
LLM_LOG_FILE = str(
    Path(_configured_log_file) if Path(_configured_log_file).is_absolute()
    else (_PROJECT_ROOT / _configured_log_file)
)

HEADERS = {}

USAGE_TOTALS = {
    "requests": 0,
    "prompt_tokens": 0,
    "completion_tokens": 0,
    "total_tokens": 0,
}

_COMMAND_CONTEXT: dict[str, str] = {}


def _refresh_provider_config() -> None:
    """Reload .env and refresh provider/model/headers without restarting the agent."""
    global NVIDIA_API_KEY, _openai_like_key, OPENROUTER_API_KEY, LLM_PROVIDER
    global PROVIDER, INVOKE_URL, ACTIVE_API_KEY, MODEL, HEADERS

    load_dotenv(override=True)

    NVIDIA_API_KEY = os.getenv("NVIDIA_API_KEY", "").strip()
    _openai_like_key = os.getenv("OPENAI_API_KEY", "").strip()
    OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY", "").strip()
    if not OPENROUTER_API_KEY and _openai_like_key.startswith("sk-or-"):
        OPENROUTER_API_KEY = _openai_like_key

    LLM_PROVIDER = os.getenv("LLM_PROVIDER", "auto").strip().lower()
    PROVIDER, INVOKE_URL, ACTIVE_API_KEY = _resolve_provider()

    default_model = "openai/gpt-4o-mini" if PROVIDER in {"openrouter", "none"} else "microsoft/phi-3.5-vision-instruct"
    MODEL = os.getenv("AGENT_MODEL", default_model).strip() or default_model

    HEADERS = {}
    if PROVIDER != "none":
        HEADERS = {
            "Authorization": f"Bearer {ACTIVE_API_KEY}",
            "Accept": "application/json",
            "Content-Type": "application/json",
        }
        if PROVIDER == "openrouter":
            HEADERS["HTTP-Referer"] = os.getenv("OPENROUTER_HTTP_REFERER", "http://localhost")
            HEADERS["X-Title"] = os.getenv("OPENROUTER_APP_TITLE", "AiAgent")


_refresh_provider_config()


def _truncate_text(text: str, limit: int = 2000) -> str:
    value = str(text or "")
    return value if len(value) <= limit else value[:limit] + "...<truncated>"


def set_command_context(command_id: str, command: str) -> None:
    _COMMAND_CONTEXT["command_id"] = str(command_id or "").strip()
    _COMMAND_CONTEXT["command"] = str(command or "").strip()


def clear_command_context() -> None:
    _COMMAND_CONTEXT.clear()


def log_command_event(event: str, payload: dict | None = None) -> None:
    _append_llm_log(event, payload or {})


def _append_llm_log(event: str, payload: dict) -> None:
    """Write a structured JSON line to llm_steps.log."""
    try:
        Path(LLM_LOG_FILE).parent.mkdir(parents=True, exist_ok=True)
        payload = payload or {}
        context = {}
        command_id = _COMMAND_CONTEXT.get("command_id")
        command_text = _COMMAND_CONTEXT.get("command")
        if command_id and "command_id" not in payload:
            context["command_id"] = command_id
        if command_text and "command" not in payload:
            context["command"] = command_text
        record = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "event": event,
            "provider": PROVIDER,
            "model": MODEL,
            **context,
            **payload,
        }
        with open(LLM_LOG_FILE, "a", encoding="utf-8") as f:
            f.write(json.dumps(record, ensure_ascii=False) + "\n")
    except Exception as exc:
        print(f"[LLM] Could not write LLM log: {exc}")


def _record_usage(data: dict) -> None:
    usage = data.get("usage", {}) if isinstance(data, dict) else {}
    prompt = int(usage.get("prompt_tokens") or usage.get("input_tokens") or 0)
    completion = int(usage.get("completion_tokens") or usage.get("output_tokens") or 0)
    total = int(usage.get("total_tokens") or (prompt + completion))
    USAGE_TOTALS["requests"] += 1
    USAGE_TOTALS["prompt_tokens"] += prompt
    USAGE_TOTALS["completion_tokens"] += completion
    USAGE_TOTALS["total_tokens"] += total


def get_usage_totals() -> dict[str, int]:
    return dict(USAGE_TOTALS)


def _ensure_provider_ready() -> None:
    _refresh_provider_config()
    if PROVIDER == "none":
        raise RuntimeError(
            "No LLM provider configured. Set OPENROUTER_API_KEY (recommended) "
            "or NVIDIA_API_KEY in .env."
        )


# ── SCREENSHOT ────────────────────────────────────────────────

def get_screenshot_base64() -> str:
    """Capture screen, resize to half, save as screen.png, return base64."""
    img = ImageGrab.grab()
    w, h = img.size
    img = img.resize((w // 2, h // 2), Image.Resampling.LANCZOS)
    img.save("screen.png")
    with open("screen.png", "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


# ── SYSTEM PROMPT ─────────────────────────────────────────────

_SYSTEM_PROMPT = """\
You are a Desktop Automation AI Agent running on Windows.
Your job is NOT to chat. Your job is to CONTROL a computer safely and deterministically.

You receive:
1. USER_COMMAND (natural language goal)
2. CURRENT_SCREENSHOT (image of current screen state — taken RIGHT NOW)
3. ACTION_HISTORY (list of ALL past steps with their results)

You MUST return ONLY a valid JSON object. No explanation. No markdown. No extra text. Just raw JSON.

OUTPUT FORMAT (STRICT):
{
  "goal": "short description of the overall task",
  "plan": ["step1", "step2", "..."],
  "actions": [
    {
      "action": "open_app | open_file | open_url | click | double_click | right_click | click_text | type_text | paste_text | press_key | scroll | wait | screenshot | drag | move_mouse | delete_rows",
      "app": "app name if open_app",
      "path": "file path if open_file",
      "url": "url if open_url",
      "x": number,
      "y": number,
      "text": "text to type or paste",
      "keys": ["ctrl", "s"],
      "seconds": number,
      "label": "screenshot label",
      "amount": number
    }
  ],
  "validation": {
    "expected_result": "what should change on screen after these actions",
    "method": "visual | text | url"
  },
  "status": "continue | done | failed",
  "error": "reason if failed, else null"
}

MEMORY RULES (critical):
- ACTION_HISTORY contains ALL steps already completed — read it before deciding what to do next
- NEVER repeat an action that succeeded in ACTION_HISTORY
- If an action FAILED in history, try a different approach — not the same action again
- Return status=done when the GOAL is fully achieved based on what you see on screen
- Return status=failed if stuck (same failure 2+ times) or if goal is impossible

LOOP PREVENTION:
- If ACTION_HISTORY shows the same action failing 2+ times → return status=failed with explanation
- If screen looks identical to a previous screenshot step → try a different action or return done

SAFETY GUARDRAILS — DO NOT:
- Delete files or format drives without explicit confirmation
- Run unknown executables
- Expose credentials or API keys
- Perform irreversible destructive actions without user confirmation

BROWSER NAVIGATION pattern (always follow when navigating):
1. open_app chrome/brave/edge
2. wait 2 seconds
3. press_key ctrl+l (focus address bar)
4. type_text the URL
5. press_key enter
6. wait 3 seconds for page to load
7. screenshot to verify

EMAIL SEND:
- After Gmail compose is open, use press_key with keys ["ctrl","enter"] to send
- NEVER use click_text "Send" — OCR is unreliable for this

PART CODE SEARCH SPECIALIZATION:
If command involves electronic part codes (e.g. 0603YG105ZAT2A):
1. Open browser to Google or Mouser/Digi-Key
2. Search the part code
3. Screenshot to verify results loaded
4. Return status=done (agent will extract data from clipboard)
"""


# ── MAIN FUNCTION: Get Action Plan ────────────────────────────

def get_action_plan(command: str, action_history: list | None = None) -> list:
    """
    Calls LLM with:
    - Current screenshot (taken right now)
    - The original command
    - Full action_history (all steps done so far + their results)

    Returns a flat list of action dicts for the executor.
    The agentic loop in agent.py calls this repeatedly, growing the history.
    """
    _ensure_provider_ready()
    history_count = len(action_history or [])
    print(f"[LLM] Calling {PROVIDER}:{MODEL} | history={history_count} steps | taking screenshot...")

    # Build history text for the prompt
    history_text = ""
    if action_history:
        history_text = "\n\nACTION_HISTORY (steps already completed — DO NOT repeat successes):\n"
        for i, step in enumerate(action_history, 1):
            success_str = "SUCCESS" if step.get("success") else "FAILED"
            verified_str = ""
            if "verified" in step:
                verified_str = f" | verified={step['verified']}"
            history_text += f"  Step {i} [{success_str}{verified_str}]: {step.get('action')} — {step.get('note', '')}\n"

    user_message = f"USER_COMMAND: {command}{history_text}"

    max_retries = 3

    for attempt in range(max_retries):
        try:
            # Take fresh screenshot for this call
            screenshot_b64 = get_screenshot_base64()

            payload = {
                "model": MODEL,
                "messages": [
                    {"role": "system", "content": _SYSTEM_PROMPT},
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": user_message},
                            {
                                "type": "image_url",
                                "image_url": {"url": f"data:image/png;base64,{screenshot_b64}"},
                            },
                        ],
                    },
                ],
                "max_tokens": 3000,
                "temperature": 0.10,
                "top_p": 0.70,
                "stream": False,
            }

            response = requests.post(INVOKE_URL, headers=HEADERS, json=payload, timeout=60)
            response.raise_for_status()

            data = response.json()
            _record_usage(data)

            raw = data["choices"][0]["message"]["content"].strip()
            raw_for_log = raw
            print(f"[LLM] Raw response (first 300 chars): {raw[:300]}...")

            # Strip markdown fences
            raw = re.sub(r"```json|```", "", raw).strip()
            raw = raw.replace("\\\\\\\\", "\\\\")

            # Parse with truncation recovery
            try:
                parsed = json.loads(raw)
            except json.JSONDecodeError:
                print("[LLM] JSON truncated — attempting recovery...")
                last_close = raw.rfind("}]")
                if last_close != -1:
                    raw_try = raw[: last_close + 2] + "}"
                    try:
                        parsed = json.loads(raw_try)
                    except Exception:
                        raw = raw[: last_close + 2]
                        parsed = None
                else:
                    last_brace = raw.rfind("}")
                    if last_brace != -1:
                        raw = raw[: last_brace + 1]
                    parsed = None

                if parsed is None:
                    try:
                        parsed = json.loads(raw)
                    except json.JSONDecodeError:
                        raise Exception("Could not recover JSON from truncated response.")

            # Handle both formats
            if isinstance(parsed, list):
                actions = parsed
                print(f"[LLM] Got {len(actions)} actions (legacy list format)")
                _append_llm_log("action_plan", {
                    "command": command,
                    "attempt": attempt + 1,
                    "history_count": history_count,
                    "response_format": "legacy_list",
                    "actions_count": len(actions),
                    "actions": actions,
                    "raw_preview": _truncate_text(raw_for_log),
                })
                return actions

            if isinstance(parsed, dict):
                status = parsed.get("status", "continue")
                error = parsed.get("error")
                actions = parsed.get("actions", [])
                goal = parsed.get("goal", "")
                plan = parsed.get("plan", [])

                print(f"[LLM] Goal: {goal}")
                print(f"[LLM] Plan: {plan}")
                print(f"[LLM] Status: {status} | Actions: {len(actions)}")

                _append_llm_log("action_plan", {
                    "command": command,
                    "attempt": attempt + 1,
                    "history_count": history_count,
                    "response_format": "structured",
                    "status": status,
                    "error": error,
                    "goal": goal,
                    "plan": plan,
                    "actions_count": len(actions),
                    "actions": actions,
                    "raw_preview": _truncate_text(raw_for_log),
                })

                if status == "failed":
                    reason = error or "Model returned status=failed"
                    raise Exception(f"Agent safety stop: {reason}")

                if status == "done" and not actions:
                    print("[LLM] Model says task is already done.")
                    return []

                if not actions:
                    raise Exception("Model returned no actions in structured response.")

                return actions

            raise Exception(f"Unexpected model response type: {type(parsed)}")

        except requests.exceptions.HTTPError as e:
            status_code = e.response.status_code if e.response else 0
            _append_llm_log("action_plan_error", {
                "command": command,
                "attempt": attempt + 1,
                "history_count": history_count,
                "status_code": status_code,
                "error": str(e),
            })

            if status_code == 401:
                key_name = "OPENROUTER_API_KEY" if PROVIDER == "openrouter" else "NVIDIA_API_KEY"
                msg = (
                    f"[LLM] API key rejected (401 Unauthorized).\n"
                    f"  Fix: Update {key_name} in your .env file.\n"
                    f"  OpenRouter keys: https://openrouter.ai/keys\n"
                    f"  NVIDIA keys: https://build.nvidia.com/"
                )
                print(msg)
                raise RuntimeError(f"401 Unauthorized — invalid {key_name}. Check .env file.") from e

            if status_code == 429:
                if attempt < max_retries - 1:
                    wait_seconds = 30 * (attempt + 1)
                    print(f"[LLM] Rate limit. Waiting {wait_seconds}s before retry {attempt + 2}/{max_retries}...")
                    time.sleep(wait_seconds)
                else:
                    raise
            elif status_code == 410:
                raise RuntimeError("Model endpoint returned HTTP 410 Gone. Switch provider or model in .env.") from e
            elif status_code == 404:
                raise RuntimeError(f"Model '{MODEL}' not found on provider '{PROVIDER}'.") from e
            elif status_code == 400:
                raise RuntimeError("Provider rejected request. Verify AGENT_MODEL supports this input format.") from e
            else:
                raise

        except Exception as e:
            error_str = str(e)
            _append_llm_log("action_plan_error", {
                "command": command,
                "attempt": attempt + 1,
                "history_count": history_count,
                "error": error_str,
            })
            if "429" in error_str and attempt < max_retries - 1:
                wait_seconds = 30 * (attempt + 1)
                print(f"[LLM] Quota hit. Waiting {wait_seconds}s...")
                time.sleep(wait_seconds)
            else:
                raise

    raise Exception("Failed to generate action plan after all retries.")


# ── VERIFY STEP ───────────────────────────────────────────────

def verify_step(command: str, step_description: str) -> bool | None:
    """
    Take a fresh screenshot and ask LLM if the last action succeeded.
    Returns True (success), False (failure), or None (skipped/error).
    Logs result to llm_steps.log.
    """
    try:
        _ensure_provider_ready()
        print(f"[LLM] Verifying step: '{step_description}'")

        screenshot_b64 = get_screenshot_base64()

        prompt = f"""\
Look at this screenshot carefully.
Overall task: {command}
Last completed action: {step_description}

Did the action succeed based on what you see on screen?
Answer with ONLY one word: SUCCESS or FAILURE
"""

        payload = {
            "model": MODEL,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{screenshot_b64}"}},
                    ],
                }
            ],
            "max_tokens": 10,
            "temperature": 0.10,
            "stream": False,
        }

        response = requests.post(INVOKE_URL, headers=HEADERS, json=payload, timeout=60)
        response.raise_for_status()
        data = response.json()
        _record_usage(data)
        result = data["choices"][0]["message"]["content"].strip().upper()
        verified = "SUCCESS" in result

        print(f"[LLM] Verification result: {result} -> {'OK' if verified else 'FAIL'}")
        _append_llm_log("verify_step", {
            "command": command,
            "step_description": step_description,
            "result": result,
            "verified": verified,
        })
        return verified

    except Exception as exc:
        err_str = str(exc)
        print(f"[LLM] Verification skipped: {err_str}")
        if "401" in err_str or "Unauthorized" in err_str:
            key_name = "OPENROUTER_API_KEY" if PROVIDER == "openrouter" else "NVIDIA_API_KEY"
            print(f"[LLM] Fix: Update {key_name} in .env | OpenRouter keys: https://openrouter.ai/keys")
        _append_llm_log("verify_step_error", {
            "command": command,
            "step_description": step_description,
            "error": err_str,
        })
        return None


# ── EXCEL STRUCTURED PLAN ─────────────────────────────────────

def get_excel_action_plan(command: str) -> list:
    _ensure_provider_ready()
    print(f"[LLM] Asking {PROVIDER}:{MODEL} for structured Excel actions...")

    prompt = f"""\
You are generating Excel automation actions for a Windows desktop agent.
User command: {command}

Return ONLY a valid JSON array. No markdown. No explanation.
Each item must use one of these action names only:
open_workbook, select_sheet, aggregate, set_cell_value, set_range_value,
clear_contents, delete_rows, sort_sheet, filter_column, clear_filter, save_workbook

Examples:
[{{"action":"open_workbook","reference":"C:\\\\Users\\\\DELL\\\\Downloads\\\\Orders.csv"}}]
[{{"action":"open_workbook","reference":"C:\\\\Users\\\\DELL\\\\Downloads\\\\Orders.csv"}},{{"action":"delete_rows","start_row":2,"end_row":10,"workbook":"C:\\\\Users\\\\DELL\\\\Downloads\\\\Orders.csv"}}]

Rules:
- Keep action list short and specific.
- Do not use vague actions.
- Preserve exact workbook path if provided.
"""

    for attempt in range(2):
        payload = {
            "model": MODEL,
            "messages": [{"role": "user", "content": prompt}],
            "max_tokens": 1024,
            "temperature": 0.10,
            "top_p": 0.70,
            "stream": False,
        }

        response = requests.post(INVOKE_URL, headers=HEADERS, json=payload, timeout=60)
        response.raise_for_status()
        data = response.json()
        _record_usage(data)
        raw = data["choices"][0]["message"]["content"].strip()
        raw_for_log = raw
        raw = re.sub(r"```json|```", "", raw).strip()
        raw = raw.replace("\\\\\\\\", "\\\\")

        actions = json.loads(raw)
        valid, note = validate_excel_actions(actions)
        if valid:
            print(f"[LLM] Got {len(actions)} structured Excel actions")
            _append_llm_log("excel_action_plan", {
                "command": command,
                "attempt": attempt + 1,
                "actions_count": len(actions),
                "actions": actions,
                "raw_preview": _truncate_text(raw_for_log),
            })
            return actions

        _append_llm_log("excel_action_plan_error", {
            "command": command,
            "attempt": attempt + 1,
            "error": note,
            "raw_preview": _truncate_text(raw_for_log),
        })
        prompt += f"\nYour last answer was invalid because: {note}\nReturn a corrected JSON array only."

    raise Exception("Excel action plan invalid. Make the Excel request more specific.")


# ── EMAIL POLISH ──────────────────────────────────────────────

def generate_professional_email(intent_text: str, recipient_email: str | None = None) -> dict:
    _ensure_provider_ready()
    print("[LLM] Polishing email draft...")

    recipient_name = _recipient_name_from_email(recipient_email) if recipient_email else ""
    recipient_context = (
        f"Recipient email: {recipient_email}\nPreferred name: {recipient_name}\n"
        if recipient_email
        else "Recipient not provided. Use generic greeting.\n"
    )

    prompt = f"""\
Rewrite this rough note as a professional email.

User intent: {intent_text}
{recipient_context}

Return ONLY valid JSON: {{"subject":"...","body":"..."}}

Rules:
- Professional, polite, concise tone.
- Do not invent facts.
- Body must be plain text, no markdown.
- Never use placeholders like [Name] or {{name}}.
- Use recipient name naturally if known, else generic greeting.
"""

    payload = {
        "model": MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": 400,
        "temperature": 0.20,
        "top_p": 0.80,
        "stream": False,
    }

    try:
        response = requests.post(INVOKE_URL, headers=HEADERS, json=payload, timeout=60)
        response.raise_for_status()
        data = response.json()
        _record_usage(data)
        raw = data["choices"][0]["message"]["content"].strip()
        raw_for_log = raw
        raw = re.sub(r"```json|```", "", raw).strip()

        parsed = json.loads(raw)
        subject = str(parsed.get("subject", "Quick update")).strip() or "Quick update"
        body = str(parsed.get("body", "")).strip()
        body = _sanitize_placeholder_text(body, recipient_name)

        _append_llm_log("email_polish", {
            "intent_text": _truncate_text(intent_text),
            "recipient_email": recipient_email,
            "subject": subject,
            "body_preview": _truncate_text(body),
            "raw_preview": _truncate_text(raw_for_log),
        })
        return {"subject": subject, "body": body}

    except Exception as exc:
        err_str = str(exc)
        if "401" in err_str or "Unauthorized" in err_str:
            key_name = "OPENROUTER_API_KEY" if PROVIDER == "openrouter" else "NVIDIA_API_KEY"
            print(f"[LLM] Email polish failed (401). Fix: Update {key_name} in .env")
        _append_llm_log("email_polish_error", {
            "intent_text": _truncate_text(intent_text),
            "recipient_email": recipient_email,
            "error": err_str,
        })
        raise


def _recipient_name_from_email(email: str | None) -> str:
    if not email or "@" not in email:
        return ""
    local = email.split("@", 1)[0]
    local = re.sub(r"[._-]+", " ", local)
    local = re.sub(r"\s+", " ", local).strip()
    return " ".join(part.capitalize() for part in local.split(" ")) if local else ""


def _sanitize_placeholder_text(text: str, recipient_name: str) -> str:
    placeholder_pattern = re.compile(
        r"\[(recipient'?s?\s*name|name)\]|<\s*name\s*>|\{\{\s*name\s*\}\}|\[\s*name\s*\]",
        re.IGNORECASE,
    )
    replacement = recipient_name if recipient_name else "there"
    cleaned = placeholder_pattern.sub(replacement, text)
    return re.sub(r"\s+", " ", cleaned).strip()
