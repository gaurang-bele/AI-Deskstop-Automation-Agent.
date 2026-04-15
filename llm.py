# ============================================================
#  llm.py — The Brain: Talks to NVIDIA API
# ============================================================
#
#  FIXES IN THIS VERSION:
#    1. max_tokens increased 1024 → 2048 (fixes truncated JSON)
#    2. Truncation recovery added (salvages partial JSON)
#    3. type_text rule added (model was typing "Ctrl+A" as text)
#    4. All previous fixes kept (rules, backslash cleanup etc.)
# ============================================================


import os
import json
import re
import base64
import time
import requests
from PIL import ImageGrab, Image
from dotenv import load_dotenv
from excel_automation import validate_excel_actions

load_dotenv()

INVOKE_URL = "https://integrate.api.nvidia.com/v1/chat/completions"
NVIDIA_API_KEY = os.getenv("NVIDIA_API_KEY")

MODEL = "microsoft/phi-3.5-vision-instruct"
# If this hits rate limits, try:
#   "meta/llama-3.2-11b-vision-instruct"
#   "microsoft/phi-3-vision-128k-instruct"

HEADERS = {
    "Authorization": f"Bearer {NVIDIA_API_KEY}",
    "Accept": "application/json",
    "Content-Type": "application/json"
}


# ── FUNCTION 1: Take a Screenshot ────────────────────────────

def get_screenshot_base64() -> str:
    # Capture screen → resize to half → encode to base64
    # WHY RESIZE: cuts token usage ~75%, still readable by model

    img = ImageGrab.grab()

    w, h = img.size
    img = img.resize((w // 2, h // 2), Image.Resampling.LANCZOS)
    # Image.LANCZOS → best quality downscaling filter

    img.save("screen.png")
    # Save so you can open screen.png to see what model sees

    with open("screen.png", "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")
    # "rb"        → read binary mode (images are binary)
    # b64encode() → bytes → base64 bytes
    # .decode()   → base64 bytes → UTF-8 string


# ── FUNCTION 2: Get Action Plan from NVIDIA Model ────────────

def get_action_plan(command: str) -> list:
    print("[LLM] Capturing screen and asking NVIDIA model for action plan...")

    prompt = f"""
You are a desktop automation AI agent running on Windows.
The user wants you to do this task: {command}

I have attached a screenshot of the current screen state.
Look at it carefully to understand what is open and where things are.

Return ONLY a valid JSON array of actions to perform this task.
No explanation. No markdown. No extra text. Just the raw JSON array.

Available action types and their exact format:

1. Open an application:
   {{"action": "open_app", "app": "notepad"}}
   Valid app names: notepad, excel, word, chrome, edge, firefox, calculator, explorer, powershell, vscode
   IMPORTANT: If user says "edge" or "microsoft edge", use app="edge", NOT "chrome"
   IMPORTANT: If user says "chrome" or "google", use app="chrome"
   IMPORTANT: Always use the EXACT browser the user mentions!

2. Open a specific file by full path:
   {{"action": "open_file", "path": "C:\\Users\\dell\\Desktop\\myfile.csv"}}
   {{"action": "open_file", "path": "C:\\Users\\dell\\Desktop\\myfile.xlsx"}}
   {{"action": "open_file", "path": "C:\\Users\\dell\\Desktop\\myfile.pdf"}}
   IMPORTANT: Use SINGLE backslashes in the path. Never double backslashes.
   IMPORTANT: Use the exact file extension the user mentions (.csv, .xlsx, .pdf etc.)
   IMPORTANT: Preserve the EXACT filename the user provides! Do NOT modify spaces, parentheses, or special characters.
   Example: "extracted_table (4).csv" must stay as "extracted_table (4).csv", NOT "extracted_table_4.csv"
   NEVER assume .xlsx if the user says csv. NEVER assume .csv if the user says excel.
   If the user doesn't provide a full path, use C:\\Users\\DELL\\Downloads\\ as the default folder.

3. Left click at coordinates:
   {{"action": "click", "x": 500, "y": 300}}

4. Double click:
   {{"action": "double_click", "x": 500, "y": 300}}

5. Right click:
   {{"action": "right_click", "x": 500, "y": 300}}

6. Click a button or label by its visible text (preferred over coordinates):
   {{"action": "click_text", "text": "Save"}}
   {{"action": "click_text", "text": "Delete"}}
   {{"action": "click_text", "text": "OK"}}
   {{"action": "click_text", "text": "Yes"}}
   Use this whenever clicking a named button, menu item, or label.

7. Type actual text content (keyboard input):
   {{"action": "type_text", "text": "Hello World"}}
   WARNING: type_text physically types characters. NEVER use it for shortcuts.
   WRONG: {{"action": "type_text", "text": "Ctrl+A"}}
   WRONG: {{"action": "type_text", "text": "Ctrl+S"}}
   IMPORTANT: When user asks to type a long paragraph, copy the ENTIRE text exactly word for word into the "text" field. NEVER shorten, summarize or cut it. The full text must appear in the JSON.
```

8. Paste text from clipboard (preferred for long text, emails, URLs, and repeated content):
    {{"action": "paste_text", "text": "Long email body or message"}}
    Use this instead of type_text when the content is long or needs to be pasted into a field.

9. Open a URL directly:
    {{"action": "open_url", "url": "https://example.com"}}
    Use this for websites and internet navigation when no specific browser steps are needed.

10. Drag with the mouse:
    {{"action": "drag", "start_x": 100, "start_y": 200, "end_x": 300, "end_y": 400}}
    Use this for drag-and-drop, range selection, and slider movement.

11. Move the mouse pointer:
    {{"action": "move_mouse", "x": 500, "y": 300}}

12. Press keyboard shortcut (use this for ALL shortcuts):
   {{"action": "press_key", "keys": ["ctrl", "s"]}}
   {{"action": "press_key", "keys": ["ctrl", "a"]}}
   {{"action": "press_key", "keys": ["ctrl", "z"]}}
   {{"action": "press_key", "keys": ["delete"]}}
   {{"action": "press_key", "keys": ["enter"]}}
   {{"action": "press_key", "keys": ["shift", "space"]}}

13. Scroll the mouse wheel:
   {{"action": "scroll", "x": 500, "y": 300, "amount": -3}}
   (negative = scroll down, positive = scroll up)

14. Wait for something to load:
    {{"action": "wait", "seconds": 2}}

15. Take a verification screenshot:
    {{"action": "screenshot", "label": "after_open_file"}}

16. Delete rows (use when user says delete rows):
    {{"action": "delete_rows", "count": 3, "app": "excel"}}
    {{"action": "delete_rows", "count": 3, "app": "notepad"}}
    "count" = how many rows to delete from the top.
    "app"   = which app the file is open in (excel or notepad).
    Use "excel" for .xlsx and .csv files opened in Excel.
    Use "notepad" for .txt files opened in Notepad.

RULES — follow every one of these strictly:
- NEVER open notepad, wordpad, or any app the user did not ask for
- If user says "open notepad and write/type something", ALWAYS use open_app with app="notepad" followed by type_text. NEVER use open_file for this.
- If the user wants to compose a long email, message, or paragraph, prefer paste_text over type_text.
- open_file is ONLY for files that already exist on disk (.csv, .xlsx, .pdf, .txt files with real paths)
- If user mentions a file (.csv, .xlsx, .pdf etc.), use open_file with the full path
- Do NOT use open_app before open_file — open_file opens the correct app automatically
- File paths MUST use single backslashes: C:\\Users\\dell\\Desktop\\file.csv
- Always add a wait action after open_file (files need 3 seconds to load)
- Prefer click_text over click for buttons and menu items
- NEVER use type_text for keyboard shortcuts — always use press_key instead
- Use screenshot action after important steps to verify they worked
- Keep the action plan SHORT and focused — only include necessary steps
- For internet tasks, prefer open_url or browser navigation plus click_text and paste_text instead of inventing extra app-specific actions.
- Return ONLY the raw JSON array, no explanation, no markdown
- CALCULATOR: After typing an equation in calculator, ALWAYS press Enter or click "=" to get the result:
  {{"action": "press_key", "keys": ["enter"]}}
  OR {{"action": "click_text", "text": "="}}
  Without this step, the calculation will NOT be performed!
- BROWSER NAVIGATION: To go to a website in Chrome/Edge/Firefox:
  1. First open the browser the user mentioned: {{"action": "open_app", "app": "edge"}} or {{"action": "open_app", "app": "chrome"}}
  2. Wait for it to load: {{"action": "wait", "seconds": 2}}
  3. Press Ctrl+L to focus address bar: {{"action": "press_key", "keys": ["ctrl", "l"]}}
  4. Type the URL: {{"action": "type_text", "text": "gmail.com"}}
  5. Press Enter: {{"action": "press_key", "keys": ["enter"]}}
  IMPORTANT: Use the EXACT browser the user mentioned (edge, chrome, firefox). Do NOT substitute one for another!
  NEVER use open_file for URLs or websites!
- SEARCH QUERIES: When user says "search for X", "look up X", "find X", or "google X":
  1. Open the browser: {{"action": "open_app", "app": "chrome"}}
  2. Wait: {{"action": "wait", "seconds": 2}}
  3. Focus address bar: {{"action": "press_key", "keys": ["ctrl", "l"]}}
  4. Type Google search URL with query: {{"action": "type_text", "text": "google.com/search?q=cat images"}}
  5. Press Enter: {{"action": "press_key", "keys": ["enter"]}}
  Replace "cat images" with whatever the user wants to search for.
  NEVER type example.com or random URLs! Always use google.com/search?q=QUERY
"""

    max_retries = 3

    for attempt in range(max_retries):
        try:
            screenshot_b64 = get_screenshot_base64()

            payload = {
                "model": MODEL,
                "messages": [
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "text",
                                "text": prompt
                            },
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:image/png;base64,{screenshot_b64}"
                                    # "data:image/png;base64," prefix required
                                    # tells API this is inline base64, not a URL
                                }
                            }
                        ]
                    }
                ],

                "max_tokens": 2048,
                # ↑ INCREASED from 1024 → 2048
                # WHY: 1024 was cutting off mid-JSON for complex tasks
                # causing "Unterminated string" JSONDecodeError
                # 2048 gives enough room for detailed action plans

                "temperature": 0.10,
                # ↑ Low = deterministic, follows rules strictly
                # High = creative but unpredictable JSON output

                "top_p": 0.70,
                # ↑ Only consider top 70% probability tokens
                # Keeps output focused and consistent

                "stream": False
                # ↑ MUST be False — streaming breaks json.loads()
                # We need the complete response at once
            }

            response = requests.post(
                INVOKE_URL,
                headers=HEADERS,
                json=payload,
                timeout=60
                # timeout=60 → vision models can be slow, 60s is safe
            )

            response.raise_for_status()
            # ↑ Raises exception for 4xx/5xx HTTP errors
            # catches 429 (rate limit), 401 (bad key), 500 (server error)

            data = response.json()

            # Extract the model's text reply
            raw = data["choices"][0]["message"]["content"].strip()
            print(f"[LLM] Raw model response: {raw[:200]}...")

            # Remove markdown fences if model added them
            # e.g. ```json [...] ``` → [...]
            raw = re.sub(r"```json|```", "", raw).strip()

            # Fix double backslashes in paths
            # Even after instructions, model sometimes returns C:\\\\Users\\\\
            # This cleans it to C:\\Users\\ before json.loads()
            raw = raw.replace("\\\\\\\\", "\\\\")

            # ── PARSE JSON WITH TRUNCATION RECOVERY ──────────
            # WHY: if max_tokens is still not enough, the response
            # gets cut off mid-JSON. Instead of crashing, we try
            # to salvage the complete actions that were returned.
            try:
                actions = json.loads(raw)
                # ↑ Best case — full valid JSON parsed successfully

            except json.JSONDecodeError:
                print("[LLM] ⚠ JSON was truncated — attempting recovery...")

                # Strategy: find the last COMPLETE action in the list
                # A complete action ends with }
                # The full list ends with }]
                # We cut off everything after the last complete }

                last_close = raw.rfind("}]")
                # rfind() → finds LAST occurrence of "}]" (end of array)

                if last_close != -1:
                    # Found a proper array ending — use it
                    raw = raw[:last_close + 2]
                else:
                    # No proper ending — find last complete action
                    last_brace = raw.rfind("}")
                    if last_brace != -1:
                        # Close the array manually after last complete action
                        raw = raw[:last_brace + 1] + "]"
                    else:
                        raise Exception(
                            "Model response was completely malformed. "
                            "Try a simpler command or increase max_tokens further."
                        )

                # Try parsing the recovered JSON
                try:
                    actions = json.loads(raw)
                    print(f"[LLM] ✓ Recovered {len(actions)} actions from truncated response")
                except json.JSONDecodeError:
                    raise Exception(
                        "Could not recover JSON from truncated response. "
                        "Try a simpler/shorter command."
                    )
            # ─────────────────────────────────────────────────

            print(f"[LLM] Got {len(actions)} actions to execute")
            return actions

        except requests.exceptions.HTTPError as e:
            status_code = e.response.status_code if e.response else 0

            if status_code == 429:
                # Rate limit → wait and retry
                if attempt < max_retries - 1:
                    wait_seconds = 30 * (attempt + 1)
                    # 30s first retry, 60s second, 90s third
                    print(f"[LLM] ⚠ Rate limit. Waiting {wait_seconds}s before retry {attempt + 2}/{max_retries}...")
                    time.sleep(wait_seconds)
                else:
                    print("[LLM] ❌ All retries failed. Rate limit still active.")
                    raise
            elif status_code == 401:
                # Bad API key — no point retrying
                print("[LLM] ❌ Invalid API key. Check NVIDIA_API_KEY in .env")
                raise
            else:
                raise

        except Exception as e:
            error_str = str(e)
            if "429" in error_str and attempt < max_retries - 1:
                wait_seconds = 30 * (attempt + 1)
                print(f"[LLM] ⚠ Quota hit. Waiting {wait_seconds}s...")
                time.sleep(wait_seconds)
            else:
                raise

    raise Exception("Failed to generate action plan after all retries.")


# ── FUNCTION 3: Verify a Step Worked ─────────────────────────

def verify_step(command: str, step_description: str) -> bool:
    # After a screenshot action, ask model to confirm step succeeded
    # RETURNS: True = SUCCESS, False = FAILURE

    print(f"[LLM] Verifying: '{step_description}'")

    screenshot_b64 = get_screenshot_base64()

    prompt = f"""
Look at this screenshot carefully.

The agent was trying to complete this overall task: {command}
The last action it performed was: {step_description}

Did the action succeed based on what you see on screen?
Answer with ONLY one word: SUCCESS or FAILURE
Do not write anything else.
"""

    payload = {
        "model": MODEL,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/png;base64,{screenshot_b64}"
                        }
                    }
                ]
            }
        ],
        "max_tokens": 10,
        # ↑ Only need SUCCESS or FAILURE — 10 tokens is enough
        "temperature": 0.10,
        "stream": False
    }

    response = requests.post(
        INVOKE_URL,
        headers=HEADERS,
        json=payload,
        timeout=60
    )
    response.raise_for_status()

    data = response.json()
    result = data["choices"][0]["message"]["content"].strip().upper()
    print(f"[LLM] Verification result: {result}")

    return "SUCCESS" in result


def get_excel_action_plan(command: str) -> list:
    print("[LLM] Asking NVIDIA model for structured Excel actions...")

    prompt = f"""
You are generating Excel automation actions for a Windows desktop agent.
The user command is: {command}

Return ONLY a valid JSON array. No markdown. No explanation.
Each item must use one of these action names only:
open_workbook, select_sheet, aggregate, set_cell_value, set_range_value,
clear_contents, delete_rows, sort_sheet, filter_column, clear_filter,
save_workbook

Examples:
[{{"action":"open_workbook","reference":"C:\\Users\\DELL\\Downloads\\Orders.csv"}}]
[{{"action":"open_workbook","reference":"C:\\Users\\DELL\\Downloads\\Orders.csv"}},{{"action":"delete_rows","start_row":2,"end_row":10,"workbook":"C:\\Users\\DELL\\Downloads\\Orders.csv"}}]
[{{"action":"open_workbook","reference":"C:\\Users\\DELL\\Downloads\\Orders.csv"}},{{"action":"aggregate","function":"average","source_type":"column","source":"Amount","workbook":"C:\\Users\\DELL\\Downloads\\Orders.csv"}}]
[{{"action":"open_workbook","reference":"C:\\Users\\DELL\\Downloads\\Orders.csv"}},{{"action":"set_cell_value","target":"D2","value":"Approved","workbook":"C:\\Users\\DELL\\Downloads\\Orders.csv"}}]

Rules:
- Keep the action list short and specific.
- Do not use vague actions like edit, analyze, formula_task, or UI clicks.
- If the user provided a workbook path, preserve it exactly.
- If the request is Excel-related, return only structured Excel actions.
"""

    for _ in range(2):
        payload = {
            "model": MODEL,
            "messages": [{"role": "user", "content": prompt}],
            "max_tokens": 1024,
            "temperature": 0.10,
            "top_p": 0.70,
            "stream": False
        }

        response = requests.post(
            INVOKE_URL,
            headers=HEADERS,
            json=payload,
            timeout=60
        )
        response.raise_for_status()

        data = response.json()
        raw = data["choices"][0]["message"]["content"].strip()
        raw = re.sub(r"```json|```", "", raw).strip()
        raw = raw.replace("\\\\\\\\", "\\\\")

        actions = json.loads(raw)
        valid, note = validate_excel_actions(actions)
        if valid:
            print(f"[LLM] Got {len(actions)} structured Excel actions")
            return actions

        prompt += f"\nYour last answer was invalid because: {note}\nReturn a corrected JSON array only."

    raise Exception("The Excel action plan was too vague or invalid. Please make the Excel request more specific.")


def generate_professional_email(intent_text: str, recipient_email: str | None = None) -> dict:
    print("[LLM] Asking NVIDIA model to professionalize email draft...")

    recipient_name = _recipient_name_from_email(recipient_email) if recipient_email else ""
    recipient_context = (
        f"Recipient email: {recipient_email}\nPreferred recipient name: {recipient_name}\n"
        if recipient_email
        else "Recipient is not provided. Use a generic greeting only.\n"
    )

    prompt = f"""
You are an assistant that rewrites rough notes into a professional email draft.

User intent/details:
{intent_text}

{recipient_context}

Return ONLY valid JSON object with this exact shape:
{{"subject":"...","body":"..."}}

Rules:
- Keep tone professional, polite, concise.
- Do not invent facts.
- If leave/sick context exists, produce a suitable leave message.
- Body must be plain text, no markdown.
- Never output placeholders like [Recipient Name], [Name], <name>, {{name}}, or similar.
- If recipient name is known, use it naturally in greeting.
- If recipient name is unknown, do not use any placeholder; use a generic greeting.
"""

    payload = {
        "model": MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": 400,
        "temperature": 0.20,
        "top_p": 0.80,
        "stream": False,
    }

    response = requests.post(
        INVOKE_URL,
        headers=HEADERS,
        json=payload,
        timeout=60,
    )
    response.raise_for_status()

    data = response.json()
    raw = data["choices"][0]["message"]["content"].strip()
    raw = re.sub(r"```json|```", "", raw).strip()

    parsed = json.loads(raw)
    subject = str(parsed.get("subject", "Quick update")).strip() or "Quick update"
    body = str(parsed.get("body", "")).strip()
    body = _sanitize_placeholder_text(body, recipient_name)
    return {"subject": subject, "body": body}


def _recipient_name_from_email(email: str | None) -> str:
    if not email or "@" not in email:
        return ""
    local = email.split("@", 1)[0]
    local = re.sub(r"[._-]+", " ", local)
    local = re.sub(r"\s+", " ", local).strip()
    if not local:
        return ""
    return " ".join(part.capitalize() for part in local.split(" "))


def _sanitize_placeholder_text(text: str, recipient_name: str) -> str:
    cleaned = text
    placeholder_pattern = re.compile(
        r"\[(recipient'?s?\s*name|name)\]|<\s*name\s*>|\{\{\s*name\s*\}\}|\[\s*name\s*\]",
        re.IGNORECASE,
    )
    replacement = recipient_name if recipient_name else "there"
    cleaned = placeholder_pattern.sub(replacement, cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned
