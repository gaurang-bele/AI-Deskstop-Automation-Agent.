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

import threading
import hashlib

import logging
# logging → we set up a log file here so ALL modules write
# to the same agent.log file

import os
import re
import json
import datetime
import requests
import io
from urllib.parse import quote_plus, urlparse, parse_qs, unquote
from base64 import b64decode
# os → used to create initial files and folders on startup

# Import our own modules (the other files we built)
from watcher   import start_watching   # the file watcher
from llm       import (
    get_action_plan,
    get_excel_action_plan,
    generate_professional_email,
    get_research_findings,
    get_usage_totals,
    get_last_call_usage,
    log_external_prompt,
    verify_step,
)  # model planner
from executor  import execute_action   # action runner
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
try:
    from playwright_email import send_gmail_with_playwright
except Exception:
    send_gmail_with_playwright = None

try:
    from browser_agent import run_agentic_browse
except Exception:
    run_agentic_browse = None
try:
    from gmail_api import send_gmail_with_api
except Exception:
    send_gmail_with_api = None
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
    "attach_file",
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


def _get_email_mode() -> str:
    """Choose email delivery strategy: api, playwright, or ui."""
    mode = os.getenv("AGENT_EMAIL_MODE", "api").strip().lower()
    if mode not in {"api", "playwright", "ui"}:
        return "api"
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
    if is_excel_command(command) or is_email_command(command):
        return False

    research_terms = (
        "fetch information",
        "find information",
        "collect information",
        "research",
        "search for",
        "look up",
        "lookup",
    )
    if any(term in lower for term in research_terms):
        return True

    detail_terms = ("price", "density", "use", "uses", "application", "applications", "information")
    return ("google" in lower or "search" in lower) and any(term in lower for term in detail_terms)


def _is_agentic_browse_command(command: str) -> bool:
    lowered = (command or "").strip().lower()
    return lowered.startswith("browse:") or lowered.startswith("web:") or lowered.startswith("agentic:")


def _strip_agentic_prefix(command: str) -> str:
    lowered = (command or "").strip()
    for prefix in ("browse:", "web:", "agentic:"):
        if lowered.lower().startswith(prefix):
            return lowered[len(prefix):].strip()
    return lowered


def _parse_research_request(command: str) -> tuple[str, list[str], str]:
    lower = command.lower()

    output_file = "response.txt"
    file_match = re.search(r"\b(?:in|into|to)\s+([a-zA-Z0-9_.-]+\.txt)\b", command, re.IGNORECASE)
    if file_match:
        output_file = file_match.group(1)

    field_aliases = {
        "price": ("price", "cost", "pricing", "rate"),
        "density": ("density",),
        "use": ("use", "uses", "used for", "application", "applications"),
        "history": ("history", "background", "origin"),
        "founder": ("founder", "founded by", "creator"),
        "specifications": ("specification", "specifications", "specs", "features"),
        "voltage": ("voltage", "voltage rating", "rated voltage", "wv", "working voltage"),
        "temperature coefficient": (
            "temperature coefficient",
            "temp coefficient",
            "tempco",
            "tcc",
            "ppm/°c",
            "ppm/ c",
            "ppm/degc",
        ),
        "operating temperature": (
            "operating temperature",
            "operating temp",
            "temperature range",
            "temp range",
            "-55",
            "+125",
        ),
        "capacitance": ("capacitance",),
        "tolerance": ("tolerance",),
        "package": ("package", "size", "case"),
        "models": ("model", "models", "variants"),
        "top speed": ("top speed", "speed", "max speed"),
        "color options": ("color", "colours", "colors", "colour options", "color options"),
    }

    def _extract_custom_fields(text: str) -> list[str]:
        match = re.search(r"\bits\b\s+([^\.]+)$", text, re.IGNORECASE)
        if not match:
            match = re.search(r"\bincluding\b\s+([^\.]+)$", text, re.IGNORECASE)
        if not match:
            return []
        tail = match.group(1)
        tail = re.sub(r"\s+and\s+(?:save|write|paste|put)\b.*$", "", tail, flags=re.IGNORECASE)
        tail = tail.strip(" .")
        if not tail:
            return []
        parts = re.split(r",|\band\b", tail, flags=re.IGNORECASE)
        cleaned: list[str] = []
        for part in parts:
            field = re.sub(r"[^a-zA-Z0-9 %/°µ+-]+", " ", part).strip().lower()
            field = re.sub(r"\s+", " ", field).strip()
            if not field or len(field) < 3:
                continue
            if field not in cleaned:
                cleaned.append(field)
        return cleaned[:6]

    custom_fields = _extract_custom_fields(command)
    alias_fields = [name for name, aliases in field_aliases.items() if any(alias in lower for alias in aliases)]
    fields: list[str] = []

    if custom_fields:
        normalized: list[str] = []
        for field in custom_fields:
            mapped = None
            for canonical, aliases in field_aliases.items():
                if canonical in {"models", "top speed", "color options"}:
                    continue
                if field == canonical or any(alias == field for alias in aliases):
                    mapped = canonical
                    break
            normalized.append(mapped or field)
        fields = ["overview"] + [f for f in normalized if f != "overview"]
    elif alias_fields:
        fields = ["overview"] + alias_fields

    query = command
    for pattern in (r"\brelated to\s+(.+)", r"\babout\s+(.+)", r"\bon\s+(.+)", r"\bregarding\s+(.+)", r"\bfor\s+(.+)"):
        match = re.search(pattern, command, re.IGNORECASE)
        if match:
            query = match.group(1)
            break
    query = re.sub(r"\s+and\s+(?:save|write|paste|put)\b.*$", "", query, flags=re.IGNORECASE).strip(" .")
    query = re.sub(r"\s+in(?:to)?\s+[a-zA-Z0-9_.-]+\.txt$", "", query, flags=re.IGNORECASE).strip(" .")
    query = re.split(r"\b(?:its|including)\b", query, maxsplit=1, flags=re.IGNORECASE)[0].strip(" ,.")
    if not query:
        query = command.strip()

    if not custom_fields and not alias_fields:
        if _looks_like_part_number(query):
            fields = [
                "overview",
                "voltage",
                "temperature coefficient",
                "operating temperature",
                "capacitance",
                "tolerance",
                "package",
            ]
        else:
            fields = ["overview", "key facts", "common uses"]

    seen: set[str] = set()
    deduped: list[str] = []
    for item in fields:
        if item in seen:
            continue
        seen.add(item)
        deduped.append(item)
    fields = deduped

    return query, fields, output_file


def _browser_for_command(command: str) -> str:
    lower = command.lower()
    if "edge" in lower or "microsoft edge" in lower:
        return "edge"
    if "firefox" in lower:
        return "firefox"
    if "chrome" in lower or "google" in lower:
        return "chrome"
    return "chrome"


def _research_query_suffix(fields: list[str]) -> str:
    if any(field != "overview" for field in fields):
        specific = [f for f in fields if f != "overview"]
        return " ".join(specific[:3] + ["overview"])
    return "overview key facts common uses"


def _memory_context(limit: int = 8) -> str:
    path = "agent_memory.txt"
    if not os.path.exists(path):
        return ""
    with open(path, "r", encoding="utf-8") as f:
        lines = [line.strip() for line in f.readlines() if line.strip() and not line.startswith("#")]
    if not lines:
        return ""
    return "\n".join(lines[-limit:])


def _append_memory_entry(command: str, status: str, summary: str, usage_delta: dict | None = None) -> None:
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    usage = usage_delta or {}
    memory_line = (
        f"{timestamp} | {status} | command={command} | summary={summary} | "
        f"tokens={int(usage.get('total_tokens', 0))} | cost_usd={float(usage.get('total_cost_usd', 0.0)):.8f}"
    )
    with open("agent_memory.txt", "a", encoding="utf-8") as f:
        f.write(memory_line + "\n")

    payload = {
        "timestamp": timestamp,
        "status": status,
        "command": command,
        "summary": summary,
        "usage": usage,
    }
    with open("agent_memory.jsonl", "a", encoding="utf-8") as f:
        f.write(json.dumps(payload, ensure_ascii=False) + "\n")


def _extract_findings_from_clipboard(text: str, fields: list[str], query: str = "") -> dict[str, list[str]]:
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

        if "overview" in fields and len(findings["overview"]) < 4:
            if len(s.split()) >= 8 and (not query or query.lower() in lower):
                findings["overview"].append(s)

        if "key facts" in fields and len(findings["key facts"]) < 4:
            if any(token in lower for token in ("is", "are", "founded", "headquartered", "known for")):
                findings["key facts"].append(s)

        if "common uses" in fields and len(findings["common uses"]) < 4:
            if any(token in lower for token in ("used", "use", "application", "applications", "popular for")):
                findings["common uses"].append(s)

        for field in fields:
            if field in {"price", "density", "use", "overview", "key facts", "common uses"}:
                continue
            if field.lower() in lower and len(findings[field]) < 4:
                findings[field].append(s)

    return {k: v for k, v in findings.items() if v}


def _search_url_for_command(command: str, query: str, fields: list[str]) -> str:
    suffix = _research_query_suffix(fields)
    token = quote_plus(f"{query} {suffix}")
    if "bing" in command.lower():
        return f"bing.com/search?q={token}"
    return f"google.com/search?q={token}"


def _search_results_urls(query: str, limit: int = 3) -> list[str]:
    # Intentionally no Wikipedia/Wikidata shortcut.
    # It caused many technical / part-number queries to fail because those sites
    # rarely contain the required specifications.

    def _query_keywords(value: str) -> list[str]:
        stop_words = {
            "about", "with", "from", "that", "this", "what", "when", "where", "which", "there",
            "their", "price", "prices", "information", "details", "option", "options", "model",
            "models", "color", "colors", "speed", "top", "overview", "common", "uses", "facts",
            "and", "for", "the", "its", "diffrent", "different",
        }
        tokens = re.findall(r"[a-z0-9]+", (value or "").lower())
        filtered: list[str] = []
        for token in tokens:
            if len(token) < 3 or token in stop_words or token.isdigit():
                continue
            if token not in filtered:
                filtered.append(token)
        return filtered[:6]

    def _url_relevance_score(link: str, keywords: list[str]) -> int:
        lower = link.lower()
        if any(bad in lower for bad in ("whatsapp", "wa.me", "web.whatsapp.com", "teknogram.id")):
            return -1
        score = 0
        for token in keywords:
            if token in lower:
                score += 1
        # De-prioritize low-signal community pages, but don't fully drop them.
        if any(bad in lower for bad in ("forum", "forums", "thread", "tapatalk", "boardreader", "reddit.com", "quora.com")):
            score -= 1
        return score

    def _decode_bing_redirect(link: str) -> str:
        try:
            parsed = urlparse(link)
            if "bing.com" not in (parsed.netloc or "").lower():
                return link
            params = parse_qs(parsed.query)
            u_values = params.get("u")
            if not u_values:
                return link
            payload = u_values[0]
            if payload.startswith("a1"):
                payload = payload[2:]
            payload += "=" * ((4 - len(payload) % 4) % 4)
            decoded = b64decode(payload).decode("utf-8", errors="ignore").strip()
            if decoded.startswith("http"):
                return decoded
        except Exception:
            pass
        return link

    def _from_duckduckgo_html(search_query: str) -> list[str]:
        ddg_url = f"https://duckduckgo.com/html/?q={quote_plus(search_query)}"
        ddg_response = requests.get(
            ddg_url,
            timeout=20,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AiAgent/1.0",
                "Accept-Language": "en-US,en;q=0.9",
            },
        )
        ddg_response.raise_for_status()
        html = ddg_response.text
        raw_links = re.findall(r'href="(https?://[^"]+)"', html, flags=re.IGNORECASE)
        cleaned_links: list[str] = []
        for link in raw_links:
            parsed = urlparse(link)
            host = (parsed.netloc or "").lower()
            if "duckduckgo.com" in host:
                params = parse_qs(parsed.query)
                redirect = params.get("uddg", [])
                if redirect:
                    candidate = unquote(redirect[0])
                else:
                    continue
            else:
                candidate = link
            lower = candidate.lower()
            if _url_relevance_score(candidate, _query_keywords(search_query)) < 0:
                continue
            if candidate.startswith("http") and candidate not in cleaned_links:
                cleaned_links.append(candidate)
            if len(cleaned_links) >= max(limit * 3, 8):
                break
        return cleaned_links

    url = f"https://www.bing.com/search?q={quote_plus(query)}&format=rss"
    response = requests.get(
        url,
        timeout=20,
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AiAgent/1.0",
            "Accept": "application/rss+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
        },
    )
    response.raise_for_status()
    xml = response.text
    links = re.findall(r"<item>.*?<link>(.*?)</link>.*?</item>", xml, flags=re.DOTALL | re.IGNORECASE)

    blocked = ("caradisiac", "tapatalk", "boardreader")
    keywords = _query_keywords(query)

    scored_results: list[tuple[int, str]] = []
    results: list[str] = []
    for link in links:
        cleaned = link.strip()
        if not cleaned.startswith("http"):
            continue
        cleaned = _decode_bing_redirect(cleaned)
        lower = cleaned.lower()
        score = _url_relevance_score(cleaned, keywords)
        if score < 0:
            continue
        if cleaned not in results:
            results.append(cleaned)
            scored_results.append((score, cleaned))
        if len(results) >= max(limit * 3, 8):
            break

    relevant = [url for score, url in scored_results if score > 0][:limit]
    if relevant:
        return relevant

    # If relevance is weak (no keyword hits), avoid returning junk; try other engines first.
    max_score = max([score for score, _ in scored_results], default=-1)

    # Fallback: DuckDuckGo HTML search to avoid empty/irrelevant RSS returns.
    try:
        ddg_links = _from_duckduckgo_html(query)
        if ddg_links:
            ddg_keywords = _query_keywords(query)
            ddg_scored = sorted(
                [(_url_relevance_score(url, ddg_keywords), url) for url in ddg_links],
                key=lambda item: item[0],
                reverse=True,
            )
            ddg_positive = [url for score, url in ddg_scored if score > 0][:limit]
            if ddg_positive:
                return ddg_positive
    except Exception:
        pass

    # Last fallback: SerpAPI mirror via r.jina.ai (no API key required for simple scrape).
    try:
        jina_url = f"https://r.jina.ai/http://serpapi.com/search.json?q={quote_plus(query)}&engine=google&num={max(limit,3)}"
        jina_response = requests.get(
            jina_url,
            timeout=20,
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AiAgent/1.0"},
        )
        if jina_response.ok:
            lines = [line.strip() for line in jina_response.text.splitlines() if line.strip()]
            url_pattern = re.compile(r"https?://[^\s)\"]+")
            collected: list[str] = []
            for line in lines:
                for match in url_pattern.findall(line):
                    lower = match.lower()
                    if "serpapi.com" in lower or "google.com/search" in lower:
                        continue
                    if any(bad in lower for bad in blocked):
                        continue
                    if _url_relevance_score(match, _query_keywords(query)) <= 0:
                        continue
                    if match not in collected:
                        collected.append(match)
                    if len(collected) >= limit:
                        return collected
    except Exception:
        pass

    # Final fallback: return best-ranked Bing links only if we have *some* signal.
    if scored_results and max_score > 0:
        scored_results.sort(key=lambda item: item[0], reverse=True)
        return [url for score, url in scored_results if score > 0][:limit]

    # If nothing looks relevant, return empty rather than unrelated URLs.
    return []


def _merge_findings(base: dict[str, list[str]], incoming: dict[str, list[str]], limit: int = 8) -> dict[str, list[str]]:
    merged: dict[str, list[str]] = {k: list(v) for k, v in base.items()}
    for field, values in (incoming or {}).items():
        existing = merged.get(field, [])
        for value in values:
            if value not in existing:
                existing.append(value)
            if len(existing) >= limit:
                break
        if existing:
            merged[field] = existing[:limit]
    return merged


def _looks_like_part_number(value: str) -> bool:
    token = re.sub(r"\s+", "", (value or "").strip())
    if len(token) < 8:
        return False
    if not re.search(r"[A-Za-z]", token) or not re.search(r"\d", token):
        return False
    return bool(re.fullmatch(r"[A-Za-z0-9._/-]+", token))


def _findings_complete(findings: dict[str, list[str]], fields: list[str], min_per_field: int = 3) -> bool:
    if not fields:
        return True
    for field in fields:
        if field == "overview":
            continue
        if len(findings.get(field, [])) < min_per_field:
            return False
    return True


def _part_number_source_urls(part_number: str) -> list[str]:
    token = (part_number or "").strip()
    if not token:
        return []
    q = quote_plus(token)
    return [
        f"https://www.alldatasheet.com/view.jsp?Searchword={q}",
        f"https://octopart.com/search?q={q}",
        f"https://www.mouser.com/c/?q={q}",
        f"https://www.digikey.com/en/products/result?s={q}",
        f"https://www.murata.com/en-global/search/results?search={q}",
    ]


def _extract_text_from_url(url: str) -> str:
    if not url:
        return ""

    def _extract_text_from_pdf_bytes(data: bytes) -> str:
        try:
            from pypdf import PdfReader  # type: ignore
        except Exception:
            return ""
        try:
            reader = PdfReader(io.BytesIO(data))
            chunks: list[str] = []
            for page in reader.pages[:8]:
                text = page.extract_text() or ""
                text = text.strip()
                if text:
                    chunks.append(text)
            return "\n".join(chunks)
        except Exception:
            return ""

    normalized = url.strip()
    if not re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", normalized):
        normalized = f"https://{normalized}"
    try:
        response = requests.get(
            normalized,
            timeout=20,
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AiAgent/1.0"},
        )
        response.raise_for_status()
        content_type = (response.headers.get("Content-Type") or "").lower()
        if "pdf" in content_type or normalized.lower().endswith(".pdf"):
            pdf_text = _extract_text_from_pdf_bytes(response.content)
            pdf_text = re.sub(r"\s+", " ", pdf_text).strip()
            if pdf_text:
                return pdf_text[:20000]
        html = response.text
    except Exception:
        # Reader fallback helps with pages blocked by anti-bot / 403 protections.
        reader_url = f"https://r.jina.ai/http://{normalized.replace('https://', '').replace('http://', '')}"
        reader_resp = requests.get(
            reader_url,
            timeout=20,
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AiAgent/1.0"},
        )
        reader_resp.raise_for_status()
        html = reader_resp.text
    html = re.sub(r"(?is)<script.*?>.*?</script>", " ", html)
    html = re.sub(r"(?is)<style.*?>.*?</style>", " ", html)
    text = re.sub(r"(?s)<[^>]+>", " ", html)
    text = re.sub(r"\s+", " ", text).strip()
    return text[:20000]


def _usage_delta(before: dict[str, int | float], after: dict[str, int | float]) -> dict[str, int | float]:
    int_keys = ("requests", "prompt_tokens", "completion_tokens", "total_tokens")
    float_keys = ("prompt_cost_usd", "completion_cost_usd", "total_cost_usd")
    delta: dict[str, int | float] = {}
    for key in int_keys:
        delta[key] = max(0, int(after.get(key, 0)) - int(before.get(key, 0)))
    for key in float_keys:
        diff = float(after.get(key, 0.0)) - float(before.get(key, 0.0))
        delta[key] = round(max(0.0, diff), 8)
    return delta


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


_COMMAND_DEDUPE_LOCK = threading.Lock()
_LAST_COMMAND_FINGERPRINT = ""
_LAST_COMMAND_TS = 0.0

_EMAIL_DEDUPE_LOCK = threading.Lock()
_RECENT_EMAIL_FINGERPRINTS: dict[str, float] = {}


def _fingerprint_payload(payload: dict) -> str:
    dumped = json.dumps(payload, sort_keys=True, ensure_ascii=False, default=str)
    return hashlib.sha256(dumped.encode("utf-8")).hexdigest()


def _should_suppress_duplicate_command(command: str, window_seconds: float = 2.0) -> bool:
    global _LAST_COMMAND_FINGERPRINT, _LAST_COMMAND_TS
    now = time.monotonic()
    candidate = _fingerprint_payload({"command": (command or "").strip()})
    with _COMMAND_DEDUPE_LOCK:
        if candidate and candidate == _LAST_COMMAND_FINGERPRINT and (now - _LAST_COMMAND_TS) < window_seconds:
            return True
        _LAST_COMMAND_FINGERPRINT = candidate
        _LAST_COMMAND_TS = now
        return False


def _should_suppress_duplicate_email(details: dict, window_seconds: float = 120.0) -> bool:
    """Best-effort guard against double-trigger sends (e.g., watcher fires twice).

    If the exact same email payload is requested again inside the window,
    suppress the second send.
    """
    if not details.get("send"):
        return False

    payload = {
        "to": details.get("to") or [],
        "cc": details.get("cc") or [],
        "bcc": details.get("bcc") or [],
        "subject": details.get("subject") or "",
        "body": details.get("body") or "",
        "attachments": details.get("attachments") or [],
        "mode": _get_email_mode(),
    }
    fp = _fingerprint_payload(payload)
    now = time.monotonic()
    with _EMAIL_DEDUPE_LOCK:
        expired = [key for key, ts in _RECENT_EMAIL_FINGERPRINTS.items() if (now - ts) > window_seconds]
        for key in expired:
            _RECENT_EMAIL_FINGERPRINTS.pop(key, None)

        if fp in _RECENT_EMAIL_FINGERPRINTS:
            return True
        _RECENT_EMAIL_FINGERPRINTS[fp] = now
        return False


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

    if not os.path.exists("agent_memory.txt"):
        with open("agent_memory.txt", "w", encoding="utf-8") as f:
            f.write("# Agent command memory\n")

    if not os.path.exists("agent_memory.jsonl"):
        with open("agent_memory.jsonl", "w", encoding="utf-8") as f:
            f.write("")

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

    # Guard against rapid duplicate triggers (e.g., file watcher fires twice for one save).
    if _should_suppress_duplicate_command(command):
        print("[AGENT] 🛑 Duplicate command trigger suppressed (debounce).")
        usage_after = get_usage_totals()
        write_response(
            command=command,
            actions=[{"action": "debounce"}],
            results=[
                {
                    "step": 1,
                    "action": "debounce",
                    "params": {"action": "debounce"},
                    "success": True,
                    "note": "Duplicate trigger suppressed",
                }
            ],
            llm_usage={},
            llm_usage_total=usage_after,
            llm_last_call=get_last_call_usage(),
        )
        return

    try:
        usage_before = get_usage_totals()
        memory_context = _memory_context()
        excel_mode = _get_excel_mode()
        excel_request = is_excel_command(command)
        email_request = is_email_command(command)
        research_request = _is_web_research_command(command)
        agentic_request = _is_agentic_browse_command(command)
        actions = None

        # ── STEP 2: Build Action Plan ───────────────────────
        if agentic_request:
            goal = _strip_agentic_prefix(command)
            print("[AGENT] 🧭 Agentic browser command detected - using Playwright loop...")
            if run_agentic_browse is None:
                raise Exception("Agentic browsing requires Playwright dependencies. Install with: pip install -r requirements.txt")
            actions, results, error = run_agentic_browse(goal)
            usage_after = get_usage_totals()
            usage_delta = _usage_delta(usage_before, usage_after)
            write_response(
                command=command,
                actions=actions or [{"action": "agentic_browse", "goal": goal}],
                results=results,
                llm_usage=usage_delta,
                llm_usage_total=usage_after,
                llm_last_call=get_last_call_usage(),
            )
            status = "SUCCESS" if not error else "PARTIAL"
            summary = "agentic browse complete" if not error else f"agentic browse error: {error}"
            _append_memory_entry(command=command, status=status, summary=summary, usage_delta=usage_delta)
            return

        if research_request:
            print("[AGENT] 🌐 Web research command detected - using headless fetch flow...")
            query, fields, output_file = _parse_research_request(command)
            browser = _browser_for_command(command)
            search_query = f"{query} {_research_query_suffix(fields)}"
            if _looks_like_part_number(query) and "datasheet" not in search_query.lower():
                search_query = f"{search_query} datasheet specifications"
            # Use the enriched search_query directly so the on-screen search
            # aligns with the headless extraction query.
            if "bing" in command.lower():
                search_url = f"bing.com/search?q={quote_plus(search_query)}"
            else:
                search_url = f"google.com/search?q={quote_plus(search_query)}"
            log_external_prompt(
                "research_headless_flow",
                f"query={query} fields={fields} browser={browser} search_url={search_url}",
                {"command": command},
            )

            results = []
            step_counter = 0
            if any(token in command.lower() for token in ("open edge", "open chrome", "open firefox", "open browser", "open google")):
                success, note = execute_action({"action": "open_app", "app": browser})
                step_counter += 1
                results.append(
                    {
                        "step": step_counter,
                        "action": "open_app",
                        "params": {"action": "open_app", "app": browser},
                        "success": success,
                        "note": note,
                    }
                )
                if success:
                    nav_success, nav_note = execute_action({"action": "open_url", "url": search_url, "app": browser})
                    step_counter += 1
                    results.append(
                        {
                            "step": step_counter,
                            "action": "open_url",
                            "params": {"action": "open_url", "url": search_url, "app": browser},
                            "success": nav_success,
                            "note": nav_note or "Opened search URL",
                        }
                    )
                    if nav_success:
                        wait_success, wait_note = execute_action({"action": "wait", "seconds": 2})
                        step_counter += 1
                        results.append(
                            {
                                "step": step_counter,
                                "action": "wait",
                                "params": {"action": "wait", "seconds": 2},
                                "success": wait_success,
                                "note": wait_note or "Waited for navigation",
                            }
                        )


            urls: list[str] = []
            sources: list[dict[str, str]] = []
            findings: dict[str, list[str]] = {}
            error_msg = None
            try:
                if _looks_like_part_number(query):
                    urls.extend(_part_number_source_urls(query))
                search_urls = _search_results_urls(search_query, limit=5)
                for candidate in search_urls:
                    if candidate not in urls:
                        urls.append(candidate)
                sources = [{"title": "", "url": url} for url in urls]
            except Exception as search_error:
                error_msg = f"Headless search failed: {search_error}"

            for url in urls:
                try:
                    page_text = _extract_text_from_url(url)
                except Exception as fetch_error:
                    error_msg = error_msg or f"Fetch failed: {fetch_error}"
                    continue
                if not page_text:
                    continue
                partial = get_research_findings(query, fields, page_text)
                findings = _merge_findings(findings, partial)
                if _findings_complete(findings, fields, min_per_field=3):
                    break

            if not findings and not error_msg:
                error_msg = "Could not extract requested details from fetched sources."
            usage_after = get_usage_totals()
            usage_delta = _usage_delta(usage_before, usage_after)
            write_research_report(
                command=command,
                query=query,
                fields=fields,
                findings=findings,
                sources=sources,
                output_file=output_file,
                error=None if findings else error_msg,
                llm_usage=usage_delta,
                llm_usage_total=usage_after,
                steps=results,
            )
            success_count = sum(1 for row in results if row["success"])
            _append_memory_entry(
                command=command,
                status="SUCCESS" if findings else "PARTIAL",
                summary=f"research findings fields={len(findings)} steps={success_count}/{len(results)}",
                usage_delta=usage_delta,
            )
            return

        # Email commands use a deterministic Gmail parser to support
        # both structured and natural-language requests.
        elif email_request:
            print("[AGENT] 📧 Email command detected - building Gmail action plan...")
            details = parse_email_request(command)
            email_intent = str(details.get("intent", "compose_send") or "compose_send")
            can_polish = email_intent in {
                "compose_send",
                "schedule_send",
                "message_reply",
                "message_reply_all",
                "message_forward",
            }
            if can_polish and wants_model_polish(command):
                raw_intent = details.get("body") or details.get("raw_text") or command
                recipient_email = (details.get("to") or [None])[0]
                try:
                    polished = generate_professional_email(raw_intent, recipient_email=recipient_email)
                    details["subject"] = polished.get("subject", details.get("subject", "Quick update"))
                    details["body"] = polished.get("body", details.get("body", ""))
                    print("[AGENT] ✨ Email draft polished by model.")
                except Exception as polish_error:
                    print(f"[AGENT] ⚠ Could not polish email with model: {polish_error}")

            email_mode = _get_email_mode()

            if _should_suppress_duplicate_email(details):
                print("[AGENT] 🛑 Duplicate email send suppressed.")
                usage_after = get_usage_totals()
                usage_delta = _usage_delta(usage_before, usage_after)
                write_response(
                    command=command,
                    actions=[{"action": "email_send_guard"}],
                    results=[
                        {
                            "step": 1,
                            "action": "email_send_guard",
                            "params": {"action": "email_send_guard"},
                            "success": True,
                            "note": "Duplicate email detected within safety window; second send suppressed.",
                        }
                    ],
                    llm_usage=usage_delta,
                    llm_usage_total=usage_after,
                    llm_last_call=get_last_call_usage(),
                )
                _append_memory_entry(
                    command=command,
                    status="SUCCESS",
                    summary="duplicate email suppressed",
                    usage_delta=usage_delta,
                )
                return
            if email_mode == "api":
                print("[AGENT] 🔐 Sending email via Gmail API (OAuth)...")
                if send_gmail_with_api is None:
                    raise Exception("Gmail API mode requires Google auth libraries. Install with: pip install -r requirements.txt")
                actions, results, error = send_gmail_with_api(details)
                usage_after = get_usage_totals()
                usage_delta = _usage_delta(usage_before, usage_after)
                write_response(
                    command=command,
                    actions=actions,
                    results=results,
                    llm_usage=usage_delta,
                    llm_usage_total=usage_after,
                    llm_last_call=get_last_call_usage(),
                )
                status = "SUCCESS" if not error else "PARTIAL"
                summary = f"gmail api {email_intent} ok" if not error else f"gmail api {email_intent} error: {error}"
                _append_memory_entry(command=command, status=status, summary=summary, usage_delta=usage_delta)
                return

            if email_mode == "playwright":
                print("[AGENT] 🌐 Sending email via Playwright browser automation...")
                if send_gmail_with_playwright is None:
                    raise Exception("Playwright email mode requires Playwright. Install with: pip install -r requirements.txt")
                actions, results, error = send_gmail_with_playwright(details)
                usage_after = get_usage_totals()
                usage_delta = _usage_delta(usage_before, usage_after)
                write_response(
                    command=command,
                    actions=actions,
                    results=results,
                    llm_usage=usage_delta,
                    llm_usage_total=usage_after,
                    llm_last_call=get_last_call_usage(),
                )
                status = "SUCCESS" if not error else "PARTIAL"
                summary = "playwright email sent" if not error else f"playwright email error: {error}"
                _append_memory_entry(command=command, status=status, summary=summary, usage_delta=usage_delta)
                return

            actions = build_email_actions(details)
            valid, note = validate_email_actions(actions)
            if not valid:
                raise Exception(f"Email action plan invalid: {note}")

        # Default is hybrid: try generic desktop planner first,
        # then fallback to structured Excel only when needed.
        elif excel_request:
            print(f"[AGENT] 📊 Excel command detected (mode={excel_mode}) - trying deterministic parser first...")
            actions = parse_excel_actions_from_command(command)
            valid, note = validate_excel_actions(actions)

            if valid:
                print(f"[AGENT] ✅ Using {len(actions)} structured Excel action(s) from parser.")
            elif excel_mode in {"structured", "hybrid"}:
                print(f"[AGENT] Excel parser not specific enough ({note}) - asking model for structured Excel actions...")
                actions = get_excel_action_plan(command)
                valid, note = validate_excel_actions(actions)
                if not valid:
                    if excel_mode == "hybrid":
                        print("[AGENT] ⚠ Structured Excel plan still invalid - falling back to generic planner...")
                        actions = get_action_plan(command, memory_context=memory_context)
                        unknown = _unsupported_actions(actions)
                        if unknown:
                            raise Exception(f"Excel action plan invalid ({note}) and generic fallback returned unsupported action(s): {unknown}")
                    else:
                        raise Exception(f"Excel action plan invalid: {note}")
            else:
                print("[AGENT] Excel mode=ui - asking model for generic action plan...")
                actions = get_action_plan(command, memory_context=memory_context)
                unknown = _unsupported_actions(actions)
                if unknown:
                    raise Exception(f"Generic Excel plan returned unsupported action(s): {unknown}")
        else:
            print("[AGENT] 🧠 Asking NVIDIA for generic action plan...")
            actions = get_action_plan(command, memory_context=memory_context)

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
        usage_after = get_usage_totals()
        usage_delta = _usage_delta(usage_before, usage_after)
        write_response(
            command,
            actions,
            results,
            llm_usage=usage_delta,
            llm_usage_total=usage_after,
            llm_last_call=get_last_call_usage(),
        )
        _append_memory_entry(
            command=command,
            status="SUCCESS" if all(r.get("success") for r in results) else "PARTIAL",
            summary=f"completed_steps={sum(1 for r in results if r['success'])}/{len(actions)}",
            usage_delta=usage_delta,
        )

    except Exception as e:
        # ── CATCH ANY UNEXPECTED ERROR ────────────────────────
        # This catches: bad API key, network errors, JSON parse
        # failures, etc. We write a clear error to response.txt
        error_msg = str(e)
        print(f"[AGENT] ❌ Error: {error_msg}")
        logging.error(f"Agent error for command '{command}': {error_msg}", exc_info=True)
        # exc_info=True → also logs the full stack trace to agent.log
        write_error(command, error_msg)
        usage_after = get_usage_totals()
        usage_delta = _usage_delta(usage_before, usage_after)
        _append_memory_entry(
            command=command,
            status="ERROR",
            summary=error_msg[:220],
            usage_delta=usage_delta,
        )


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
