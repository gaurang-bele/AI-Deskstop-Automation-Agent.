import re
from typing import Any
from urllib.parse import quote_plus

EMAIL_REGEX = r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"


EMAIL_COMMAND_HINTS = (
    "email",
    "gmail",
    "mail",
    "compose",
    "send to",
    "send to model",
    "polish",
)


def is_email_command(command: str) -> bool:
    text = (command or "").lower()
    if any(hint in text for hint in EMAIL_COMMAND_HINTS):
        return True
    if "send to model" in text:
        return True
    return bool(re.search(r"\b(send|compose|write)\s+(an\s+)?email\b", text))


def parse_email_actions_from_command(command: str) -> list[dict[str, Any]]:
    details = parse_email_request(command)
    return build_email_actions(details)


def parse_email_request(command: str) -> dict[str, Any]:
    return _extract_email_details(command)


def build_email_actions(details: dict[str, Any]) -> list[dict[str, Any]]:
    to_list = details.get("to", [])

    compose_url = _build_gmail_compose_url(
        to=to_list,
        subject=details.get("subject", ""),
        body=details.get("body", ""),
        cc=details.get("cc", []),
        bcc=details.get("bcc", []),
    )

    actions: list[dict[str, Any]] = [
        {"action": "open_app", "app": details.get("browser", "chrome")},
        {"action": "wait", "seconds": 2},
        {"action": "press_key", "keys": ["ctrl", "l"]},
        {"action": "paste_text", "text": compose_url},
        {"action": "press_key", "keys": ["enter"]},
        {"action": "wait", "seconds": 4},
        {"action": "screenshot", "label": "gmail_compose_ready"},
    ]

    if details.get("send") and to_list:
        actions.extend(
            [
                {"action": "press_key", "keys": ["ctrl", "enter"]},
                {"action": "wait", "seconds": 1},
                {"action": "press_key", "keys": ["enter"]},
                {"action": "screenshot", "label": "gmail_after_send"},
            ]
        )

    return actions


def validate_email_actions(actions: list[dict[str, Any]]) -> tuple[bool, str]:
    if not isinstance(actions, list) or not actions:
        return False, "No email actions were produced."
    required = {"open_app", "wait", "press_key", "paste_text"}
    names = {item.get("action") for item in actions if isinstance(item, dict)}
    missing = sorted(required - names)
    if missing:
        return False, f"Email action plan is missing required step(s): {', '.join(missing)}."
    return True, ""


def wants_model_polish(command: str) -> bool:
    text = (command or "").lower()
    return any(phrase in text for phrase in ("send to model", "polish", "professionalize", "professionalise"))


def _extract_email_details(command: str) -> dict[str, Any]:
    text = command.strip()
    lower = text.lower()

    to_list = _extract_emails_after_keyword(text, "to")
    explicit_cc = _extract_emails_before_target(text, "cc")
    explicit_bcc = _extract_emails_before_target(text, "bcc")
    cc_list = explicit_cc if explicit_cc else _extract_emails_after_keyword(text, "cc")
    bcc_list = explicit_bcc if explicit_bcc else _extract_emails_after_keyword(text, "bcc")

    cc_list = _dedupe(cc_list)
    bcc_list = _dedupe(bcc_list)
    cc_set = {item.lower() for item in cc_list}
    bcc_set = {item.lower() for item in bcc_list}

    to_list = [item for item in _dedupe(to_list) if item.lower() not in cc_set and item.lower() not in bcc_set]

    all_emails = re.findall(EMAIL_REGEX, text)
    if not to_list and all_emails:
        to_list = [email for email in all_emails if email.lower() not in cc_set and email.lower() not in bcc_set][:1]

    subject = _extract_tagged_value(text, "subject")
    body = _extract_tagged_value(text, "body") or _extract_tagged_value(text, "message")

    if not subject:
        about_match = re.search(r"\babout\s+(.+?)(?=\s+(?:with|saying|that|to|and\s+send|send\b)|$)", text, re.IGNORECASE)
        if about_match:
            subject = about_match.group(1).strip(" .")

    if not body:
        saying_match = re.search(
            r"\b(?:saying|message\s+that|body\s+that|that)\s+(.+?)(?=\s+(?:and\s+send|send\s+it|send\s+now)\b|$)",
            text,
            re.IGNORECASE,
        )
        if saying_match:
            body = saying_match.group(1).strip()

    if not subject:
        subject = "Quick update"
    if not body:
        body = clean_email_intent_text(text)

    browser = "chrome"
    if "edge" in lower:
        browser = "edge"
    elif "firefox" in lower:
        browser = "firefox"

    send = _should_send(lower)

    return {
        "to": _dedupe(to_list),
        "cc": _dedupe(cc_list),
        "bcc": _dedupe(bcc_list),
        "subject": subject,
        "body": body,
        "browser": browser,
        "send": send,
        "model_polish": wants_model_polish(text),
        "raw_text": text,
    }


def clean_email_intent_text(text: str) -> str:
    cleaned = text
    cleaned = re.sub(r"\b(?:send\s+to\s+model|polish|professionalize|professionalise)\b", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b(?:open\s+\w+\s+go\s+to\s+gmail\s+and\s+compose\s+(?:a\s+)?mail)\b", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b(?:send|send\s+it|send\s+now|do\s+not\s+send|don't\s+send|dont\s+send)\b", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b(?:to|cc|bcc|subject|body|message)\b\s*[:=]?\s*[^\n]+", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    return cleaned


def _extract_emails_after_keyword(text: str, keyword: str) -> list[str]:
    if keyword == "to":
        text = re.split(r"\s+\bcc\b|\s+\bbcc\b", text, maxsplit=1, flags=re.IGNORECASE)[0]
    pattern = rf"\b{keyword}\b\s*[:=]?\s*(.+?)(?=\s+(?:to|cc|bcc|subject|body|message)\b|\s+and\s+send\b|\s+send\b|$)"
    match = re.search(pattern, text, re.IGNORECASE)
    if not match:
        return []
    segment = match.group(1)
    return re.findall(EMAIL_REGEX, segment)


def _extract_emails_before_target(text: str, target: str) -> list[str]:
    pattern = rf"((?:{EMAIL_REGEX}\s*(?:,|\band\b)?\s*)+)\s+to\s+{target}\b"
    matches = re.findall(pattern, text, re.IGNORECASE)
    emails: list[str] = []
    for segment in matches:
        emails.extend(re.findall(EMAIL_REGEX, segment))
    return emails


def _extract_tagged_value(text: str, keyword: str) -> str:
    tail_stop = r"(?:to|cc|bcc|subject|body|message|and\s+send|send\s+it|send\s+now|send\s+email|do\s+not\s+send|don't\s+send|dont\s+send)"
    quoted = re.search(
        rf"\b{keyword}\b\s*[:=]?\s*[\"'](.+?)[\"'](?=\s+\b{tail_stop}\b|$)",
        text,
        re.IGNORECASE,
    )
    if quoted:
        return _strip_send_tail(quoted.group(1).strip())

    unquoted = re.search(
        rf"\b{keyword}\b\s*[:=]?\s*(.+?)(?=\s+\b{tail_stop}\b|$)",
        text,
        re.IGNORECASE,
    )
    return _strip_send_tail(unquoted.group(1).strip()) if unquoted else ""


def _should_send(lower_text: str) -> bool:
    if re.search(r"\b(do\s+not|don't|dont|without)\s+send\b", lower_text):
        return False
    return bool(re.search(r"\b(send|send\s+it|send\s+now|send\s+email)\b", lower_text))


def _strip_send_tail(text: str) -> str:
    cleaned = re.sub(
        r"\s*\b(?:and\s+send|send\s+it|send\s+now|send\s+email|send|do\s+not\s+send|don't\s+send|dont\s+send)\b\s*$",
        "",
        text,
        flags=re.IGNORECASE,
    )
    return cleaned.strip(" .")


def _build_gmail_compose_url(
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
) -> str:
    parts = [
        "view=cm",
        "fs=1",
        "tf=1",
        f"su={quote_plus(subject or '')}",
        f"body={quote_plus(body or '')}",
    ]

    if to:
        parts.append(f"to={quote_plus(','.join(to))}")

    if cc:
        parts.append(f"cc={quote_plus(','.join(cc))}")
    if bcc:
        parts.append(f"bcc={quote_plus(','.join(bcc))}")

    return f"https://mail.google.com/mail/?{'&'.join(parts)}"


def _dedupe(items: list[str]) -> list[str]:
    seen: set[str] = set()
    output: list[str] = []
    for item in items:
        key = item.lower().strip()
        if not key or key in seen:
            continue
        seen.add(key)
        output.append(item.strip())
    return output
