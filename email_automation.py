import re
from typing import Any
from urllib.parse import quote_plus


EMAIL_COMMAND_HINTS = (
    "email",
    "gmail",
    "mail",
    "compose",
    "send to",
    "send to model",
    "polish",
)

AUTO_DRAFT_HINTS = (
    "by yourself",
    "on your own",
    "choose the subject",
    "select the subject",
    "decide the subject",
    "select the body",
    "main body by yourself",
    "choose subject",
    "choose body",
    "make it interesting",
    "mail should be interesting",
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
        {"action": "open_app", "app": details.get("browser", "chrome"), "url": compose_url},
        {"action": "wait", "seconds": 4},
        {"action": "screenshot", "label": "gmail_compose_ready"},
    ]

    if details.get("send") and to_list:
        actions.extend(
            [
                # FIX: Use Ctrl+Enter keyboard shortcut to send — no OCR/Tesseract needed.
                # Gmail compose accepts Ctrl+Enter as the Send shortcut in all browsers.
                {"action": "press_key", "keys": ["ctrl", "enter"]},
                {"action": "wait", "seconds": 2},
                {"action": "screenshot", "label": "gmail_after_send"},
            ]
        )

    return actions


def validate_email_actions(actions: list[dict[str, Any]]) -> tuple[bool, str]:
    if not isinstance(actions, list) or not actions:
        return False, "No email actions were produced."
    required = {"open_app", "wait", "screenshot"}
    names = {item.get("action") for item in actions if isinstance(item, dict)}
    missing = sorted(required - names)
    if missing:
        return False, f"Email action plan is missing required step(s): {', '.join(missing)}."
    return True, ""


def wants_model_polish(command: str) -> bool:
    text = (command or "").lower()
    if any(phrase in text for phrase in ("send to model", "polish", "professionalize", "professionalise")):
        return True
    return any(phrase in text for phrase in AUTO_DRAFT_HINTS)


def _extract_email_details(command: str) -> dict[str, Any]:
    text = command.strip()
    lower = text.lower()

    to_list = _extract_emails_after_keyword(text, "to")
    cc_list = _extract_emails_after_keyword(text, "cc")
    bcc_list = _extract_emails_after_keyword(text, "bcc")

    all_emails = re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", text)
    if not to_list and all_emails:
        to_list = [all_emails[0]]

    subject = _extract_tagged_value(text, "subject")
    body = _extract_tagged_value(text, "body") or _extract_tagged_value(text, "message")

    if not subject:
        about_match = re.search(
            r"\babout\s+(.+?)(?=\s+\b(?:with|saying|that|to|and\s+send|send|select|choose|decide|subject|body|mail|email)\b|$)",
            text,
            re.IGNORECASE,
        )
        if about_match:
            subject = about_match.group(1).strip(" .")

    if _looks_like_subject_instruction(subject):
        subject = ""

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
    if "brave" in lower:
        browser = "brave"
    elif "edge" in lower:
        browser = "edge"
    elif "firefox" in lower:
        browser = "firefox"

    send = _should_send(lower, has_recipient=bool(to_list))

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
    cleaned = re.sub(r"\b(?:select|choose|decide)\s+(?:the\s+)?(?:subject|body|main\s+body)\s+(?:by\s+yourself|on\s+your\s+own)?\b", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b(?:mail\s+should\s+be\s+interesting|make\s+it\s+interesting)\b", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\band\s+main\b", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", "", cleaned)
    cleaned = re.sub(r"\b(?:open|launch|start|go\s+to)\s+(?:gmail|mail|chrome|edge|brave|firefox)\b", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b(?:mail|email|gmail|compose)\b", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b(?:send|send\s+it|send\s+now|do\s+not\s+send|don't\s+send|dont\s+send)\b", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b(?:to|cc|bcc|subject|body|message)\b\s*[:=]?\s*[^\n]+", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b(?:something|anything)\s+about\b", "about", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    return cleaned


def _extract_emails_after_keyword(text: str, keyword: str) -> list[str]:
    pattern = rf"\b{keyword}\b\s*[:=]?\s*(.+?)(?=\s+\b(?:to|cc|bcc|subject|body|message|and\s+send|send\b)\b|$)"
    match = re.search(pattern, text, re.IGNORECASE)
    if not match:
        return []
    segment = match.group(1)
    return re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", segment)


def _extract_tagged_value(text: str, keyword: str) -> str:
    tail_stop = r"(?:to|cc|bcc|subject|body|message|and\s+send|send\s+it|send\s+now|send\s+email|do\s+not\s+send|don't\s+send|dont\s+send)"
    quoted = re.search(
        rf"\b{keyword}\b\s*[:=]?\s*[\"'](.+?)[\"'](?=\s+\b{tail_stop}\b|$)",
        text,
        re.IGNORECASE,
    )
    if quoted:
        return _strip_send_tail(quoted.group(1).strip())

    explicit = re.search(
        rf"\b{keyword}\b\s*[:=]\s*(.+?)(?=\s+\b{tail_stop}\b|$)",
        text,
        re.IGNORECASE,
    )
    if explicit:
        return _strip_send_tail(explicit.group(1).strip())

    with_is = re.search(
        rf"\b{keyword}\b\s+is\s+(.+?)(?=\s+\b{tail_stop}\b|$)",
        text,
        re.IGNORECASE,
    )
    return _strip_send_tail(with_is.group(1).strip()) if with_is else ""


def _should_send(lower_text: str, has_recipient: bool = False) -> bool:
    if re.search(r"\b(do\s+not|don't|dont|without)\s+send\b", lower_text):
        return False
    if re.search(r"\b(send|send\s+it|send\s+now|send\s+email)\b", lower_text):
        return True
    if not has_recipient:
        return False
    if re.search(r"\b(draft|as\s+draft|compose\s+draft)\b", lower_text):
        return False
    return bool(re.search(r"\b(mail|email|write\s+(an\s+)?email)\b", lower_text))


def _looks_like_subject_instruction(subject: str) -> bool:
    value = (subject or "").strip().lower()
    if not value:
        return False
    if value in {"and", "main", "main body", "body", "subject"}:
        return True
    return bool(re.fullmatch(r"(?:and|the|main)\s+\w+", value))


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
