import datetime
import re
from typing import Any
from urllib.parse import quote_plus

EMAIL_REGEX = r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"
WINDOWS_PATH_REGEX = r"[A-Za-z]:\\(?:[^\\/:*?\"<>|\r\n]+\\)*[^\\/:*?\"<>|\r\n]+\.[A-Za-z0-9]{1,8}"

EMAIL_COMMAND_HINTS = (
    "email",
    "gmail",
    "mail",
    "inbox",
    "unread",
    "compose",
    "reply",
    "forward",
    "thread",
    "archive",
    "star",
    "label",
    "schedule",
    "template",
    "signature",
    "send to",
    "send to model",
    "polish",
)

COMPOSE_LIKE_INTENTS = {
    "compose_send",
    "schedule_send",
    "message_reply",
    "message_reply_all",
    "message_forward",
}


def is_email_command(command: str) -> bool:
    text = (command or "").lower()
    if any(hint in text for hint in EMAIL_COMMAND_HINTS):
        return True
    if "send to model" in text:
        return True
    if re.search(r"\b(send|compose|write)\s+(an\s+)?email\b", text):
        return True
    return bool(
        re.search(
            r"\b(list|search|read|reply|forward|archive|star|label|schedule)\b.*\b(email|mail|gmail|inbox|message|thread)\b",
            text,
        )
    )


def parse_email_actions_from_command(command: str) -> list[dict[str, Any]]:
    details = parse_email_request(command)
    return build_email_actions(details)


def parse_email_request(command: str) -> dict[str, Any]:
    return _extract_email_details(command)


def build_email_actions(details: dict[str, Any]) -> list[dict[str, Any]]:
    intent = details.get("intent", "compose_send")
    browser = details.get("browser", "chrome")

    if intent != "compose_send":
        landing_url = _gmail_url_for_intent(intent)
        return [
            {"action": "open_app", "app": browser},
            {"action": "wait", "seconds": 2},
            {"action": "open_url", "url": landing_url, "app": browser},
            {"action": "wait", "seconds": 3},
            {"action": "screenshot", "label": f"gmail_{intent}"},
        ]

    to_list = details.get("to", [])
    attachments = details.get("attachments", [])

    compose_url = _build_gmail_compose_url(
        to=to_list,
        subject=details.get("subject", ""),
        body=details.get("body", ""),
        cc=details.get("cc", []),
        bcc=details.get("bcc", []),
    )
    details["compose_url"] = compose_url

    actions: list[dict[str, Any]] = [
        {"action": "open_app", "app": browser},
        {"action": "wait", "seconds": 2},
        {"action": "open_url", "url": compose_url, "app": browser},
        {"action": "wait", "seconds": 4},
        {"action": "screenshot", "label": "gmail_compose_ready"},
    ]

    for index, attachment in enumerate(attachments, start=1):
        actions.extend(
            [
                {
                    "action": "attach_file",
                    "path": attachment,
                    "compose_url": compose_url,
                    "browser": browser,
                },
                {"action": "screenshot", "label": f"gmail_attachment_{index}"},
            ]
        )

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
    required = {"open_app", "wait", "open_url"}
    names = {item.get("action") for item in actions if isinstance(item, dict)}
    missing = sorted(required - names)
    if missing:
        return False, f"Email action plan is missing required step(s): {', '.join(missing)}."
    return True, ""


def wants_model_polish(command: str) -> bool:
    text = (command or "").lower()
    return any(phrase in text for phrase in ("send to model", "polish", "professionalize", "professionalise"))


def _extract_email_details(command: str) -> dict[str, Any]:
    text = (command or "").strip()
    lower = text.lower()
    intent = _detect_email_intent(lower)

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
    html_body = _extract_tagged_value(text, "html")
    attachments = _extract_attachment_paths(text)

    if not subject:
        about_match = re.search(
            r"\babout\s+(.+?)(?=\s+(?:with|saying|that|to|and\s+send|send\b)|$)",
            text,
            re.IGNORECASE,
        )
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

    if intent in COMPOSE_LIKE_INTENTS:
        if not subject:
            subject = "Quick update"
        if not body:
            body = clean_email_intent_text(text)
        body = _strip_attachment_clause(body)
    else:
        body = body or ""

    browser = "chrome"
    if "edge" in lower:
        browser = "edge"
    elif "firefox" in lower:
        browser = "firefox"

    send = _should_send(lower)
    if intent == "schedule_send":
        send = False

    query = _extract_search_query(text)
    max_results = _extract_max_results(lower)
    message_id = _extract_message_id(text)
    thread_id = _extract_thread_id(text)
    label_name = _extract_label_name(text)
    label_id = _extract_label_id(text)
    template_name = _extract_named_entity(text, "template")
    signature_name = _extract_named_entity(text, "signature")
    schedule_at = _extract_schedule_at(text)
    schedule_id = _extract_schedule_id(text)

    if intent == "inbox_list" and not query and "unread" in lower:
        query = "is:unread"

    return {
        "intent": intent,
        "to": _dedupe(to_list),
        "cc": _dedupe(cc_list),
        "bcc": _dedupe(bcc_list),
        "subject": subject,
        "body": body,
        "html_body": html_body,
        "browser": browser,
        "attachments": _dedupe(attachments),
        "send": send,
        "model_polish": wants_model_polish(text),
        "raw_text": text,
        "query": query,
        "max_results": max_results,
        "message_id": message_id,
        "thread_id": thread_id,
        "label_name": label_name,
        "label_id": label_id,
        "template_name": template_name,
        "signature_name": signature_name,
        "schedule_at": schedule_at,
        "schedule_id": schedule_id,
        "unread_only": "unread" in lower,
        "include_body": any(token in lower for token in ("include body", "full body", "with body", "full details")),
    }


def clean_email_intent_text(text: str) -> str:
    cleaned = text
    cleaned = re.sub(r"\b(?:send\s+to\s+model|polish|professionalize|professionalise)\b", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(
        r"\b(?:open\s+\w+\s+go\s+to\s+gmail\s+and\s+compose\s+(?:a\s+)?mail)\b",
        "",
        cleaned,
        flags=re.IGNORECASE,
    )
    cleaned = re.sub(
        r"\b(?:send|send\s+it|send\s+now|do\s+not\s+send|don't\s+send|dont\s+send|reply|forward)\b",
        "",
        cleaned,
        flags=re.IGNORECASE,
    )
    cleaned = _strip_attachment_clause(cleaned)
    cleaned = re.sub(WINDOWS_PATH_REGEX, "", cleaned)
    cleaned = re.sub(r"\b(?:to|cc|bcc|subject|body|message|html|query|template|signature|label)\b\s*[:=]?\s*[^\n]+", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    return cleaned


def _detect_email_intent(lower: str) -> str:
    if any(token in lower for token in ("list templates", "show templates", "templates list")):
        return "template_list"
    if re.search(r"\b(save|create|add)\s+template\b", lower):
        return "template_save"
    if re.search(r"\b(delete|remove)\s+template\b", lower):
        return "template_delete"

    if any(token in lower for token in ("list signatures", "show signatures", "signatures list")):
        return "signature_list"
    if re.search(r"\b(save|create|add)\s+signature\b", lower):
        return "signature_save"
    if re.search(r"\b(delete|remove)\s+signature\b", lower):
        return "signature_delete"

    if re.search(r"\b(cancel|delete|remove)\s+scheduled\b", lower):
        return "schedule_cancel"
    if re.search(r"\b(list|show)\s+scheduled\b", lower):
        return "schedule_list"
    if re.search(r"\b(process|run)\s+scheduled\b", lower):
        return "schedule_process_due"
    if "schedule" in lower and any(token in lower for token in ("email", "mail", "reply", "forward", "send")):
        return "schedule_send"

    if re.search(r"\b(list|show)\s+labels\b", lower):
        return "labels_list"
    if re.search(r"\b(create|add)\s+label\b", lower) and "to" not in lower:
        return "label_create"
    if re.search(r"\b(delete|remove)\s+label\b", lower) and "from" not in lower and "message" not in lower and "thread" not in lower:
        return "label_delete"
    if re.search(r"\b(remove|clear)\s+label\b", lower):
        return "label_remove"
    if re.search(r"\b(add|apply)\s+label\b", lower):
        return "label_apply"

    if "mark unread" in lower:
        return "message_mark_unread"
    if "mark read" in lower:
        return "message_mark_read"
    if "unarchive" in lower:
        return "message_unarchive"
    if "archive" in lower:
        return "message_archive"
    if "unstar" in lower:
        return "message_unstar"
    if re.search(r"\bstar\b", lower):
        return "message_star"

    if "reply all" in lower:
        return "message_reply_all"
    if re.search(r"\breply\b", lower):
        return "message_reply"
    if re.search(r"\bforward\b", lower):
        return "message_forward"

    if "thread" in lower and any(token in lower for token in ("show", "view", "read", "open", "get", "details")):
        return "thread_get"

    if any(token in lower for token in ("read message", "show message", "view message", "open message", "message details")):
        return "message_get"

    if any(token in lower for token in ("search inbox", "search email", "search emails", "find emails", "find messages", "query inbox")):
        return "inbox_search"

    if any(token in lower for token in ("list inbox", "show inbox", "list emails", "list messages", "show unread", "unread emails", "unread messages")):
        return "inbox_list"

    return "compose_send"


def _extract_emails_after_keyword(text: str, keyword: str) -> list[str]:
    if keyword == "to":
        text = re.split(r"\s+\bcc\b|\s+\bbcc\b", text, maxsplit=1, flags=re.IGNORECASE)[0]
    pattern = rf"\b{keyword}\b\s*[:=]?\s*(.+?)(?=\s+(?:to|cc|bcc|subject|body|message|html|query|template|signature|label)\b|\s+and\s+send\b|\s+send\b|$)"
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
    tail_stop = (
        r"(?:to|cc|bcc|subject|body|message|html|query|label|template|signature|"
        r"and\s+send|send\s+it|send\s+now|send\s+email|"
        r"do\s+not\s+send|don't\s+send|dont\s+send|send\s+to\s+model|polish|professionalize|professionalise|"
        r"schedule|message\s+id|thread\s+id|limit|max|top|attach(?:ed)?|attachment)"
    )
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


def _extract_search_query(text: str) -> str:
    quoted = re.search(r"\b(?:query|search)\b\s*[:=]?\s*[\"'](.+?)[\"']", text, re.IGNORECASE)
    if quoted:
        return quoted.group(1).strip()

    patterns = (
        r"\bsearch\s+(?:inbox|emails?|messages?)\s*(?:for|about|with)?\s+(.+?)(?=\s+\b(?:top|last|latest|limit|max|label|and\b|from\b|to\b)\b|$)",
        r"\bfind\s+(?:emails?|messages?)\s*(?:for|about|with)?\s+(.+?)(?=\s+\b(?:top|last|latest|limit|max|label|and\b|from\b|to\b)\b|$)",
    )
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip(" .")

    return ""


def _extract_message_id(text: str) -> str:
    patterns = (
        r"\bmessage\s+id\b\s*[:=]?\s*([A-Za-z0-9_-]{6,})",
        r"\bmsg\s+id\b\s*[:=]?\s*([A-Za-z0-9_-]{6,})",
        r"\bmessage\b\s*[:=]\s*([A-Za-z0-9_-]{6,})",
    )
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1)
    return ""


def _extract_thread_id(text: str) -> str:
    patterns = (
        r"\bthread\s+id\b\s*[:=]?\s*([A-Za-z0-9_-]{6,})",
        r"\bconversation\s+id\b\s*[:=]?\s*([A-Za-z0-9_-]{6,})",
        r"\bthread\b\s*[:=]\s*([A-Za-z0-9_-]{6,})",
    )
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1)
    return ""


def _extract_label_name(text: str) -> str:
    quoted = re.search(r"\blabel\b\s*[:=]?\s*[\"'](.+?)[\"']", text, re.IGNORECASE)
    if quoted:
        return quoted.group(1).strip()

    pattern = (
        r"\b(?:create|add|delete|remove|apply|use)\s+label\b\s+(.+?)"
        r"(?=\s+\b(?:to|from|for|on|message|thread|id|and|send|with)\b|$)"
    )
    match = re.search(pattern, text, re.IGNORECASE)
    if match:
        return match.group(1).strip(" .")
    return ""


def _extract_label_id(text: str) -> str:
    match = re.search(r"\blabel\s+id\b\s*[:=]?\s*([A-Za-z0-9_-]{3,})", text, re.IGNORECASE)
    return match.group(1) if match else ""


def _extract_named_entity(text: str, keyword: str) -> str:
    quoted = re.search(rf"\b{keyword}\b(?:\s+name)?\s*[:=]?\s*[\"'](.+?)[\"']", text, re.IGNORECASE)
    if quoted:
        return quoted.group(1).strip()

    op_pattern = (
        rf"\b(?:save|create|add|use|delete|remove)\s+{keyword}\b\s+(.+?)"
        r"(?=\s+\b(?:to|for|with|subject|body|html|message|send|at|on|in|from|and)\b|$)"
    )
    match = re.search(op_pattern, text, re.IGNORECASE)
    if match:
        return match.group(1).strip(" .")
    return ""


def _extract_schedule_at(text: str) -> str:
    now = datetime.datetime.now()

    relative = re.search(r"\bin\s+(\d{1,4})\s+(minute|minutes|hour|hours|day|days)\b", text, re.IGNORECASE)
    if relative:
        amount = int(relative.group(1))
        unit = relative.group(2).lower()
        if "minute" in unit:
            target = now + datetime.timedelta(minutes=amount)
        elif "hour" in unit:
            target = now + datetime.timedelta(hours=amount)
        else:
            target = now + datetime.timedelta(days=amount)
        return target.replace(second=0, microsecond=0).isoformat()

    iso_like = re.search(r"\b(\d{4}-\d{2}-\d{2}[ T]\d{1,2}:\d{2}(?::\d{2})?)\b", text)
    if iso_like:
        value = iso_like.group(1).replace(" ", "T")
        try:
            dt = datetime.datetime.fromisoformat(value)
            return dt.replace(microsecond=0).isoformat()
        except ValueError:
            pass

    date_and_time = re.search(r"\bon\s+(\d{4}-\d{2}-\d{2})\s+at\s+(\d{1,2}:\d{2})\b", text, re.IGNORECASE)
    if date_and_time:
        try:
            dt = datetime.datetime.fromisoformat(f"{date_and_time.group(1)}T{date_and_time.group(2)}")
            return dt.replace(second=0, microsecond=0).isoformat()
        except ValueError:
            pass

    bare_date_time = re.search(r"\b(\d{4}-\d{2}-\d{2})\s+(\d{1,2}:\d{2})\b", text)
    if bare_date_time:
        try:
            dt = datetime.datetime.fromisoformat(f"{bare_date_time.group(1)}T{bare_date_time.group(2)}")
            return dt.replace(second=0, microsecond=0).isoformat()
        except ValueError:
            pass

    time_only = re.search(r"\bat\s+(\d{1,2}:\d{2})\b", text, re.IGNORECASE)
    if time_only:
        hour, minute = [int(part) for part in time_only.group(1).split(":", 1)]
        try:
            target = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        except ValueError:
            return ""
        if target <= now:
            target = target + datetime.timedelta(days=1)
        return target.isoformat()

    return ""


def _extract_schedule_id(text: str) -> str:
    patterns = (
        r"\bschedule\s+id\b\s*[:=]?\s*([A-Za-z0-9_-]{6,})",
        r"\bscheduled\s+id\b\s*[:=]?\s*([A-Za-z0-9_-]{6,})",
        r"\bjob\s+id\b\s*[:=]?\s*([A-Za-z0-9_-]{6,})",
    )
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1)
    return ""


def _extract_max_results(lower_text: str) -> int:
    match = re.search(r"\b(?:top|last|latest|limit|max)\s+(\d{1,3})\b", lower_text)
    if not match:
        return 10
    value = int(match.group(1))
    return max(1, min(50, value))


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


def _extract_attachment_paths(text: str) -> list[str]:
    if not text:
        return []

    segments = []
    keyword_matches = re.findall(
        r"\b(?:attach(?:ed)?|attachment|attach\s+the\s+file|attach\s+file|file\s+from)\b(.+?)(?=\s+\b(?:to|cc|bcc|subject|body|message|html|query|and\s+send|send\b|send\s+to\s+model|polish|professionalize|professionalise)\b|$)",
        text,
        re.IGNORECASE,
    )
    segments.extend(keyword_matches)
    if not segments and re.search(r"\b(?:attach|attachment)\b", text, re.IGNORECASE):
        segments.append(text)

    found: list[str] = []
    for segment in segments:
        found.extend(re.findall(WINDOWS_PATH_REGEX, segment))

    quoted = re.findall(r"[\"']([A-Za-z]:\\[^\"']+\.[A-Za-z0-9]{1,8})[\"']", text)
    found.extend(quoted)
    return _dedupe([path.replace("\\\\", "\\").strip() for path in found if path.strip()])


def _strip_attachment_clause(text: str) -> str:
    if not text:
        return text
    cleaned = re.sub(
        r"\b(?:attach(?:ed)?|attachment|attach\s+the\s+file|attach\s+file|file\s+from)\b.+?(?=\s+\b(?:to|cc|bcc|subject|body|message|html|query|and\s+send|send\b|send\s+to\s+model|polish|professionalize|professionalise)\b|$)",
        "",
        text,
        flags=re.IGNORECASE,
    )
    return re.sub(r"\s+", " ", cleaned).strip(" .")


def _gmail_url_for_intent(intent: str) -> str:
    if intent == "labels_list":
        return "https://mail.google.com/mail/u/0/#settings/labels"
    if intent in {"template_list", "template_save", "template_delete", "signature_list", "signature_save", "signature_delete"}:
        return "https://mail.google.com/mail/u/0/#settings/general"
    return "https://mail.google.com/mail/u/0/#inbox"


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
