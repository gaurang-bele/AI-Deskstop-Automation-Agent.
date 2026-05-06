# ============================================================
#  gmail_api.py - Gmail API (OAuth) operations
# ============================================================

from __future__ import annotations

import base64
import datetime
import json
import mimetypes
import os
import re
import uuid
from email.message import EmailMessage
from email.utils import getaddresses
from typing import Any

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.compose",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.labels",
]

LOCAL_ONLY_INTENTS = {
    "template_save",
    "template_list",
    "template_delete",
    "signature_save",
    "signature_list",
    "signature_delete",
    "schedule_send",
    "schedule_list",
    "schedule_cancel",
}

INTENT_ACTION_MAP = {
    "compose_send": "gmail_send_or_draft",
    "inbox_list": "gmail_inbox_list",
    "inbox_search": "gmail_inbox_search",
    "message_get": "gmail_message_get",
    "thread_get": "gmail_thread_get",
    "message_reply": "gmail_reply",
    "message_reply_all": "gmail_reply_all",
    "message_forward": "gmail_forward",
    "labels_list": "gmail_labels_list",
    "label_create": "gmail_label_create",
    "label_delete": "gmail_label_delete",
    "label_apply": "gmail_label_apply",
    "label_remove": "gmail_label_remove",
    "message_mark_read": "gmail_mark_read",
    "message_mark_unread": "gmail_mark_unread",
    "message_archive": "gmail_archive",
    "message_unarchive": "gmail_unarchive",
    "message_star": "gmail_star",
    "message_unstar": "gmail_unstar",
    "template_save": "gmail_template_save",
    "template_list": "gmail_template_list",
    "template_delete": "gmail_template_delete",
    "signature_save": "gmail_signature_save",
    "signature_list": "gmail_signature_list",
    "signature_delete": "gmail_signature_delete",
    "schedule_send": "gmail_schedule_create",
    "schedule_list": "gmail_schedule_list",
    "schedule_cancel": "gmail_schedule_cancel",
    "schedule_process_due": "gmail_schedule_process_due",
}

BUILTIN_LABELS = {
    "INBOX",
    "SPAM",
    "TRASH",
    "UNREAD",
    "STARRED",
    "IMPORTANT",
    "CATEGORY_PERSONAL",
    "CATEGORY_SOCIAL",
    "CATEGORY_PROMOTIONS",
    "CATEGORY_UPDATES",
    "CATEGORY_FORUMS",
    "SENT",
    "DRAFT",
}


def _templates_path() -> str:
    value = os.getenv("GMAIL_TEMPLATES_FILE", "gmail_templates.json").strip()
    return value or "gmail_templates.json"


def _signatures_path() -> str:
    value = os.getenv("GMAIL_SIGNATURES_FILE", "gmail_signatures.json").strip()
    return value or "gmail_signatures.json"


def _schedule_path() -> str:
    value = os.getenv("GMAIL_SCHEDULE_FILE", "gmail_scheduled_queue.json").strip()
    return value or "gmail_scheduled_queue.json"


def _load_json_file(path: str, default: Any) -> Any:
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as handle:
            return json.load(handle)
    except Exception:
        return default


def _save_json_file(path: str, payload: Any) -> None:
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2, ensure_ascii=False)


def _load_template_store() -> dict[str, Any]:
    payload = _load_json_file(_templates_path(), {"templates": {}})
    if not isinstance(payload, dict):
        return {"templates": {}}
    templates = payload.get("templates")
    if not isinstance(templates, dict):
        return {"templates": {}}
    return {"templates": templates}


def _save_template_store(store: dict[str, Any]) -> None:
    _save_json_file(_templates_path(), store)


def _load_signature_store() -> dict[str, Any]:
    payload = _load_json_file(_signatures_path(), {"signatures": {}})
    if not isinstance(payload, dict):
        return {"signatures": {}}
    signatures = payload.get("signatures")
    if not isinstance(signatures, dict):
        return {"signatures": {}}
    return {"signatures": signatures}


def _save_signature_store(store: dict[str, Any]) -> None:
    _save_json_file(_signatures_path(), store)


def _load_schedule_store() -> dict[str, Any]:
    payload = _load_json_file(_schedule_path(), {"items": []})
    if not isinstance(payload, dict):
        return {"items": []}
    items = payload.get("items")
    if not isinstance(items, list):
        return {"items": []}
    return {"items": items}


def _save_schedule_store(store: dict[str, Any]) -> None:
    _save_json_file(_schedule_path(), store)


def _normalize_name(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "").strip()).lower()


def _credentials_paths() -> tuple[str, str]:
    credentials_path = os.getenv("GMAIL_API_CREDENTIALS", "gmail_credentials.json").strip() or "gmail_credentials.json"
    token_path = os.getenv("GMAIL_API_TOKEN", "gmail_token.json").strip() or "gmail_token.json"
    return credentials_path, token_path


def _credentials_have_required_scopes(creds: Credentials) -> bool:
    existing = set(creds.scopes or [])
    if not existing:
        return False
    return set(SCOPES).issubset(existing)


def _load_credentials() -> Credentials:
    credentials_path, token_path = _credentials_paths()
    creds: Credentials | None = None

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
        if creds and not _credentials_have_required_scopes(creds):
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(credentials_path):
                raise FileNotFoundError(f"Gmail API credentials file not found: {credentials_path}")
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, "w", encoding="utf-8") as token:
            token.write(creds.to_json())

    return creds


def _record_step(
    steps: list[dict[str, Any]],
    action: str,
    success: bool,
    note: str,
    data: Any | None = None,
) -> None:
    row: dict[str, Any] = {
        "step": len(steps) + 1,
        "action": action,
        "success": success,
        "note": note,
    }
    if data is not None:
        row["data"] = data
    steps.append(row)


def _html_to_text(value: str) -> str:
    if not value:
        return ""
    text = re.sub(r"<style[\s\S]*?</style>", "", value, flags=re.IGNORECASE)
    text = re.sub(r"<script[\s\S]*?</script>", "", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _attach_files(message: EmailMessage, details: dict[str, Any]) -> None:
    for attachment in details.get("attachments", []) or []:
        path = str(attachment).strip()
        if not path:
            continue
        if not os.path.exists(path):
            raise FileNotFoundError(f"Attachment file not found: {path}")

        mime_type, _ = mimetypes.guess_type(path)
        if mime_type and "/" in mime_type:
            maintype, subtype = mime_type.split("/", 1)
        else:
            maintype, subtype = "application", "octet-stream"

        with open(path, "rb") as file_handle:
            data = file_handle.read()

        message.add_attachment(
            data,
            maintype=maintype,
            subtype=subtype,
            filename=os.path.basename(path),
        )


def _build_message(details: dict[str, Any]) -> EmailMessage:
    to_list = details.get("to", []) or []
    if not to_list:
        raise ValueError("Recipient email missing for Gmail API send.")

    message = EmailMessage()
    message["To"] = ", ".join(to_list)

    cc_list = details.get("cc", []) or []
    if cc_list:
        message["Cc"] = ", ".join(cc_list)

    bcc_list = details.get("bcc", []) or []
    if bcc_list:
        message["Bcc"] = ", ".join(bcc_list)

    subject = details.get("subject", "") or "Quick update"
    body = details.get("body", "") or ""
    html_body = details.get("html_body", "") or ""
    message["Subject"] = subject

    if html_body:
        plain_fallback = body or _html_to_text(html_body) or "(HTML email)"
        message.set_content(plain_fallback)
        message.add_alternative(html_body, subtype="html")
    else:
        message.set_content(body)

    _attach_files(message, details)
    return message


def _encode_message(message: EmailMessage) -> str:
    return base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")


def _send_or_draft_encoded(
    service: Any,
    encoded_message: str,
    send_now: bool,
    thread_id: str | None = None,
) -> tuple[str, str]:
    if send_now:
        payload: dict[str, Any] = {"raw": encoded_message}
        if thread_id:
            payload["threadId"] = thread_id
        response = service.users().messages().send(userId="me", body=payload).execute()
        return "send", response.get("id", "")

    draft_payload: dict[str, Any] = {"raw": encoded_message}
    if thread_id:
        draft_payload["threadId"] = thread_id
    response = service.users().drafts().create(userId="me", body={"message": draft_payload}).execute()
    return "draft", response.get("id", "")


def _headers_to_dict(headers: list[dict[str, Any]] | None) -> dict[str, str]:
    output: dict[str, str] = {}
    for header in headers or []:
        name = str(header.get("name", "")).strip().lower()
        if not name:
            continue
        output[name] = str(header.get("value", "")).strip()
    return output


def _decode_part_data(value: str) -> str:
    if not value:
        return ""
    padding = "=" * ((4 - len(value) % 4) % 4)
    try:
        decoded = base64.urlsafe_b64decode((value + padding).encode("utf-8"))
        return decoded.decode("utf-8", errors="replace")
    except Exception:
        return ""


def _extract_payload_bodies(payload: dict[str, Any] | None) -> tuple[str, str]:
    plain_parts: list[str] = []
    html_parts: list[str] = []

    def walk(part: dict[str, Any]) -> None:
        mime_type = str(part.get("mimeType", "")).lower()
        body = part.get("body") if isinstance(part.get("body"), dict) else {}
        data = str(body.get("data", "")) if isinstance(body, dict) else ""

        if data and mime_type == "text/plain":
            plain_parts.append(_decode_part_data(data))
        elif data and mime_type == "text/html":
            html_parts.append(_decode_part_data(data))

        for child in part.get("parts", []) or []:
            if isinstance(child, dict):
                walk(child)

    if isinstance(payload, dict):
        walk(payload)

    return "\n".join([part for part in plain_parts if part.strip()]).strip(), "\n".join(
        [part for part in html_parts if part.strip()]
    ).strip()


def _extract_email_addresses(value: str) -> list[str]:
    addresses = [addr for _, addr in getaddresses([value or ""]) if addr]
    return _dedupe_emails(addresses)


def _dedupe_emails(items: list[str]) -> list[str]:
    seen: set[str] = set()
    output: list[str] = []
    for item in items:
        cleaned = str(item).strip()
        key = cleaned.lower()
        if not cleaned or key in seen:
            continue
        seen.add(key)
        output.append(cleaned)
    return output


def _summarize_message(service: Any, message_id: str) -> dict[str, Any]:
    item = service.users().messages().get(
        userId="me",
        id=message_id,
        format="metadata",
        metadataHeaders=["From", "To", "Subject", "Date"],
    ).execute()
    headers = _headers_to_dict(item.get("payload", {}).get("headers", []))
    return {
        "id": item.get("id", ""),
        "thread_id": item.get("threadId", ""),
        "from": headers.get("from", ""),
        "to": headers.get("to", ""),
        "subject": headers.get("subject", ""),
        "date": headers.get("date", ""),
        "snippet": item.get("snippet", ""),
        "label_ids": item.get("labelIds", []),
    }


def _find_label_id_by_name(service: Any, label_name: str) -> str:
    wanted = (label_name or "").strip().lower()
    if not wanted:
        return ""
    labels = service.users().labels().list(userId="me").execute().get("labels", [])
    for label in labels:
        if str(label.get("name", "")).strip().lower() == wanted:
            return str(label.get("id", "")).strip()
    return ""


def _resolve_label_id(service: Any, label_name: str, label_id: str) -> str:
    if (label_id or "").strip():
        return label_id.strip()

    name = (label_name or "").strip()
    if not name:
        raise ValueError("Label name or label id is required.")

    upper = name.upper()
    if upper in BUILTIN_LABELS:
        return upper

    found = _find_label_id_by_name(service, name)
    if found:
        return found

    raise ValueError(f"Label not found: {name}")


def _resolve_modify_target(details: dict[str, Any]) -> tuple[str, str]:
    message_id = str(details.get("message_id", "")).strip()
    thread_id = str(details.get("thread_id", "")).strip()

    if message_id:
        return "message", message_id
    if thread_id:
        return "thread", thread_id

    raise ValueError("A message_id or thread_id is required for this operation.")


def _modify_target_labels(
    service: Any,
    target_kind: str,
    target_id: str,
    add_label_ids: list[str] | None = None,
    remove_label_ids: list[str] | None = None,
) -> dict[str, Any]:
    payload = {
        "addLabelIds": add_label_ids or [],
        "removeLabelIds": remove_label_ids or [],
    }
    if target_kind == "message":
        return service.users().messages().modify(userId="me", id=target_id, body=payload).execute()
    return service.users().threads().modify(userId="me", id=target_id, body=payload).execute()


def _parse_datetime(value: str) -> datetime.datetime:
    if not value:
        raise ValueError("Missing datetime value.")
    try:
        return datetime.datetime.fromisoformat(value)
    except ValueError as exc:
        raise ValueError(
            "Invalid datetime format. Use ISO style, for example: 2026-05-01T10:30"
        ) from exc


def _parse_datetime_safe(value: str) -> datetime.datetime | None:
    if not value:
        return None
    try:
        return datetime.datetime.fromisoformat(value)
    except ValueError:
        return None


def _apply_template_and_signature(details: dict[str, Any]) -> dict[str, Any]:
    prepared = dict(details)

    template_name = str(prepared.get("template_name", "")).strip()
    if template_name:
        template_store = _load_template_store()
        template = template_store.get("templates", {}).get(_normalize_name(template_name))
        if not isinstance(template, dict):
            raise ValueError(f"Template not found: {template_name}")

        if not prepared.get("subject"):
            prepared["subject"] = template.get("subject", "") or "Quick update"
        if not prepared.get("body"):
            prepared["body"] = template.get("body", "")
        if not prepared.get("html_body"):
            prepared["html_body"] = template.get("html_body", "")

    signature_name = str(prepared.get("signature_name", "")).strip()
    if signature_name:
        signature_store = _load_signature_store()
        signature = signature_store.get("signatures", {}).get(_normalize_name(signature_name))
        if not isinstance(signature, dict):
            raise ValueError(f"Signature not found: {signature_name}")

        sig_text = str(signature.get("text", "")).strip()
        sig_html = str(signature.get("html", "")).strip()

        if sig_text:
            body = str(prepared.get("body", "")).rstrip()
            prepared["body"] = f"{body}\n\n{sig_text}".strip() if body else sig_text

        if sig_html:
            html_body = str(prepared.get("html_body", "")).strip()
            prepared["html_body"] = f"{html_body}<br><br>{sig_html}" if html_body else sig_html

    return prepared


def _build_schedulable_payload(details: dict[str, Any]) -> dict[str, Any]:
    payload = {
        "to": list(details.get("to") or []),
        "cc": list(details.get("cc") or []),
        "bcc": list(details.get("bcc") or []),
        "subject": str(details.get("subject", "") or "Quick update"),
        "body": str(details.get("body", "") or ""),
        "html_body": str(details.get("html_body", "") or ""),
        "attachments": list(details.get("attachments") or []),
        "template_name": str(details.get("template_name", "") or ""),
        "signature_name": str(details.get("signature_name", "") or ""),
        "thread_id": str(details.get("thread_id", "") or ""),
        "send": True,
    }
    return _apply_template_and_signature(payload)


def _handle_compose(service: Any, details: dict[str, Any], steps: list[dict[str, Any]]) -> None:
    prepared = _apply_template_and_signature(details)
    message = _build_message(prepared)
    encoded = _encode_message(message)
    mode, item_id = _send_or_draft_encoded(
        service,
        encoded,
        bool(prepared.get("send")),
        str(prepared.get("thread_id", "")).strip() or None,
    )

    if mode == "send":
        note = f"Message sent (id={item_id})" if item_id else "Message sent"
        _record_step(steps, "gmail_send", True, note)
    else:
        note = f"Draft created (id={item_id})" if item_id else "Draft created"
        _record_step(steps, "gmail_draft", True, note)


def _list_messages(service: Any, query: str, max_results: int) -> list[dict[str, Any]]:
    cleaned_query = query.strip()
    response = service.users().messages().list(
        userId="me",
        q=cleaned_query or None,
        maxResults=max(1, min(50, int(max_results or 10))),
    ).execute()

    output: list[dict[str, Any]] = []
    for row in response.get("messages", []) or []:
        message_id = str(row.get("id", "")).strip()
        if not message_id:
            continue
        output.append(_summarize_message(service, message_id))
    return output


def _handle_inbox_list(service: Any, details: dict[str, Any], steps: list[dict[str, Any]]) -> None:
    query = str(details.get("query", "")).strip()
    if not query and details.get("unread_only"):
        query = "is:unread"

    messages = _list_messages(service, query=query, max_results=int(details.get("max_results", 10) or 10))
    _record_step(
        steps,
        "gmail_inbox_list",
        True,
        f"Fetched {len(messages)} message(s).",
        {"messages": messages},
    )


def _handle_inbox_search(service: Any, details: dict[str, Any], steps: list[dict[str, Any]]) -> None:
    query = str(details.get("query", "")).strip()
    if not query:
        raise ValueError("Search query missing. Add query text, for example: search inbox query \"invoice\".")

    messages = _list_messages(service, query=query, max_results=int(details.get("max_results", 10) or 10))
    _record_step(
        steps,
        "gmail_inbox_search",
        True,
        f"Search returned {len(messages)} message(s).",
        {"query": query, "messages": messages},
    )


def _handle_message_get(service: Any, details: dict[str, Any], steps: list[dict[str, Any]]) -> None:
    message_id = str(details.get("message_id", "")).strip()
    if not message_id:
        raise ValueError("message_id is required for reading an email.")

    item = service.users().messages().get(userId="me", id=message_id, format="full").execute()
    headers = _headers_to_dict(item.get("payload", {}).get("headers", []))
    plain_body, html_body = _extract_payload_bodies(item.get("payload", {}))

    payload = {
        "id": item.get("id", ""),
        "thread_id": item.get("threadId", ""),
        "label_ids": item.get("labelIds", []),
        "from": headers.get("from", ""),
        "to": headers.get("to", ""),
        "subject": headers.get("subject", ""),
        "date": headers.get("date", ""),
        "snippet": item.get("snippet", ""),
    }

    if details.get("include_body"):
        payload["plain_body"] = plain_body
        payload["html_body"] = html_body

    _record_step(steps, "gmail_message_get", True, f"Fetched message {message_id}.", payload)


def _resolve_thread_id(service: Any, details: dict[str, Any]) -> str:
    thread_id = str(details.get("thread_id", "")).strip()
    if thread_id:
        return thread_id

    message_id = str(details.get("message_id", "")).strip()
    if not message_id:
        return ""

    item = service.users().messages().get(userId="me", id=message_id, format="minimal").execute()
    return str(item.get("threadId", "")).strip()


def _handle_thread_get(service: Any, details: dict[str, Any], steps: list[dict[str, Any]]) -> None:
    thread_id = _resolve_thread_id(service, details)
    if not thread_id:
        raise ValueError("thread_id (or message_id to resolve thread) is required.")

    thread = service.users().threads().get(
        userId="me",
        id=thread_id,
        format="metadata",
        metadataHeaders=["From", "To", "Subject", "Date"],
    ).execute()

    summarized: list[dict[str, Any]] = []
    for message in thread.get("messages", []) or []:
        headers = _headers_to_dict(message.get("payload", {}).get("headers", []))
        summarized.append(
            {
                "id": message.get("id", ""),
                "from": headers.get("from", ""),
                "to": headers.get("to", ""),
                "subject": headers.get("subject", ""),
                "date": headers.get("date", ""),
                "snippet": message.get("snippet", ""),
            }
        )

    _record_step(
        steps,
        "gmail_thread_get",
        True,
        f"Fetched thread {thread_id} with {len(summarized)} message(s).",
        {"thread_id": thread_id, "messages": summarized},
    )


def _resolve_source_message_for_response(service: Any, details: dict[str, Any]) -> dict[str, Any]:
    message_id = str(details.get("message_id", "")).strip()
    if message_id:
        return service.users().messages().get(userId="me", id=message_id, format="full").execute()

    thread_id = str(details.get("thread_id", "")).strip()
    if not thread_id:
        raise ValueError("message_id or thread_id is required.")

    thread = service.users().threads().get(userId="me", id=thread_id, format="full").execute()
    messages = thread.get("messages", []) or []
    if not messages:
        raise ValueError(f"No messages found in thread {thread_id}.")
    return messages[-1]


def _get_my_email(service: Any) -> str:
    profile = service.users().getProfile(userId="me").execute()
    return str(profile.get("emailAddress", "")).strip().lower()


def _handle_reply(
    service: Any,
    details: dict[str, Any],
    steps: list[dict[str, Any]],
    *,
    reply_all: bool,
) -> None:
    source = _resolve_source_message_for_response(service, details)
    headers = _headers_to_dict(source.get("payload", {}).get("headers", []))

    source_subject = headers.get("subject", "")
    subject = str(details.get("subject", "")).strip() or source_subject
    if not subject.lower().startswith("re:"):
        subject = f"Re: {subject or 'Quick update'}"

    my_email = _get_my_email(service)
    to_candidates = _extract_email_addresses(headers.get("reply-to", "") or headers.get("from", ""))
    cc_candidates: list[str] = []

    if reply_all:
        cc_candidates.extend(_extract_email_addresses(headers.get("to", "")))
        cc_candidates.extend(_extract_email_addresses(headers.get("cc", "")))

    to_list = _dedupe_emails([item for item in to_candidates if item.lower() != my_email])
    cc_list = _dedupe_emails(
        [
            item
            for item in cc_candidates
            if item.lower() != my_email and item.lower() not in {value.lower() for value in to_list}
        ]
    )

    if not to_list:
        raise ValueError("Could not resolve recipient for reply. Provide message_id from a thread with valid headers.")

    body = str(details.get("body", "")).strip() or "Thanks for the update."

    reply_details = dict(details)
    reply_details["to"] = to_list
    reply_details["cc"] = cc_list
    reply_details["bcc"] = []
    reply_details["subject"] = subject
    reply_details["body"] = body

    prepared = _apply_template_and_signature(reply_details)
    message = _build_message(prepared)

    message_id_header = headers.get("message-id", "")
    references_header = headers.get("references", "")
    if message_id_header:
        message["In-Reply-To"] = message_id_header
    references_value = " ".join(part for part in (references_header, message_id_header) if part).strip()
    if references_value:
        message["References"] = references_value

    encoded = _encode_message(message)
    mode, item_id = _send_or_draft_encoded(
        service,
        encoded,
        bool(prepared.get("send")),
        str(source.get("threadId", "")).strip() or None,
    )

    if mode == "send":
        action = "gmail_reply_all_send" if reply_all else "gmail_reply_send"
        note = f"Reply sent (id={item_id})" if item_id else "Reply sent"
    else:
        action = "gmail_reply_all_draft" if reply_all else "gmail_reply_draft"
        note = f"Reply draft created (id={item_id})" if item_id else "Reply draft created"

    _record_step(steps, action, True, note)


def _handle_forward(service: Any, details: dict[str, Any], steps: list[dict[str, Any]]) -> None:
    source = _resolve_source_message_for_response(service, details)
    headers = _headers_to_dict(source.get("payload", {}).get("headers", []))
    plain_body, _ = _extract_payload_bodies(source.get("payload", {}))

    to_list = list(details.get("to") or [])
    if not to_list:
        raise ValueError("Forward requires recipients in 'to'.")

    source_subject = headers.get("subject", "") or "Message"
    subject = str(details.get("subject", "")).strip() or source_subject
    if not subject.lower().startswith("fwd:"):
        subject = f"Fwd: {subject}"

    intro = str(details.get("body", "")).strip()
    forwarded_block = (
        "--- Forwarded message ---\n"
        f"From: {headers.get('from', '')}\n"
        f"Date: {headers.get('date', '')}\n"
        f"Subject: {headers.get('subject', '')}\n"
        f"To: {headers.get('to', '')}\n\n"
        f"{plain_body}"
    )

    combined_body = f"{intro}\n\n{forwarded_block}".strip() if intro else forwarded_block

    forward_details = dict(details)
    forward_details["subject"] = subject
    forward_details["body"] = combined_body
    forward_details["to"] = to_list

    prepared = _apply_template_and_signature(forward_details)
    message = _build_message(prepared)

    encoded = _encode_message(message)
    mode, item_id = _send_or_draft_encoded(service, encoded, bool(prepared.get("send")), None)

    if mode == "send":
        note = f"Forwarded email sent (id={item_id})" if item_id else "Forwarded email sent"
        _record_step(steps, "gmail_forward_send", True, note)
    else:
        note = f"Forward draft created (id={item_id})" if item_id else "Forward draft created"
        _record_step(steps, "gmail_forward_draft", True, note)


def _handle_labels_list(service: Any, steps: list[dict[str, Any]]) -> None:
    labels = service.users().labels().list(userId="me").execute().get("labels", [])
    mapped = [
        {
            "id": row.get("id", ""),
            "name": row.get("name", ""),
            "type": row.get("type", ""),
            "message_list_visibility": row.get("messageListVisibility", ""),
            "label_list_visibility": row.get("labelListVisibility", ""),
        }
        for row in labels
    ]
    _record_step(steps, "gmail_labels_list", True, f"Fetched {len(mapped)} label(s).", {"labels": mapped})


def _handle_label_create(service: Any, details: dict[str, Any], steps: list[dict[str, Any]]) -> None:
    label_name = str(details.get("label_name", "")).strip()
    if not label_name:
        raise ValueError("label_name is required to create a label.")

    existing = _find_label_id_by_name(service, label_name)
    if existing:
        _record_step(steps, "gmail_label_create", True, f"Label already exists (id={existing}).")
        return

    response = service.users().labels().create(
        userId="me",
        body={
            "name": label_name,
            "labelListVisibility": "labelShow",
            "messageListVisibility": "show",
        },
    ).execute()

    _record_step(
        steps,
        "gmail_label_create",
        True,
        f"Created label {label_name} (id={response.get('id', '')}).",
        {"label": response},
    )


def _handle_label_delete(service: Any, details: dict[str, Any], steps: list[dict[str, Any]]) -> None:
    label_id = _resolve_label_id(
        service,
        str(details.get("label_name", "")).strip(),
        str(details.get("label_id", "")).strip(),
    )
    if label_id in BUILTIN_LABELS:
        raise ValueError("Built-in labels cannot be deleted.")

    service.users().labels().delete(userId="me", id=label_id).execute()
    _record_step(steps, "gmail_label_delete", True, f"Deleted label {label_id}.")


def _handle_label_apply_remove(
    service: Any,
    details: dict[str, Any],
    steps: list[dict[str, Any]],
    *,
    remove: bool,
) -> None:
    label_id = _resolve_label_id(
        service,
        str(details.get("label_name", "")).strip(),
        str(details.get("label_id", "")).strip(),
    )
    target_kind, target_id = _resolve_modify_target(details)

    if remove:
        _modify_target_labels(service, target_kind, target_id, add_label_ids=[], remove_label_ids=[label_id])
        _record_step(steps, "gmail_label_remove", True, f"Removed label {label_id} from {target_kind} {target_id}.")
    else:
        _modify_target_labels(service, target_kind, target_id, add_label_ids=[label_id], remove_label_ids=[])
        _record_step(steps, "gmail_label_apply", True, f"Applied label {label_id} to {target_kind} {target_id}.")


def _handle_builtin_label_toggle(
    service: Any,
    details: dict[str, Any],
    steps: list[dict[str, Any]],
    *,
    action: str,
    add_label_ids: list[str] | None = None,
    remove_label_ids: list[str] | None = None,
) -> None:
    target_kind, target_id = _resolve_modify_target(details)
    _modify_target_labels(
        service,
        target_kind,
        target_id,
        add_label_ids=add_label_ids or [],
        remove_label_ids=remove_label_ids or [],
    )
    _record_step(steps, action, True, f"Updated {target_kind} {target_id}.")


def _handle_template_save(details: dict[str, Any], steps: list[dict[str, Any]]) -> None:
    template_name = str(details.get("template_name", "")).strip()
    if not template_name:
        raise ValueError("Template name missing. Example: save template followup ...")

    subject = str(details.get("subject", "") or "Quick update")
    body = str(details.get("body", "") or "")
    html_body = str(details.get("html_body", "") or "")

    if not body and not html_body:
        raise ValueError("Template body is empty. Provide body or html content.")

    store = _load_template_store()
    key = _normalize_name(template_name)
    store.setdefault("templates", {})[key] = {
        "name": template_name,
        "subject": subject,
        "body": body,
        "html_body": html_body,
        "updated_at": datetime.datetime.now().replace(microsecond=0).isoformat(),
    }
    _save_template_store(store)

    _record_step(steps, "gmail_template_save", True, f"Saved template '{template_name}'.")


def _handle_template_list(steps: list[dict[str, Any]]) -> None:
    store = _load_template_store()
    templates = store.get("templates", {})
    names = sorted([value.get("name", key) for key, value in templates.items() if isinstance(value, dict)])
    _record_step(steps, "gmail_template_list", True, f"Found {len(names)} template(s).", {"templates": names})


def _handle_template_delete(details: dict[str, Any], steps: list[dict[str, Any]]) -> None:
    template_name = str(details.get("template_name", "")).strip()
    if not template_name:
        raise ValueError("Template name missing for delete.")

    store = _load_template_store()
    key = _normalize_name(template_name)
    templates = store.setdefault("templates", {})
    if key not in templates:
        raise ValueError(f"Template not found: {template_name}")

    templates.pop(key, None)
    _save_template_store(store)
    _record_step(steps, "gmail_template_delete", True, f"Deleted template '{template_name}'.")


def _handle_signature_save(details: dict[str, Any], steps: list[dict[str, Any]]) -> None:
    signature_name = str(details.get("signature_name", "")).strip()
    if not signature_name:
        raise ValueError("Signature name missing. Example: save signature work body \"Regards...\"")

    text_value = str(details.get("body", "") or "")
    html_value = str(details.get("html_body", "") or "")
    if not text_value and not html_value:
        raise ValueError("Signature content is empty. Provide body or html content.")

    store = _load_signature_store()
    key = _normalize_name(signature_name)
    store.setdefault("signatures", {})[key] = {
        "name": signature_name,
        "text": text_value,
        "html": html_value,
        "updated_at": datetime.datetime.now().replace(microsecond=0).isoformat(),
    }
    _save_signature_store(store)

    _record_step(steps, "gmail_signature_save", True, f"Saved signature '{signature_name}'.")


def _handle_signature_list(steps: list[dict[str, Any]]) -> None:
    store = _load_signature_store()
    signatures = store.get("signatures", {})
    names = sorted([value.get("name", key) for key, value in signatures.items() if isinstance(value, dict)])
    _record_step(steps, "gmail_signature_list", True, f"Found {len(names)} signature(s).", {"signatures": names})


def _handle_signature_delete(details: dict[str, Any], steps: list[dict[str, Any]]) -> None:
    signature_name = str(details.get("signature_name", "")).strip()
    if not signature_name:
        raise ValueError("Signature name missing for delete.")

    store = _load_signature_store()
    key = _normalize_name(signature_name)
    signatures = store.setdefault("signatures", {})
    if key not in signatures:
        raise ValueError(f"Signature not found: {signature_name}")

    signatures.pop(key, None)
    _save_signature_store(store)
    _record_step(steps, "gmail_signature_delete", True, f"Deleted signature '{signature_name}'.")


def _handle_schedule_create(details: dict[str, Any], steps: list[dict[str, Any]]) -> None:
    schedule_at = str(details.get("schedule_at", "")).strip()
    if not schedule_at:
        raise ValueError(
            "Schedule time missing. Include a time, for example: schedule email ... at 2026-05-01T10:30"
        )

    send_at = _parse_datetime(schedule_at)
    now = datetime.datetime.now()
    if send_at <= now:
        raise ValueError("Scheduled time must be in the future.")

    payload = _build_schedulable_payload(details)

    job_id = uuid.uuid4().hex[:12]
    store = _load_schedule_store()
    items = store.setdefault("items", [])
    items.append(
        {
            "id": job_id,
            "send_at": send_at.replace(microsecond=0).isoformat(),
            "created_at": now.replace(microsecond=0).isoformat(),
            "details": payload,
        }
    )
    items.sort(key=lambda item: str(item.get("send_at", "")))
    _save_schedule_store(store)

    _record_step(
        steps,
        "gmail_schedule_create",
        True,
        f"Scheduled email {job_id} for {send_at.replace(microsecond=0).isoformat()}.",
        {"schedule_id": job_id, "send_at": send_at.replace(microsecond=0).isoformat()},
    )


def _handle_schedule_list(steps: list[dict[str, Any]]) -> None:
    store = _load_schedule_store()
    items = store.get("items", [])
    compact = []
    for item in items:
        compact.append(
            {
                "id": item.get("id", ""),
                "send_at": item.get("send_at", ""),
                "to": (item.get("details", {}) or {}).get("to", []),
                "subject": (item.get("details", {}) or {}).get("subject", ""),
            }
        )

    _record_step(steps, "gmail_schedule_list", True, f"Found {len(compact)} scheduled email(s).", {"items": compact})


def _handle_schedule_cancel(details: dict[str, Any], steps: list[dict[str, Any]]) -> None:
    schedule_id = str(details.get("schedule_id", "")).strip()
    if not schedule_id:
        raise ValueError("Schedule id is required to cancel. Example: cancel scheduled id abc123")

    store = _load_schedule_store()
    items = store.get("items", [])

    kept: list[dict[str, Any]] = []
    removed: list[dict[str, Any]] = []
    for item in items:
        item_id = str(item.get("id", ""))
        if item_id == schedule_id or item_id.startswith(schedule_id):
            removed.append(item)
        else:
            kept.append(item)

    if not removed:
        raise ValueError(f"No scheduled email found for id: {schedule_id}")

    store["items"] = kept
    _save_schedule_store(store)

    _record_step(
        steps,
        "gmail_schedule_cancel",
        True,
        f"Canceled {len(removed)} scheduled email(s).",
        {"canceled_ids": [item.get("id", "") for item in removed]},
    )


def _process_due_scheduled_messages(
    service: Any,
    steps: list[dict[str, Any]],
    *,
    record_when_empty: bool,
) -> int:
    store = _load_schedule_store()
    items = store.get("items", [])

    if not items:
        if record_when_empty:
            _record_step(steps, "gmail_schedule_process_due", True, "No scheduled emails are queued.")
        return 0

    now = datetime.datetime.now()
    sent_count = 0
    remaining: list[dict[str, Any]] = []

    for item in items:
        send_at = _parse_datetime_safe(str(item.get("send_at", "")))
        if send_at is None or send_at > now:
            remaining.append(item)
            continue

        payload = item.get("details", {}) if isinstance(item.get("details"), dict) else {}
        payload = dict(payload)
        payload["send"] = True

        try:
            message = _build_message(payload)
            encoded = _encode_message(message)
            _, message_id = _send_or_draft_encoded(
                service,
                encoded,
                True,
                str(payload.get("thread_id", "")).strip() or None,
            )
            sent_count += 1
            note = f"Sent scheduled email {item.get('id', '')} (id={message_id})" if message_id else f"Sent scheduled email {item.get('id', '')}"
            _record_step(steps, "gmail_schedule_send_due", True, note)
        except Exception as exc:
            item["last_error"] = str(exc)
            item["last_attempt_at"] = now.replace(microsecond=0).isoformat()
            remaining.append(item)
            _record_step(steps, "gmail_schedule_send_due", False, f"Failed to send scheduled email {item.get('id', '')}: {exc}")

    store["items"] = remaining
    _save_schedule_store(store)

    if record_when_empty and sent_count == 0:
        _record_step(steps, "gmail_schedule_process_due", True, "No scheduled emails were due at this time.")

    return sent_count


def send_gmail_with_api(details: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]], str | None]:
    """
    Return (planned_actions, step_results, error_message)
    """
    intent = str(details.get("intent", "compose_send") or "compose_send").strip()
    action_name = INTENT_ACTION_MAP.get(intent, f"gmail_{intent}")

    actions: list[dict[str, Any]] = [{"action": action_name}]
    steps: list[dict[str, Any]] = []

    try:
        local_only = intent in LOCAL_ONLY_INTENTS
        service = None

        if not local_only:
            creds = _load_credentials()
            actions.insert(0, {"action": "gmail_oauth"})
            _record_step(steps, "gmail_oauth", True, "OAuth token ready")
            service = build("gmail", "v1", credentials=creds)

            if intent not in {"schedule_process_due"}:
                _process_due_scheduled_messages(service, steps, record_when_empty=False)

        if intent == "compose_send":
            _handle_compose(service, details, steps)
        elif intent == "inbox_list":
            _handle_inbox_list(service, details, steps)
        elif intent == "inbox_search":
            _handle_inbox_search(service, details, steps)
        elif intent == "message_get":
            _handle_message_get(service, details, steps)
        elif intent == "thread_get":
            _handle_thread_get(service, details, steps)
        elif intent == "message_reply":
            _handle_reply(service, details, steps, reply_all=False)
        elif intent == "message_reply_all":
            _handle_reply(service, details, steps, reply_all=True)
        elif intent == "message_forward":
            _handle_forward(service, details, steps)
        elif intent == "labels_list":
            _handle_labels_list(service, steps)
        elif intent == "label_create":
            _handle_label_create(service, details, steps)
        elif intent == "label_delete":
            _handle_label_delete(service, details, steps)
        elif intent == "label_apply":
            _handle_label_apply_remove(service, details, steps, remove=False)
        elif intent == "label_remove":
            _handle_label_apply_remove(service, details, steps, remove=True)
        elif intent == "message_mark_read":
            _handle_builtin_label_toggle(service, details, steps, action="gmail_mark_read", remove_label_ids=["UNREAD"])
        elif intent == "message_mark_unread":
            _handle_builtin_label_toggle(service, details, steps, action="gmail_mark_unread", add_label_ids=["UNREAD"])
        elif intent == "message_archive":
            _handle_builtin_label_toggle(service, details, steps, action="gmail_archive", remove_label_ids=["INBOX"])
        elif intent == "message_unarchive":
            _handle_builtin_label_toggle(service, details, steps, action="gmail_unarchive", add_label_ids=["INBOX"])
        elif intent == "message_star":
            _handle_builtin_label_toggle(service, details, steps, action="gmail_star", add_label_ids=["STARRED"])
        elif intent == "message_unstar":
            _handle_builtin_label_toggle(service, details, steps, action="gmail_unstar", remove_label_ids=["STARRED"])
        elif intent == "template_save":
            _handle_template_save(details, steps)
        elif intent == "template_list":
            _handle_template_list(steps)
        elif intent == "template_delete":
            _handle_template_delete(details, steps)
        elif intent == "signature_save":
            _handle_signature_save(details, steps)
        elif intent == "signature_list":
            _handle_signature_list(steps)
        elif intent == "signature_delete":
            _handle_signature_delete(details, steps)
        elif intent == "schedule_send":
            _handle_schedule_create(details, steps)
        elif intent == "schedule_list":
            _handle_schedule_list(steps)
        elif intent == "schedule_cancel":
            _handle_schedule_cancel(details, steps)
        elif intent == "schedule_process_due":
            _process_due_scheduled_messages(service, steps, record_when_empty=True)
        else:
            raise ValueError(f"Unsupported Gmail intent: {intent}")

        return actions, steps, None
    except Exception as exc:
        _record_step(steps, "gmail_error", False, str(exc))
        return actions, steps, str(exc)
