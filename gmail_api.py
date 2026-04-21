# ============================================================
#  gmail_api.py — Gmail API (OAuth) email sender
# ============================================================

from __future__ import annotations

import base64
import mimetypes
import os
from email.message import EmailMessage
from typing import Any

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.compose",
]


def _credentials_paths() -> tuple[str, str]:
    credentials_path = os.getenv("GMAIL_API_CREDENTIALS", "gmail_credentials.json").strip()
    token_path = os.getenv("GMAIL_API_TOKEN", "gmail_token.json").strip()
    return credentials_path, token_path


def _load_credentials() -> Credentials:
    credentials_path, token_path = _credentials_paths()
    creds: Credentials | None = None

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(credentials_path):
                raise FileNotFoundError(
                    f"Gmail API credentials file not found: {credentials_path}"
                )
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, "w", encoding="utf-8") as token:
            token.write(creds.to_json())

    return creds


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
    message["Subject"] = subject
    message.set_content(body)

    for attachment in details.get("attachments", []) or []:
        path = attachment.strip()
        if not path:
            continue
        if not os.path.exists(path):
            raise FileNotFoundError(f"Attachment file not found: {path}")
        mime_type, _ = mimetypes.guess_type(path)
        if mime_type and "/" in mime_type:
            maintype, subtype = mime_type.split("/", 1)
        else:
            maintype, subtype = "application", "octet-stream"
        with open(path, "rb") as f:
            data = f.read()
        message.add_attachment(data, maintype=maintype, subtype=subtype, filename=os.path.basename(path))

    return message


def _encode_message(message: EmailMessage) -> str:
    return base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")


def send_gmail_with_api(details: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]], str | None]:
    """
    Return (planned_actions, step_results, error_message)
    """
    actions = [{"action": "gmail_oauth"}, {"action": "gmail_send" if details.get("send") else "gmail_draft"}]
    steps: list[dict[str, Any]] = []

    try:
        creds = _load_credentials()
        steps.append(
            {"step": len(steps) + 1, "action": "gmail_oauth", "success": True, "note": "OAuth token ready"}
        )
        service = build("gmail", "v1", credentials=creds)
        message = _build_message(details)
        encoded = _encode_message(message)

        if details.get("send"):
            response = service.users().messages().send(userId="me", body={"raw": encoded}).execute()
            message_id = response.get("id", "")
            steps.append(
                {
                    "step": len(steps) + 1,
                    "action": "gmail_send",
                    "success": True,
                    "note": f"Message sent (id={message_id})" if message_id else "Message sent",
                }
            )
        else:
            response = service.users().drafts().create(userId="me", body={"message": {"raw": encoded}}).execute()
            draft_id = response.get("id", "")
            steps.append(
                {
                    "step": len(steps) + 1,
                    "action": "gmail_draft",
                    "success": True,
                    "note": f"Draft created (id={draft_id})" if draft_id else "Draft created",
                }
            )

        return actions, steps, None
    except Exception as exc:
        steps.append(
            {
                "step": len(steps) + 1,
                "action": "gmail_error",
                "success": False,
                "note": str(exc),
            }
        )
        return actions, steps, str(exc)
