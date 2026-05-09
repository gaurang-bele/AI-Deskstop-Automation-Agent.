# ============================================================
#  playwright_email.py — DOM-based Gmail automation (Playwright)
# ============================================================

from __future__ import annotations

import logging
import os
import re
from pathlib import Path
from typing import Any

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


LOG_PATH = "browser_agent.log"
logger = logging.getLogger("browser_agent")
if not logger.handlers:
    logger.setLevel(logging.INFO)
    handler = logging.FileHandler(LOG_PATH, encoding="utf-8")
    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    handler.setFormatter(formatter)
    logger.addHandler(handler)


def _log_event(action: str, detail: str, level: str = "INFO") -> None:
    message = f"[{action}] {detail}"
    getattr(logger, level.lower(), logger.info)(message)


def _resolve_chrome_executable() -> str | None:
    candidates = []
    pf_x86 = os.environ.get("PROGRAMFILES(X86)")
    pf = os.environ.get("PROGRAMFILES")
    if pf_x86:
        candidates.append(os.path.join(pf_x86, "Google", "Chrome", "Application", "chrome.exe"))
    if pf:
        candidates.append(os.path.join(pf, "Google", "Chrome", "Application", "chrome.exe"))
    for path in candidates:
        if os.path.exists(path):
            return path
    return None


def _default_user_data_dir() -> str:
    local = os.environ.get("LOCALAPPDATA")
    if local:
        return os.path.join(local, "AiAgent", "PlaywrightProfile")
    return str(Path(".playwright-profile").resolve())


def _resolve_user_data_dir() -> str:
    override = os.environ.get("PLAYWRIGHT_USER_DATA_DIR", "").strip()
    if override:
        return override
    return _default_user_data_dir()


def _classify_browser_failure(text: str) -> str:
    lowered = (text or "").lower()
    if any(token in lowered for token in ("captcha", "recaptcha", "i'm not a robot")):
        return "CAPTCHA"
    if any(token in lowered for token in ("two-factor", "otp", "verification code", "authenticator")):
        return "OTP_MFA"
    if any(token in lowered for token in ("couldn't sign you in", "could not sign you in", "this browser or app may not be secure")):
        return "GOOGLE_SIGNIN_BLOCKED"
    if any(token in lowered for token in ("session expired", "sign in", "login", "log in")):
        return "SESSION_EXPIRED"
    if any(token in lowered for token in ("cookies", "consent", "privacy")):
        return "COOKIE_CONSENT"
    if any(token in lowered for token in ("automation", "bot detected", "unusual traffic")):
        return "AUTOMATION_DETECTED"
    return "UNKNOWN"


def _record_step(steps: list[dict[str, Any]], action: str, success: bool, note: str) -> None:
    steps.append(
        {
            "step": len(steps) + 1,
            "action": action,
            "success": success,
            "note": note,
        }
    )


def _friendly_playwright_error(exc: Exception) -> str:
    message = str(exc)
    lowered = message.lower()
    if "executable doesn't exist" in lowered or "playwright install" in lowered:
        return "Playwright browser binaries missing. Run: python -m playwright install"
    if "couldn't sign you in" in lowered or "this browser or app may not be secure" in lowered:
        return (
            "Google blocked sign-in in this browser. Open Gmail once in the dedicated Playwright "
            "profile, finish any login there manually, then rerun the command."
        )
    if "user data dir" in lowered or "singletonlock" in lowered or "profile" in lowered and "in use" in lowered:
        return (
            "Chrome profile appears to be in use. Close Chrome and retry, or set "
            "PLAYWRIGHT_USER_DATA_DIR to a separate folder."
        )
    return message


def _page_text_for_error(page) -> str:
    try:
        return page.inner_text("body")
    except Exception:
        try:
            return page.content()
        except Exception:
            return ""


def _launch_context(pw, user_data_dir: str, profile_dir: str, chrome_path: str | None):
    return pw.chromium.launch_persistent_context(
        user_data_dir=user_data_dir,
        headless=False,
        executable_path=chrome_path,
        args=[
            f"--profile-directory={profile_dir}",
            "--no-first-run",
            "--no-default-browser-check",
            "--start-maximized",
        ],
        viewport=None,
    )


def _ensure_selector(page, selectors: list[str], timeout: int = 15000):
    last_error = None
    for selector in selectors:
        try:
            return page.wait_for_selector(selector, timeout=timeout), selector
        except PlaywrightTimeoutError as exc:
            last_error = exc
            continue
    if last_error:
        raise last_error
    raise PlaywrightTimeoutError("No selectors provided")


def _fill_textarea(page, selectors: list[str], value: str) -> None:
    handle, selector = _ensure_selector(page, selectors)
    handle.click()
    handle.fill(value)
    _log_event("FILL", f"Filled selector {selector}")


def send_gmail_with_playwright(details: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]], str | None]:
    """
    Return (planned_actions, step_results, error_message)
    """
    actions: list[dict[str, Any]] = [
        {"action": "pw_launch", "browser": "chrome"},
        {"action": "pw_goto", "url": details.get("compose_url")},
        {"action": "pw_fill", "target": "to"},
        {"action": "pw_fill", "target": "subject"},
        {"action": "pw_fill", "target": "body"},
        {"action": "pw_attach", "target": "file"},
    ]
    if details.get("send"):
        actions.append({"action": "pw_send"})

    steps: list[dict[str, Any]] = []
    error_message: str | None = None

    to_list = details.get("to", [])
    subject = details.get("subject", "")
    body = details.get("body", "")
    attachments = details.get("attachments", [])
    compose_url = details.get("compose_url", "")

    if not compose_url:
        error_message = "Compose URL is missing for Playwright email flow."
        _record_step(steps, "pw_goto", False, error_message)
        return actions, steps, error_message

    chrome_path = _resolve_chrome_executable()
    user_data_dir = _resolve_user_data_dir()
    profile_dir_env = os.environ.get("PLAYWRIGHT_PROFILE_DIR", "").strip()
    profile_dir = profile_dir_env or "Default"

    _log_event("PLAYWRIGHT", f"Launching with user data dir: {user_data_dir}")

    try:
        print("[AGENT] 🌐 Playwright: launching browser...")
        with sync_playwright() as pw:
            browser = pw.chromium
            context = _launch_context(pw, user_data_dir, profile_dir, chrome_path)
            context.set_default_timeout(30000)
            context.set_default_navigation_timeout(30000)
            page = context.pages[0] if context.pages else context.new_page()

            _record_step(steps, "pw_launch", True, "Playwright browser launched")
            print("[AGENT] 🌐 Playwright: opening Gmail compose...")
            page.goto(compose_url, wait_until="domcontentloaded", timeout=30000)
            page.bring_to_front()
            if page.url.startswith("about:blank") or page.url.startswith("chrome-error://"):
                raise PlaywrightTimeoutError(f"Gmail compose did not load (current URL: {page.url})")
            _record_step(steps, "pw_goto", True, "Opened Gmail compose URL")

            try:
                _ensure_selector(
                    page,
                    ["textarea[name='to']", "textarea[aria-label='To']", "div[aria-label='Message Body']"],
                    timeout=20000,
                )
            except Exception:
                page_text = _page_text_for_error(page)
                category = _classify_browser_failure(page_text)
                raise PlaywrightTimeoutError(
                    f"Gmail compose not ready (category={category}). Make sure Gmail is logged in."
                )

            if to_list:
                _fill_textarea(page, ["textarea[name='to']", "textarea[aria-label='To']"], ", ".join(to_list))
            _record_step(steps, "pw_fill_to", True, "Recipient field set")

            if subject:
                _fill_textarea(page, ["input[name='subjectbox']"], subject)
            _record_step(steps, "pw_fill_subject", True, "Subject field set")

            if body:
                _fill_textarea(page, ["div[aria-label='Message Body']", "div[role='textbox']"], body)
            _record_step(steps, "pw_fill_body", True, "Body field set")

            if attachments:
                attach_path = attachments[0]
                if not os.path.exists(attach_path):
                    raise FileNotFoundError(f"Attachment file not found: {attach_path}")

                file_input = page.locator("input[type='file']")
                if file_input.count() == 0:
                    # Attempt to reveal attachment control.
                    for selector in (
                        "div[command='Files']",
                        "div[aria-label*='Attach']",
                        "button[aria-label*='Attach']",
                    ):
                        if page.locator(selector).count():
                            page.locator(selector).first.click()
                            break
                file_input = page.locator("input[type='file']")
                if file_input.count() == 0:
                    raise PlaywrightTimeoutError("Attachment input not found in Gmail compose.")

                file_input.first.set_input_files(attach_path)
                filename = Path(attach_path).name
                try:
                    page.wait_for_selector(f"text={filename}", timeout=20000)
                except PlaywrightTimeoutError:
                    _log_event("ATTACH", f"Attachment label not found for {filename}", "WARNING")

                _record_step(steps, "pw_attach", True, f"Attached file {filename}")

            if details.get("send"):
                send_selectors = [
                    "div[role='button'][data-tooltip^='Send']",
                    "div[role='button'][aria-label^='Send']",
                    "div[role='button'][aria-label*='Send']",
                ]
                _, selector = _ensure_selector(page, send_selectors, timeout=15000)
                page.locator(selector).first.click()
                _record_step(steps, "pw_send", True, "Send button clicked")

            context.close()
            return actions, steps, None

    except Exception as exc:
        error_message = _friendly_playwright_error(exc)
        _log_event("ERROR", error_message, "ERROR")
        _record_step(steps, "pw_error", False, error_message)
        return actions, steps, error_message
