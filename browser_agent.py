# ============================================================
#  browser_agent.py — Agentic browsing loop with Playwright
# ============================================================

from __future__ import annotations

import logging
import os
import time
from dataclasses import dataclass, field
import base64
from pathlib import Path
from typing import Any

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from llm import plan_browser_action, validate_browser_goal


LOG_PATH = "browser_agent.log"
logger = logging.getLogger("browser_agent")
if not logger.handlers:
    logger.setLevel(logging.INFO)
    handler = logging.FileHandler(LOG_PATH, encoding="utf-8")
    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    handler.setFormatter(formatter)
    logger.addHandler(handler)


def _log(action: str, detail: str, level: str = "INFO") -> None:
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


def _classify_failure(dom_text: str) -> str:
    text = (dom_text or "").lower()
    if any(token in text for token in ("captcha", "recaptcha", "i'm not a robot")):
        return "CAPTCHA"
    if any(token in text for token in ("two-factor", "otp", "verification code", "authenticator")):
        return "OTP_MFA"
    if any(token in text for token in ("session expired", "sign in", "login", "log in")):
        return "SESSION_EXPIRED"
    if any(token in text for token in ("cookies", "consent", "privacy")):
        return "COOKIE_CONSENT"
    if any(token in text for token in ("automation", "bot detected", "unusual traffic")):
        return "AUTOMATION_DETECTED"
    return "UNKNOWN"


@dataclass
class BrowserAction:
    action: str
    selector: str | None = None
    text: str | None = None
    url: str | None = None
    keys: str | None = None
    seconds: float | None = None
    note: str | None = None


@dataclass
class BrowserAgentConfig:
    max_steps: int = 12
    headless: bool = False
    browser: str = "chrome"
    user_data_dir: str = field(default_factory=lambda: os.environ.get("PLAYWRIGHT_USER_DATA_DIR", "").strip() or _default_user_data_dir())
    profile_dir: str = field(default_factory=lambda: os.environ.get("PLAYWRIGHT_PROFILE_DIR", "").strip() or "Default")


class BrowserController:
    def __init__(self, config: BrowserAgentConfig) -> None:
        self.config = config
        self.context = None
        self.page = None

    def launch(self) -> None:
        chrome_path = _resolve_chrome_executable()
        _log("BROWSER", f"Launching {self.config.browser} with profile {self.config.profile_dir}")
        self._playwright = sync_playwright().start()
        try:
            self.context = self._playwright.chromium.launch_persistent_context(
                user_data_dir=self.config.user_data_dir,
                headless=self.config.headless,
                executable_path=chrome_path,
                args=[f"--profile-directory={self.config.profile_dir}", "--start-maximized"],
                viewport=None,
            )
        except Exception as exc:
            if chrome_path is None:
                raise RuntimeError(
                    "Playwright browser not found. Install browsers via: playwright install"
                ) from exc
            else:
                raise
        self.page = self.context.pages[0] if self.context.pages else self.context.new_page()

    def close(self) -> None:
        try:
            if self.context:
                self.context.close()
        finally:
            if hasattr(self, "_playwright"):
                self._playwright.stop()

    def execute(self, action: BrowserAction) -> tuple[bool, str]:
        if not self.page:
            return False, "Browser page not initialized"

        try:
            if action.action == "navigate" and action.url:
                self.page.goto(action.url, wait_until="domcontentloaded")
                return True, f"Navigated to {action.url}"
            if action.action == "click" and action.selector:
                self.page.locator(action.selector).first.click()
                return True, f"Clicked {action.selector}"
            if action.action == "fill" and action.selector is not None:
                self.page.locator(action.selector).first.fill(action.text or "")
                return True, f"Filled {action.selector}"
            if action.action == "press" and action.keys:
                self.page.keyboard.press(action.keys)
                return True, f"Pressed {action.keys}"
            if action.action == "wait":
                time.sleep(action.seconds or 1.0)
                return True, f"Waited {action.seconds or 1.0}s"
            if action.action == "done":
                return True, action.note or "Goal achieved"
        except PlaywrightTimeoutError as exc:
            return False, f"Timeout: {exc}"
        except Exception as exc:
            return False, str(exc)

        return False, "Invalid action"


class PagePerception:
    @staticmethod
    def snapshot(page) -> dict[str, str]:
        url = page.url
        title = page.title()
        dom_text = ""
        try:
            dom_text = page.inner_text("body")
        except Exception:
            dom_text = page.content()
        dom_text = (dom_text or "")[:16000]
        screenshot = page.screenshot(full_page=True)
        screenshot_b64 = base64.b64encode(screenshot).decode("utf-8")
        return {
            "url": url,
            "title": title,
            "dom_text": dom_text,
            "screenshot_b64": screenshot_b64,
        }


def run_agentic_browse(command: str, config: BrowserAgentConfig | None = None) -> tuple[list[dict[str, Any]], list[dict[str, Any]], str | None]:
    cfg = config or BrowserAgentConfig()
    controller = BrowserController(cfg)
    steps: list[dict[str, Any]] = []
    error_message: str | None = None

    try:
        controller.launch()
        _log("BROWSER", "Browser launched")
        history: list[dict[str, str]] = []
        last_error: str | None = None
        failure_count = 0

        for _ in range(cfg.max_steps):
            perception = PagePerception.snapshot(controller.page)
            screenshot_b64 = perception["screenshot_b64"]

            action_payload = plan_browser_action(
                goal=command,
                url=perception["url"],
                title=perception["title"],
                dom_text=perception["dom_text"],
                screenshot_b64=screenshot_b64,
                history=history,
                last_error=last_error,
            )

            action = BrowserAction(**action_payload)
            success, note = controller.execute(action)
            steps.append({"step": len(steps) + 1, "action": action.action, "success": success, "note": note})
            history.append({"action": action.action, "note": note})
            if action.action == "done":
                return [], steps, None

            if not success:
                last_error = note
                failure_count += 1
                category = _classify_failure(perception["dom_text"])
                error_message = f"{note} (category={category})"
                _log("ERROR", error_message, "ERROR")
                if failure_count >= 2:
                    break
                continue

            try:
                updated = PagePerception.snapshot(controller.page)
                if validate_browser_goal(
                    goal=command,
                    url=updated["url"],
                    title=updated["title"],
                    dom_text=updated["dom_text"],
                    screenshot_b64=updated["screenshot_b64"],
                ):
                    steps.append({"step": len(steps) + 1, "action": "validate", "success": True, "note": "Goal verified"})
                    return [], steps, None
            except Exception as exc:
                _log("VALIDATE", f"Validator failed: {exc}", "WARNING")

        if not error_message:
            error_message = "Agentic browser loop reached max steps without completion."
            _log("ERROR", error_message, "WARNING")
        return [], steps, error_message
    finally:
        controller.close()
