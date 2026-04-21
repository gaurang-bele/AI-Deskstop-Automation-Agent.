# ============================================================
#  executor.py — The Hands: Executes Every Action
# ============================================================
#
#  IMPROVEMENTS IN THIS VERSION:
#    1. APP CATEGORIES - Separate handling for browsers, office, system apps
#    2. DETAILED LOGGING - Shows exactly what's happening at each step
#    3. TEST MODE - Verify features work without breaking changes
#
#  SAFETY:
#    - pyautogui.FAILSAFE = True → move mouse to top-left to stop
#    - pyautogui.PAUSE = 0.3 → 0.3s gap between every action
# ============================================================


# ── IMPORTS ──────────────────────────────────────────────────

import pyautogui
import subprocess
import time
import logging
import os
import json
import re
import unittest
from datetime import datetime

import pytesseract
from PIL import ImageGrab
import pygetwindow as gw
from excel_automation import EXCEL_ACTIONS, execute_excel_action, run_excel_smoke_checks
from llm import locate_ui_target

try:
    import tkinter as tk
except ImportError:
    tk = None

try:
    import win32gui
    import win32con
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False
    print("[EXECUTOR] win32gui not found - install pywin32 for better window focusing")


# ── LOGGING SETUP ────────────────────────────────────────────

# Configure detailed logging
logging.basicConfig(
    filename='agent_detailed.log',
    level=logging.DEBUG,
    format='%(asctime)s | %(levelname)s | %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

def log_action(action_type: str, details: str, status: str = "INFO"):
    """Enhanced logging with timestamps and categories"""
    timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
    level = getattr(logging, status.upper(), logging.INFO)
    msg = f"[{action_type}] {details}"
    logging.log(level, msg)
    print(f"[{timestamp}] {msg}")


def _set_clipboard_text(text: str) -> bool:
    if tk is None:
        return False

    root = None
    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        root.clipboard_clear()
        root.clipboard_append(text)
        root.update_idletasks()
        root.update()
        return True
    except Exception as exc:
        log_action("CLIPBOARD", f"Could not set clipboard text: {exc}", "WARNING")
        return False
    finally:
        if root is not None:
            try:
                root.destroy()
            except Exception:
                pass


def _paste_text(text: str) -> bool:
    if _set_clipboard_text(text):
        pyautogui.hotkey("ctrl", "v")
        return True

    pyautogui.write(text, interval=0.03)
    return True


def get_clipboard_text() -> str:
    if tk is None:
        return ""

    root = None
    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        root.update_idletasks()
        root.update()
        value = root.clipboard_get()
        return str(value)
    except Exception as exc:
        log_action("CLIPBOARD", f"Could not read clipboard text: {exc}", "WARNING")
        return ""
    finally:
        if root is not None:
            try:
                root.destroy()
            except Exception:
                pass


# ── GLOBAL STATE ─────────────────────────────────────────────

_last_opened_app = None
_last_opened_category = None


# ── TESSERACT PATH ────────────────────────────────────────────

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


# ── SAFETY SETTINGS ───────────────────────────────────────────

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3


# ── APP CATEGORIES ───────────────────────────────────────────
# Each category has its own timing and focus settings

APP_CATEGORIES = {
    "browser": {
        "apps": ["chrome", "google", "edge", "msedge", "firefox", "brave"],
        "load_time": 3.0,       # Seconds to wait after opening
        "focus_retries": 3,     # Number of focus attempts
        "focus_delay": 0.5,     # Delay between focus attempts
        "action_delay": 0.2,    # Delay after focusing before action
        "window_keywords": ["chrome", "edge", "firefox", "brave", "google"]
    },
    "office": {
        "apps": ["excel", "word", "msword", "powerpoint", "outlook"],
        "load_time": 2.5,
        "focus_retries": 3,
        "focus_delay": 0.5,
        "action_delay": 0.4,    # Office apps need more time
        "window_keywords": ["excel", "word", "powerpoint", "outlook"]
    },
    "system": {
        "apps": ["notepad", "calculator", "explorer", "file explorer", "paint"],
        "load_time": 1.0,
        "focus_retries": 2,
        "focus_delay": 0.3,
        "action_delay": 0.2,
        "window_keywords": ["notepad", "calculator", "explorer", "paint"]
    },
    "dev": {
        "apps": ["vscode", "code", "visual studio"],
        "load_time": 3.0,
        "focus_retries": 3,
        "focus_delay": 0.5,
        "action_delay": 0.3,
        "window_keywords": ["visual studio code", "code"]
    },
    "communication": {
        "apps": ["discord", "spotify", "zoom", "teams"],
        "load_time": 3.0,
        "focus_retries": 2,
        "focus_delay": 0.5,
        "action_delay": 0.3,
        "window_keywords": ["discord", "spotify", "zoom", "teams"]
    }
}


def get_app_category(app_name: str) -> tuple:
    """Get the category and settings for an app"""
    app_lower = app_name.lower()
    for category, settings in APP_CATEGORIES.items():
        if app_lower in settings["apps"]:
            return category, settings
    # Default settings for unknown apps
    return "unknown", {
        "load_time": 2.0,
        "focus_retries": 2,
        "focus_delay": 0.5,
        "action_delay": 0.3,
        "window_keywords": [app_lower]
    }


# ── APP NAME MAP ──────────────────────────────────────────────

APP_MAP = {
    # System apps
    "notepad":       "notepad.exe",
    "calculator":    "calc.exe",
    "explorer":      "explorer.exe",
    "file explorer": "explorer.exe",
    "paint":         "mspaint.exe",

    # Office apps
    "excel":         "excel.exe",
    "word":          "winword.exe",
    "msword":        "winword.exe",
    "powerpoint":    "powerpnt.exe",
    "outlook":       "outlook.exe",

    # Browsers
    "chrome":        'start chrome --profile-directory="Profile 3"',
    "google":        'start chrome --profile-directory="Profile 3"',
    "edge":          "start microsoft-edge:",
    "msedge":        "start microsoft-edge:",
    "firefox":       "start firefox",
    "brave":         "start brave",

    # Dev tools
    "vscode":        "code",
    "code":          "code",
    "visual studio": "devenv.exe",

    # Communication
    "discord":       "start discord:",
    "spotify":       "start spotify:",
    "zoom":          "start zoommtg:",
    "teams":         "start msteams:",

    # Special
    "search":        "start ms-search:",
    "youtube":       "start https://youtube.com",
}

# Window title keywords for focusing
WINDOW_TITLES = {
    "chrome":     "Chrome",
    "edge":       "Edge",
    "firefox":    "Firefox",
    "notepad":    "Notepad",
    "excel":      "Excel",
    "word":       "Word",
    "calculator": "Calculator",
    "explorer":   "File Explorer",
    "vscode":     "Visual Studio Code",
    "paint":      "Paint",
    "outlook":    "Outlook",
    "powerpoint": "PowerPoint",
    "teams":      "Teams",
    "discord":    "Discord",
    "spotify":    "Spotify",
    "brave":      "Brave",
    "zoom":       "Zoom",
}

# Create screenshots folder
os.makedirs("screenshots", exist_ok=True)


def _resolve_edge_command() -> str:
    candidates = []
    pf_x86 = os.environ.get("PROGRAMFILES(X86)")
    pf = os.environ.get("PROGRAMFILES")
    if pf_x86:
        candidates.append(os.path.join(pf_x86, "Microsoft", "Edge", "Application", "msedge.exe"))
    if pf:
        candidates.append(os.path.join(pf, "Microsoft", "Edge", "Application", "msedge.exe"))
    for path in candidates:
        if os.path.exists(path):
            return f'"{path}"'
    return APP_MAP.get("edge", "start microsoft-edge:")


def _open_url_in_browser(url: str, browser: str) -> bool:
    browser_name = (browser or "").lower().strip()
    if not browser_name:
        return False

    try:
        if browser_name in {"chrome", "google"}:
            cmd = f'start chrome --profile-directory="Profile 3" "{url}"'
        elif browser_name in {"edge", "msedge"}:
            edge_cmd = _resolve_edge_command()
            if edge_cmd.startswith('"'):
                cmd = f'{edge_cmd} "{url}"'
            else:
                cmd = f'start microsoft-edge:{url}'
        elif browser_name == "firefox":
            cmd = f'start firefox "{url}"'
        elif browser_name == "brave":
            cmd = f'start brave "{url}"'
        else:
            return False

        subprocess.Popen(cmd, shell=True)
        return True
    except Exception as exc:
        log_action("OPEN_URL", f"Browser-specific URL launch failed for {browser_name}: {exc}", "WARNING")
        return False


# ── CATEGORY-SPECIFIC FOCUS FUNCTIONS ────────────────────────

def focus_browser(app_name: str, settings: dict) -> bool:
    """Focus a browser window - handles profile names in titles"""
    log_action("FOCUS", f"Focusing browser: {app_name}")

    all_windows = gw.getAllWindows()
    matching = []

    for win in all_windows:
        if not win.title:
            continue
        title_lower = win.title.lower()
        # Browsers often have page title + browser name
        for keyword in settings["window_keywords"]:
            if keyword in title_lower:
                matching.append(win)
                break

    if matching:
        # For browsers, prefer windows with shorter titles (main window vs tabs)
        matching.sort(key=lambda w: len(w.title) if w.title else 999)
        return _activate_window(matching[0], settings)

    log_action("FOCUS", f"No browser window found for {app_name}", "WARNING")
    return False


def focus_office(app_name: str, settings: dict) -> bool:
    """Focus an Office app - handles document names in titles"""
    log_action("FOCUS", f"Focusing Office app: {app_name}")

    title_keyword = WINDOW_TITLES.get(app_name.lower(), app_name)
    all_windows = gw.getAllWindows()
    matching = []

    for win in all_windows:
        if win.title and title_keyword.lower() in win.title.lower():
            matching.append(win)

    if matching:
        # For Office, prefer most recently used (first in list)
        return _activate_window(matching[0], settings)

    log_action("FOCUS", f"No Office window found for {app_name}", "WARNING")
    return False


def focus_system(app_name: str, settings: dict) -> bool:
    """Focus a system app - straightforward title matching"""
    log_action("FOCUS", f"Focusing system app: {app_name}")

    title_keyword = WINDOW_TITLES.get(app_name.lower(), app_name)
    all_windows = gw.getAllWindows()

    for win in all_windows:
        if win.title and title_keyword.lower() in win.title.lower():
            return _activate_window(win, settings)

    log_action("FOCUS", f"No system window found for {app_name}", "WARNING")
    return False


def focus_generic(app_name: str, settings: dict) -> bool:
    """Focus any app - fallback for unknown categories"""
    log_action("FOCUS", f"Focusing generic app: {app_name}")

    title_keyword = WINDOW_TITLES.get(app_name.lower(), app_name)
    all_windows = gw.getAllWindows()

    # Debug output
    log_action("DEBUG", f"Looking for '{title_keyword}' in {len(all_windows)} windows")

    for win in all_windows:
        if win.title and title_keyword.lower() in win.title.lower():
            return _activate_window(win, settings)

    return False


def _activate_window(win, settings: dict) -> bool:
    """Actually activate/focus a window"""
    try:
        if HAS_WIN32:
            hwnd = win._hWnd
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.2)
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(settings["action_delay"])
            log_action("FOCUS", f"Activated (win32): '{win.title}'", "DEBUG")
            return True
    except Exception as e:
        log_action("FOCUS", f"win32 failed: {e}, trying pygetwindow", "WARNING")

    # Fallback to pygetwindow
    try:
        if win.isMinimized:
            win.restore()
            time.sleep(0.2)
        win.activate()
        time.sleep(settings["action_delay"])
        log_action("FOCUS", f"Activated (pygetwindow): '{win.title}'", "DEBUG")
        return True
    except Exception as e:
        log_action("FOCUS", f"pygetwindow failed: {e}", "ERROR")
        return False


def focus_app_by_category(app_name: str) -> bool:
    """Main focus function - routes to category-specific handler"""
    global _last_opened_category

    category, settings = get_app_category(app_name)
    _last_opened_category = category

    log_action("FOCUS", f"App '{app_name}' is category '{category}'")

    # Route to category-specific focus function
    focus_funcs = {
        "browser": focus_browser,
        "office": focus_office,
        "system": focus_system,
        "dev": focus_generic,
        "communication": focus_generic,
        "unknown": focus_generic
    }

    focus_func = focus_funcs.get(category, focus_generic)

    # Try multiple times with category-specific settings
    for attempt in range(settings["focus_retries"]):
        if focus_func(app_name, settings):
            return True
        log_action("FOCUS", f"Attempt {attempt + 1}/{settings['focus_retries']} failed, retrying...")
        time.sleep(settings["focus_delay"])

    return False


# Legacy function for compatibility
def focus_app(app_name: str) -> bool:
    """Legacy focus function - now routes to category-based system"""
    return focus_app_by_category(app_name)


def _focus_last_opened_app(action_label: str, strict: bool = False) -> tuple[bool, str]:
    if not _last_opened_app:
        return True, ""

    _, settings = get_app_category(_last_opened_app)
    focused = focus_app_by_category(_last_opened_app)
    if not focused:
        time.sleep(settings["focus_delay"])
        focused = focus_app_by_category(_last_opened_app)

    if focused:
        time.sleep(settings["action_delay"])
        return True, ""

    if strict:
        log_action(action_label, f"Could not focus {_last_opened_app}", "ERROR")
        return False, f"Could not focus {_last_opened_app}"

    log_action(
        action_label,
        f"Could not focus {_last_opened_app}; continuing with current foreground window",
        "WARNING",
    )
    return True, ""


# ── HELPER: Find Text on Screen Using OCR ────────────────────

def find_text_on_screen(text: str):
    log_action("OCR", f"Searching screen for text: '{text}'")

    img = ImageGrab.grab()
    data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)

    for i, word in enumerate(data["text"]):
        if not word.strip():
            continue
        if text.lower() in word.lower():
            conf = int(data["conf"][i])
            if conf < 60:
                continue
            x = data["left"][i] + data["width"][i] // 2
            y = data["top"][i] + data["height"][i] // 2
            log_action("OCR", f"Found '{word}' at ({x}, {y}) — confidence: {conf}%")
            return (x, y)

    log_action("OCR", f"Text '{text}' not found on screen", "WARNING")
    return None


def _is_file_dialog_open() -> bool:
    # Language-independent check: native Windows file pickers are commonly #32770 dialogs.
    if HAS_WIN32:
        found_dialog = {"value": False}

        def _enum_handler(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return True
            class_name = win32gui.GetClassName(hwnd)
            if class_name == "#32770":
                found_dialog["value"] = True
                return False
            return True

        try:
            win32gui.EnumWindows(_enum_handler, None)
            if found_dialog["value"]:
                return True
        except Exception:
            pass

    try:
        titles = [title.strip().lower() for title in gw.getAllTitles() if title and title.strip()]
    except Exception:
        return False

    dialog_markers = (
        "open",
        "choose file",
        "file upload",
        "upload file",
    )
    return any(any(marker in title for marker in dialog_markers) for title in titles)


def _focus_gmail_compose_window() -> bool:
    try:
        windows = gw.getAllWindows()
    except Exception:
        return False

    compose_markers = ("compose mail", "gmail")
    candidates = [
        win for win in windows
        if getattr(win, "title", None)
        and any(marker in win.title.lower() for marker in compose_markers)
    ]
    if not candidates:
        return False

    # Prefer explicit compose windows first.
    candidates.sort(key=lambda w: 0 if "compose" in w.title.lower() else 1)
    return _activate_window(candidates[0], APP_CATEGORIES["browser"])


def _get_active_browser_window():
    try:
        windows = gw.getAllWindows()
    except Exception:
        return None

    browser_markers = ("google chrome", "gmail")
    active = []
    for win in windows:
        title = getattr(win, "title", "")
        if not title:
            continue
        lower = title.lower()
        if any(marker in lower for marker in browser_markers):
            active.append(win)

    if not active:
        return None

    # Prefer compose tab/window if available, otherwise take current browser.
    active.sort(key=lambda w: 0 if "compose" in w.title.lower() else 1)
    return active[0]


def _window_bounds(win) -> tuple[int, int, int, int] | None:
    try:
        left = int(win.left)
        top = int(win.top)
        width = int(win.width)
        height = int(win.height)
        return left, top, width, height
    except Exception:
        return None


def _find_text_on_screen_in_region(text: str, region: tuple[int, int, int, int], min_conf: int = 60):
    left, top, width, height = region
    right = left + width
    bottom = top + height
    log_action("OCR", f"Searching '{text}' in region ({left},{top})-({right},{bottom})")

    img = ImageGrab.grab()
    data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)

    for i, word in enumerate(data["text"]):
        if not word.strip():
            continue
        if text.lower() not in word.lower():
            continue
        conf = int(data["conf"][i])
        if conf < min_conf:
            continue
        x = data["left"][i] + data["width"][i] // 2
        y = data["top"][i] + data["height"][i] // 2
        if left <= x <= right and top <= y <= bottom:
            log_action("OCR", f"Found '{word}' at ({x}, {y}) in region — confidence: {conf}%")
            return (x, y)

    log_action("OCR", f"Text '{text}' not found in region", "WARNING")
    return None


def _open_attach_dialog_by_layout_guess() -> bool:
    win = _get_active_browser_window()
    if not win:
        log_action("ATTACH", "Layout fallback skipped: no active browser window", "WARNING")
        return False

    bounds = _window_bounds(win)
    if not bounds:
        log_action("ATTACH", "Layout fallback skipped: could not read browser window bounds", "WARNING")
        return False
    left, top, width, height = bounds

    log_action(
        "ATTACH",
        f"Layout fallback on window '{_safe_console_text(getattr(win, 'title', ''))}' "
        f"at ({left},{top}) size {width}x{height}",
    )

    # Constrain to compose-toolbar area (avoid random clicks in page/header).
    x_ratios = (0.62, 0.66, 0.70, 0.74, 0.78, 0.82)
    y_ratios = (0.78, 0.82, 0.86, 0.90)
    candidates = [(left + int(width * xr), top + int(height * yr)) for yr in y_ratios for xr in x_ratios]

    for x, y in candidates:
        log_action("ATTACH", f"Layout fallback click at ({x}, {y})", "DEBUG")
        pyautogui.click(x, y)
        time.sleep(1.2)
        if _is_file_dialog_open():
            log_action("ATTACH", "File dialog detected after layout fallback click")
            return True
    log_action("ATTACH", "Layout fallback clicks did not open a file dialog", "WARNING")
    return False


def _open_gmail_attach_dialog() -> bool:
    win = _get_active_browser_window()
    bounds = _window_bounds(win) if win else None
    search_region = None
    if bounds:
        left, top, width, height = bounds
        search_region = (
            left + int(width * 0.52),
            top + int(height * 0.50),
            int(width * 0.46),
            int(height * 0.48),
        )

    # 1) Try clicking explicit attach labels (depends on UI language/accessibility text).
    for label in ("Attach files", "Attach"):
        coords = _find_text_on_screen_in_region(label, search_region) if search_region else find_text_on_screen(label)
        if coords:
            pyautogui.click(coords[0], coords[1])
            time.sleep(1.0)
            if _is_file_dialog_open():
                return True

    # 2) If labels are not OCR-visible, click near Send button where paperclip usually lives.
    for send_label in ("Send",):
        send_coords = _find_text_on_screen_in_region(send_label, search_region) if search_region else find_text_on_screen(send_label)
        if send_coords:
            attach_guess = (send_coords[0] + 110, send_coords[1])
            pyautogui.click(attach_guess[0], attach_guess[1])
            time.sleep(1.0)
            if _is_file_dialog_open():
                return True

    # 3) Last resort: click likely compose-toolbar paperclip positions.
    if _open_attach_dialog_by_layout_guess():
        return True

    # 4) Vision fallback: ask LLM to locate paperclip/attach button.
    try:
        llm_point = locate_ui_target("Gmail compose toolbar paperclip attach-files button")
    except Exception as exc:
        log_action("ATTACH", f"LLM locate failed: {exc}", "WARNING")
        llm_point = None

    if llm_point:
        if search_region:
            rx, ry, rw, rh = search_region
            if not (rx <= llm_point[0] <= rx + rw and ry <= llm_point[1] <= ry + rh):
                log_action("ATTACH", f"LLM point {llm_point} outside compose region; skipping click", "WARNING")
                llm_point = None
        if llm_point:
            log_action("ATTACH", f"LLM fallback click at {llm_point}")
            pyautogui.click(llm_point[0], llm_point[1])
            time.sleep(1.2)
            if _is_file_dialog_open():
                log_action("ATTACH", "File dialog detected after LLM fallback click")
                return True

    return False


def _run_excel_unit_suite() -> tuple[str, str]:
    suite = unittest.defaultTestLoader.loadTestsFromName("test_excel_automation")
    result = unittest.TestResult()
    suite.run(result)
    total = result.testsRun
    failures = len(result.failures) + len(result.errors)

    if failures:
        return "FAIL", f"{total} tests, {failures} failed"
    return "PASS", f"{total} tests passed"


def _safe_console_text(text: str) -> str:
    return str(text).encode("ascii", errors="replace").decode("ascii")


# ── TEST MODE ────────────────────────────────────────────────

def run_tests():
    """
    Test mode to verify all features work correctly.
    Run with: python -c "from executor import run_tests; run_tests()"
    """
    print("\n" + "=" * 60)
    print("  EXECUTOR TEST MODE")
    print("=" * 60 + "\n")

    results = []

    # Test 1: Category detection
    print("Test 1: App Category Detection")
    test_apps = {
        "chrome": "browser",
        "excel": "office",
        "notepad": "system",
        "vscode": "dev",
        "discord": "communication"
    }
    for app, expected in test_apps.items():
        category, _ = get_app_category(app)
        status = "PASS" if category == expected else "FAIL"
        results.append(("Category Detection", app, status))
        print(f"  {status}: {app} -> {category} (expected: {expected})")

    # Test 2: Window list
    print("\nTest 2: Window Detection")
    try:
        windows = gw.getAllWindows()
        print(f"  PASS: Found {len(windows)} windows")
        results.append(("Window Detection", "getAllWindows", "PASS"))
        print("  First 5 windows:")
        for w in windows[:5]:
            if w.title:
                print(f"    - '{_safe_console_text(w.title)}'")
    except Exception as e:
        print(f"  FAIL: {e}")
        results.append(("Window Detection", "getAllWindows", "FAIL"))

    # Test 3: pyautogui
    print("\nTest 3: PyAutoGUI")
    try:
        pos = pyautogui.position()
        print(f"  PASS: Mouse at {pos}")
        results.append(("PyAutoGUI", "position", "PASS"))
    except Exception as e:
        print(f"  FAIL: {e}")
        results.append(("PyAutoGUI", "position", "FAIL"))

    # Test 4: win32gui
    print("\nTest 4: Win32GUI")
    if HAS_WIN32:
        print("  PASS: win32gui available")
        results.append(("Win32GUI", "import", "PASS"))
    else:
        print("  WARNING: win32gui not available (install pywin32)")
        results.append(("Win32GUI", "import", "WARNING"))

    # Test 5: OCR (Tesseract)
    print("\nTest 5: Tesseract OCR")
    try:
        img = ImageGrab.grab()
        data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
        word_count = len([w for w in data["text"] if w.strip()])
        print(f"  PASS: OCR found {word_count} words on screen")
        results.append(("Tesseract OCR", "image_to_data", "PASS"))
    except Exception as e:
        print(f"  FAIL: {e}")
        results.append(("Tesseract OCR", "image_to_data", "FAIL"))

    # Test 6: Excel unit suite
    print("\nTest 6: Excel Unit Tests")
    try:
        status, note = _run_excel_unit_suite()
        print(f"  {status}: {note}")
        results.append(("Excel Unit Tests", note, status))
    except Exception as e:
        print(f"  FAIL: {e}")
        results.append(("Excel Unit Tests", str(e), "FAIL"))

    # Test 7: Excel smoke checks
    print("\nTest 7: Excel Smoke Checks")
    sample_workbook = os.getenv("AIAGENT_EXCEL_TEST_FILE")
    try:
        excel_results = run_excel_smoke_checks(sample_workbook=sample_workbook)
        for name, note, status in excel_results:
            print(f"  {status}: {name} -> {note}")
            results.append((f"Excel Smoke::{name}", note, status))
    except Exception as e:
        print(f"  FAIL: {e}")
        results.append(("Excel Smoke", str(e), "FAIL"))

    # Summary
    print("\n" + "=" * 60)
    print("  TEST SUMMARY")
    print("=" * 60)
    passed = sum(1 for r in results if r[2] == "PASS")
    failed = sum(1 for r in results if r[2] == "FAIL")
    warnings = sum(1 for r in results if r[2] == "WARNING")
    print(f"  PASSED: {passed}")
    print(f"  FAILED: {failed}")
    print(f"  WARNINGS: {warnings}")
    print("=" * 60 + "\n")

    return results


def test_action(action_type: str, **kwargs):
    """
    Test a single action without side effects where possible.
    Usage: python -c "from executor import test_action; test_action('focus_app', app='chrome')"
    """
    print(f"\n[TEST] Testing action: {action_type}")
    print(f"[TEST] Parameters: {kwargs}")

    if action_type == "focus_app":
        app = kwargs.get("app", "notepad")
        category, settings = get_app_category(app)
        print(f"[TEST] App '{app}' → Category '{category}'")
        print(f"[TEST] Settings: load_time={settings['load_time']}, retries={settings['focus_retries']}")
        result = focus_app_by_category(app)
        print(f"[TEST] Result: {'SUCCESS' if result else 'FAILED'}")
        return result

    elif action_type == "find_text":
        text = kwargs.get("text", "Start")
        coords = find_text_on_screen(text)
        print(f"[TEST] Result: {coords if coords else 'NOT FOUND'}")
        return coords

    else:
        print(f"[TEST] Unknown action type: {action_type}")
        return None


# ── MAIN EXECUTE FUNCTION ─────────────────────────────────────

def execute_action(action: dict) -> tuple[bool, str]:
    """
    Execute a single action from the model's plan.
    Returns (success: bool, note: str)
    """
    global _last_opened_app, _last_opened_category

    act = action.get("action", "")
    log_action("EXECUTE", f"Action: {act} | Params: {json.dumps(action)}")

    try:
        # ── OPEN APP ──────────────────────────────────────────
        if act == "open_app":
            app_name = action.get("app", "").lower()
            if app_name in {"edge", "msedge"}:
                exe = _resolve_edge_command()
            else:
                exe = APP_MAP.get(app_name, app_name)
            category, settings = get_app_category(app_name)

            log_action("OPEN_APP", f"Opening '{app_name}' (category: {category})")

            subprocess.Popen(exe, shell=True)

            # Wait based on category
            log_action("OPEN_APP", f"Waiting {settings['load_time']}s for {category} app to load")
            time.sleep(settings["load_time"])

            # Store for later focus
            _last_opened_app = app_name
            _last_opened_category = category

            # Focus with category-specific handling
            focused = focus_app_by_category(app_name)

            if not focused:
                log_action("OPEN_APP", f"App opened but window not focused", "WARNING")

            return True, f"Opened {exe}"

        # ── OPEN FILE ─────────────────────────────────────────
        elif act == "open_file":
            path = action.get("path", "").replace("\\\\", "\\")

            log_action("OPEN_FILE", f"Opening file: {path}")

            if not os.path.exists(path):
                log_action("OPEN_FILE", f"File not found: {path}", "ERROR")
                return False, f"File not found: '{path}'"

            # Determine what app will open this file
            ext = os.path.splitext(path)[1].lower()
            if ext in ['.xlsx', '.xls', '.csv']:
                _last_opened_app = "excel"
                _last_opened_category = "office"
            elif ext in ['.doc', '.docx']:
                _last_opened_app = "word"
                _last_opened_category = "office"
            elif ext == '.txt':
                _last_opened_app = "notepad"
                _last_opened_category = "system"
            else:
                _last_opened_app = None
                _last_opened_category = "unknown"

            os.startfile(path)

            # Wait based on category
            _, settings = get_app_category(_last_opened_app or "unknown")
            log_action("OPEN_FILE", f"Waiting {settings['load_time'] + 1}s for file to open")
            time.sleep(settings["load_time"] + 1)

            return True, f"Opened file: {path}"

        # ── LEFT CLICK ────────────────────────────────────────
        elif act in EXCEL_ACTIONS:
            log_action("EXCEL", f"Running shared Excel action: {act}")
            success, note = execute_excel_action(action)
            if success:
                _last_opened_app = "excel"
                _last_opened_category = "office"
                return True, note
            return False, note

        elif act == "click":
            x, y = action["x"], action["y"]
            log_action("CLICK", f"Clicking at ({x}, {y})")
            pyautogui.click(x, y)
            return True, f"Clicked ({x}, {y})"

        # ── DOUBLE CLICK ──────────────────────────────────────
        elif act == "double_click":
            x, y = action["x"], action["y"]
            log_action("CLICK", f"Double-clicking at ({x}, {y})")
            pyautogui.doubleClick(x, y)
            return True, f"Double-clicked ({x}, {y})"

        # ── RIGHT CLICK ───────────────────────────────────────
        elif act == "right_click":
            x, y = action["x"], action["y"]
            log_action("CLICK", f"Right-clicking at ({x}, {y})")
            pyautogui.rightClick(x, y)
            return True, f"Right-clicked ({x}, {y})"

        # ── CLICK TEXT (OCR-based) ─────────────────────────────
        elif act == "click_text":
            text_to_find = action.get("text", "")
            log_action("CLICK_TEXT", f"Looking for text: '{text_to_find}'")

            coords = find_text_on_screen(text_to_find)

            if coords:
                pyautogui.click(coords[0], coords[1])
                return True, f"Clicked text '{text_to_find}' at {coords}"
            else:
                return False, f"Text '{text_to_find}' not found on screen"

        # ── ATTACH FILE (Gmail compose) ──────────────────────────
        elif act == "attach_file":
            path = action.get("path", "").replace("\\\\", "\\").strip().strip('"')
            compose_url = action.get("compose_url", "").strip()
            browser_hint = action.get("browser", "").strip().lower()
            if not path:
                return False, "Missing attachment path"
            if not os.path.exists(path):
                return False, f"Attachment file not found: '{path}'"

            log_action("ATTACH", f"Attaching file: {path}")

            focused, note = _focus_last_opened_app("ATTACH", strict=False)
            if not focused:
                log_action("ATTACH", note, "WARNING")

            # Always reopen compose URL in the intended browser before attach to avoid cross-browser drift.
            if compose_url:
                launched = _open_url_in_browser(compose_url, browser_hint or _last_opened_app or "chrome")
                if not launched:
                    os.startfile(compose_url)
                time.sleep(4.0)

            # Ensure we are on Gmail compose, not a generic tab.
            compose_focused = _focus_gmail_compose_window()

            if not compose_focused:
                return False, "Could not focus Gmail compose window."

            time.sleep(0.5)

            _open_gmail_attach_dialog()

            if not _is_file_dialog_open():
                return False, "Could not open attachment dialog."

            _paste_text(path)
            time.sleep(0.2)
            pyautogui.press("enter")
            time.sleep(2.0)
            return True, f"Attached file: {path}"

        # ── TYPE TEXT ─────────────────────────────────────────
        elif act == "type_text":
            text = action["text"]
            log_action("TYPE", f"Typing: '{text[:50]}...' (len={len(text)})")

            focused, note = _focus_last_opened_app("TYPE")
            if not focused:
                return False, note

            pyautogui.write(text, interval=0.03)
            return True, f"Typed: '{text}'"

        # ── PASTE TEXT ───────────────────────────────────────
        elif act == "paste_text":
            text = action["text"]
            log_action("PASTE", f"Pasting: '{text[:50]}...' (len={len(text)})")

            focused, note = _focus_last_opened_app("PASTE")
            if not focused:
                return False, note

            _paste_text(text)
            return True, f"Pasted text len={len(text)}"

        # ── OPEN URL ─────────────────────────────────────────
        elif act == "open_url":
            url = action.get("url", "").strip()
            target_browser = action.get("app", "").strip().lower()
            if not url:
                return False, "Missing URL"

            if not re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", url):
                url = f"https://{url}"

            log_action("OPEN_URL", f"Opening URL: {url}")
            launched = _open_url_in_browser(url, target_browser or _last_opened_app or "")
            if not launched:
                os.startfile(url)
                target_browser = target_browser or _last_opened_app or "chrome"
            time.sleep(2.5)
            _last_opened_app = target_browser or "chrome"
            _last_opened_category = "browser"
            return True, f"Opened URL: {url}"

        # ── DRAG ─────────────────────────────────────────────
        elif act == "drag":
            start_x = action["start_x"]
            start_y = action["start_y"]
            end_x = action["end_x"]
            end_y = action["end_y"]
            duration = action.get("duration", 0.3)
            log_action("DRAG", f"Dragging from ({start_x}, {start_y}) to ({end_x}, {end_y})")
            pyautogui.moveTo(start_x, start_y)
            pyautogui.dragTo(end_x, end_y, duration=duration, button="left")
            return True, f"Dragged to ({end_x}, {end_y})"

        # ── MOVE MOUSE ───────────────────────────────────────
        elif act == "move_mouse":
            x = action["x"]
            y = action["y"]
            duration = action.get("duration", 0.2)
            log_action("MOUSE", f"Moving to ({x}, {y})")
            pyautogui.moveTo(x, y, duration=duration)
            return True, f"Moved mouse to ({x}, {y})"

        # ── PRESS KEY / SHORTCUT ──────────────────────────────
        elif act == "press_key":
            keys = action.get("keys", [])
            log_action("KEY", f"Pressing: {keys}")

            focused, note = _focus_last_opened_app("KEY")
            if not focused:
                return False, note

            if isinstance(keys, list) and len(keys) > 0:
                pyautogui.hotkey(*keys)
            else:
                pyautogui.press(keys)
            return True, f"Pressed: {keys}"

        # ── SCROLL ────────────────────────────────────────────
        elif act == "scroll":
            x = action.get("x", 500)
            y = action.get("y", 300)
            amount = action.get("amount", -3)
            log_action("SCROLL", f"Scrolling {amount} at ({x}, {y})")
            pyautogui.scroll(amount, x=x, y=y)
            return True, f"Scrolled {amount} at ({x},{y})"

        # ── WAIT ──────────────────────────────────────────────
        elif act == "wait":
            seconds = action.get("seconds", 1)
            log_action("WAIT", f"Waiting {seconds}s")
            time.sleep(seconds)
            return True, f"Waited {seconds}s"

        # ── SCREENSHOT ────────────────────────────────────────
        elif act == "screenshot":
            label = action.get("label", "step")
            filepath = f"screenshots/{label}.png"
            log_action("SCREENSHOT", f"Saving to {filepath}")
            pyautogui.screenshot(filepath)
            return True, f"Screenshot saved: {filepath}"

        # ── DELETE ROWS ───────────────────────────────────────
        elif act == "delete_rows":
            num_rows = action.get("count", 1)
            app = action.get("app", "excel").lower()

            log_action("DELETE_ROWS", f"Deleting {num_rows} rows in {app}")

            # Get category-specific settings
            category, settings = get_app_category(app)

            # Focus the app first with proper settings
            focused = focus_app_by_category(app)
            if not focused:
                time.sleep(settings["focus_delay"])
                focused = focus_app_by_category(app)
                if not focused:
                    log_action("DELETE_ROWS", f"Could not focus {app}", "ERROR")
                    return False, f"Could not focus {app} window"

            # Extra delay for Office apps
            time.sleep(settings["action_delay"])

            if app == "excel":
                # Go to beginning
                pyautogui.hotkey("ctrl", "Home")
                time.sleep(0.4)

                # Use Go To dialog (Ctrl+G)
                pyautogui.hotkey("ctrl", "g")
                time.sleep(0.5)  # Extra time for Office dialog

                # Type row range
                pyautogui.typewrite(f"1:{num_rows}", interval=0.05)
                time.sleep(0.3)

                # Select rows
                pyautogui.press("enter")
                time.sleep(0.4)

                # Close dialog
                pyautogui.press("escape")
                time.sleep(0.3)

                # Delete rows
                pyautogui.hotkey("ctrl", "-")
                time.sleep(0.5)

                # Confirm
                pyautogui.press("enter")
                time.sleep(0.3)

            elif app == "notepad":
                pyautogui.press("home")
                time.sleep(0.2)
                for _ in range(num_rows):
                    pyautogui.hotkey("shift", "down")
                    time.sleep(0.1)
                pyautogui.press("delete")
            else:
                # Generic approach
                pyautogui.press("home")
                time.sleep(0.2)
                for _ in range(num_rows):
                    pyautogui.hotkey("shift", "down")
                    time.sleep(0.1)
                pyautogui.press("delete")

            return True, f"Deleted {num_rows} row(s) in {app}"

        # ── UNKNOWN ACTION ────────────────────────────────────
        else:
            log_action("EXECUTE", f"Unknown action type: '{act}'", "ERROR")
            return False, f"Unknown action type: '{act}'"

    except Exception as e:
        error_msg = str(e)
        log_action("EXECUTE", f"Action failed [{act}]: {error_msg}", "ERROR")
        return False, f"Error: {error_msg}"


# ── RUN TESTS ON IMPORT (optional) ────────────────────────────

if __name__ == "__main__":
    print("Running executor tests...")
    run_tests()
