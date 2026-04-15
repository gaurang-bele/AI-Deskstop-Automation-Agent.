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
    "edge":          "start msedge",
    "msedge":        "start msedge",
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

        # ── TYPE TEXT ─────────────────────────────────────────
        elif act == "type_text":
            text = action["text"]
            log_action("TYPE", f"Typing: '{text[:50]}...' (len={len(text)})")

            # Focus the target app first
            if _last_opened_app:
                _, settings = get_app_category(_last_opened_app)
                focused = focus_app_by_category(_last_opened_app)
                if not focused:
                    time.sleep(settings["focus_delay"])
                    focused = focus_app_by_category(_last_opened_app)
                    if not focused:
                        log_action("TYPE", f"Could not focus {_last_opened_app}", "ERROR")
                        return False, f"Could not focus {_last_opened_app}"
                time.sleep(settings["action_delay"])

            pyautogui.write(text, interval=0.03)
            return True, f"Typed: '{text}'"

        # ── PASTE TEXT ───────────────────────────────────────
        elif act == "paste_text":
            text = action["text"]
            log_action("PASTE", f"Pasting: '{text[:50]}...' (len={len(text)})")

            if _last_opened_app:
                _, settings = get_app_category(_last_opened_app)
                focused = focus_app_by_category(_last_opened_app)
                if not focused:
                    time.sleep(settings["focus_delay"])
                    focused = focus_app_by_category(_last_opened_app)
                    if not focused:
                        log_action("PASTE", f"Could not focus {_last_opened_app}", "ERROR")
                        return False, f"Could not focus {_last_opened_app}"
                time.sleep(settings["action_delay"])

            _paste_text(text)
            return True, f"Pasted text len={len(text)}"

        # ── OPEN URL ─────────────────────────────────────────
        elif act == "open_url":
            url = action.get("url", "").strip()
            if not url:
                return False, "Missing URL"

            if not re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", url):
                url = f"https://{url}"

            log_action("OPEN_URL", f"Opening URL: {url}")
            os.startfile(url)
            time.sleep(2.5)
            _last_opened_app = "chrome"
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

            # Focus the target app first
            if _last_opened_app:
                _, settings = get_app_category(_last_opened_app)
                focused = focus_app_by_category(_last_opened_app)
                if not focused:
                    time.sleep(settings["focus_delay"])
                    focused = focus_app_by_category(_last_opened_app)
                    if not focused:
                        log_action("KEY", f"Could not focus {_last_opened_app}", "ERROR")
                        return False, f"Could not focus {_last_opened_app}"
                time.sleep(settings["action_delay"])

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
