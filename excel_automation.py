import contextlib
import logging
import math
import os
import re
import statistics
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

try:
    import pygetwindow as gw
except ImportError:
    gw = None

try:
    import pythoncom
    import win32com.client as win32
    import win32con
    import win32gui
    HAS_EXCEL_COM = True
except ImportError:
    pythoncom = None
    win32 = None
    win32con = None
    win32gui = None
    HAS_EXCEL_COM = False


SUPPORTED_WORKBOOK_EXTENSIONS = (".xlsx", ".xls", ".xlsm", ".csv")

EXCEL_FUNCTION_ALIASES = {
    "sum": "SUM",
    "total": "SUM",
    "average": "AVERAGE",
    "avg": "AVERAGE",
    "mean": "AVERAGE",
    "min": "MIN",
    "minimum": "MIN",
    "max": "MAX",
    "maximum": "MAX",
    "count": "COUNT",
    "counta": "COUNTA",
    "median": "MEDIAN",
    "mode": "MODE",
    "product": "PRODUCT",
    "stdev": "STDEV.S",
    "std dev": "STDEV.S",
    "standard deviation": "STDEV.S",
    "variance": "VAR.S",
    "var": "VAR.S",
}

FUNCTION_PATTERN = "|".join(
    sorted((re.escape(name) for name in EXCEL_FUNCTION_ALIASES), key=len, reverse=True)
)

CLAUSE_SPLIT_RE = re.compile(
    r"(?:\.\s+|[;\n]+|\s+\bthen\b\s+|\s+\band\b\s+(?=(?:open|load|use|select|switch|go|find|"
    r"calculate|get|show|compute|sum|total|average|avg|mean|min|max|count|counta|median|"
    r"mode|product|stdev|variance|set|edit|change|update|delete|insert|sort|filter|clear|"
    r"replace|create|rename|duplicate|copy|autofit|freeze|unfreeze|write|apply|put|save)\b))",
    re.IGNORECASE,
)

COMMAND_HINTS = (
    "workbook",
    "spreadsheet",
    ".xlsx",
    ".xls",
    ".xlsm",
    ".csv",
    "formula",
    "cell ",
    "cells ",
    "column ",
    "columns ",
    "row ",
    "rows ",
    "sheet ",
    "worksheet",
    "autofit",
    "freeze panes",
)

ACTION_KEYWORDS = (
    "sum",
    "total",
    "average",
    "mean",
    "min",
    "minimum",
    "max",
    "maximum",
    "count",
    "median",
    "mode",
    "product",
    "stdev",
    "variance",
    "sort",
    "filter",
    "replace",
    "rename",
    "delete",
    "insert",
    "edit",
    "update",
    "set",
    "save",
    "open",
    "freeze",
)

HEADER_SCAN_ROWS = 10
SEARCH_DEPTH_LIMIT = 3

EXCEL_ACTIONS = {
    "open_workbook",
    "select_sheet",
    "aggregate",
    "set_cell_value",
    "set_range_value",
    "clear_contents",
    "delete_rows",
    "sort_sheet",
    "filter_column",
    "clear_filter",
    "save_workbook",
}

EXCEL_ACTION_REQUIRED_FIELDS = {
    "open_workbook": {"reference"},
    "select_sheet": {"sheet"},
    "aggregate": {"function", "source_type", "source"},
    "set_cell_value": {"target", "value"},
    "set_range_value": {"target", "value"},
    "clear_contents": {"target"},
    "delete_rows": {"start_row", "end_row"},
    "sort_sheet": {"column", "order"},
    "filter_column": {"column", "operator", "value"},
}


@dataclass
class ExcelOperation:
    action: str
    params: dict[str, Any] = field(default_factory=dict)
    description: str = ""
    start: int = 0


@dataclass
class ExcelCommandResult:
    handled: bool
    actions: list[dict[str, Any]]
    results: list[dict[str, Any]]


@dataclass
class ResolvedColumn:
    name: str
    index: int
    letter: str
    header_row: int | None = None


def log_excel(action_type: str, details: str, status: str = "INFO") -> None:
    timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
    level = getattr(logging, status.upper(), logging.INFO)
    message = f"[EXCEL::{action_type}] {details}"
    logging.log(level, message)
    print(f"[{timestamp}] {message}")


def normalize_text(value: Any) -> str:
    text = str(value or "").strip().lower()
    return re.sub(r"[^a-z0-9]+", "", text)


def clean_phrase(text: str | None) -> str:
    cleaned = str(text or "").strip().strip("\"'")
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip(" .,:;")


def is_excel_command(command: str) -> bool:
    text = command.lower()
    if any(hint in text for hint in COMMAND_HINTS):
        return True
    if re.search(r"\bopen\s+.+\s+excel\b", text):
        return True
    if "excel" in text and any(keyword in text for keyword in ACTION_KEYWORDS):
        return True
    if any(keyword in text for keyword in ("cell ", "column ", "sheet ", "row ")):
        if any(action in text for action in ACTION_KEYWORDS):
            return True
    return False


def column_letter_to_index(letter: str) -> int:
    value = 0
    for char in letter.upper():
        if not ("A" <= char <= "Z"):
            raise ValueError(f"Invalid column letter: {letter}")
        value = value * 26 + (ord(char) - ord("A") + 1)
    return value


def index_to_column_letter(index: int) -> str:
    if index < 1:
        raise ValueError("Excel columns start at 1")

    letters: list[str] = []
    current = index
    while current:
        current, remainder = divmod(current - 1, 26)
        letters.append(chr(ord("A") + remainder))
    return "".join(reversed(letters))


def is_cell_address(value: str) -> bool:
    return bool(re.fullmatch(r"[A-Za-z]{1,3}\d+", clean_phrase(value)))


def is_range_address(value: str) -> bool:
    return bool(re.fullmatch(r"[A-Za-z]{1,3}\d+:[A-Za-z]{1,3}\d+", clean_phrase(value)))


def is_column_reference(value: str) -> bool:
    cleaned = clean_phrase(value)
    return bool(re.fullmatch(r"[A-Za-z]{1,3}", cleaned) or re.fullmatch(r"[A-Za-z]{1,3}:[A-Za-z]{1,3}", cleaned))


def format_value(value: Any) -> str:
    if isinstance(value, float):
        rounded = round(value, 6)
        if rounded.is_integer():
            return str(int(rounded))
        return f"{rounded:.6f}".rstrip("0").rstrip(".")
    return str(value)


def parse_literal_value(text: str) -> Any:
    cleaned = clean_phrase(text)
    lower = cleaned.lower()
    if cleaned.startswith("="):
        return cleaned
    if lower in {"true", "yes"}:
        return True
    if lower in {"false", "no"}:
        return False
    if lower in {"blank", "empty", "null", "none"}:
        return ""
    if re.fullmatch(r"-?\d+", cleaned):
        return int(cleaned)
    if re.fullmatch(r"-?\d+\.\d+", cleaned):
        return float(cleaned)
    return cleaned


def iter_walk_limited(root: Path, max_depth: int = SEARCH_DEPTH_LIMIT):
    base_parts = len(root.parts)
    for dirpath, dirnames, filenames in os.walk(root):
        current = Path(dirpath)
        depth = len(current.parts) - base_parts
        if depth >= max_depth:
            dirnames[:] = []
        yield current, filenames


def get_search_roots() -> list[Path]:
    roots: list[Path] = []
    cwd = Path.cwd()
    home = Path.home()
    candidates = [cwd, home / "Downloads", home / "Desktop", home / "Documents"]

    seen: set[str] = set()
    for candidate in candidates:
        try:
            resolved = str(candidate.resolve())
        except OSError:
            continue
        if resolved not in seen and candidate.exists():
            seen.add(resolved)
            roots.append(candidate)
    return roots


def find_workbook_path(reference: str | None) -> Path | None:
    cleaned = clean_phrase(reference)
    if not cleaned:
        return None

    direct = Path(cleaned)
    if direct.exists():
        return direct.resolve()
    # If we were given an absolute path but it doesn't exist, it might be missing
    # the workbook extension (common when the parser/LLM extracts a stem).
    # Try common Excel extensions before giving up.
    if direct.is_absolute() and not direct.exists():
        if direct.suffix.lower() not in SUPPORTED_WORKBOOK_EXTENSIONS:
            for ext in SUPPORTED_WORKBOOK_EXTENSIONS:
                candidate = direct.with_suffix(ext)
                if candidate.exists():
                    return candidate.resolve()
        return None

    expected_names: set[str] = set()
    lower = cleaned.lower()
    if any(lower.endswith(ext) for ext in SUPPORTED_WORKBOOK_EXTENSIONS):
        expected_names.add(Path(cleaned).name.lower())
    else:
        for ext in SUPPORTED_WORKBOOK_EXTENSIONS:
            expected_names.add(f"{cleaned.lower()}{ext}")

    expected_stem = Path(cleaned).stem.lower()
    for root in get_search_roots():
        for current_dir, filenames in iter_walk_limited(root):
            filename_map = {name.lower(): name for name in filenames}
            for expected_name in expected_names:
                if expected_name in filename_map:
                    return (current_dir / filename_map[expected_name]).resolve()

            for name in filenames:
                lowered = name.lower()
                if lowered.endswith(SUPPORTED_WORKBOOK_EXTENSIONS) and Path(name).stem.lower() == expected_stem:
                    return (current_dir / name).resolve()
    return None


def normalize_function_name(function_name: str) -> str:
    normalized = clean_phrase(function_name).lower()
    if normalized not in EXCEL_FUNCTION_ALIASES:
        raise ValueError(f"Unsupported Excel function '{function_name}'.")
    return normalized


@contextlib.contextmanager
def excel_com_session():
    if not HAS_EXCEL_COM:
        raise RuntimeError("Excel automation requires pywin32 and Microsoft Excel on Windows.")
    if pythoncom is None:
        raise RuntimeError("Excel COM support is unavailable.")

    pythoncom.CoInitialize()
    try:
        yield
    finally:
        pythoncom.CoUninitialize()


def parse_excel_actions_from_command(command: str) -> list[dict[str, Any]]:
    parser = ExcelCommandParser()
    operations = parser.parse(command)
    return [{"action": operation.action, **operation.params} for operation in operations]


def validate_excel_actions(actions: list[dict[str, Any]]) -> tuple[bool, str]:
    if not isinstance(actions, list) or not actions:
        return False, "No Excel actions were produced."

    for index, action in enumerate(actions, start=1):
        if not isinstance(action, dict):
            return False, f"Action {index} is not an object."

        action_name = action.get("action")
        if action_name not in EXCEL_ACTIONS:
            return False, f"Action {index} has unsupported Excel action '{action_name}'."

        required_fields = EXCEL_ACTION_REQUIRED_FIELDS.get(action_name, set())
        missing = [field for field in required_fields if action.get(field) in (None, "")]
        if missing:
            return False, f"Action {index} is missing required field(s): {', '.join(missing)}."

    return True, ""


class ExcelCommandParser:
    def parse(self, command: str) -> list[ExcelOperation]:
        clauses = self._split_clauses(command)
        operations: list[ExcelOperation] = []
        current_workbook: str | None = None
        current_sheet: str | None = None
        offset = 0

        for clause in clauses:
            start = command.lower().find(clause.lower(), offset)
            if start == -1:
                start = offset
            offset = start + len(clause)

            parsed = self._parse_clause(clause, start)
            if not parsed:
                continue

            for operation in parsed:
                if operation.action == "open_workbook":
                    current_workbook = operation.params.get("reference")
                elif operation.action == "select_sheet":
                    current_sheet = operation.params.get("sheet")
                else:
                    if not operation.params.get("workbook"):
                        operation.params["workbook"] = current_workbook
                    if not operation.params.get("sheet"):
                        operation.params["sheet"] = current_sheet
                operations.append(operation)

        operations.sort(key=lambda item: item.start)
        return operations

    def _split_clauses(self, command: str) -> list[str]:
        return [clean_phrase(part) for part in CLAUSE_SPLIT_RE.split(command) if clean_phrase(part)]

    def _parse_clause(self, clause: str, start: int) -> list[ExcelOperation]:
        operations: list[ExcelOperation] = []
        sheet_name = self._extract_sheet_reference(clause)
        workbook_ref = self._extract_workbook_reference(clause)

        if self._is_sheet_selection(clause):
            if sheet_name:
                operations.append(
                    ExcelOperation(
                        action="select_sheet",
                        params={"sheet": sheet_name},
                        description=f"Select sheet {sheet_name}",
                        start=start,
                    )
                )
            return operations

        save_as = re.search(r"\bsave(?:\s+(?:the\s+)?workbook)?\s+as\s+(.+)$", clause, re.IGNORECASE)
        if save_as:
            operations.append(
                ExcelOperation(
                    action="save_workbook",
                    params={"path": clean_phrase(save_as.group(1))},
                    description="Save workbook as",
                    start=start,
                )
            )
            return operations

        if re.search(r"\bsave(?:\s+the\s+workbook|\s+file)?\b", clause, re.IGNORECASE):
            operations.append(
                ExcelOperation(
                    action="save_workbook",
                    params={},
                    description="Save workbook",
                    start=start,
                )
            )
            return operations

        open_ref = self._extract_open_reference(clause)
        if open_ref:
            operations.append(
                ExcelOperation(
                    action="open_workbook",
                    params={"reference": open_ref},
                    description=f"Open workbook {open_ref}",
                    start=start,
                )
            )
            if sheet_name:
                operations.append(
                    ExcelOperation(
                        action="select_sheet",
                        params={"sheet": sheet_name},
                        description=f"Select sheet {sheet_name}",
                        start=start + 1,
                    )
                )
            return operations

        return self._parse_operation_clause(clause, start, workbook_ref, sheet_name)

    def _parse_operation_clause(
        self,
        clause: str,
        start: int,
        workbook_ref: str | None,
        sheet_name: str | None,
    ) -> list[ExcelOperation]:
        operations: list[ExcelOperation] = []

        aggregate = re.search(
            rf"\b(?:find|calculate|get|show|compute|what(?:'s| is)?\s+the)?\s*({FUNCTION_PATTERN})\b(.+)$",
            clause,
            re.IGNORECASE,
        )
        if aggregate:
            source_type, source = self._extract_source_reference(aggregate.group(2))
            if source:
                operations.append(
                    ExcelOperation(
                        action="aggregate",
                        params={
                            "function": clean_phrase(aggregate.group(1)),
                            "source_type": source_type,
                            "source": source,
                            "workbook": workbook_ref,
                            "sheet": sheet_name,
                        },
                        description="Calculate aggregate",
                        start=start,
                    )
                )
                return operations

        set_range = re.search(
            r"\b(?:set|edit|change|update)\s+(?:range\s+)?([A-Za-z]{1,3}\d+:[A-Za-z]{1,3}\d+)\s+(?:to|=)\s+(.+)$",
            clause,
            re.IGNORECASE,
        )
        if set_range:
            operations.append(
                ExcelOperation(
                    action="set_range_value",
                    params={
                        "target": clean_phrase(set_range.group(1)),
                        "value": parse_literal_value(set_range.group(2)),
                        "workbook": workbook_ref,
                        "sheet": sheet_name,
                    },
                    description="Update range",
                    start=start,
                )
            )
            return operations

        set_cell = re.search(
            r"\b(?:set|edit|change|update)\s+(?:cell\s+)?([A-Za-z]{1,3}\d+)\s+(?:to|=)\s+(.+)$",
            clause,
            re.IGNORECASE,
        )
        if set_cell:
            operations.append(
                ExcelOperation(
                    action="set_cell_value",
                    params={
                        "target": clean_phrase(set_cell.group(1)),
                        "value": parse_literal_value(set_cell.group(2)),
                        "workbook": workbook_ref,
                        "sheet": sheet_name,
                    },
                    description="Update cell",
                    start=start,
                )
            )
            return operations

        clear_target = re.search(
            r"\bclear\s+(?:contents\s+of\s+)?(?:cell|range|column)?\s*(.+)$",
            clause,
            re.IGNORECASE,
        )
        # IMPORTANT: handle "clear filter" before "clear contents".
        # Otherwise "clear filter" matches the clear-contents regex.
        if re.search(r"\b(?:clear|remove|reset)\s+filter\b|\bshow\s+all\s+rows\b", clause, re.IGNORECASE):
            operations.append(
                ExcelOperation(
                    action="clear_filter",
                    params={"workbook": workbook_ref, "sheet": sheet_name},
                    description="Clear filter",
                    start=start,
                )
            )
            return operations

        if clear_target:
            operations.append(
                ExcelOperation(
                    action="clear_contents",
                    params={
                        "target": self._strip_location_phrases(clear_target.group(1)),
                        "workbook": workbook_ref,
                        "sheet": sheet_name,
                    },
                    description="Clear contents",
                    start=start,
                )
            )
            return operations

        delete_rows = re.search(r"\bdelete\s+rows?\s+(\d+)(?:\s*(?:to|-)\s*(\d+))?$", clause, re.IGNORECASE)
        delete_rows_count = re.search(r"\bdelete\s+(\d+)\s+rows?(?:\s*(?:to|-)\s*(\d+))?$", clause, re.IGNORECASE)
        if delete_rows or delete_rows_count:
            match = delete_rows or delete_rows_count
            assert match is not None
            if delete_rows_count and not delete_rows:
                count = int(match.group(1))
                end_row = int(match.group(2) or count)
                start_row = 1
            else:
                start_row = int(match.group(1))
                end_row = int(match.group(2) or match.group(1))
            operations.append(
                ExcelOperation(
                    action="delete_rows",
                    params={
                        "start_row": start_row,
                        "end_row": end_row,
                        "workbook": workbook_ref,
                        "sheet": sheet_name,
                    },
                    description="Delete rows",
                    start=start,
                )
            )
            return operations

        sort_match = re.search(
            r"\bsort(?:\s+the\s+data)?\s+by\s+column\s+(.+?)(?:\s+(ascending|descending|asc|desc|a to z|z to a))?$",
            clause,
            re.IGNORECASE,
        )
        if sort_match:
            operations.append(
                ExcelOperation(
                    action="sort_sheet",
                    params={
                        "column": self._strip_location_phrases(sort_match.group(1)),
                        "order": clean_phrase(sort_match.group(2) or "ascending"),
                        "workbook": workbook_ref,
                        "sheet": sheet_name,
                    },
                    description="Sort sheet",
                    start=start,
                )
            )
            return operations

        filter_match = re.search(
            r"\bfilter\s+column\s+(.+?)\s+(equals|=|is|contains|starts with|ends with|greater than or equal to|less than or equal to|greater than|less than|>=|<=|>|<|for)\s+(.+)$",
            clause,
            re.IGNORECASE,
        )
        if filter_match:
            operations.append(
                ExcelOperation(
                    action="filter_column",
                    params={
                        "column": self._strip_location_phrases(filter_match.group(1)),
                        "operator": clean_phrase(filter_match.group(2)),
                        "value": parse_literal_value(filter_match.group(3)),
                        "workbook": workbook_ref,
                        "sheet": sheet_name,
                    },
                    description="Filter column",
                    start=start,
                )
            )
            return operations

        return operations

    def _is_sheet_selection(self, clause: str) -> bool:
        return bool(re.search(r"\b(?:use|select|switch to|go to)\s+sheet\b", clause, re.IGNORECASE))

    def _extract_sheet_reference(self, clause: str) -> str | None:
        patterns = [
            r"\b(?:use|select|switch to|go to)\s+sheet\s+(.+)$",
            r"\b(?:on|in|from)\s+sheet\s+(.+)$",
        ]
        for pattern in patterns:
            match = re.search(pattern, clause, re.IGNORECASE)
            if match:
                return clean_phrase(match.group(1))
        return None

    def _extract_open_reference(self, clause: str) -> str | None:
        if not re.search(r"\b(?:open|load)\b", clause, re.IGNORECASE):
            return None

        path_match = re.search(r"([A-Za-z]:\\[^,;]+?\.(?:xlsx|xls|xlsm|csv))", clause, re.IGNORECASE)
        if path_match:
            return clean_phrase(path_match.group(1))

        file_match = re.search(r"([^\s,;]+?\.(?:xlsx|xls|xlsm|csv))", clause, re.IGNORECASE)
        if file_match:
            return clean_phrase(file_match.group(1))

        match = re.search(r"\b(?:open|load)\s+(.+)$", clause, re.IGNORECASE)
        if not match:
            return None

        ref = match.group(1)
        ref = re.sub(r"\bon\s+sheet\s+.+$", "", ref, flags=re.IGNORECASE)
        ref = re.sub(r"\bin\s+sheet\s+.+$", "", ref, flags=re.IGNORECASE)
        ref = re.sub(r"\b(?:workbook|spreadsheet|excel(?:\s+file)?|file)\b", "", ref, flags=re.IGNORECASE)
        ref = clean_phrase(ref)
        return ref or None

    def _extract_workbook_reference(self, clause: str) -> str | None:
        path_match = re.search(r"([A-Za-z]:\\[^,;]+?\.(?:xlsx|xls|xlsm|csv))", clause, re.IGNORECASE)
        if path_match:
            return clean_phrase(path_match.group(1))

        file_match = re.search(r"([^\s,;]+?\.(?:xlsx|xls|xlsm|csv))", clause, re.IGNORECASE)
        if file_match:
            return clean_phrase(file_match.group(1))

        workbook_match = re.search(
            r"\b(?:in|from|using)\s+(?:the\s+)?(.+?)\s+(?:workbook|spreadsheet)\b",
            clause,
            re.IGNORECASE,
        )
        if workbook_match:
            return clean_phrase(workbook_match.group(1))
        return None

    def _strip_location_phrases(self, text: str | None) -> str:
        if not text:
            return ""
        stripped = re.sub(r"\b(?:on|in|from)\s+sheet\s+.+$", "", text, flags=re.IGNORECASE)
        stripped = re.sub(
            r"\b(?:in|from|using)\s+(?:the\s+)?(.+?)\s+(?:workbook|spreadsheet)\b.*$",
            "",
            stripped,
            flags=re.IGNORECASE,
        )
        return clean_phrase(stripped)

    def _extract_source_reference(self, text: str) -> tuple[str | None, str | None]:
        source_patterns = [
            (r"\b(?:of|for)\s+(?:the\s+)?column\s+(.+)$", "column"),
            (r"\b(?:of|for)\s+(?:the\s+)?range\s+([A-Za-z]{1,3}\d+:[A-Za-z]{1,3}\d+)\b", "range"),
            (r"\b(?:of|for)\s+([A-Za-z]{1,3}\d+:[A-Za-z]{1,3}\d+)\b", "range"),
            (r"\b(?:of|for)\s+([A-Za-z]{1,3})\b", "column"),
        ]
        for pattern, source_type in source_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return source_type, self._strip_location_phrases(match.group(1))
        return None, None


class ExcelAutomation:
    def __init__(self):
        pass

    def try_handle_command(self, command: str) -> ExcelCommandResult:
        actions = parse_excel_actions_from_command(command)
        if not actions:
            return ExcelCommandResult(handled=False, actions=[], results=[])

        results: list[dict[str, Any]] = []

        for step, action in enumerate(actions, start=1):
            success = False
            note = ""
            try:
                note = self.execute_action(action)
                success = True
            except Exception as exc:
                note = str(exc)
                log_excel("ERROR", f"{action.get('action')} failed: {note}", "ERROR")

            results.append(
                {
                    "step": step,
                    "action": action.get("action", "unknown"),
                    "params": action,
                    "success": success,
                    "note": note,
                }
            )

            if not success:
                break

        return ExcelCommandResult(handled=True, actions=actions, results=results)

    def _ensure_excel(self):
        if not HAS_EXCEL_COM:
            raise RuntimeError("Excel automation requires pywin32 and Microsoft Excel on Windows.")
        if win32 is None:
            raise RuntimeError("Excel COM support is unavailable.")

        try:
            excel_app = win32.GetActiveObject("Excel.Application")
            log_excel("CONNECT", "Attached to running Excel instance")
        except Exception:
            excel_app = win32.Dispatch("Excel.Application")
            log_excel("CONNECT", "Started a new Excel instance")

        excel_app.Visible = True
        excel_app.DisplayAlerts = False
        self._bring_excel_to_front(excel_app)
        return excel_app

    def _bring_excel_to_front(self, excel_app) -> None:
        try:
            excel_app.Visible = True
            excel_app.WindowState = -4137  # xlMaximized
            excel_app.Activate()
            if getattr(excel_app, "ActiveWindow", None) is not None:
                excel_app.ActiveWindow.WindowState = -4137  # xlMaximized
            self._activate_excel_window(excel_app)
        except Exception as exc:
            log_excel("FOCUS", f"Could not force Excel to front: {exc}", "WARNING")

    def _activate_excel_window(self, excel_app) -> None:
        if gw is None:
            return

        titles = {"excel"}
        try:
            for workbook in excel_app.Workbooks:
                titles.add(normalize_text(workbook.Name))
                titles.add(normalize_text(Path(workbook.Name).stem))
        except Exception:
            pass

        for window in gw.getAllWindows():
            title = normalize_text(window.title)
            if not title:
                continue
            if not any(keyword and keyword in title for keyword in titles):
                continue

            try:
                if win32gui is not None and win32con is not None:
                    hwnd = int(window._hWnd)
                    if win32gui.IsIconic(hwnd):
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                        time.sleep(0.2)
                    else:
                        win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
                    win32gui.SetForegroundWindow(hwnd)
                    return
            except Exception:
                pass

            try:
                if window.isMinimized:
                    window.restore()
                    time.sleep(0.2)
                window.activate()
                return
            except Exception:
                continue

        try:
            hwnd = int(getattr(excel_app, "Hwnd", 0) or 0)
            if hwnd and win32gui is not None and win32con is not None:
                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                else:
                    win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
                win32gui.SetForegroundWindow(hwnd)
        except Exception as exc:
            log_excel("FOCUS", f"Fallback window activation failed: {exc}", "WARNING")

    def _get_workbook(self, reference: str | None = None):
        excel = self._ensure_excel()

        if reference:
            cleaned = clean_phrase(reference)
            for workbook in excel.Workbooks:
                if self._workbook_matches(workbook, cleaned):
                    workbook.Activate()
                    self._bring_excel_to_front(excel)
                    log_excel("WORKBOOK", f"Using open workbook '{workbook.Name}'")
                    return workbook

            path = find_workbook_path(cleaned)
            if not path:
                raise FileNotFoundError(f"Workbook '{cleaned}' was not found in common folders.")

            workbook = excel.Workbooks.Open(str(path))
            workbook.Activate()
            self._bring_excel_to_front(excel)
            log_excel("WORKBOOK", f"Opened workbook '{workbook.Name}'")
            time.sleep(0.5)
            return workbook

        if excel.Workbooks.Count == 0:
            raise RuntimeError("No Excel workbook is open. Mention a workbook name or path first.")

        workbook = excel.ActiveWorkbook
        if workbook is None:
            workbook = excel.Workbooks(1)
        workbook.Activate()
        self._bring_excel_to_front(excel)
        return workbook

    def _workbook_matches(self, workbook, reference: str) -> bool:
        if normalize_text(workbook.Name) == normalize_text(reference):
            return True
        if normalize_text(Path(workbook.Name).stem) == normalize_text(Path(reference).stem):
            return True
        try:
            if normalize_text(str(workbook.FullName)) == normalize_text(reference):
                return True
        except Exception:
            return False
        return False

    def _get_sheet(self, workbook, sheet_name: str | None = None):
        if sheet_name:
            target = normalize_text(sheet_name)
            for sheet in workbook.Worksheets:
                if normalize_text(sheet.Name) == target:
                    sheet.Activate()
                    self._bring_excel_to_front(workbook.Parent)
                    log_excel("SHEET", f"Using sheet '{sheet.Name}'")
                    return sheet
            raise ValueError(f"Sheet '{sheet_name}' was not found in workbook '{workbook.Name}'.")

        sheet = workbook.ActiveSheet
        sheet.Activate()
        self._bring_excel_to_front(workbook.Parent)
        return sheet

    def _used_range_bounds(self, sheet) -> tuple[int, int, int, int]:
        used = sheet.UsedRange
        first_row = int(used.Row)
        first_col = int(used.Column)
        last_row = first_row + int(used.Rows.Count) - 1
        last_col = first_col + int(used.Columns.Count) - 1
        return first_row, first_col, max(last_row, first_row), max(last_col, first_col)

    def execute_action(self, action: dict[str, Any]) -> str:
        with excel_com_session():
            return self._execute_action_internal(action)

    def _execute_action_internal(self, action: dict[str, Any]) -> str:
        action_name = action.get("action", "")
        params = action

        if action_name == "open_workbook":
            workbook = self._get_workbook(params.get("reference"))
            return f"Opened workbook '{workbook.Name}'."

        if action_name == "select_sheet":
            workbook = self._get_workbook(params.get("workbook"))
            sheet = self._get_sheet(workbook, params.get("sheet"))
            return f"Selected sheet '{sheet.Name}'."

        if action_name == "aggregate":
            workbook = self._get_workbook(params.get("workbook"))
            sheet = self._get_sheet(workbook, params.get("sheet"))
            result = self._aggregate(sheet, params["function"], params["source_type"], params["source"])
            func_name = EXCEL_FUNCTION_ALIASES[normalize_function_name(params["function"])]
            return (
                f"{func_name} of {params['source_type']} '{params['source']}' "
                f"on sheet '{sheet.Name}' is {format_value(result)}."
            )

        if action_name == "set_cell_value":
            workbook = self._get_workbook(params.get("workbook"))
            sheet = self._get_sheet(workbook, params.get("sheet"))
            sheet.Range(params["target"]).Value = params["value"]
            return f"Updated cell {params['target']} on '{sheet.Name}' to {format_value(params['value'])} (workbook not saved yet)."

        if action_name == "set_range_value":
            workbook = self._get_workbook(params.get("workbook"))
            sheet = self._get_sheet(workbook, params.get("sheet"))
            sheet.Range(params["target"]).Value = params["value"]
            return f"Updated range {params['target']} on '{sheet.Name}' (workbook not saved yet)."

        if action_name == "clear_contents":
            workbook = self._get_workbook(params.get("workbook"))
            sheet = self._get_sheet(workbook, params.get("sheet"))
            clear_range = self._resolve_target_range(sheet, params["target"])
            sheet.Range(clear_range).ClearContents()
            return f"Cleared contents in {clear_range} on '{sheet.Name}' (workbook not saved yet)."

        if action_name == "delete_rows":
            workbook = self._get_workbook(params.get("workbook"))
            sheet = self._get_sheet(workbook, params.get("sheet"))
            sheet.Rows(f"{params['start_row']}:{params['end_row']}").Delete()
            return f"Deleted row(s) {params['start_row']}:{params['end_row']} on '{sheet.Name}' (workbook not saved yet)."

        if action_name == "sort_sheet":
            workbook = self._get_workbook(params.get("workbook"))
            sheet = self._get_sheet(workbook, params.get("sheet"))
            self._sort_sheet(sheet, params["column"], params["order"])
            return f"Sorted sheet '{sheet.Name}' by column '{params['column']}' {params['order']} (workbook not saved yet)."

        if action_name == "filter_column":
            workbook = self._get_workbook(params.get("workbook"))
            sheet = self._get_sheet(workbook, params.get("sheet"))
            self._apply_filter(sheet, params["column"], params["operator"], params["value"])
            return (
                f"Applied filter on column '{params['column']}' in '{sheet.Name}' using "
                f"{params['operator']} {format_value(params['value'])}."
            )

        if action_name == "clear_filter":
            workbook = self._get_workbook(params.get("workbook"))
            sheet = self._get_sheet(workbook, params.get("sheet"))
            self._clear_filter(sheet)
            return f"Cleared filters on '{sheet.Name}'."

        if action_name == "save_workbook":
            workbook = self._get_workbook(params.get("workbook"))
            target_path = params.get("path")
            if target_path:
                resolved = self._resolve_save_path(workbook, target_path)
                workbook.SaveAs(str(resolved))
                return f"Saved workbook as '{resolved}'."
            workbook.Save()
            return f"Saved workbook '{workbook.Name}'."

        raise ValueError(f"Unsupported Excel action '{action_name}'")

    def _resolve_save_path(self, workbook, target_path: str) -> Path:
        candidate = Path(target_path)
        if candidate.is_absolute():
            return candidate
        try:
            base_dir = Path(workbook.Path) if workbook.Path else Path.cwd()
        except Exception:
            base_dir = Path.cwd()
        return (base_dir / candidate).resolve()

    def _resolve_column(self, sheet, target: str) -> ResolvedColumn:
        cleaned = clean_phrase(target)
        if not cleaned:
            raise ValueError("Column reference is missing.")

        first_row, first_col, last_row, last_col = self._used_range_bounds(sheet)
        if is_column_reference(cleaned):
            letter = cleaned.split(":")[0].upper()
            index = column_letter_to_index(letter)
            return ResolvedColumn(name=cleaned, index=index, letter=index_to_column_letter(index))

        if cleaned.isdigit():
            index = int(cleaned)
            return ResolvedColumn(name=cleaned, index=index, letter=index_to_column_letter(index))

        wanted = normalize_text(cleaned)
        best_match: ResolvedColumn | None = None
        for row in range(first_row, min(last_row, first_row + HEADER_SCAN_ROWS - 1) + 1):
            for col in range(first_col, last_col + 1):
                value = clean_phrase(sheet.Cells(row, col).Value)
                if not value:
                    continue
                normalized = normalize_text(value)
                if normalized == wanted:
                    return ResolvedColumn(name=cleaned, index=col, letter=index_to_column_letter(col), header_row=row)
                if wanted in normalized or normalized in wanted:
                    best_match = ResolvedColumn(name=cleaned, index=col, letter=index_to_column_letter(col), header_row=row)

        if best_match:
            return best_match
        if re.fullmatch(r"[A-Za-z]{1,3}", cleaned):
            index = column_letter_to_index(cleaned)
            return ResolvedColumn(name=cleaned, index=index, letter=index_to_column_letter(index))
        raise ValueError(f"Could not resolve column '{cleaned}' on sheet '{sheet.Name}'.")

    def _get_data_range_for_column(self, sheet, column: ResolvedColumn) -> str:
        first_row, _, last_row, _ = self._used_range_bounds(sheet)
        if column.header_row is not None:
            data_start = column.header_row + 1
        else:
            data_start = self._guess_data_start_row(sheet, column.index, first_row, last_row)

        if last_row < data_start:
            data_start = last_row
        return f"{column.letter}{data_start}:{column.letter}{last_row}"

    def _guess_data_start_row(self, sheet, column_index: int, first_row: int, last_row: int) -> int:
        if last_row <= first_row:
            return first_row
        first_value = sheet.Cells(first_row, column_index).Value
        second_value = sheet.Cells(first_row + 1, column_index).Value
        if isinstance(first_value, str) and first_value.strip() and self._to_number(second_value) is not None:
            return first_row + 1
        return first_row

    def _resolve_target_range(self, sheet, target: str) -> str:
        cleaned = clean_phrase(target)
        if is_range_address(cleaned) or is_cell_address(cleaned):
            return cleaned
        column = self._resolve_column(sheet, cleaned)
        return self._get_data_range_for_column(sheet, column)

    def _iter_range_values(self, excel_range) -> list[Any]:
        values = excel_range.Value
        if isinstance(values, tuple):
            flattened: list[Any] = []
            for row in values:
                if isinstance(row, tuple):
                    flattened.extend(list(row))
                else:
                    flattened.append(row)
            return flattened
        return [values]

    def _to_number(self, value: Any) -> float | None:
        if value is None or value == "":
            return None
        if isinstance(value, bool):
            return float(int(value))
        if isinstance(value, (int, float)):
            return float(value)
        cleaned = str(value).strip().replace(",", "").replace("$", "").replace("%", "")
        if not cleaned:
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None

    def _aggregate(self, sheet, function_name: str, source_type: str, source: str) -> float:
        function_key = EXCEL_FUNCTION_ALIASES[normalize_function_name(function_name)]
        excel_range = self._resolve_source_range(sheet, source_type, source)
        raw_values = self._iter_range_values(excel_range)

        if function_key == "COUNTA":
            values = [value for value in raw_values if value not in (None, "")]
        else:
            values = [number for number in (self._to_number(value) for value in raw_values) if number is not None]

        if not values:
            raise ValueError(f"No usable values found for {function_key} in {source_type} '{source}'.")
        if function_key == "SUM":
            return float(sum(values))
        if function_key == "AVERAGE":
            return float(sum(values) / len(values))
        if function_key == "MIN":
            return float(min(values))
        if function_key == "MAX":
            return float(max(values))
        if function_key in {"COUNT", "COUNTA"}:
            return float(len(values))
        if function_key == "MEDIAN":
            return float(statistics.median(values))
        if function_key == "MODE":
            return float(statistics.multimode(values)[0])
        if function_key == "PRODUCT":
            return float(math.prod(values))
        if function_key == "STDEV.S":
            return float(statistics.stdev(values)) if len(values) > 1 else 0.0
        if function_key == "VAR.S":
            return float(statistics.variance(values)) if len(values) > 1 else 0.0
        raise ValueError(f"Aggregate '{function_name}' is not supported yet.")

    def _resolve_source_range(self, sheet, source_type: str, source: str):
        if source_type == "range":
            return sheet.Range(clean_phrase(source))
        if source_type == "column":
            column = self._resolve_column(sheet, source)
            return sheet.Range(self._get_data_range_for_column(sheet, column))
        raise ValueError(f"Unsupported source type '{source_type}'.")

    def _sort_sheet(self, sheet, column_name: str, order: str) -> None:
        column = self._resolve_column(sheet, column_name)
        first_row, first_col, last_row, last_col = self._used_range_bounds(sheet)
        sort_start_row = column.header_row or first_row
        header_mode = 1 if column.header_row else 2
        order_value = 2 if order.lower() in {"descending", "desc", "z to a"} else 1

        sort_range = sheet.Range(sheet.Cells(sort_start_row, first_col), sheet.Cells(last_row, last_col))
        key_start = (column.header_row or sort_start_row) + (1 if column.header_row else 0)
        key_range = sheet.Range(sheet.Cells(key_start, column.index), sheet.Cells(last_row, column.index))
        sort_range.Sort(Key1=key_range, Order1=order_value, Header=header_mode)

    def _apply_filter(self, sheet, column_name: str, operator: str, value: Any) -> None:
        column = self._resolve_column(sheet, column_name)
        _, first_col, _, _ = self._used_range_bounds(sheet)
        field_index = column.index - first_col + 1
        criteria = self._build_filter_criteria(operator, value)
        sheet.UsedRange.AutoFilter(Field=field_index, Criteria1=criteria)

    def _build_filter_criteria(self, operator: str, value: Any) -> str:
        op = operator.lower()
        text = str(value)
        if op in {"equals", "=", "is", "for"}:
            return text
        if op == "contains":
            return f"*{text}*"
        if op == "starts with":
            return f"{text}*"
        if op == "ends with":
            return f"*{text}"
        if op in {">", ">=", "<", "<="}:
            return f"{op}{text}"
        mapping = {
            "greater than": f">{text}",
            "less than": f"<{text}",
            "greater than or equal to": f">={text}",
            "less than or equal to": f"<={text}",
        }
        if op in mapping:
            return mapping[op]
        raise ValueError(f"Unsupported filter operator '{operator}'.")

    def _clear_filter(self, sheet) -> None:
        if sheet.FilterMode:
            sheet.ShowAllData()
        elif sheet.AutoFilterMode:
            sheet.AutoFilterMode = False


_EXCEL_AUTOMATION = ExcelAutomation()


def execute_excel_action(action: dict[str, Any]) -> tuple[bool, str]:
    try:
        note = _EXCEL_AUTOMATION.execute_action(action)
        return True, note
    except Exception as exc:
        error = str(exc)
        log_excel("ERROR", f"{action.get('action', 'unknown')} failed: {error}", "ERROR")
        return False, error


def run_excel_smoke_checks(sample_workbook: str | None = None) -> list[tuple[str, str, str]]:
    results: list[tuple[str, str, str]] = []

    if HAS_EXCEL_COM:
        results.append(("Excel COM import", "pywin32", "PASS"))
    else:
        results.append(("Excel COM import", "pywin32", "FAIL"))
        return results

    try:
        with excel_com_session():
            pass
        results.append(("Excel COM init", "CoInitialize/CoUninitialize", "PASS"))
    except Exception as exc:
        results.append(("Excel COM init", str(exc), "FAIL"))
        return results

    parser_actions = parse_excel_actions_from_command(
        "open C:\\Users\\DELL\\Downloads\\Orders_Table (2).csv and delete rows 2 to 10"
    )
    parser_ok, parser_note = validate_excel_actions(parser_actions)
    results.append(("Excel parser to DSL", parser_note or f"{len(parser_actions)} actions", "PASS" if parser_ok else "FAIL"))

    live_target = clean_phrase(sample_workbook) if sample_workbook else ""
    if not live_target:
        candidate = find_workbook_path("Orders_Table (2).csv")
        live_target = str(candidate) if candidate else ""

    if live_target:
        try:
            success, note = execute_excel_action({"action": "open_workbook", "reference": live_target})
            results.append(("Excel live open", note, "PASS" if success else ("FAIL" if sample_workbook else "WARNING")))
        except Exception as exc:
            results.append(("Excel live open", str(exc), "FAIL" if sample_workbook else "WARNING"))
    else:
        results.append(("Excel live open", "No sample workbook found", "WARNING"))

    return results


def try_handle_excel_command(command: str) -> ExcelCommandResult:
    return _EXCEL_AUTOMATION.try_handle_command(command)