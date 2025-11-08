import os
import datetime
import itertools
import re
import sqlite3
from typing import Any, Dict, List, Optional

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Path to your service account key file
SERVICE_ACCOUNT_FILE = 'sheetsense-477619-ad0eb7d32908.json'

# Scopes required for Google Sheets and Drive APIs
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def create_drive_service():
    """Create Google Drive API service"""
    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, 
        scopes=SCOPES
    )
    return build('drive', 'v3', credentials=credentials)

def create_sheets_service():
    """Create Google Sheets API service"""
    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, 
        scopes=SCOPES
    )
    return build('sheets', 'v4', credentials=credentials)

def discover_google_sheets():
    """Discover all Google Sheets in your Drive"""
    try:
        drive_service = create_drive_service()
        results = drive_service.files().list(
            q="mimeType='application/vnd.google-apps.spreadsheet'",
            fields='files(id, name, webViewLink, createdTime)'
        ).execute()
        
        sheets = results.get('files', [])
        if not sheets:
            print('No Google Sheets found.')
            return []
        
        print(f'Found {len(sheets)} Google Sheets:')
        print('-' * 80)
        
        for i, sheet in enumerate(sheets, 1):
            print(f"{i}. {sheet['name']}")
            print(f"   ID: {sheet['id']}")
            print(f"   URL: {sheet['webViewLink']}")
            print(f"   Created: {sheet['createdTime']}")
            print()
        
        return sheets
    except Exception as e:
        print(f"Error discovering sheets: {e}")
        return []


def column_index_to_letter(index: int) -> str:
    """Convert zero-based column index to Excel-style letters (0 -> A)."""
    if index < 0:
        raise ValueError("Column index cannot be negative")
    letters = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def column_letter_to_index(letter: str) -> int:
    if not letter:
        raise ValueError("Column letter must be provided")
    letter = letter.upper()
    result = 0
    for char in letter:
        if not ('A' <= char <= 'Z'):
            raise ValueError(f"Invalid column letter: {letter}")
        result = result * 26 + (ord(char) - 64)
    return result - 1


COLOR_NAME_MAP = {
    "red": {"red": 1.0, "green": 0.0, "blue": 0.0},
    "green": {"red": 0.0, "green": 0.6, "blue": 0.0},
    "blue": {"red": 0.1, "green": 0.3, "blue": 0.9},
    "yellow": {"red": 1.0, "green": 0.96, "blue": 0.4},
    "orange": {"red": 1.0, "green": 0.6, "blue": 0.0},
    "purple": {"red": 0.6, "green": 0.2, "blue": 0.8},
    "gray": {"red": 0.6, "green": 0.6, "blue": 0.6},
    "grey": {"red": 0.6, "green": 0.6, "blue": 0.6},
    "white": {"red": 1.0, "green": 1.0, "blue": 1.0},
    "black": {"red": 0.0, "green": 0.0, "blue": 0.0},
}


def normalize_color(
    color_value: str | Dict[str, float] | None,
    fallback: Dict[str, float],
) -> Dict[str, float]:
    if not color_value:
        return fallback
    if isinstance(color_value, str):
        normalized = COLOR_NAME_MAP.get(color_value.lower())
        if not normalized:
            raise ValueError(
                f"Unsupported color name '{color_value}'. "
                f"Use RGB dict or one of {list(COLOR_NAME_MAP)}"
            )
        return normalized
    if isinstance(color_value, dict):
        result: Dict[str, float] = {}
        for channel in ("red", "green", "blue"):
            if channel in color_value and color_value[channel] is not None:
                val = float(color_value[channel])
                if val > 1:
                    val = val / 255
                result[channel] = min(max(val, 0.0), 1.0)
        if not result:
            return fallback
        return result
    raise ValueError("Color must be a dict or a supported color name string")


def _normalize_header(value: str | None) -> str:
    return (value or "").strip().lower()


def find_header_index(
    headers: List[str],
    requested: str,
    sheet_label: str,
    required: bool = True,
) -> tuple[int | None, str | None]:
    """Locate a header ignoring case/spaces; optionally allow missing."""
    normalized_target = _normalize_header(requested)
    for idx, header in enumerate(headers):
        if _normalize_header(header) == normalized_target:
            return idx, header
    if required:
        raise ValueError(
            f"Column '{requested}' not found in sheet '{sheet_label}'."
        )
    return None, None


def get_sheet_tabs_and_a1(spreadsheet_id: str) -> List[Dict[str, str]]:
    """Return each subsheet title and the value stored in its A1 cell."""
    sheets_service = create_sheets_service()

    try:
        spreadsheet = sheets_service.spreadsheets().get(
            spreadsheetId=spreadsheet_id
        ).execute()
    except HttpError as exc:
        raise RuntimeError(f"Unable to read spreadsheet metadata: {exc}") from exc

    result: List[Dict[str, str]] = []
    for sheet in spreadsheet.get("sheets", []):
        sheet_props = sheet.get("properties", {})
        title = sheet_props.get("title", "Untitled")
        range_name = f"{title}!A1"

        try:
            response = sheets_service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=range_name
            ).execute()
        except HttpError as exc:
            raise RuntimeError(f"Unable to read range {range_name}: {exc}") from exc

        values = response.get("values", [])
        a1_value = values[0][0] if values and values[0] else ""
        result.append({
            "title": title,
            "sheetId": sheet_props.get("sheetId"),
            "a1_value": a1_value
        })

    return result


def list_subsheets(spreadsheet_id: str) -> List[Dict[str, str]]:
    """Expose subsheet metadata (title + A1 value) for a spreadsheet."""
    try:
        return get_sheet_tabs_and_a1(spreadsheet_id)
    except Exception as exc:
        print(f"Error listing subsheets: {exc}")
        return []


def format_subsheet_summary(subsheets: List[Dict[str, str]]) -> str:
    """Format a human-readable summary for subsheet metadata."""
    if not subsheets:
        return "No subsheets found."

    lines = []
    for sheet in subsheets:
        value_display = sheet["a1_value"] if sheet["a1_value"] else "[empty]"
        lines.append(f"- {sheet['title']}: {value_display}")
    return "\n".join(lines)


def get_sheet_properties(spreadsheet_id: str, sheet_name: str) -> Dict[str, Any]:
    sheets_service = create_sheets_service()
    metadata = sheets_service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)))"
    ).execute()
    for sheet in metadata.get("sheets", []):
        props = sheet.get("properties", {})
        if props.get("title") == sheet_name:
            return props
    raise ValueError(f"Sheet '{sheet_name}' not found in spreadsheet")


_A1_RANGE_RE = re.compile(
    r"^(?:(?P<sheet>[^!]+)!)?"
    r"(?P<start_col>[A-Za-z]+)?(?P<start_row>\d+)?"
    r"(?:\:(?P<end_col>[A-Za-z]+)?(?P<end_row>\d+)?)?$"
)

_SHEET_CELL_REF_RE = re.compile(
    r"(?P<sheet>'[^']+'|[A-Za-z0-9 _-]+)"
    r"!\$?(?P<col>[A-Za-z]+)\$?(?P<row>\d+)"
)


def a1_to_grid_range(range_cells: str, sheet_props: Dict[str, Any]) -> Dict[str, int]:
    match = _A1_RANGE_RE.match(range_cells.strip())
    if not match:
        raise ValueError(f"Invalid A1 range: {range_cells}")

    grid_props = sheet_props.get("gridProperties", {})
    total_rows = grid_props.get("rowCount", 1000)
    total_cols = grid_props.get("columnCount", 26)

    start_col = match.group("start_col")
    start_row = match.group("start_row")
    end_col = match.group("end_col")
    end_row = match.group("end_row")

    start_col_idx = column_letter_to_index(start_col) if start_col else 0
    end_col_idx = (
        column_letter_to_index(end_col) + 1 if end_col else
        (start_col_idx + 1 if start_col else total_cols)
    )

    start_row_idx = int(start_row) - 1 if start_row else 0
    end_row_idx = (
        int(end_row) if end_row else
        (start_row_idx + 1 if start_row else total_rows)
    )

    return {
        "sheetId": sheet_props.get("sheetId"),
        "startRowIndex": max(start_row_idx, 0),
        "endRowIndex": min(end_row_idx, total_rows),
        "startColumnIndex": max(start_col_idx, 0),
        "endColumnIndex": min(end_col_idx, total_cols),
    }


def get_sheet_headers(spreadsheet_id: str, sheet_name: str) -> List[str]:
    """Fetch the header row (first row) for a given sheet."""
    sheets_service = create_sheets_service()
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!1:1"
    ).execute()
    values = result.get('values', [])
    return values[0] if values else []


def get_cell_value(spreadsheet_id: str, sheet_name: str, cell: str) -> str:
    sheets_service = create_sheets_service()
    range_ref = f"{sheet_name}!{cell}"
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_ref
    ).execute()
    values = result.get('values', [])
    if values and values[0]:
        return str(values[0][0])
    return ""


def _fetch_sheet_values(spreadsheet_id: str, sheet_name: str) -> List[List[str]]:
    """Fetch up to the first 1000 rows/columns from a sheet."""
    sheets_service = create_sheets_service()
    range_all = f"{sheet_name}!A1:ZZZ1000"
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_all
    ).execute()
    return result.get('values', [])


def _get_range_values(spreadsheet_id: str, sheet_name: str, range_cells: str) -> List[List[str]]:
    sheets_service = create_sheets_service()
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!{range_cells}"
    ).execute()
    return result.get('values', [])


def _iter_cells(values: List[List[str]]) -> List[str]:
    for row in values:
        for cell in row:
            yield cell


def _coerce_number(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _is_non_empty(value: Any) -> bool:
    return bool(str(value).strip()) if value is not None else False


def sum_range(spreadsheet_id: str, sheet_name: str, range_cells: str) -> float:
    values = _get_range_values(spreadsheet_id, sheet_name, range_cells)
    nums = [_coerce_number(val) for val in _iter_cells(values)]
    return sum(num for num in nums if num is not None)


def average_range(spreadsheet_id: str, sheet_name: str, range_cells: str) -> float:
    values = _get_range_values(spreadsheet_id, sheet_name, range_cells)
    nums = [_coerce_number(val) for val in _iter_cells(values)]
    filtered = [num for num in nums if num is not None]
    if not filtered:
        raise ValueError("No numeric values found for AVERAGE")
    return sum(filtered) / len(filtered)


def count_range(spreadsheet_id: str, sheet_name: str, range_cells: str) -> int:
    values = _get_range_values(spreadsheet_id, sheet_name, range_cells)
    nums = [_coerce_number(val) for val in _iter_cells(values)]
    return sum(1 for num in nums if num is not None)


def counta_range(spreadsheet_id: str, sheet_name: str, range_cells: str) -> int:
    values = _get_range_values(spreadsheet_id, sheet_name, range_cells)
    return sum(1 for val in _iter_cells(values) if _is_non_empty(val))


def min_range(spreadsheet_id: str, sheet_name: str, range_cells: str) -> float:
    values = _get_range_values(spreadsheet_id, sheet_name, range_cells)
    nums = [_coerce_number(val) for val in _iter_cells(values)]
    filtered = [num for num in nums if num is not None]
    if not filtered:
        raise ValueError("No numeric values found for MIN")
    return min(filtered)


def max_range(spreadsheet_id: str, sheet_name: str, range_cells: str) -> float:
    values = _get_range_values(spreadsheet_id, sheet_name, range_cells)
    nums = [_coerce_number(val) for val in _iter_cells(values)]
    filtered = [num for num in nums if num is not None]
    if not filtered:
        raise ValueError("No numeric values found for MAX")
    return max(filtered)


def _parse_criterion(criterion: str) -> tuple[str, str]:
    criterion = criterion.strip()
    for op in ("<>", "!=", ">=", "<=", ">", "<", "="):
        if criterion.startswith(op):
            return op, criterion[len(op):].strip()
    return "=", criterion


def _compare(value: Any, op: str, target: str) -> bool:
    value_num = _coerce_number(value)
    target_num = _coerce_number(target)
    if value_num is not None and target_num is not None:
        left, right = value_num, target_num
    else:
        left = str(value).strip().lower()
        right = target.strip().lower()

    if op in ("=", "=="):
        return left == right
    if op in ("<>", "!="):
        return left != right
    if op == ">":
        return left > right
    if op == "<":
        return left < right
    if op == ">=":
        return left >= right
    if op == "<=":
        return left <= right
    return False


def sumif_range(
    spreadsheet_id: str,
    sheet_name: str,
    criteria_range: str,
    criterion: str,
    sum_range_cells: Optional[str] = None,
) -> float:
    crit_values = _get_range_values(spreadsheet_id, sheet_name, criteria_range)
    sum_values = (
        _get_range_values(spreadsheet_id, sheet_name, sum_range_cells)
        if sum_range_cells else crit_values
    )
    crit_flat = list(_iter_cells(crit_values))
    sum_flat = list(_iter_cells(sum_values))
    op, target = _parse_criterion(criterion)
    total = 0.0
    for idx, value in enumerate(crit_flat):
        if idx >= len(sum_flat):
            break
        if _compare(value, op, target):
            num = _coerce_number(sum_flat[idx])
            if num is not None:
                total += num
    return total


def countif_range(
    spreadsheet_id: str,
    sheet_name: str,
    criteria_range: str,
    criterion: str,
) -> int:
    crit_values = _get_range_values(spreadsheet_id, sheet_name, criteria_range)
    crit_flat = list(_iter_cells(crit_values))
    op, target = _parse_criterion(criterion)
    return sum(1 for value in crit_flat if _compare(value, op, target))


def if_condition(condition: str, value_if_true: str, value_if_false: str) -> str:
    match = re.match(r"\s*(.+?)\s*(<=|>=|<>|!=|=|==|>|<)\s*(.+)\s*", condition)
    if not match:
        raise ValueError("Condition must be in the form 'value1 > value2'")
    left_raw, op, right_raw = match.groups()
    result = _compare(left_raw, op, right_raw)
    return value_if_true if result else value_if_false


def match_position(
    spreadsheet_id: str,
    sheet_name: str,
    lookup_value: Any,
    lookup_range: str,
    match_type: str = "exact",
) -> int:
    """Return 1-based index within the lookup_range matching lookup_value."""
    values = list(_iter_cells(_get_range_values(spreadsheet_id, sheet_name, lookup_range)))
    if not values:
        raise ValueError(f"Lookup range '{lookup_range}' is empty")

    def normalize(val: Any) -> str:
        return str(val).strip().lower()

    lookup_normalized = normalize(lookup_value)
    numbers = [_coerce_number(v) for v in values]

    if match_type in ("exact", "0", 0, None):
        for idx, value in enumerate(values, start=1):
            if normalize(value) == lookup_normalized:
                return idx
        raise ValueError(f"Value '{lookup_value}' not found in range {lookup_range}")

    sorted_values = [
        (idx + 1, num) for idx, num in enumerate(numbers) if num is not None
    ]
    if not sorted_values:
        raise ValueError("Lookup range must contain numeric values for approximate MATCH")

    lookup_num = _coerce_number(lookup_value)
    if lookup_num is None:
        raise ValueError("Lookup value must be numeric for approximate MATCH")

    if str(match_type) in ("-1", "less_than"):
        # Largest value less than or equal to lookup_num
        eligible = [pair for pair in sorted_values if pair[1] <= lookup_num]
        if not eligible:
            raise ValueError("No values less than or equal to lookup value")
        return max(eligible, key=lambda pair: pair[1])[0]
    else:
        # Smallest value greater than or equal
        eligible = [pair for pair in sorted_values if pair[1] >= lookup_num]
        if not eligible:
            raise ValueError("No values greater than or equal to lookup value")
        return min(eligible, key=lambda pair: pair[1])[0]


def index_match_lookup(
    spreadsheet_id: str,
    sheet_name: str,
    array_range: str,
    row_lookup_value: Any | None = None,
    row_lookup_range: str | None = None,
    column_lookup_value: Any | None = None,
    column_lookup_range: str | None = None,
    row_num: int | None = None,
    col_num: int | None = None,
    match_type: str = "exact",
) -> Any:
    """Flexible INDEX/MATCH helper supporting row/column lookups."""
    array_values = _get_range_values(spreadsheet_id, sheet_name, array_range)
    if not array_values:
        raise ValueError(f"Array range '{array_range}' returned no data")

    def resolve_index(
        provided_num: int | None,
        lookup_value: Any | None,
        lookup_range: str | None,
        axis_label: str,
    ) -> int:
        if provided_num is not None:
            return provided_num
        if lookup_value is None or not lookup_range:
            raise ValueError(f"{axis_label} lookup parameters missing")
        return match_position(
            spreadsheet_id,
            sheet_name,
            lookup_value,
            lookup_range,
            match_type=match_type,
        )

    total_columns = max(len(row) for row in array_values)
    resolved_row = resolve_index(
        row_num, row_lookup_value, row_lookup_range, "Row"
    )
    resolved_col = resolve_index(
        col_num, column_lookup_value, column_lookup_range, "Column"
    )

    if resolved_row < 1 or resolved_col < 1:
        raise ValueError("Row and column indexes must be >= 1")
    if resolved_row > len(array_values):
        raise ValueError("Row index out of bounds for array range")
    if resolved_col > total_columns:
        raise ValueError("Column index out of bounds for array range")

    target_row = array_values[resolved_row - 1]
    return target_row[resolved_col - 1] if len(target_row) >= resolved_col else ""


def sumproduct_range(
    spreadsheet_id: str,
    sheet_name: str,
    ranges: List[str],
) -> float:
    if len(ranges) < 2:
        raise ValueError("SUMPRODUCT requires at least two ranges")
    flattened = [
        [_coerce_number(val) or 0.0 for val in _iter_cells(_get_range_values(spreadsheet_id, sheet_name, rng))]
        for rng in ranges
    ]
    lengths = {len(arr) for arr in flattened}
    if len(lengths) != 1:
        raise ValueError("All SUMPRODUCT ranges must have the same size")
    total = 0.0
    for values in zip(*flattened):
        product = 1.0
        for val in values:
            product *= val
        total += product
    return total


def len_text(value: str) -> int:
    return len(value or "")


def left_text(value: str, num_chars: int | None = None) -> str:
    if num_chars is None:
        num_chars = 1
    return (value or "")[:max(num_chars, 0)]


def right_text(value: str, num_chars: int | None = None) -> str:
    if num_chars is None:
        num_chars = 1
    return (value or "")[-max(num_chars, 0):] if num_chars else ""


def mid_text(value: str, start_num: int, num_chars: int) -> str:
    start_idx = max(start_num - 1, 0)
    end_idx = start_idx + max(num_chars, 0)
    return (value or "")[start_idx:end_idx]


def _convert_column_spec(column_spec: Any, headers: List[str]) -> int:
    if isinstance(column_spec, int):
        return column_spec - 1
    idx, _ = find_header_index(headers, str(column_spec), "range")
    return idx


def sort_range_data(
    spreadsheet_id: str,
    sheet_name: str,
    range_cells: str,
    instructions: List[Dict[str, Any]],
) -> List[List[str]]:
    values = _get_range_values(spreadsheet_id, sheet_name, range_cells)
    if not values:
        return []
    headers = values[0]
    rows = values[1:]
    for instruction in reversed(instructions or []):
        col_idx = _convert_column_spec(instruction.get("column", 1), headers)
        ascending = instruction.get("ascending", True)
        rows.sort(key=lambda row: row[col_idx] if len(row) > col_idx else "", reverse=not ascending)
    return [headers] + rows


def _row_matches_conditions(row: List[str], headers: List[str], conditions: List[Dict[str, Any]]) -> bool:
    for condition in conditions or []:
        col_idx = _convert_column_spec(condition.get("column", 1), headers)
        value = row[col_idx] if len(row) > col_idx else ""
        criterion = condition.get("criterion", "")
        op, target = _parse_criterion(criterion)
        if not _compare(value, op, target):
            return False
    return True


def filter_range_data(
    spreadsheet_id: str,
    sheet_name: str,
    range_cells: str,
    conditions: List[Dict[str, Any]],
) -> List[List[str]]:
    values = _get_range_values(spreadsheet_id, sheet_name, range_cells)
    if not values:
        return []
    headers = values[0]
    filtered = [
        row for row in values[1:] if _row_matches_conditions(row, headers, conditions)
    ]
    return [headers] + filtered


def unique_range_data(
    spreadsheet_id: str,
    sheet_name: str,
    range_cells: str,
) -> List[List[str]]:
    values = _get_range_values(spreadsheet_id, sheet_name, range_cells)
    if not values:
        return []
    headers = values[0]
    seen = set()
    unique_rows = []
    for row in values[1:]:
        key = tuple(row)
        if key in seen:
            continue
        seen.add(key)
        unique_rows.append(row)
    return [headers] + unique_rows


def format_table(values: List[List[Any]], max_rows: int = 10) -> str:
    if not values:
        return "[no data]"
    lines = []
    for row in values[:max_rows]:
        lines.append(" | ".join(str(cell) for cell in row))
    remaining = len(values) - max_rows
    if remaining > 0:
        lines.append(f"... (+{remaining} more row(s))")
    return "\n".join(lines)


def apply_conditional_formatting(
    spreadsheet_id: str,
    sheet_name: str,
    range_cells: str,
    rule: Dict[str, Any],
) -> str:
    sheets_service = create_sheets_service()
    sheet_props = get_sheet_properties(spreadsheet_id, sheet_name)
    grid_range = a1_to_grid_range(range_cells, sheet_props)

    rule_type = (rule.get("type") or "text_contains").lower()
    condition_type_map = {
        "text_contains": "TEXT_CONTAINS",
        "text_equals": "TEXT_EQ",
        "equals": "TEXT_EQ",
        "number_greater": "NUMBER_GREATER",
        "greater_than": "NUMBER_GREATER",
        "number_greater_eq": "NUMBER_GREATER_THAN_EQ",
        "number_less": "NUMBER_LESS",
        "less_than": "NUMBER_LESS",
        "number_eq": "NUMBER_EQ",
        "custom_formula": "CUSTOM_FORMULA",
    }
    condition_type = condition_type_map.get(rule_type)
    if not condition_type:
        raise ValueError(f"Unsupported conditional format type: {rule_type}")

    def substitute_external_refs(expression: str) -> str:
        def replace(match: re.Match) -> str:
            raw_sheet = match.group("sheet")
            ref_sheet = raw_sheet.strip("'")
            if ref_sheet == sheet_name:
                return match.group(0)
            col = match.group("col")
            row = match.group("row")
            value = get_cell_value(spreadsheet_id, ref_sheet, f"{col}{row}")
            numeric = _coerce_number(value)
            if numeric is not None:
                return str(numeric)
            escaped = value.replace('"', '""')
            return f'"{escaped}"'
        return _SHEET_CELL_REF_RE.sub(replace, expression)

    condition: Dict[str, Any] = {"type": condition_type}
    if condition_type == "CUSTOM_FORMULA":
        formula = rule.get("formula")
        if not formula:
            raise ValueError("Custom formula rule requires 'formula'")
        formula = substitute_external_refs(formula)
        condition["values"] = [{"userEnteredValue": formula}]
    else:
        if "value" not in rule:
            raise ValueError("This rule requires a 'value' field")
        value = rule["value"]
        if isinstance(value, (int, float)) and "valueRelative" in rule:
            raise ValueError("Use either 'value' or 'valueRelative', not both")
        if isinstance(value, str) and value.startswith("="):
            formula_value = substitute_external_refs(value)
            condition["values"] = [{"userEnteredValue": formula_value}]
        else:
            condition["values"] = [{"userEnteredValue": str(value)}]

    formatting: Dict[str, Any] = {}
    background = normalize_color(
        rule.get("backgroundColor"),
        {"red": 1.0, "green": 0.9, "blue": 0.6},
    )
    formatting["backgroundColor"] = background
    text_format: Dict[str, Any] = {}
    if rule.get("bold") is not None:
        text_format["bold"] = bool(rule["bold"])
    if rule.get("italic") is not None:
        text_format["italic"] = bool(rule["italic"])
    if rule.get("textColor"):
        text_format["foregroundColor"] = normalize_color(
            rule.get("textColor"),
            {"red": 0.0, "green": 0.0, "blue": 0.0},
        )
    if text_format:
        formatting["textFormat"] = text_format

    request_body = {
        "requests": [
            {
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [grid_range],
                        "booleanRule": {
                            "condition": condition,
                            "format": formatting,
                        },
                    },
                    "index": rule.get("index", 0),
                }
            }
        ]
    }

    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=request_body
    ).execute()

    return (
        f"Applied conditional formatting on {sheet_name}!{range_cells} "
        f"with rule '{condition_type}'"
    )


def arrayformula_write(
    spreadsheet_id: str,
    sheet_name: str,
    destination: str,
    formula_body: str,
) -> str:
    sheets_service = create_sheets_service()
    formula_body = formula_body.strip()
    if formula_body.startswith("="):
        formula_body = formula_body[1:].lstrip()
    if not formula_body.upper().startswith("ARRAYFORMULA("):
        formula_body = f"ARRAYFORMULA({formula_body})"
    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!{destination}",
        valueInputOption='USER_ENTERED',
        body={'values': [[f"={formula_body}"]]}
    ).execute()
    return f"Wrote ={formula_body} to {sheet_name}!{destination}"


def query_range_data(
    spreadsheet_id: str,
    sheet_name: str,
    range_cells: str,
    query_string: str,
    headers: int = 1,
) -> List[List[str]]:
    values = _get_range_values(spreadsheet_id, sheet_name, range_cells)
    if not values:
        return []
    header_row = values[0] if headers else [f"Col{i+1}" for i in range(len(values[0]))]
    sanitized = []
    mapping = {}
    for idx, header in enumerate(header_row):
        safe = re.sub(r"[^0-9a-zA-Z_]", "_", header or f"Col{idx+1}")
        if not safe:
            safe = f"Col{idx+1}"
        count = 1
        original_safe = safe
        while safe in sanitized:
            safe = f"{original_safe}_{count}"
            count += 1
        sanitized.append(safe)
        mapping[header] = safe

    # Replace column names in query with sanitized tokens (simple replacement)
    normalized_query = query_string
    for original, safe in mapping.items():
        if original:
            normalized_query = re.sub(rf"\b{re.escape(original)}\b", safe, normalized_query)

    conn = sqlite3.connect(":memory:")
    try:
        placeholders = ",".join("?" for _ in sanitized)
        conn.execute(
            f"CREATE TABLE data ({', '.join(f'{col} TEXT' for col in sanitized)})"
        )
        for row in values[headers:]:
            padded = row + [""] * (len(sanitized) - len(row))
            conn.execute(f"INSERT INTO data VALUES ({placeholders})", padded)
        cursor = conn.execute(normalized_query.replace("`", ""))
        result_rows = cursor.fetchall()
        column_names = [description[0] for description in cursor.description]
        return [column_names] + [list(row) for row in result_rows]
    finally:
        conn.close()


def today_value() -> str:
    return datetime.date.today().isoformat()


def now_value() -> str:
    return datetime.datetime.now().isoformat(timespec="seconds")


def join_values(delimiter: str, values: List[str] | None = None, range_values: List[List[str]] | None = None) -> str:
    items: List[str] = []
    if values:
        items.extend(values)
    if range_values:
        items.extend(str(val) for val in _iter_cells(range_values))
    return (delimiter or "").join(str(item) for item in items if item is not None)


def split_text(text: str, delimiter: str) -> List[str]:
    return (text or "").split(delimiter or ",")


def merge_sheet_column_by_key(
    spreadsheet_id: str,
    source_sheet: str,
    source_key_column: str,
    source_value_column: str,
    target_sheet: str,
    target_key_column: str | None = None,
    target_value_column: str | None = None,
) -> int:
    """Perform a VLOOKUP-style merge from one sheet to another.

    Args:
        spreadsheet_id: Spreadsheet identifier.
        source_sheet: Sheet providing the lookup table.
        source_key_column: Header name for the unique key (e.g., Email).
        source_value_column: Header to pull data from (e.g., Food).
        target_sheet: Sheet that will receive the merged column.
        target_key_column: Header name containing the key in the target sheet
            (defaults to source_key_column).
        target_value_column: Header that should receive the merged value
            (defaults to source_value_column).

    Returns:
        Number of target rows updated.
    """
    target_key_column = target_key_column or source_key_column
    target_value_column = target_value_column or source_value_column

    sheets_service = create_sheets_service()

    source_values = _fetch_sheet_values(spreadsheet_id, source_sheet)
    target_values = _fetch_sheet_values(spreadsheet_id, target_sheet)

    if not source_values:
        raise ValueError(f"Source sheet '{source_sheet}' is empty")
    if not target_values:
        raise ValueError(f"Target sheet '{target_sheet}' is empty")

    source_headers = source_values[0]
    target_headers = target_values[0]

    source_key_idx, _ = find_header_index(
        source_headers, source_key_column, source_sheet
    )
    source_val_idx, _ = find_header_index(
        source_headers, source_value_column, source_sheet
    )
    target_key_idx, _ = find_header_index(
        target_headers, target_key_column, target_sheet
    )

    target_value_idx, existing_header = find_header_index(
        target_headers, target_value_column, target_sheet, required=False
    )
    if target_value_idx is None:
        target_headers = target_headers + [target_value_column]
        target_value_idx = len(target_headers) - 1
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{target_sheet}!1:1",
            valueInputOption='RAW',
            body={'values': [target_headers]}
        ).execute()
    else:
        target_value_column = existing_header or target_value_column

    source_lookup = {}
    for row in source_values[1:]:
        key = row[source_key_idx] if len(row) > source_key_idx else ""
        if not key:
            continue
        value = row[source_val_idx] if len(row) > source_val_idx else ""
        if value:
            source_lookup[key] = value

    updates = []
    updated_rows = 0
    column_letter = column_index_to_letter(target_value_idx)
    for offset, row in enumerate(target_values[1:], start=2):
        key = row[target_key_idx] if len(row) > target_key_idx else ""
        if not key or key not in source_lookup:
            continue
        value = source_lookup[key]
        range_ref = f"{target_sheet}!{column_letter}{offset}"
        updates.append({'range': range_ref, 'values': [[value]]})
        updated_rows += 1

    if updates:
        sheets_service.spreadsheets().values().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'valueInputOption': 'RAW', 'data': updates}
        ).execute()

    return updated_rows

def read_sheet_sample(spreadsheet_id, sheet_name=None):
    """Read a sample of data from a specific sheet"""
    try:
        sheets_service = create_sheets_service()
        
        # First get sheet metadata to see available sheets
        metadata = sheets_service.spreadsheets().get(
            spreadsheetId=spreadsheet_id
        ).execute()
        
        print(f"Spreadsheet: {metadata['properties']['title']}")
        print("Available sheets:")
        available_sheets = []
        for sheet in metadata['sheets']:
            sheet_title = sheet['properties']['title']
            print(f"  - {sheet_title}")
            available_sheets.append(sheet_title)
        
        # Use first sheet if no sheet name specified
        if sheet_name is None:
            sheet_name = available_sheets[0]
            print(f"Using first sheet: {sheet_name}")
        
        # Read data from specified sheet
        range_name = f"{sheet_name}!A1:Z10"  # Read first 10 rows
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        if values:
            print(f"\nSample data from '{sheet_name}':")
            for row in values[:5]:  # Show first 5 rows
                print(row)
        else:
            print(f"No data found in '{sheet_name}'")
        
        return values
    except Exception as e:
        print(f"Error reading sheet: {e}")
        return []

# Core agent functions (80/20 rule)
def write_cell(spreadsheet_id, sheet_name, cell, value):
    """Put a value in a specific cell"""
    try:
        sheets_service = create_sheets_service()
        result = sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f'{sheet_name}!{cell}',
            valueInputOption='RAW',
            body={'values': [[value]]}
        ).execute()
        print(f"✓ Added '{value}' to {sheet_name}!{cell}")
        return result.get('updatedCells', 0)
    except Exception as e:
        print(f"✗ Error writing to {cell}: {e}")
        return 0

def read_range(spreadsheet_id, sheet_name, range_cells):
    """Read data from a range of cells"""
    try:
        sheets_service = create_sheets_service()
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f'{sheet_name}!{range_cells}'
        ).execute()
        values = result.get('values', [])
        print(f"✓ Read {len(values)} rows from {sheet_name}!{range_cells}")
        return values
    except Exception as e:
        print(f"✗ Error reading range: {e}")
        return []

def append_row(spreadsheet_id, sheet_name, data):
    """Add a new row to the end of the sheet"""
    try:
        sheets_service = create_sheets_service()
        result = sheets_service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=f'{sheet_name}!A:A',
            valueInputOption='RAW',
            insertDataOption='INSERT_ROWS',
            body={'values': [data]}
        ).execute()
        print(f"✓ Added new row to {sheet_name}")
        return result.get('updatedCells', 0)
    except Exception as e:
        print(f"✗ Error appending row: {e}")
        return 0

def find_replace(spreadsheet_id, sheet_name, find_text, replace_text):
    """Replace all occurrences of text in the sheet"""
    try:
        sheets_service = create_sheets_service()
        
        # Get sheet ID for the replace request
        metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheet_id = None
        for sheet in metadata['sheets']:
            if sheet['properties']['title'] == sheet_name:
                sheet_id = sheet['properties']['sheetId']
                break
        
        if sheet_id is None:
            print(f"✗ Sheet '{sheet_name}' not found")
            return 0
        
        request = {
            'findReplace': {
                'find': find_text,
                'replacement': replace_text,
                'sheetId': sheet_id,
                'allSheets': False
            }
        }
        
        result = sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': [request]}
        ).execute()
        
        replaced_count = result['replies'][0]['findReplace']['occurrencesChanged']
        print(f"✓ Replaced {replaced_count} occurrences of '{find_text}' with '{replace_text}'")
        return replaced_count
    except Exception as e:
        print(f"✗ Error in find/replace: {e}")
        return 0

# Intelligent Agent using Gemini
import json
import google.generativeai as genai
from dotenv import load_dotenv

load_dotenv()

MAX_PLAN_STEPS = 5

class SheetsAgent:
    def __init__(self):
        # Configure Gemini API
        genai.configure(api_key=os.getenv('GEMINI_API_KEY'))
        self.model = genai.GenerativeModel('gemini-2.5-pro')

        self.sheets = discover_google_sheets()
        self.default_sheet_id = None
        self.default_sheet_name = None
        self.subsheet_cache: List[Dict[str, str]] = []
        self.sheet_headers: Dict[str, List[str]] = {}

        if self.sheets:
            self.default_sheet_id = self.sheets[0]['id']
            self._refresh_subsheet_cache()
            if self.subsheet_cache:
                self.default_sheet_name = self.subsheet_cache[0]['title']
    
    def execute_command(self, user_prompt):
        """Execute a natural language command on Google Sheets"""
        if not self.default_sheet_id:
            return "❌ No Google Sheets found. Please share a sheet with the service account first."

        available_sheet_names = ", ".join(
            sheet["title"] for sheet in self.subsheet_cache
        ) if self.subsheet_cache else "None"

        try:
            plan = self._generate_plan(user_prompt, available_sheet_names)
            execution_log = self._run_plan(plan)
            return self._summarize_execution(user_prompt, plan, execution_log)
        except Exception as e:
            return f"❌ Error processing command: {str(e)}"

    def _execute_function(self, command):
        """Execute the parsed function call"""
        func_name = command.get("function")
        params = command.get("params", {})

        def resolve_sheet_name():
            sheet_name = params.get("sheet_name")
            if sheet_name:
                if not any(sheet["title"] == sheet_name for sheet in self.subsheet_cache):
                    self._refresh_subsheet_cache()
                    if not any(sheet["title"] == sheet_name for sheet in self.subsheet_cache):
                        raise ValueError(f"Sheet '{sheet_name}' not found.")
                return sheet_name
            if self.default_sheet_name:
                return self.default_sheet_name
            raise ValueError("No default subsheet available. Please list subsheets first.")
        
        try:
            if func_name == "write_cell":
                target_sheet = resolve_sheet_name()
                result = write_cell(
                    self.default_sheet_id, 
                    target_sheet, 
                    params["cell"], 
                    params["value"]
                )
                return f"✅ Successfully wrote '{params['value']}' to cell {target_sheet}!{params['cell']}"
            
            elif func_name == "read_range":
                target_sheet = resolve_sheet_name()
                result = read_range(
                    self.default_sheet_id, 
                    target_sheet, 
                    params["range_cells"]
                )
                if result:
                    return (
                        f"✅ Data from {target_sheet}!{params['range_cells']}:\n"
                        + "\n".join([str(row) for row in result[:10]])
                    )
                else:
                    return "❌ No data found in that range"
            
            elif func_name == "append_row":
                target_sheet = resolve_sheet_name()
                result = append_row(
                    self.default_sheet_id, 
                    target_sheet, 
                    params["data"]
                )
                return f"✅ Successfully added new row to {target_sheet} with data: {params['data']}"
            
            elif func_name == "find_replace":
                target_sheet = resolve_sheet_name()
                result = find_replace(
                    self.default_sheet_id, 
                    target_sheet, 
                    params["find_text"], 
                    params["replace_text"]
                )
                return f"✅ Replaced {result} occurrences of '{params['find_text']}' with '{params['replace_text']}'"
            elif func_name == "sum":
                target_sheet = resolve_sheet_name()
                total = sum_range(
                    self.default_sheet_id,
                    target_sheet,
                    params["range_cells"],
                )
                return f"✅ SUM of {target_sheet}!{params['range_cells']} = {total}"
            elif func_name == "average":
                target_sheet = resolve_sheet_name()
                avg = average_range(
                    self.default_sheet_id,
                    target_sheet,
                    params["range_cells"],
                )
                return f"✅ AVERAGE of {target_sheet}!{params['range_cells']} = {avg}"
            elif func_name == "count":
                target_sheet = resolve_sheet_name()
                count_value = count_range(
                    self.default_sheet_id,
                    target_sheet,
                    params["range_cells"],
                )
                return f"✅ COUNT of numeric cells in {target_sheet}!{params['range_cells']} = {count_value}"
            elif func_name == "counta":
                target_sheet = resolve_sheet_name()
                count_value = counta_range(
                    self.default_sheet_id,
                    target_sheet,
                    params["range_cells"],
                )
                return f"✅ COUNTA of {target_sheet}!{params['range_cells']} = {count_value}"
            elif func_name == "min":
                target_sheet = resolve_sheet_name()
                value = min_range(
                    self.default_sheet_id,
                    target_sheet,
                    params["range_cells"],
                )
                return f"✅ MIN of {target_sheet}!{params['range_cells']} = {value}"
            elif func_name == "max":
                target_sheet = resolve_sheet_name()
                value = max_range(
                    self.default_sheet_id,
                    target_sheet,
                    params["range_cells"],
                )
                return f"✅ MAX of {target_sheet}!{params['range_cells']} = {value}"
            elif func_name == "sumif":
                target_sheet = resolve_sheet_name()
                total = sumif_range(
                    self.default_sheet_id,
                    target_sheet,
                    params["criteria_range"],
                    params["criterion"],
                    params.get("sum_range"),
                )
                return (
                    f"✅ SUMIF {target_sheet}!{params['criteria_range']} | criterion '{params['criterion']}'"
                    f" = {total}"
                )
            elif func_name == "countif":
                target_sheet = resolve_sheet_name()
                count_value = countif_range(
                    self.default_sheet_id,
                    target_sheet,
                    params["criteria_range"],
                    params["criterion"],
                )
                return (
                    f"✅ COUNTIF {target_sheet}!{params['criteria_range']} | criterion '{params['criterion']}'"
                    f" = {count_value}"
                )
            elif func_name == "if":
                result_value = if_condition(
                    params["condition"],
                    params["value_if_true"],
                    params["value_if_false"],
                )
                return f"✅ IF result: {result_value}"
            elif func_name == "match":
                target_sheet = resolve_sheet_name()
                index = match_position(
                    self.default_sheet_id,
                    target_sheet,
                    params["lookup_value"],
                    params["lookup_range"],
                    params.get("match_type", "exact"),
                )
                return f"✅ MATCH found at position {index}"
            elif func_name == "index_match":
                target_sheet = resolve_sheet_name()
                value = index_match_lookup(
                    self.default_sheet_id,
                    target_sheet,
                    params["array_range"],
                    params.get("row_lookup_value"),
                    params.get("row_lookup_range"),
                    params.get("column_lookup_value"),
                    params.get("column_lookup_range"),
                    params.get("row_num"),
                    params.get("col_num"),
                    params.get("match_type", "exact"),
                )
                return f"✅ INDEX/MATCH result: {value}"
            elif func_name == "sumproduct":
                target_sheet = resolve_sheet_name()
                total = sumproduct_range(
                    self.default_sheet_id,
                    target_sheet,
                    params["ranges"],
                )
                return f"✅ SUMPRODUCT = {total}"
            elif func_name == "len":
                return f"✅ LEN = {len_text(params.get('text', ''))}"
            elif func_name == "left":
                return f"✅ LEFT = {left_text(params.get('text', ''), params.get('num_chars'))}"
            elif func_name == "right":
                return f"✅ RIGHT = {right_text(params.get('text', ''), params.get('num_chars'))}"
            elif func_name == "mid":
                return f"✅ MID = {mid_text(params.get('text', ''), params.get('start_num', 1), params.get('num_chars', 1))}"
            elif func_name == "sort":
                target_sheet = resolve_sheet_name()
                sorted_values = sort_range_data(
                    self.default_sheet_id,
                    target_sheet,
                    params["range_cells"],
                    params.get("instructions", []),
                )
                return "✅ SORT result:\n" + format_table(sorted_values)
            elif func_name == "filter":
                target_sheet = resolve_sheet_name()
                filtered = filter_range_data(
                    self.default_sheet_id,
                    target_sheet,
                    params["range_cells"],
                    params.get("conditions", []),
                )
                return "✅ FILTER result:\n" + format_table(filtered)
            elif func_name == "unique":
                target_sheet = resolve_sheet_name()
                unique_values = unique_range_data(
                    self.default_sheet_id,
                    target_sheet,
                    params["range_cells"],
                )
                return "✅ UNIQUE result:\n" + format_table(unique_values)
            elif func_name == "arrayformula":
                target_sheet = params.get("sheet_name") or resolve_sheet_name()
                result = arrayformula_write(
                    self.default_sheet_id,
                    target_sheet,
                    params["destination"],
                    params["formula_body"],
                )
                return f"✅ {result}"
            elif func_name == "query":
                target_sheet = resolve_sheet_name()
                query_result = query_range_data(
                    self.default_sheet_id,
                    target_sheet,
                    params["range_cells"],
                    params["query_string"],
                    params.get("headers", 1),
                )
                return "✅ QUERY result:\n" + format_table(query_result)
            elif func_name == "today":
                return f"✅ TODAY = {today_value()}"
            elif func_name == "now":
                return f"✅ NOW = {now_value()}"
            elif func_name == "join":
                delimiter = params.get("delimiter", ", ")
                values = params.get("values")
                range_cells = params.get("range_cells")
                range_values = None
                if range_cells:
                    target_sheet = resolve_sheet_name()
                    range_values = _get_range_values(
                        self.default_sheet_id,
                        target_sheet,
                        range_cells,
                    )
                joined = join_values(delimiter, values, range_values)
                return f"✅ JOIN result: {joined}"
            elif func_name == "split":
                result_items = split_text(
                    params.get("text", ""),
                    params.get("delimiter", ","),
                )
                return "✅ SPLIT result: " + ", ".join(result_items)
            elif func_name == "conditional_format":
                target_sheet = params.get("sheet_name") or resolve_sheet_name()
                message = apply_conditional_formatting(
                    self.default_sheet_id,
                    target_sheet,
                    params["range_cells"],
                    params.get("rule", {}),
                )
                return f"✅ {message}"
            elif func_name == "merge_sheet_column":
                updated = merge_sheet_column_by_key(
                    self.default_sheet_id,
                    params["source_sheet"],
                    params["source_key_column"],
                    params["source_value_column"],
                    params["target_sheet"],
                    params.get("target_key_column"),
                    params.get("target_value_column"),
                )
                self._refresh_subsheet_cache()
                if updated:
                    return (
                        f"✅ Merged column '{params['source_value_column']}' "
                        f"into {params['target_sheet']} ({updated} row(s) updated)"
                    )
                return "⚠️ Merge completed but no matching rows were found."
            elif func_name == "list_subsheets":
                subsheets = list_subsheets(self.default_sheet_id)
                self.subsheet_cache = subsheets
                if subsheets and not self.default_sheet_name:
                    self.default_sheet_name = subsheets[0]["title"]
                self._refresh_sheet_headers()
                summary = format_subsheet_summary(subsheets)
                return "✅ Subsheet overview:\n" + summary
            else:
                return f"❌ Unknown function: {func_name}"
                
        except Exception as e:
            return f"❌ Error executing {func_name}: {str(e)}"

    def _generate_plan(self, user_prompt: str, available_sheet_names: str) -> Dict[str, Any]:
        headers_snapshot = json.dumps(self.sheet_headers)
        instruction = f"""You are a thoughtful Google Sheets automation agent.

Available functions:
1. write_cell(cell, value, sheet_name optional)
2. read_range(range_cells, sheet_name optional)
3. append_row(data, sheet_name optional)
4. find_replace(find_text, replace_text, sheet_name optional)
5. list_subsheets()
6. merge_sheet_column(source_sheet, source_key_column, source_value_column, target_sheet, target_key_column optional, target_value_column optional)
7. sum(range_cells, sheet_name optional)
8. average(range_cells, sheet_name optional)
9. count(range_cells, sheet_name optional)
10. counta(range_cells, sheet_name optional)
11. min(range_cells, sheet_name optional)
12. max(range_cells, sheet_name optional)
13. sumif(criteria_range, criterion, sum_range optional, sheet_name optional)
14. countif(criteria_range, criterion, sheet_name optional)
15. if(condition, value_if_true, value_if_false)
16. match(lookup_value, lookup_range, match_type optional)
17. index_match(array_range, row_lookup..., column_lookup..., row_num optional, col_num optional)
18. sumproduct(ranges list, sheet_name optional)
19. len(text) / left(text, num_chars) / right(text, num_chars) / mid(text, start_num, num_chars)
20. sort(range_cells, instructions, sheet_name optional)
21. filter(range_cells, conditions, sheet_name optional)
22. unique(range_cells, sheet_name optional)
23. arrayformula(destination, formula_body, sheet_name optional)
24. query(range_cells, query_string, headers optional, sheet_name optional)
25. today() / now()
26. join(delimiter, values list optional, range_cells optional)
27. split(text, delimiter)
28. conditional_format(range_cells, rule, sheet_name optional)

Current spreadsheet ID: {self.default_sheet_id}
Default subsheet: {self.default_sheet_name}
Available subsheets: {available_sheet_names}
Sheet columns (by tab): {headers_snapshot}

Think about the user's goal and craft a multi-step plan (up to {MAX_PLAN_STEPS} steps).
Each plan step must be a JSON object with:
- "step": number
- "type": "think" or "action"
- "reason": short explanation
- If type == "action", include "function" and "params" (object of arguments for the function names above).

When the user references a column using natural language (e.g., "favorite food"), map it to the closest real column name listed in the sheet columns above (e.g., "Food"). Always pass the actual header names in function parameters.
For SUMIF/COUNTIF criteria, use spreadsheet-style notation such as ">10", "=Approved", or "Pending" (defaults to equality if no operator is provided).
For IF, supply a condition like "Total > 1000" along with the true/false return strings.
For conditional_format, include a "rule" object with a supported type (text_contains, greater_than, custom_formula, etc.), a "value" or "formula" as needed, and optional color/formatting fields.

Return JSON with keys: "thought" (overall reasoning), "plan" (array of steps), "final_goal" (summary of intended outcome).
Do not execute anything yourself.

User request: {user_prompt}
"""

        response = self.model.generate_content(instruction)
        response_text = response.text.strip()
        if response_text.startswith('```json'):
            response_text = response_text[len('```json'):].strip()
        if response_text.startswith('```'):
            response_text = response_text.split('```', 2)[1].strip()

        def extract_json_block(raw: str) -> str:
            start = raw.find('{')
            if start == -1:
                raise json.JSONDecodeError("Missing opening brace", raw, 0)
            depth = 0
            for idx in range(start, len(raw)):
                char = raw[idx]
                if char == '{':
                    depth += 1
                elif char == '}':
                    depth -= 1
                    if depth == 0:
                        return raw[start:idx + 1]
            raise json.JSONDecodeError("Unbalanced braces", raw, len(raw))

        def try_parse(text: str) -> Dict[str, Any]:
            snippet = extract_json_block(text)
            return json.loads(snippet)

        try:
            plan = try_parse(response_text)
        except json.JSONDecodeError:
            cleaned = response_text.replace('\n', ' ').replace('\r', ' ')
            plan = try_parse(cleaned)
        if "plan" not in plan or not isinstance(plan["plan"], list):
            raise ValueError("Planning response missing plan array")
        plan["plan"] = plan["plan"][:MAX_PLAN_STEPS]
        return plan

    def _run_plan(self, plan: Dict[str, Any]) -> List[Dict[str, Any]]:
        execution_log: List[Dict[str, Any]] = []
        for idx, step in enumerate(plan.get("plan", []), start=1):
            step_type = step.get("type") or ("action" if step.get("function") else "think")
            reason = step.get("reason") or step.get("description") or ""
            if step_type == "think" and not step.get("function"):
                execution_log.append({
                    "step": idx,
                    "type": "thought",
                    "reason": reason
                })
                continue

            function_name = step.get("function")
            params = step.get("params", {})
            command = {"function": function_name, "params": params}
            try:
                output = self._execute_function(command)
                execution_log.append({
                    "step": idx,
                    "type": "action",
                    "function": function_name,
                    "params": params,
                    "reason": reason,
                    "status": "success",
                    "output": output
                })
            except Exception as exc:
                execution_log.append({
                    "step": idx,
                    "type": "action",
                    "function": function_name,
                    "params": params,
                    "reason": reason,
                    "status": "error",
                    "output": str(exc)
                })
                break

        return execution_log

    def _summarize_execution(
        self,
        user_prompt: str,
        plan: Dict[str, Any],
        execution_log: List[Dict[str, Any]],
    ) -> str:
        total_steps = len(plan.get("plan") or [])
        successes = [entry for entry in execution_log if entry.get("status") == "success"]
        errors = [entry for entry in execution_log if entry.get("status") == "error"]

        if errors:
            first_error = errors[0]
            step_number = first_error.get("step", "?")
            function_name = first_error.get("function") or "step"
            error_text = first_error.get("error") or first_error.get("output") or "an error occurred"
            return (
                f"Ran {len(successes)} of {total_steps} steps. "
                f"Stopped at step {step_number} ({function_name}) because {error_text}."
            )

        if total_steps == 0:
            return "No steps were required."

        return f"Completed {total_steps} steps successfully."

    def _refresh_subsheet_cache(self):
        if not self.default_sheet_id:
            self.subsheet_cache = []
            return
        self.subsheet_cache = list_subsheets(self.default_sheet_id)
        self._refresh_sheet_headers()

    def _refresh_sheet_headers(self):
        self.sheet_headers = {}
        if not self.default_sheet_id:
            return
        for sheet in self.subsheet_cache:
            title = sheet.get("title")
            if not title:
                continue
            try:
                self.sheet_headers[title] = get_sheet_headers(
                    self.default_sheet_id,
                    title
                )
            except Exception:
                self.sheet_headers[title] = []

def chat_interface():
    """Simple chat interface for testing the agent"""
    agent = SheetsAgent()
    
    print("🤖 SheetSense Agent Ready!")
    print("💡 Try commands like:")
    print("   - 'Put Hello in cell A1'")
    print("   - 'Show me the data in A1:E5'") 
    print("   - 'Add a new employee row with John, Doe, Developer'")
    print("   - 'Replace all Manager with Director'")
    print("   - 'List all tabs' or 'Use the Summary sheet'")
    print("   - 'Merge the Foods column into Names by matching Email'")
    print("   - 'What is the sum of Sales!C:C?'")
    print("   - 'Average the Revenue column where Status is Approved'")
    print("   - 'Filter the Pipeline sheet for Region = West and sort by Amount desc'")
    print("   - 'Use INDEX/MATCH to pull the phone for Alex'")
    print("   - 'Show unique values from the Department column'")
    print("   - 'Highlight overdue tasks in red if Due Date < TODAY()'")
    print("   - Type 'quit' to exit\n")
    
    while True:
        user_input = input("👤 You: ").strip()
        
        if user_input.lower() in ['quit', 'exit', 'q']:
            print("👋 Goodbye!")
            break
            
        if not user_input:
            continue
            
        print("🤖 Agent:", agent.execute_command(user_input))
        print()

if __name__ == "__main__":
    print("=== Google Sheets Discovery ===\n")
    
    # Discover all Google Sheets
    sheets = discover_google_sheets()
    
    # If sheets found, offer to start the agent
    if sheets:
        print(f"\nFound sheet: '{sheets[0]['name']}'")
        subsheets = list_subsheets(sheets[0]['id'])
        print("\nAvailable subsheets:")
        print(format_subsheet_summary(subsheets))
        print("\n🤖 Starting SheetSense Agent...")
        chat_interface()
    else:
        print("❌ No sheets found. Share a sheet with the service account first.")
