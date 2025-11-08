import argparse
from typing import Dict, List

from googleapiclient.errors import HttpError

from sheets_agent import create_sheets_service, discover_google_sheets


def get_sheet_tabs_and_a1(spreadsheet_id: str) -> List[Dict[str, str]]:
    """Return each sheet tab title and the value that lives in its A1 cell."""
    service = create_sheets_service()

    try:
        spreadsheet = service.spreadsheets().get(
            spreadsheetId=spreadsheet_id
        ).execute()
    except HttpError as exc:
        raise RuntimeError(f"Unable to read spreadsheet metadata: {exc}") from exc

    result: List[Dict[str, str]] = []
    for sheet in spreadsheet.get("sheets", []):
        title = sheet["properties"]["title"]
        range_name = f"{title}!A1"

        try:
            response = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=range_name,
            ).execute()
        except HttpError as exc:
            raise RuntimeError(f"Unable to read range {range_name}: {exc}") from exc

        values = response.get("values", [])
        a1_value = values[0][0] if values and values[0] else ""
        result.append({"title": title, "a1_value": a1_value})

    return result


def resolve_spreadsheet_id(explicit_id: str | None) -> str:
    """Pick a spreadsheet ID, defaulting to the first result from discovery."""
    if explicit_id:
        return explicit_id

    sheets = discover_google_sheets()
    if not sheets:
        raise RuntimeError("No Google Sheets found for the configured API key.")

    first_sheet = sheets[0]
    print(
        "No spreadsheet ID supplied; reusing the first sheet discovered via "
        f"discover_sheets.py: {first_sheet['name']} ({first_sheet['id']})"
    )
    return first_sheet["id"]


def main(spreadsheet_id: str | None):
    sheet_id = resolve_spreadsheet_id(spreadsheet_id)
    sheets = get_sheet_tabs_and_a1(sheet_id)

    if not sheets:
        print("No sheet tabs found in the selected spreadsheet.")
        return

    print(f"Found {len(sheets)} sheet tab(s) in spreadsheet {sheet_id}:")
    print("-" * 60)
    for sheet in sheets:
        value_display = sheet["a1_value"] if sheet["a1_value"] else "[empty]"
        print(f"{sheet['title']}: {value_display}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="List all sheet tabs in a spreadsheet and dump the A1 value."
    )
    parser.add_argument(
        "--spreadsheet-id",
        help="Explicit Google Sheet ID; omit to reuse the first sheet from "
        "discover_sheets.py.",
    )
    args = parser.parse_args()

    main(args.spreadsheet_id)
