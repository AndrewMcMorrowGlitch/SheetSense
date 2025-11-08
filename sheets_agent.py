import os
from typing import Dict, List

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

class SheetsAgent:
    def __init__(self):
        # Configure Gemini API
        genai.configure(api_key=os.getenv('GEMINI_API_KEY'))
        self.model = genai.GenerativeModel('gemini-2.5-pro')

        self.sheets = discover_google_sheets()
        self.default_sheet_id = None
        self.default_sheet_name = None
        self.subsheet_cache: List[Dict[str, str]] = []

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

        # Create function calling prompt
        system_prompt = f"""You are a Google Sheets agent. You have access to these functions:

1. write_cell(cell, value, sheet_name optional) - Write a value to a specific cell (e.g., "A1", "B5")
2. read_range(range_cells, sheet_name optional) - Read data from a range (e.g., "A1:C10", "A:A")  
3. append_row(data, sheet_name optional) - Add a new row with data as a list
4. find_replace(find_text, replace_text, sheet_name optional) - Replace all occurrences of text
5. list_subsheets() - Return every sheet tab name plus the value in cell A1

Current sheet: "{self.sheets[0]['name']}" (ID: {self.default_sheet_id})
Default subsheet: "{self.default_sheet_name}"
Available subsheets: {available_sheet_names}

Based on the user's request, determine which function to call and with what parameters.
Return ONLY a JSON object with the function call, like:
{{"function": "write_cell", "params": {{"cell": "A1", "value": "Hello"}}}}

If the request is unclear or impossible, return:
{{"error": "explanation of the problem"}}

User request: {user_prompt}"""

        try:
            # Get Gemini's response
            response = self.model.generate_content(system_prompt)
            response_text = response.text.strip()
            
            # Parse the JSON response
            if response_text.startswith('```json'):
                response_text = response_text.replace('```json', '').replace('```', '').strip()
            
            command = json.loads(response_text)
            
            if "error" in command:
                return f"❌ {command['error']}"
            
            # Execute the function
            return self._execute_function(command)
            
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
            elif func_name == "list_subsheets":
                subsheets = list_subsheets(self.default_sheet_id)
                self.subsheet_cache = subsheets
                if subsheets and not self.default_sheet_name:
                    self.default_sheet_name = subsheets[0]["title"]
                summary = format_subsheet_summary(subsheets)
                return "✅ Subsheet overview:\n" + summary
            else:
                return f"❌ Unknown function: {func_name}"
                
        except Exception as e:
            return f"❌ Error executing {func_name}: {str(e)}"

    def _refresh_subsheet_cache(self):
        if not self.default_sheet_id:
            self.subsheet_cache = []
            return
        self.subsheet_cache = list_subsheets(self.default_sheet_id)

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
