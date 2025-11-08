import os
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

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

if __name__ == "__main__":
    print("=== Google Sheets Discovery ===\n")
    
    # Discover all Google Sheets
    sheets = discover_google_sheets()
    
    # If sheets found, offer to read from the first one
    if sheets:
        print(f"\nWould you like to read sample data from '{sheets[0]['name']}'?")
        print("Uncomment the lines below and run again:")
        print(f"# read_sheet_sample('{sheets[0]['id']}')")
        print("\n# Example agent functions:")
        print(f"# write_cell('{sheets[0]['id']}', 'Data', 'P1', 'Agent Test')")
        print(f"# append_row('{sheets[0]['id']}', 'Data', ['E12345', 'John Doe', 'Developer'])")
        print(f"# find_replace('{sheets[0]['id']}', 'Data', 'Manager', 'Director')")