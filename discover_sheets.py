import os
from googleapiclient.discovery import build
from dotenv import load_dotenv

load_dotenv()

GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY")

def create_drive_service():
    """Create Google Drive API service"""
    return build('drive', 'v3', developerKey=GOOGLE_API_KEY)

def create_sheets_service():
    """Create Google Sheets API service"""
    return build('sheets', 'v4', developerKey=GOOGLE_API_KEY)

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

def read_sheet_sample(spreadsheet_id, sheet_name="Sheet1"):
    """Read a sample of data from a specific sheet"""
    try:
        sheets_service = create_sheets_service()
        
        # First get sheet metadata to see available sheets
        metadata = sheets_service.spreadsheets().get(
            spreadsheetId=spreadsheet_id
        ).execute()
        
        print(f"Spreadsheet: {metadata['properties']['title']}")
        print("Available sheets:")
        for sheet in metadata['sheets']:
            print(f"  - {sheet['properties']['title']}")
        
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

if __name__ == "__main__":
    print("=== Google Sheets Discovery ===\n")
    
    # Discover all Google Sheets
    sheets = discover_google_sheets()
    
    # If sheets found, offer to read from the first one
    if sheets:
        print(f"\nWould you like to read sample data from '{sheets[0]['name']}'?")
        print("Uncomment the lines below and run again:")
        print(f"# read_sheet_sample('{sheets[0]['id']}')")