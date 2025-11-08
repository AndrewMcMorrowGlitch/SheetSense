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

if __name__ == "__main__":
    print("=== Google Sheets Discovery ===\n")
    
    # Discover all Google Sheets
    sheets = discover_google_sheets()
    
    # If sheets found, offer to read from the first one
    if sheets:
        print(f"\nWould you like to read sample data from '{sheets[0]['name']}'?")
        print("Uncomment the lines below and run again:")
        print(f"# read_sheet_sample('{sheets[0]['id']}')")