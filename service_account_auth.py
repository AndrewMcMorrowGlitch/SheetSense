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

def create_authenticated_services():
    """Create authenticated Google Sheets and Drive services using service account"""
    # Load credentials from service account file
    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, 
        scopes=SCOPES
    )
    
    # Build authenticated services
    sheets_service = build('sheets', 'v4', credentials=credentials)
    drive_service = build('drive', 'v3', credentials=credentials)
    
    return sheets_service, drive_service

def list_shared_spreadsheets():
    """List all spreadsheets shared with the service account"""
    try:
        _, drive_service = create_authenticated_services()
        
        results = drive_service.files().list(
            q="mimeType='application/vnd.google-apps.spreadsheet'",
            fields='files(id, name, webViewLink, owners)'
        ).execute()
        
        sheets = results.get('files', [])
        if not sheets:
            print('No spreadsheets shared with service account.')
            print(f'Service account email: sheetsense-service-account@sheetsense-477619.iam.gserviceaccount.com')
            print('Share a Google Sheet with this email to test!')
            return []
        
        print(f'Found {len(sheets)} shared spreadsheets:')
        print('-' * 80)
        
        for i, sheet in enumerate(sheets, 1):
            print(f"{i}. {sheet['name']}")
            print(f"   ID: {sheet['id']}")
            print(f"   URL: {sheet['webViewLink']}")
            print()
        
        return sheets
    except Exception as e:
        print(f"Error listing spreadsheets: {e}")
        return []

def read_sheet_data(spreadsheet_id, range_name="Sheet1!A1:Z10"):
    """Read data from a specific spreadsheet"""
    try:
        sheets_service, _ = create_authenticated_services()
        
        # Get spreadsheet metadata
        metadata = sheets_service.spreadsheets().get(
            spreadsheetId=spreadsheet_id
        ).execute()
        
        print(f"Spreadsheet: {metadata['properties']['title']}")
        print("Available sheets:")
        for sheet in metadata['sheets']:
            print(f"  - {sheet['properties']['title']}")
        
        # Read data
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        if values:
            print(f"\nData from '{range_name}':")
            for row in values[:5]:  # Show first 5 rows
                print(row)
            if len(values) > 5:
                print(f"... and {len(values) - 5} more rows")
        else:
            print(f"No data found in range '{range_name}'")
        
        return values
    except Exception as e:
        print(f"Error reading sheet: {e}")
        return []

if __name__ == "__main__":
    print("=== Google Sheets Service Account Test ===\n")
    
    # Test authentication and list shared sheets
    sheets = list_shared_spreadsheets()
    
    # If you want to test reading from a specific sheet:
    # read_sheet_data('your-spreadsheet-id-here', 'Sheet1!A1:Z10')