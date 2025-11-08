import os
from googleapiclient.discovery import build
from dotenv import load_dotenv

load_dotenv()

GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY")

def create_sheets_service():
    return build('sheets', 'v4', developerKey=GOOGLE_API_KEY)

# Example usage
if __name__ == "__main__":
    service = create_sheets_service()
    print("Google Sheets API service created successfully!")
    
    # Example: Get spreadsheet metadata
    # spreadsheet_id = "your-spreadsheet-id"
    # result = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    # print(f"Spreadsheet title: {result.get('properties', {}).get('title')}")