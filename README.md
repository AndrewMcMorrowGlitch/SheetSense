# SheetSense

Google Sheets API integration with service account authentication.

## Setup

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Add Google credentials:**
   - Get the service account JSON file from your team
   - Copy it to the project root directory
   - File should be named: `sheetsense-477619-ad0eb7d32908.json`

3. **Share Google Sheets:**
   - Share your Google Sheets with: `sheetsense-service-account@sheetsense-477619.iam.gserviceaccount.com`
   - Give it Editor permissions

4. **Test:**
   ```bash
   python service_account_auth.py
   ```

⚠️ **Never commit the JSON file to git** - it's already in `.gitignore`