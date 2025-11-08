# SheetSense Google Sheets Add-on Setup Guide

This guide will help you set up the SheetSense AI assistant as a sidebar in your Google Sheets.

## Architecture

```
Google Sheets (Browser)
  └── Apps Script Sidebar (Chat UI)
      └── HTTP Requests
          └── Python Flask Server (localhost:5000)
              └── SheetsAgent (Gemini AI)
                  └── Google Sheets API
```

## Prerequisites

- Python 3.8+ installed
- Google Sheets access
- Service account credentials file (`sheetsense-477619-ad0eb7d32908.json`)
- Gemini API key

## Part 1: Start the Python Backend Server

### 1. Install Flask dependencies

```bash
pip install flask flask-cors
```

Or install all dependencies:

```bash
pip install -r requirements.txt
```

### 2. Verify your .env file

Make sure your `.env` file contains:

```
GEMINI_API_KEY=AIzaSyBtdPSKOEcViKyis80uAKDB_pJELCk8HQk
PORT=5000
```

### 3. Start the Flask server

```bash
python server.py
```

You should see:

```
Starting SheetSense API Server on port 5000...
Server will be available at: http://localhost:5000

Available endpoints:
  - GET  /api/health
  - POST /api/execute-command
  - GET  /api/sheets
  - GET  /api/sheets/<id>/subsheets
  - GET  /api/agent-info
```

**Keep this terminal window open!** The server must be running for the add-on to work.

### 4. Test the server (Optional)

Open a new terminal and run:

```bash
curl http://localhost:5000/api/health
```

You should get: `{"message":"SheetSense API server is running","status":"ok"}`

## Part 2: Set Up Google Apps Script

### 1. Open your Google Sheet

Go to your Google Sheet (the one shared with your service account).

### 2. Open Apps Script Editor

- Click **Extensions** → **Apps Script**
- This will open the Apps Script editor in a new tab

### 3. Delete existing code

- Delete all the default code in `Code.gs`

### 4. Add the Apps Script files

You need to add 3 files to your Apps Script project:

#### File 1: Code.gs

- Click on `Code.gs` in the left sidebar
- Copy the entire contents of `apps-script/Code.gs` from this project
- Paste it into the Apps Script editor

#### File 2: Sidebar.html

- Click the **+** next to **Files** in the left sidebar
- Select **HTML**
- Name it: `Sidebar`
- Copy the entire contents of `apps-script/Sidebar.html` from this project
- Paste it into the new HTML file

#### File 3: appsscript.json

- Click on the **Project Settings** (gear icon) in the left sidebar
- Check the box: **"Show 'appsscript.json' manifest file in editor"**
- Go back to the **Editor** tab
- Click on `appsscript.json` in the file list
- Replace its contents with the contents of `apps-script/appsscript.json` from this project

### 5. Save the project

- Click the disk icon or press `Ctrl+S` (Windows) / `Cmd+S` (Mac)
- Give your project a name: "SheetSense AI Assistant"

## Part 3: Test the Add-on

### 1. Refresh your Google Sheet

- Go back to your Google Sheet tab
- Refresh the page (F5 or Ctrl+R)

### 2. You should see a new menu

After a few seconds, you should see a new menu item: **SheetSense**

### 3. Check server status

- Click **SheetSense** → **Check Server Status**
- You should get a success message confirming the server is running
- If you get an error, make sure `python server.py` is running

### 4. Open the chat sidebar

- Click **SheetSense** → **Open Chat Assistant**
- A sidebar should appear on the right side of your Google Sheet

### 5. Try some commands

Type these commands in the chat:

- "List all tabs"
- "Show me the data in A1:E5"
- "Put 'Hello World' in cell A1"
- "Add a new row with John, Doe, Engineer"

## Troubleshooting

### "Cannot connect to server" error

**Problem:** The sidebar shows "Cannot connect to SheetSense server"

**Solutions:**
1. Make sure `python server.py` is running in a terminal
2. Check that it's running on port 5000 (check the terminal output)
3. Try accessing http://localhost:5000/api/health in your browser

### "CORS error" in Apps Script

**Problem:** CORS-related errors in the Apps Script logs

**Solution:** Make sure `flask-cors` is installed:
```bash
pip install flask-cors
```

### "No sheets found" error

**Problem:** Agent can't find your Google Sheets

**Solutions:**
1. Verify the service account credentials file exists: `sheetsense-477619-ad0eb7d32908.json`
2. Make sure your Google Sheet is shared with: `sheetsense-service-account@sheetsense-477619.iam.gserviceaccount.com`
3. Give it **Editor** permissions

### Apps Script menu doesn't appear

**Problem:** The "SheetSense" menu doesn't show up

**Solutions:**
1. Refresh the Google Sheet page
2. Wait 10-15 seconds for the script to load
3. Check the Apps Script editor for any errors (View → Logs)
4. Make sure you saved all files in the Apps Script editor

### Gemini API errors

**Problem:** Commands fail with API errors

**Solutions:**
1. Verify your Gemini API key in `.env`
2. Check that the key is valid and has quota remaining
3. Make sure `google-generativeai` is installed: `pip install google-generativeai`

## How It Works

1. **You type a command** in the sidebar (e.g., "Put Hello in A1")
2. **Sidebar JavaScript** calls the Apps Script function `executeCommand()`
3. **Apps Script** makes an HTTP POST request to `http://localhost:5000/api/execute-command`
4. **Flask server** receives the request and passes it to `SheetsAgent`
5. **SheetsAgent** sends your command to Gemini AI to interpret
6. **Gemini AI** returns a structured command (e.g., `write_cell`)
7. **SheetsAgent** executes the command using Google Sheets API
8. **Result** travels back through Flask → Apps Script → Sidebar
9. **You see the response** in the chat interface

## Development Tips

### Viewing Server Logs

Watch the terminal where `python server.py` is running to see incoming requests and responses.

### Viewing Apps Script Logs

In the Apps Script editor:
- Click **View** → **Logs**
- Or press `Ctrl+Enter` after running a function

### Testing Server Endpoints

Use curl or Postman to test endpoints:

```bash
# Test health check
curl http://localhost:5000/api/health

# Test command execution
curl -X POST http://localhost:5000/api/execute-command \
  -H "Content-Type: application/json" \
  -d '{"command": "list all tabs"}'

# Get sheets list
curl http://localhost:5000/api/sheets

# Get agent info
curl http://localhost:5000/api/agent-info
```

### Modifying the UI

To change the sidebar appearance:
1. Edit `apps-script/Sidebar.html` locally
2. Copy the updated contents
3. Paste into the Apps Script editor's `Sidebar.html`
4. Save and refresh your Google Sheet

## Next Steps

### Deploy to Production

For others to use your add-on:

1. **Deploy Flask server** to a cloud provider:
   - Heroku: https://www.heroku.com/
   - Railway: https://railway.app/
   - Render: https://render.com/

2. **Update API_BASE_URL** in `Code.gs`:
   ```javascript
   const API_BASE_URL = 'https://your-app.herokuapp.com/api';
   ```

3. **Publish the Add-on** (optional):
   - In Apps Script: **Deploy** → **New deployment**
   - Choose **Add-on**
   - Follow Google's add-on publication process

### Add More Features

You can extend the agent by:
- Adding more operations to `sheets_agent.py`
- Updating the Gemini AI prompt with new functions
- Creating new API endpoints in `server.py`
- Enhancing the UI in `Sidebar.html`

## Support

If you encounter issues:
1. Check the terminal running `server.py` for errors
2. Check Apps Script logs (View → Logs)
3. Verify all prerequisites are met
4. Make sure all files are saved and the server is running

Enjoy using SheetSense! 🎉
