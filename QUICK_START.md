# Quick Start Guide - SheetSense Google Sheets Add-on

## 🚀 Running the Add-on (5 minutes)

### Step 1: Start the Python Server

Open a terminal in this directory and run:

```bash
python server.py
```

You should see:
```
Starting SheetSense API Server on port 5000...
Server will be available at: http://localhost:5000
 * Running on http://127.0.0.1:5000
```

**Keep this terminal open!**

---

### Step 2: Add Apps Script to Your Google Sheet

1. **Open your Google Sheet** (the one shared with your service account)

2. **Open Apps Script Editor:**
   - Click **Extensions** → **Apps Script**

3. **Add the 3 files:**

   **File 1: Code.gs**
   - Click on `Code.gs` in the left sidebar
   - Delete existing code
   - Copy ALL contents from: `apps-script/Code.gs`
   - Paste into Apps Script editor

   **File 2: Sidebar (HTML)**
   - Click the **+** next to **Files**
   - Select **HTML**
   - Name it: `Sidebar`
   - Copy ALL contents from: `apps-script/Sidebar.html`
   - Paste into the HTML file

   **File 3: appsscript.json**
   - Click **Project Settings** (gear icon)
   - Check: **"Show 'appsscript.json' manifest file in editor"**
   - Go back to **Editor** tab
   - Click `appsscript.json`
   - Replace contents with: `apps-script/appsscript.json`

4. **Save the project:**
   - Click the disk icon or `Ctrl+S`
   - Name it: "SheetSense AI Assistant"

---

### Step 3: Use the Chat Interface

1. **Refresh your Google Sheet**
   - Go back to your Google Sheet tab
   - Press F5 or reload the page

2. **Wait 10 seconds** for the script to load

3. **Look for the "SheetSense" menu** at the top

4. **Click: SheetSense → Open Chat Assistant**
   - A sidebar will appear on the right

5. **Try these commands:**
   - "List all tabs"
   - "Show me the data in A1:E5"
   - "Put 'Hello World' in cell A1"
   - "Add a new row with Test, Data, Here"

---

## 📁 Project Structure

```
SheetSense/
├── server.py                    ← Flask API server
├── sheets_agent.py              ← AI agent with Gemini
├── apps-script/
│   ├── Code.gs                 ← Apps Script backend
│   ├── Sidebar.html            ← Chat UI
│   └── appsscript.json         ← Manifest
├── SETUP_GUIDE.md              ← Detailed setup instructions
└── QUICK_START.md              ← This file
```

---

## ✅ Verification Checklist

Before opening the sidebar, verify:

- [ ] Python server is running (`python server.py`)
- [ ] Server shows "Running on http://127.0.0.1:5000"
- [ ] All 3 Apps Script files are created and saved
- [ ] Google Sheet page has been refreshed
- [ ] "SheetSense" menu appears at the top

---

## ❓ Troubleshooting

**"SheetSense" menu doesn't appear**
- Wait 10-15 seconds after refreshing
- Check Apps Script editor for errors (View → Logs)
- Make sure you saved all files

**"Cannot connect to server" in sidebar**
- Check that `python server.py` is running
- Visit http://localhost:5000/api/health in your browser
- Should return: `{"status":"ok"}`

**Commands not working**
- Verify your `.env` file has `GEMINI_API_KEY`
- Check that service account credentials file exists
- Make sure your sheet is shared with the service account

---

## 🎯 Next Steps

For more details, see [SETUP_GUIDE.md](SETUP_GUIDE.md)

For development and advanced features, check the full documentation.

Enjoy your AI-powered Google Sheets assistant! 🎉
