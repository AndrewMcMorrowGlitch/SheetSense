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

## API Usage

### Live Production API
🚀 **Production Server**: https://sheetsense-ugeg.onrender.com/api

- **Interactive Docs**: https://sheetsense-ugeg.onrender.com/api/docs
- **Health Check**: https://sheetsense-ugeg.onrender.com/api/health

### Local Development
Run the web API server locally:
```bash
python api.py
```

### Key Endpoints

- **POST /execute-command**: Execute single command, get JSON response
- **GET /execute-stream**: Execute command with real-time streaming updates  
- **GET /health**: Check API and agent status
- **GET /docs**: Interactive API documentation

### Example Commands

- "Put Hello in cell A1"
- "Show me data in A1:E5"
- "Add a new row with John, Doe, Developer"
- "Replace Manager with Director"

### Example Usage

**Production API:**
```bash
# Execute command via POST
curl -X POST "https://sheetsense-ugeg.onrender.com/api/execute-command" \
     -H "Content-Type: application/json" \
     -d '{"command": "Put Hello in cell A1"}'

# Stream command execution
curl "https://sheetsense-ugeg.onrender.com/api/execute-stream?command=Show%20me%20data%20in%20A1:E5"

# Health check
curl "https://sheetsense-ugeg.onrender.com/api/health"
```

**Local Development:**
```bash
# Execute command via POST
curl -X POST "http://localhost:5000/api/execute-command" \
     -H "Content-Type: application/json" \
     -d '{"command": "Put Hello in cell A1"}'

# Stream command execution
curl "http://localhost:5000/api/execute-stream?command=Show%20me%20data%20in%20A1:E5"
```