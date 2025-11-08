/**
 * SheetSense - AI-powered Google Sheets Assistant
 *
 * This add-on provides a chat interface to interact with your sheets
 * using natural language commands powered by Gemini AI.
 */

// Configuration
const API_BASE_URL = 'https://sheetsense-ugeg.onrender.com/api';

/**
 * Creates custom menu when the spreadsheet is opened
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('SheetSense')
    .addItem('Open Chat Assistant', 'showSidebar')
    .addSeparator()
    .addItem('Check Server Status', 'checkServerStatus')
    .addItem('About', 'showAbout')
    .addToUi();
}

/**
 * Shows the chat sidebar
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('SheetSense AI Assistant')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Check if the local Python server is running
 */
function checkServerStatus() {
  try {
    const response = UrlFetchApp.fetch(API_BASE_URL + '/health', {
      method: 'get',
      muteHttpExceptions: true
    });

    const data = JSON.parse(response.getContentText());

    if (data.status === 'healthy' || data.status === 'ok') {
      SpreadsheetApp.getUi().alert(
        'Server Status',
        'SheetSense server is running!\\n\\nAgent ready: ' + data.agent_ready + '\\nAvailable sheets: ' + data.available_sheets,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      throw new Error('Server returned unexpected status: ' + data.status);
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      'Server Offline',
      'Cannot connect to SheetSense server.\\n\\n' +
      'Please make sure you have started the server:\\n' +
      'python server.py\\n\\n' +
      'Error: ' + e.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Show about dialog
 */
function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'About SheetSense',
    'SheetSense AI Assistant v1.0\\n\\n' +
    'An AI-powered tool to interact with your Google Sheets using natural language.\\n\\n' +
    'Powered by Gemini AI',
    ui.ButtonSet.OK
  );
}

/**
 * Execute a natural language command
 * Called from the sidebar
 */
function executeCommand(userCommand) {
  try {
    const payload = {
      command: userCommand
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(API_BASE_URL + '/execute-command', options);
    const data = JSON.parse(response.getContentText());

    if (data.success) {
      return {
        success: true,
        result: data.result
      };
    } else {
      return {
        success: false,
        error: data.error || 'Unknown error occurred'
      };
    }
  } catch (e) {
    return {
      success: false,
      error: 'Cannot connect to server at ' + API_BASE_URL + '\\n\\nError: ' + e.message
    };
  }
}

/**
 * Get list of available sheets
 */
function getSheetsList() {
  try {
    const response = UrlFetchApp.fetch(API_BASE_URL + '/sheets', {
      method: 'get',
      muteHttpExceptions: true
    });

    const data = JSON.parse(response.getContentText());

    if (data.success) {
      return {
        success: true,
        sheets: data.sheets
      };
    } else {
      return {
        success: false,
        error: data.error || 'Could not fetch sheets'
      };
    }
  } catch (e) {
    return {
      success: false,
      error: 'Server connection error: ' + e.message
    };
  }
}

/**
 * Get agent information
 */
function getAgentInfo() {
  try {
    const response = UrlFetchApp.fetch(API_BASE_URL + '/agent-info', {
      method: 'get',
      muteHttpExceptions: true
    });

    const data = JSON.parse(response.getContentText());

    if (data.success) {
      return {
        success: true,
        info: data.info
      };
    } else {
      return {
        success: false,
        error: data.error
      };
    }
  } catch (e) {
    return {
      success: false,
      error: 'Server connection error: ' + e.message
    };
  }
}

/**
 * Get the current spreadsheet ID
 */
function getCurrentSpreadsheetId() {
  return SpreadsheetApp.getActiveSpreadsheet().getId();
}

/**
 * Get the current sheet name
 */
function getCurrentSheetName() {
  return SpreadsheetApp.getActiveSheet().getName();
}
