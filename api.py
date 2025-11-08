"""
SheetSense API - Natural Language Google Sheets Agent

A FastAPI web service that provides natural language interface to Google Sheets.

## Quick Start

1. Install dependencies: pip install -r requirements.txt
2. Set up Google Sheets API credentials (service account JSON)
3. Run: python api.py
4. API docs: http://localhost:8000/docs

## Key Endpoints

- POST /execute-command: Execute single command, get JSON response
- GET /execute-stream: Execute command with real-time streaming updates
- GET /health: Check API and agent status

## Example Commands

- "Put Hello in cell A1"
- "Show me data in A1:E5" 
- "Add a new row with John, Doe, Developer"
- "Replace Manager with Director"
"""

from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import json
import time
from typing import Dict, Any
import asyncio
import logging

from sheets_agent import SheetsAgent

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="SheetSense API",
    description="Natural language Google Sheets agent API",
    version="1.0.0",
    root_path="/api"
)

# Add CORS middleware for browser clients
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Configure appropriately for production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Request/Response models
class CommandRequest(BaseModel):
    command: str

class CommandResponse(BaseModel):
    result: str
    success: bool
    execution_time: float

class HealthResponse(BaseModel):
    status: str
    agent_ready: bool
    available_sheets: int

# Global agent instance
agent = None

def get_agent():
    """Get or create agent instance"""
    global agent
    if agent is None:
        try:
            agent = SheetsAgent()
            logger.info("SheetsAgent initialized successfully")
        except Exception as e:
            logger.error(f"Failed to initialize SheetsAgent: {e}")
            raise HTTPException(status_code=500, detail="Failed to initialize agent")
    return agent

@app.on_event("startup")
async def startup_event():
    """Initialize agent on startup"""
    get_agent()
    logger.info("SheetSense API started")

@app.get("/health", response_model=HealthResponse)
async def health_check():
    """Health check endpoint"""
    try:
        current_agent = get_agent()
        return HealthResponse(
            status="healthy",
            agent_ready=bool(current_agent.default_sheet_id),
            available_sheets=len(current_agent.sheets) if current_agent.sheets else 0
        )
    except Exception as e:
        return HealthResponse(
            status="unhealthy",
            agent_ready=False,
            available_sheets=0
        )

@app.post("/execute-command", response_model=CommandResponse)
async def execute_command(request: CommandRequest):
    """Execute a single command and return the result"""
    start_time = time.time()
    
    try:
        current_agent = get_agent()
        result = current_agent.execute_command(request.command)
        execution_time = time.time() - start_time
        
        # Determine if command was successful based on result
        success = not result.startswith("❌")
        
        return CommandResponse(
            result=result,
            success=success,
            execution_time=execution_time
        )
        
    except Exception as e:
        execution_time = time.time() - start_time
        logger.error(f"Command execution failed: {e}")
        
        return CommandResponse(
            result=f"❌ Internal error: {str(e)}",
            success=False,
            execution_time=execution_time
        )

async def stream_command_execution(command: str):
    """Stream command execution with real-time updates"""
    
    def format_sse_data(data: Dict[str, Any]) -> str:
        """Format data for Server-Sent Events"""
        return f"data: {json.dumps(data)}\n\n"
    
    start_time = time.time()
    
    try:
        # Send start event
        yield format_sse_data({
            "type": "start",
            "message": f"🤖 Executing: {command}",
            "timestamp": time.time()
        })
        
        # Small delay to show streaming effect
        await asyncio.sleep(0.1)
        
        current_agent = get_agent()
        
        # Send progress event
        yield format_sse_data({
            "type": "progress", 
            "message": "⚡ Processing command with Gemini...",
            "timestamp": time.time()
        })
        
        await asyncio.sleep(0.1)
        
        # Execute command
        result = current_agent.execute_command(command)
        execution_time = time.time() - start_time
        success = not result.startswith("❌")
        
        # Send result event
        yield format_sse_data({
            "type": "result",
            "message": result,
            "success": success,
            "execution_time": execution_time,
            "timestamp": time.time()
        })
        
        # Send completion event
        yield format_sse_data({
            "type": "complete",
            "message": "✨ Command completed",
            "timestamp": time.time()
        })
        
    except Exception as e:
        execution_time = time.time() - start_time
        logger.error(f"Streaming command execution failed: {e}")
        
        # Send error event
        yield format_sse_data({
            "type": "error",
            "message": f"❌ Internal error: {str(e)}",
            "success": False,
            "execution_time": execution_time,
            "timestamp": time.time()
        })

@app.get("/execute-stream")
async def execute_command_stream(command: str):
    """Execute a command with streaming updates via Server-Sent Events"""
    if not command:
        raise HTTPException(status_code=400, detail="Command parameter is required")
    
    return StreamingResponse(
        stream_command_execution(command),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "Access-Control-Allow-Origin": "*",
        }
    )

@app.get("/")
async def root():
    """Root endpoint with API information"""
    return {
        "message": "SheetSense API - Natural Language Google Sheets Agent",
        "version": "1.0.0",
        "endpoints": {
            "health": "/health",
            "execute": "/execute-command",
            "stream": "/execute-stream?command=YOUR_COMMAND",
            "docs": "/docs"
        },
        "example_commands": [
            "Put Hello in cell A1",
            "Show me data in A1:E5",
            "List all tabs",
            "Add a new row with John, Doe, Developer",
            "Replace Manager with Director"
        ]
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=5000, log_level="info")