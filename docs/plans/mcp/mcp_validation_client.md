# Plan: MCP Validation Test Client

**Goal**: Create a robust test client script (`test/mcp_client.py`) that acts as an MCP Host, checks connection to the ExcelTamer MCP server, and performs an end-to-end validation using an LLM (OpenAI compatible).

## 1. Features
- **Configuration**:
  - Reads `OPENAI_API_KEY`, `OPENAI_BASE_URL` from `.env` or environment variables.
  - Accepts CLI arguments for model selection (default: `gpt-5-nano`).
  - Accepts CLI arguments for transport mode (Stdio vs SSE).
- **Transport Support**:
  - **Stdio**: Spawns the MCP server subprocess directly (default for local testing).
  - **HTTP SSE**: Connects to a running server instance (useful for debugging remote/containerized setups).
- **Interactive/Automated Modes**:
  - **Automated Mode**: Runs a predefined script (e.g., "Open file, read cell A1, close").
  - **Interactive Mode**: Allows user to chat with the LLM which uses the MCP tools.

## 2. Dependencies
- `mcp` (Official Python SDK)
- `openai` (Official Python SDK)
- `python-dotenv`
- `typer` or `argparse` (for CLI args)

## 3. Implementation Details

### `test/mcp_client.py` structure

```python
import asyncio
import os
import argparse
from dotenv import load_dotenv
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from mcp.client.sse import sse_client
from openai import AsyncOpenAI

# ... implementation ...
```

### Configuration Logic
- Load `.env`.
- Parse arguments:
  - `--model`: default `"gpt-5-nano"`
  - `--url`: Optional base URL (override `OPENAI_BASE_URL`).
  - `--transport`: `"stdio"` (default) or `"sse"`.
  - `--port`: if SSE, target port (default `8123`).
  - `--file`: Path to Excel file to test with.

### Execution Flow (Automated)
1. **Initialize**: Setup OpenAI client and MCP Client.
2. **Connect**: Establish connection to MCP server.
3. **Discover**: Call `list_tools()` to verify available tools.
4. **Interact**: 
   - Send System Prompt: "You are an Excel Assistant with access to these tools: [Tool Defs]..."
   - Send User Message: "Open the workbook at {file_path} and tell me what is in the first sheet."
5. **Tool Loop**:
   - Parse LLM response for tool calls.
   - Execute tool on MCP server.
   - Send result back to LLM.
   - Repeat until final answer.

## 4. Work Items

1.  [ ] Create `test/mcp_client.py`.
    -   Implement argument parsing.
    -   Implement OpenAI client setup.
    -   Implement MCP Client Session management (Stdio & SSE).
    -   Implement Chat Loop (LLM -> Tool -> LLM).
2.  [ ] Create `test/fixtures/simple.xlsx` (if not exists) for reliable testing.
3.  [ ] documentation: Add usage instructions to `docs/mcp_server.md`.

## 5. Usage Example

```bash
# Run with default model (gpt-5-nano) and stdio transport
python test/mcp_client.py --file test/example.xlsx

# Run with specific model and SSE
python test/mcp_client.py --model gpt-4o --transport sse --port 8080 --file test/example.xlsx
```
