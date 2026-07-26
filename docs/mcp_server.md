# ExcelTamer MCP Server Guide

This guide details the Model Context Protocol (MCP) server integration for ExcelTamer. This server exposes Excel automation capabilities to AI agents (like Claude Desktop, Cursor, or custom MCP clients) in a safe, structured way.

## Overview

The ExcelTamer MCP server provides a standardized interface to interact with local Excel workbooks. It supports:
*   **Lifecycle Management**: Open, close, save, and save-as operations.
*   **Reading**: Structure inspection, cell querying, and range reading (with limit enforcement).
*   **Writing**: Single cell updates, batch updates, and 2D range writes.
*   **Search**: Finding values or patterns across sheets.
*   **Safety**: Path sandboxing, checkpoints (undo/rollback), and audit logging.
*   **Resources**: Direct access to workbook metadata.
*   **Prompts**: Pre-packaged workflows for common tasks.

## Installation

The MCP server is part of the `ExcelTamer` package.

1.  **Install dependencies**:
    ```bash
    pip install -r requirements.txt
    ```
    Ensure `mcp` and `openpyxl` are installed.

2.  **Verify installation**:
    ```bash
    python -m ExcelTamer.mcp.main
    ```
    This should start the server in STDIO mode (you won't see output because it waits for JSON-RPC input).

## Configuration

The server is configured via environment variables.

| Variable | Description | Default |
| :--- | :--- | :--- |
| `EXCELTAMER_MCP_ALLOWED_ROOTS` | Comma-separated list of allowed directory paths. Access outside these roots is blocked. | Current Working Directory |
| `EXCELTAMER_MCP_MAX_CELLS_READ` | Global limit for cell reads per request. | `20000` |
| `EXCELTAMER_MCP_MAX_CELLS_WRITE` | Global limit for cell writes per request. | `5000` |
| `EXCELTAMER_MCP_AUDIT_LOG_DIR` | Directory to store audit logs (`audit.jsonl`). | `./.exceltamer_mcp_logs` |

## Running with Claude Desktop

To use ExcelTamer with Claude Desktop, add the following to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "exceltamer": {
      "command": "python",
      "args": [
        "-m",
        "ExcelTamer.mcp.main"
      ],
      "env": {
        "EXCELTAMER_MCP_ALLOWED_ROOTS": "C:\\Users\\YourName\\Documents\\ExcelFiles",
        "PYTHONPATH": "path/to/ExcelTamer/repo" 
      }
    }
  }
}
```
*Note: If installed via pip, you don't need `PYTHONPATH`.*

## Running in HTTP Mode (SSE)

The server can also run in HTTP mode using Server-Sent Events (SSE), which is useful for remote access or clients that prefer HTTP over Stdio.

```bash
python -m ExcelTamer.mcp.main --port 8080
```

This will run the server on `http://0.0.0.0:8080`.
- SSE Endpoint: `/sse`
- POST Messages Endpoint: `/messages`

## Tools Reference

### Lifecycle
*   **`excel.open_workbook(path)`**: Opens a workbook and returns a `workbook_id`. All other tools require this ID.
*   **`excel.close(workbook_id)`**: Closes the workbook release resources.
*   **`excel.save(workbook_id)`**: Saves changes to the current file.
*   **`excel.save_as(workbook_id, output_path)`**: Saves to a new file.

### Reading
*   **`excel.get_structure(workbook_id)`**: Returns sheets, dimensions, and named ranges.
*   **`excel.query_cell(workbook_id, sheet, cell)`**: Returns value, formula, and visible text.
*   **`excel.read_range(workbook_id, sheet, range_a1, max_rows, max_cols)`**: Returns a 2D array of values.
*   **`excel.read_sheet_preview(workbook_id, sheet)`**: Quick look at the top-left of a sheet.

### Writing
*   **`excel.change_cell_value(workbook_id, sheet, cell, value)`**: Update a single cell.
*   **`excel.batch_update_cells(workbook_id, updates)`**: Efficiently update multiple non-contiguous cells. `updates` is a list of `{sheet, cell, value}`.
*   **`excel.write_range(workbook_id, sheet, start_cell, values)`**: Write a 2D matrix starting at a cell.

### Search
*   **`excel.search(workbook_id, query, sheet, scope, match_mode)`**: Search for values.
    *   `choice`: "values", "formulas", "both"
    *   `match_mode`: "contains", "exact", "regex"

### Safety & History
*   **`excel.checkpoint_create(workbook_id, name)`**: Save a snapshot of the current state.
*   **`excel.checkpoint_rollback(workbook_id, name)`**: Revert the workbook to a saved snapshot.
*   **`excel.preview_diff(workbook_id)`**: Show a summary of recent actions (audit log) for the session.

## Resources

*   **`excel://workbooks`**: JSON list of currently open workbook IDs and filenames.
*   **`excel://workbooks/{id}/summary`**: Returns the structure of a specific workbook.

## Prompts

*   **`safe-edit`**: A guided workflow for safely editing files (Inspection -> Checkpoint -> Edit -> Verify).
*   **`financial-extract`**: A guide for extracting time-series financial metrics.

## Development

The MCP server code is located in `ExcelTamer/mcp/`.

*   `server.py`: Main MCP server definition and tool registration.
*   `engine/`: Core logic implementation.
*   `safety.py`: Path validation logic.
*   `audit.py`: Logging implementation.

To run tests:
```bash
python test/test_mcp_smoke.py
```
