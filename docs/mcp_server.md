# ExcelTamer MCP Server Guide

ExcelTamer exposes Microsoft Excel automation through Model Context Protocol
tools, resources, and prompts. It does not embed or call a model provider.

## Installation

ExcelTamer requires Windows, Microsoft Excel, and Python 3.11 or newer.

```powershell
pip install ExcelTamer
```

The installed package provides the `exceltamer-mcp` command.

## Transports

Stdio is the default and is recommended for local MCP clients:

```powershell
exceltamer-mcp
```

SSE can be enabled by supplying a port:

```powershell
$env:EXCELTAMER_MCP_ALLOWED_ROOTS = "C:\Users\you\Documents\Excel"
exceltamer-mcp --port 8123
```

Configure an SSE-capable MCP client to connect to the server:

```json
{
  "mcpServers": {
    "exceltamer": {
      "url": "http://127.0.0.1:8123/sse"
    }
  }
}
```

Client configuration field names can vary. The SSE stream endpoint is
`http://127.0.0.1:8123/sse`, and client messages are posted to
`http://127.0.0.1:8123/messages`. `python -m ExcelTamer.mcp.main` supports the
same arguments.

## Configuration

| Variable | Default | Purpose |
| --- | --- | --- |
| `EXCELTAMER_MCP_ALLOWED_ROOTS` | Current directory | Comma-separated filesystem roots the server may access |
| `EXCELTAMER_MCP_DEFAULT_MODE` | `ro` | Default workbook mode: `ro` or `rw` |
| `EXCELTAMER_MCP_MAX_CELLS_READ` | `20000` | Maximum cells considered by one read |
| `EXCELTAMER_MCP_MAX_CELLS_WRITE` | `5000` | Maximum cells accepted by one write |
| `EXCELTAMER_MCP_AUDIT_LOG_DIR` | `./.exceltamer_mcp_logs` | Directory containing `audit.jsonl` |

Example client configuration:

```json
{
  "mcpServers": {
    "exceltamer": {
      "command": "exceltamer-mcp",
      "env": {
        "EXCELTAMER_MCP_ALLOWED_ROOTS": "C:\\Users\\you\\Documents\\Excel",
        "EXCELTAMER_MCP_DEFAULT_MODE": "ro"
      }
    }
  }
}
```

## Quick-start workflow

After adding the server configuration to an MCP client, restart the client so
it discovers ExcelTamer. You can then describe the workbook task in natural
language; the client selects and invokes the `excel.*` tools.

### Inspect a workbook

Ask the client:

```text
Open C:\Users\you\Documents\Excel\budget.xlsx in read-only mode. List the
worksheets, preview the first 10 rows of the first sheet, summarize what the
workbook contains, and close it when finished.
```

The expected tool sequence is:

1. `excel.open_workbook` with mode `ro`
2. `excel.get_structure`
3. `excel.read_sheet_preview` or `excel.read_range`
4. `excel.close`

`excel.open_workbook` returns a `workbook_id`. Every subsequent workbook tool
requires that identifier, so the MCP client must reuse it until the workbook is
closed.

### Edit a workbook safely

Ask the client:

```text
Open C:\Users\you\Documents\Excel\budget.xlsx in read-write mode. Create a
checkpoint, update cell B4 on Sheet1 to 120, read the cell back to verify the
change, save the workbook, and close it. If verification fails, roll back to
the checkpoint.
```

The expected tool sequence is:

1. `excel.open_workbook` with mode `rw`
2. `excel.checkpoint_create`
3. One or more write tools
4. A read tool to verify the result
5. `excel.save` and `excel.close`, or `excel.checkpoint_rollback` on failure

The MCP-native `safe-edit` prompt provides the same checkpoint-first workflow.
If the workbook path is outside `EXCELTAMER_MCP_ALLOWED_ROOTS`, opening it is
rejected before Excel is started.

## Tools

### Workbook lifecycle

| Tool | Purpose |
| --- | --- |
| `excel.open_workbook` | Open a validated workbook path and return a `workbook_id` |
| `excel.close` | Close a session workbook |
| `excel.save` | Save the current workbook |
| `excel.save_as` | Save to another validated path |

### Inspection and reading

| Tool | Purpose |
| --- | --- |
| `excel.get_structure` | Return sheets, dimensions, used ranges, and named ranges |
| `excel.query_cell` | Return a cell's value, formula, and rendered text |
| `excel.read_range` | Return a bounded rectangular value matrix |
| `excel.read_sheet_preview` | Return a bounded top-left sheet preview |

### Writing and search

| Tool | Purpose |
| --- | --- |
| `excel.change_cell_value` | Write one value or formula |
| `excel.batch_update_cells` | Write non-contiguous cells in one call |
| `excel.write_range` | Write a two-dimensional matrix |
| `excel.search` | Search values using contains, exact, or regex matching |

### Checkpoints and history

| Tool | Purpose |
| --- | --- |
| `excel.checkpoint_create` | Save a temporary checkpoint copy |
| `excel.checkpoint_rollback` | Restore a named checkpoint |
| `excel.preview_diff` | Return recent audited write actions |

All tools after `excel.open_workbook` require its returned `workbook_id`.

## Resources and prompts

`excel://workbooks` returns the workbooks currently held by the server session.
`excel://workbooks/{id}/summary` returns structure for a specific workbook.

The server exposes two MCP-native prompts:

- `safe-edit`: inspect, checkpoint, edit, and verify a workbook safely.
- `financial-extract`: inspect labels and extract time-series metrics reliably.

## Validation

Protocol-only validation does not require a workbook:

```powershell
python test/mcp_client.py
python -m unittest discover -s test -p "test_*.py" -v
```

To validate actual Excel automation:

```powershell
python test/mcp_client.py --file test/fixtures/simple.xlsx
```

Use `--transport sse --port 8123` when validating a separately running SSE
server.

## Architecture

- `ExcelTamer/mcp/server.py` defines the MCP protocol surface.
- `ExcelTamer/mcp/engine/` implements lifecycle, read, write, search, and
  checkpoint operations.
- `ExcelTamer/mcp/excel.py` is the internal xlwings backend.
- `sessions.py`, `safety.py`, and `audit.py` manage workbook handles, path
  controls, and write history.
