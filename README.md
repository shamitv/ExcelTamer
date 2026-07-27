# ExcelTamer

ExcelTamer is a Model Context Protocol (MCP) server for safe, structured
automation of Microsoft Excel. It exposes workbook lifecycle, reading, writing,
search, checkpoint, resource, and prompt capabilities to any MCP-compatible
client.

**New to ExcelTamer?** Follow the [quick start](#quick-start-use-exceltamer-with-an-mcp-client)
below or read the
[complete MCP usage guide](https://github.com/shamitv/ExcelTamer/blob/main/docs/mcp_server.md).

**Building or contributing?** Read the
[developer guide](https://github.com/shamitv/ExcelTamer/blob/main/docs/developer_guide.md).

## Quick start: use ExcelTamer with an MCP client

### 1. Install ExcelTamer

ExcelTamer requires Windows with Microsoft Excel installed and Python 3.11 or
newer.

```powershell
pip install ExcelTamer
```

### 2. Add ExcelTamer to your MCP client

Choose either stdio or Streamable HTTP, depending on the transports supported
by your MCP client.

#### Stdio

For the default stdio transport, add this server definition to your MCP
client's configuration, replacing the allowed root with the directory
containing your workbooks:

```json
{
  "mcpServers": {
    "exceltamer": {
      "command": "exceltamer-mcp",
      "env": {
        "EXCELTAMER_MCP_ALLOWED_ROOTS": "C:\\Users\\you\\Documents\\Excel"
      }
    }
  }
}
```

`EXCELTAMER_MCP_ALLOWED_ROOTS` limits which workbook paths the server may
access. The server defaults to read-only workbook mode and records write
operations in an audit log.

#### Streamable HTTP

To use Streamable HTTP, start ExcelTamer separately in PowerShell. Set server
environment variables in the same shell:

```powershell
$env:EXCELTAMER_MCP_ALLOWED_ROOTS = "C:\Users\you\Documents\Excel"
exceltamer-mcp --port 8123
```

Then configure an HTTP-capable MCP client to connect to the MCP endpoint:

```json
{
  "mcpServers": {
    "exceltamer": {
      "type": "http",
      "url": "http://127.0.0.1:8123/mcp"
    }
  }
}
```

Client configuration field names can vary; use
`http://127.0.0.1:8123/mcp` as the server URL.

### 3. Restart the client and ask it to use Excel

For inspection, try:

```text
Open C:\Users\you\Documents\Excel\budget.xlsx in read-only mode. List the
worksheets, preview the first 10 rows of the first sheet, summarize what the
workbook contains, and close it when finished.
```

For a safe edit, try:

```text
Open C:\Users\you\Documents\Excel\budget.xlsx in read-write mode. Create a
checkpoint, update cell B4 on Sheet1 to 120, read the cell back to verify the
change, save the workbook, and close it. If verification fails, roll back to
the checkpoint.
```

Your MCP client handles the `excel.*` tool calls and carries the returned
`workbook_id` through the workflow.

If the workbook is already open in Excel, focus its Excel window and try:

```text
List the workbooks currently open in Excel, attach the active workbook,
describe its worksheets and used ranges, and then detach from it without
closing the workbook or Excel.
```

That follows the `list → attach → operate → detach` workflow:

1. `excel.list_open_workbooks` discovers workbooks in all visible Excel
   applications.
2. `excel.attach_workbook` registers the active workbook and returns its
   `workbook_id`.
3. Read, write, checkpoint, and save tools operate on that identifier.
4. `excel.close` removes an attached workbook from the MCP session but leaves
   both the workbook and Excel open.

> **Security warning:** Discovery and attachment deliberately bypass
> `EXCELTAMER_MCP_ALLOWED_ROOTS`. They can expose every workbook open in the
> same Windows user session, including unsaved workbooks and files outside the
> configured roots. Use these tools only with a trusted local MCP client.
> Checkpoint creation is supported for attached workbooks, but checkpoint
> rollback is rejected because rollback would have to close and reopen the
> user's workbook.

## Run the server manually

Start the default stdio transport:

```powershell
exceltamer-mcp
```

The module form remains available:

```powershell
python -m ExcelTamer.mcp.main
```

For Streamable HTTP transport:

```powershell
exceltamer-mcp --port 8123
```

### Command-line options

```text
exceltamer-mcp [-h] [--port PORT]
```

| Option | Description |
| --- | --- |
| `-h`, `--help` | Show the command help and exit |
| `--port PORT` | Run the Streamable HTTP transport on the specified integer port |

When `--port` is omitted, the server uses stdio transport. The module form
accepts the same options:

```powershell
python -m ExcelTamer.mcp.main --port 8123
```

### Environment variables

ExcelTamer reads these variables when the server process starts. Restart the
server or MCP client after changing them.

| Variable | Default | Effect |
| --- | --- | --- |
| `EXCELTAMER_MCP_ALLOWED_ROOTS` | Current working directory | Comma-separated directory roots from which workbook paths may be opened |
| `EXCELTAMER_MCP_DEFAULT_MODE` | `ro` | Default workbook mode: `ro` for read-only or `rw` for read-write |
| `EXCELTAMER_MCP_MAX_CELLS_READ` | `20000` | Maximum number of cells considered by one read operation |
| `EXCELTAMER_MCP_MAX_CELLS_WRITE` | `5000` | Maximum number of cells accepted by one write operation |
| `EXCELTAMER_MCP_AUDIT_LOG_DIR` | `./.exceltamer_mcp_logs` | Directory in which write audit records are stored as `audit.jsonl` |

Set these values in the MCP client configuration's `env` object, as shown in
the quick start, or in the shell before starting `exceltamer-mcp`.

## Capabilities

- Discover all open Excel workbooks and attach the active one without
  reopening or taking ownership of it
- Open, inspect, save, save-as, close, and detach workbooks
- Read cells, ranges, sheet previews, and workbook structure
- Capture a worksheet's used range or an explicit A1 range as a PNG
- Write cells, batches, and rectangular ranges
- Search workbook values with exact, contains, or regex matching
- Create and roll back checkpoints
- Inspect recent write history
- Discover workbook resources and MCP-native workflow prompts

ExcelTamer 0.4.0 exposes 18 MCP tools, one resource, and two prompts.

## Validation

Run the automated MCP smoke tests:

```powershell
python -m unittest discover -s test -p "test_*.py" -v
```

Run the standalone protocol client:

```powershell
python test/mcp_client.py
```

Pass `--file path\to\workbook.xlsx` to additionally validate opening,
inspecting, and closing a real workbook.
