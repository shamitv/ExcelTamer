# ExcelTamer

ExcelTamer is a Model Context Protocol (MCP) server for safe, structured
automation of Microsoft Excel. It exposes workbook lifecycle, reading, writing,
search, checkpoint, resource, and prompt capabilities to any MCP-compatible
client.

## Requirements

- Windows with Microsoft Excel installed
- Python 3.11 or newer

## Installation

```powershell
pip install .
```

## Run the server

Start the default stdio transport:

```powershell
exceltamer-mcp
```

The module form remains available:

```powershell
python -m ExcelTamer.mcp.main
```

For SSE transport:

```powershell
exceltamer-mcp --port 8123
```

## MCP client configuration

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

## Capabilities

- Open, inspect, save, save-as, and close workbooks
- Read cells, ranges, sheet previews, and workbook structure
- Write cells, batches, and rectangular ranges
- Search workbook values with exact, contains, or regex matching
- Create and roll back checkpoints
- Inspect recent write history
- Discover workbook resources and MCP-native workflow prompts

See the [MCP server guide](docs/mcp_server.md) for configuration, transports,
and the complete protocol reference.

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
