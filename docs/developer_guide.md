# ExcelTamer Developer Guide

This guide covers the Windows development workflow for ExcelTamer: setting up
an editable checkout, navigating the MCP implementation, running automated and
live validation, and building release artifacts. For MCP client configuration
and the public protocol reference, see the [MCP server guide](mcp_server.md).

## Prerequisites

- Windows
- Python 3.11, 3.12, 3.13, or 3.14
- Git
- Microsoft Excel for live workbook validation

Microsoft Excel is not required to run the automated smoke tests or build the
package.

## Set up a development environment

Clone the repository and enter its directory:

```powershell
git clone https://github.com/shamitv/ExcelTamer.git
Set-Location ExcelTamer
```

Create and activate a virtual environment:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python --version
```

The reported Python version must be between 3.11 and 3.14. Install ExcelTamer
in editable mode, then install the build-validation tools:

```powershell
python -m pip install --upgrade pip
python -m pip install -e .
python -m pip install build twine
```

`build` and `twine` are development tools and are intentionally not installed
as ExcelTamer runtime dependencies.

## Repository layout

| Path | Purpose |
| --- | --- |
| `pyproject.toml` | Package metadata, supported Python versions, runtime dependencies, and the `exceltamer-mcp` entry point |
| `ExcelTamer/mcp/main.py` | CLI parsing and stdio/SSE transport selection |
| `ExcelTamer/mcp/server.py` | MCP tools, resources, prompts, request dispatch, and transport hosts |
| `ExcelTamer/mcp/engine/` | Workbook lifecycle, read, write, search, checkpoint, and audit-history operations |
| `ExcelTamer/mcp/excel.py` | Internal xlwings backend |
| `ExcelTamer/mcp/config.py` | Environment-driven server settings |
| `ExcelTamer/mcp/safety.py` | Workbook path validation |
| `ExcelTamer/mcp/audit.py` and `sessions.py` | Write-audit persistence and active workbook sessions |
| `ExcelTamer/mcp/prompts/` | Markdown bodies for the packaged MCP prompts |
| `test/test_mcp_smoke.py` | Automated protocol and engine smoke tests |
| `test/mcp_client.py` | Standalone stdio/SSE validation client |
| `docs/excel/` | Internal Excel behavior, constraints, and threading documentation |

The MCP protocol is the public integration boundary. Treat the xlwings backend
and engine modules as internal implementation details.

## Run the development server

After the editable install, the `exceltamer-mcp` command runs the current
checkout.

### Stdio

Run the default stdio transport:

```powershell
exceltamer-mcp
```

The equivalent module command is:

```powershell
python -m ExcelTamer.mcp.main
```

The standalone client starts its own stdio server subprocess, performs MCP
discovery, retrieves the packaged prompts, and exits:

```powershell
python test/mcp_client.py
```

### SSE

Start the SSE server in one PowerShell window. Environment variables must be
set in the server process:

```powershell
$env:EXCELTAMER_MCP_ALLOWED_ROOTS = (Get-Location).Path
exceltamer-mcp --port 8123
```

In a second activated PowerShell window, connect the validation client:

```powershell
python test/mcp_client.py --transport sse --port 8123
```

The SSE stream is available at `http://127.0.0.1:8123/sse`; client messages are
posted to `http://127.0.0.1:8123/messages`. See the
[configuration reference](mcp_server.md#configuration) for every environment
variable that influences the server.

## Test changes

Run the automated suite from the repository root:

```powershell
python -m unittest discover -s test -p "test_*.py" -v
```

The current suite runs seven tests covering the MCP surface, packaged prompts,
range normalization, engine read/search/write behavior, and a real stdio
handshake. These tests use fakes where workbook behavior is needed and do not
launch Microsoft Excel.

For optional live validation, use the included example workbook:

```powershell
python test/mcp_client.py --file .\test\example.xlsx
```

This opens the workbook through Excel in read-only mode, retrieves its
structure, and closes it. To run the same check against an already running SSE
server:

```powershell
python test/mcp_client.py --transport sse --port 8123 --file .\test\example.xlsx
```

The workbook must be under one of `EXCELTAMER_MCP_ALLOWED_ROOTS`. The default
allowed root is the current working directory.

## Build and verify packages

Run the following workflow from a clean repository checkout with the virtual
environment activated.

### 1. Confirm the package version

The version used in both artifact names comes from `pyproject.toml`:

```powershell
Select-String -Path .\pyproject.toml -Pattern '^version = '
git status --short
```

The current version is `0.2.1`. Review unexpected working-tree changes before
building.

### 2. Run the automated tests

```powershell
python -m unittest discover -s test -p "test_*.py" -v
```

### 3. Build the source and binary distributions

```powershell
python -m build
```

For version `0.2.1`, this creates:

- `dist\exceltamer-0.2.1.tar.gz` — source distribution
- `dist\exceltamer-0.2.1-py3-none-any.whl` — binary wheel

The `dist/` directory is ignored by Git. If it contains artifacts from other
versions, verify only the files for the version being prepared.

### 4. Check PyPI metadata and README rendering

```powershell
python -m twine check `
  .\dist\exceltamer-0.2.1.tar.gz `
  .\dist\exceltamer-0.2.1-py3-none-any.whl
```

Both artifacts must report `PASSED`.

### 5. Generate optional SHA-256 hashes

```powershell
Get-FileHash -Algorithm SHA256 `
  .\dist\exceltamer-0.2.1.tar.gz, `
  .\dist\exceltamer-0.2.1-py3-none-any.whl
```

Record these hashes when artifacts are transferred or retained for a release.

Uploading to TestPyPI or PyPI, managing credentials, and publishing a release
are intentionally outside the scope of this guide.
