# Workbook Structure & Lifecycle

This document describes how ExcelTamer models an Excel workbook and the lifecycle of opening, inspecting, saving, and closing workbooks.

---

## Core Class: `ExcelAutomation`

**File:** [`ExcelAutomation.py`](file:///d:/work/ExcelTamer/ExcelTamer/ExcelAutomation.py)

`ExcelAutomation` is the foundational class that wraps the [xlwings](https://www.xlwings.org/) library to interact with a live Excel application instance via COM automation.

### Construction

```python
class ExcelAutomation:
    def __init__(self, file_path: str = None):
        self.app = xw.apps.active if xw.apps else xw.App(visible=True)

        if file_path:
            self.wb = self.app.books.open(file_path)
        else:
            self.wb = self.app.books.active if self.app.books else self.app.books.add()
```

| Scenario | Behaviour |
|---|---|
| `file_path` provided | Opens the specified workbook in the active Excel app (or starts a new one) |
| `file_path=None` | Attaches to the currently active workbook, or creates a new blank workbook |

> [!IMPORTANT]
> `ExcelAutomation` always connects to a **live Excel process**. It does not operate on files in-memory.

### Key Attributes

| Attribute | Type | Description |
|---|---|---|
| `self.app` | `xw.App` | Reference to the Excel application process |
| `self.wb` | `xw.Book` | Reference to the open workbook |

### Sheet & Workbook Management

The core class provides synchronous methods for managing sheets and workbooks. These are heavily utilised internally, and some are exposed via tools/resources:

- `list_open_workbooks()`: Returns a list of absolute paths for all workbooks currently open in the Excel process.
- `list_sheets()`: Returns a list of sheet names in the active workbook.
- `add_sheet(sheet_name: str)`: Creates a new worksheet with the given name.
- `remove_sheet(sheet_name: str)`: Deletes the specified worksheet.
- `list_named_ranges()`: Returns a dictionary mapping range names to their A1 addresses (e.g. `{"Revenue": "$B$5:$F$5"}`).

---

## Workbook Structure Inspection

The `get_structure()` method provides a complete overview of the workbook:

```python
def get_structure(self) -> list[dict]:
```

**Returns** a list of dictionaries, one per sheet:

```json
[
  {
    "Sheet Name": "Income Statement",
    "Rows": 150,
    "Columns": 30,
    "Range": "$A$1:$AD$150",
    "Named Ranges": [
      {"Name": "Revenue", "Refers To": "$B$5:$F$5"}
    ]
  }
]
```

Each entry contains:

| Field | Description |
|---|---|
| `Sheet Name` | Name of the worksheet |
| `Rows` / `Columns` | Dimensions of the used range |
| `Range` | A1-notation address of the used range |
| `Named Ranges` | List of named ranges scoped to that sheet |

---

## Workbook Lifecycle (MCP Server)

When accessed via the MCP server, workbooks follow a session-based lifecycle managed by the engine modules.

### Opening

**Tool:** `excel.open_workbook`
**Engine:** [`workbook.py`](file:///d:/work/ExcelTamer/ExcelTamer/mcp/engine/workbook.py) → `open_workbook(path, mode)`

1. The file path is validated against `ALLOWED_ROOTS` (see [Safety](#path-validation--safety)).
2. An `ExcelAutomation` instance is created, which opens the workbook in Excel.
3. The instance is stored in the global `SessionState` singleton and assigned a UUID (`workbook_id`).
4. Metadata (filename, sheets, mode) is returned to the caller.

```
Client ──▶ excel.open_workbook(path, mode)
               │
               ├─ validate_path(path)         # safety.py
               ├─ ExcelAutomation(file_path)  # opens in Excel
               ├─ session.add_workbook(...)   # stores with UUID
               └─ returns { workbook_id, filename, sheets, mode }
```

### Saving

| Tool | Behaviour |
|---|---|
| `excel.save` | Saves to the current file location |
| `excel.save_as` | Saves to a new path (validated for safety) |

### Closing

**Tool:** `excel.close`

Closes the workbook in Excel **without** quitting the application so that other workbooks remain open. Removes the workbook from the session registry.

### MCP Resources & Prompts

Beyond tools, the server exposes metadata and workflows via MCP's native capabilities:

**Resources:**
- `excel://workbooks` — Returns a JSON list of all `{id, name}` pairs for currently open workbooks.
- `excel://workbooks/{id}/summary` — Returns the structural summary of a specific workbook (identical to calling `excel.get_structure`).

**Prompts:**
- `safe-edit` — Provides an LLM with instructions on how to use checkpoints to safely edit an Excel file.
- `financial-extract` — Provides an LLM with instructions on how to reliably extract temporal financial metrics.

---

## Session Management

**File:** [`sessions.py`](file:///d:/work/ExcelTamer/ExcelTamer/mcp/sessions.py)

```python
class SessionState:          # Singleton pattern
    open_workbooks: Dict[str, ExcelAutomation]

    def add_workbook(automation) -> str     # returns UUID
    def get_workbook(wb_id) -> ExcelAutomation
    def remove_workbook(wb_id)
    def clear()
```

- Uses the **Singleton** pattern (`__new__` override) — one global session per server process.
- Each open workbook is keyed by a generated UUID string (`workbook_id`).
- All engine functions resolve workbooks by calling `session.get_workbook(workbook_id)`.

---

## Path Validation & Safety

**File:** [`safety.py`](file:///d:/work/ExcelTamer/ExcelTamer/mcp/safety.py)

Before any file is opened or saved, `validate_path()` checks that the resolved absolute path falls within the configured `ALLOWED_ROOTS`:

```python
def validate_path(request_path: str, allow_write: bool = False) -> Path:
    abs_path = Path(request_path).resolve()
    # checks abs_path.relative_to(allowed_root) for each root
```

- Resolves symlinks to prevent directory traversal attacks.
- Raises `SecurityError` if the path is outside all allowed roots.
- `ALLOWED_ROOTS` is configured via the `EXCELTAMER_MCP_ALLOWED_ROOTS` environment variable (comma-separated).

---

## Configuration

**File:** [`config.py`](file:///d:/work/ExcelTamer/ExcelTamer/mcp/config.py)

| Variable | Env Var | Default | Description |
|---|---|---|---|
| `ALLOWED_ROOTS` | `EXCELTAMER_MCP_ALLOWED_ROOTS` | Current working directory | Comma-separated allowed filesystem roots |
| `MAX_CELLS_READ` | `EXCELTAMER_MCP_MAX_CELLS_READ` | `20000` | Maximum cells returned in a single read |
| `MAX_CELLS_WRITE` | `EXCELTAMER_MCP_MAX_CELLS_WRITE` | `5000` | Maximum cells in a single write operation |
| `DEFAULT_MODE` | `EXCELTAMER_MCP_DEFAULT_MODE` | `ro` | Default open mode (read-only) |
| `AUDIT_LOG_DIR` | `EXCELTAMER_MCP_AUDIT_LOG_DIR` | `./.exceltamer_mcp_logs` | Directory for audit log files |
