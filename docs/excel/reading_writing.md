# Reading & Writing Cells and Ranges

This document covers how ExcelTamer reads from and writes to Excel workbooks, at both the core (`ExcelAutomation`) and MCP tool levels.

---

## Reading Data

### Single-Cell Query

Retrieve the value, formula, and visible text of a specific cell.

**Core method:** `ExcelAutomation.query_cell(sheet_name, cell)`

```python
result = automation.query_cell("Sheet1", "B5")
# {'Value': 42000.0, 'Formula': '=SUM(B2:B4)', 'VisibleText': '$42,000'}
```

**MCP tool:** `excel.query_cell`

The engine layer ([`read.py`](file:///d:/work/ExcelTamer/ExcelTamer/mcp/engine/read.py)) normalises keys and adds the `entry_type`:

```json
{
  "sheet": "Sheet1",
  "cell": "B5",
  "value": 42000.0,
  "formula": "=SUM(B2:B4)",
  "visible_text": "$42,000",
  "entry_type": "<class 'float'>"
}
```

---

### Range as DataFrame

Read a rectangular range and return it as a pandas DataFrame, with Excel column letters as headers and an extra `RowNumber` column.

**Core method:** `ExcelAutomation.get_range_as_dataframe(sheet_name, cell_range=None)`

```python
df = automation.get_range_as_dataframe("Sheet1", "A1:C5")
# Columns: ['RowNumber', 'A', 'B', 'C']
```

| Behaviour | Description |
|---|---|
| `cell_range=None` | Defaults to the entire used range of the sheet |
| Column headers | Actual Excel column letters (A, B, …, AH), **not** data from row 1 |
| `RowNumber` column | Inserted at position 0, contains actual Excel row indices |

The internal helper `get_dataframe_with_excel_headers_impl(sheet, range)` handles column-letter extraction and row-number mapping.

---

### Range as Markdown

**Core method:** `ExcelAutomation.get_range_as_markdown(sheet_name, cell_range=None)`

Converts the DataFrame to a Markdown table via `df.to_markdown(index=True)`. Useful for LLM consumption.

**LangChain tool:** `excel_range_or_sheet_as_markdown` — exposes this as an agent tool.

---

### Range as 2D Array (MCP)

**MCP tool:** `excel.read_range`
**Engine:** [`read.py`](file:///d:/work/ExcelTamer/ExcelTamer/mcp/engine/read.py) → `read_range()`

Returns a JSON-friendly structure:

```json
{
  "range_address": "A1:C5",
  "shape": [5, 3],
  "headers": ["A", "B", "C"],
  "values": [[1, "Name", 100], ...],
  "truncated": false,
  "warnings": []
}
```

**Safety limits:**

| Limit | Default | What happens |
|---|---|---|
| `MAX_CELLS_READ` (global) | 20,000 | Truncates with a warning |
| `max_rows` (per-request) | 1,000 | Truncates rows |
| `max_cols` (per-request) | 100 | Truncates columns |

The `RowNumber` column is **dropped** from the MCP response for a clean matrix output.

---

### Sheet Preview

**MCP tool:** `excel.read_sheet_preview`

A convenience wrapper around `read_range` with lower defaults:

| Parameter | Default |
|---|---|
| `rows` | 50 |
| `cols` | 20 |

Reads the top-left portion of the sheet — ideal for quick inspection.

---

## Writing Data

### Single-Cell Write

**Core method:** `ExcelAutomation.write_cell(sheet_name, cell, value)`

```python
automation.write_cell("Sheet1", "B5", 50000)
automation.write_cell("Sheet1", "B6", "=B5*1.1")  # formulas work too
```

**MCP tool:** `excel.change_cell_value`
**Engine:** [`write.py`](file:///d:/work/ExcelTamer/ExcelTamer/mcp/engine/write.py)

Every write is logged to the [audit trail](checkpoints_and_audit.md):

```python
log_write("change_cell_value", workbook_id, {
    "sheet": sheet, "cell": cell, "value_preview": str(value)[:50]
})
```

---

### Batch Cell Update

**MCP tool:** `excel.batch_update_cells`

Write multiple non-contiguous cells in a single tool call:

```json
{
  "workbook_id": "abc-123",
  "updates": [
    {"sheet": "Sheet1", "cell": "A1", "value": "Revenue"},
    {"sheet": "Sheet1", "cell": "B1", "value": 100000},
    {"sheet": "Sheet2", "cell": "C5", "value": "=Sheet1!B1*0.3"}
  ]
}
```

- Enforces `MAX_CELLS_WRITE` (default: 5,000).
- Iterates through updates internally — cells can span different sheets.
- Supports a `formula` key as an alias for `value`.
- Logs total count to audit.

---

### Range Write

**MCP tool:** `excel.write_range`

Write a contiguous 2D block of values starting at a specific cell:

```json
{
  "workbook_id": "abc-123",
  "sheet": "Sheet1",
  "start_cell": "A1",
  "values": [
    ["Name", "Q1", "Q2"],
    ["Revenue", 100, 200],
    ["Costs", 50, 75]
  ]
}
```

- Uses xlwings' native range-assignment (`ws.range(start_cell).value = values`), which is highly performant for block writes.
- Enforces `MAX_CELLS_WRITE`.
- Returns `written_cells` count and shape.

---

## Data Type Handling

xlwings automatically maps between Python and Excel types:

| Python Type | Excel Representation |
|---|---|
| `int` / `float` | Numeric cell |
| `str` starting with `=` | Formula |
| `str` (other) | Text cell |
| `datetime` | Date/time cell |
| `None` | Empty cell |
| `bool` | Boolean (TRUE/FALSE) |

The MCP `read_range` replaces `NaN` with `None` for JSON serialization.
