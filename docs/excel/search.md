# Search Functionality

ExcelTamer provides search capabilities at two levels: a simple value-matching search in the core layer, and a more flexible search engine in the MCP server.

---

## Core Search: `ExcelAutomation`

**File:** [`ExcelAutomation.py`](file:///d:/work/ExcelTamer/ExcelTamer/ExcelAutomation.py)

### `find_all_cells_by_value(value, sheet_name=None, search_whole_workbook=False)`

Finds all cells containing an **exact** match.

| Parameter | Default | Description |
|---|---|---|
| `value` | — | The value to search for |
| `sheet_name` | `None` | Specific sheet to search (uses active sheet if `None`) |
| `search_whole_workbook` | `False` | If `True`, iterates all sheets |

**Returns:** `list[tuple[str, str, int]]` — each tuple is `(sheet_name, column_letter, row_number)`.

**How it works:**
1. Reads the entire used range into a DataFrame via `get_dataframe_with_excel_headers_impl()`.
2. Uses `df.isin([value]).stack()` to find matching cells.
3. Converts DataFrame indices back to Excel coordinates (sheet, column letter, row).

### `find_metric_value(sheet_name, metric_name, time_period)`

A specialised search for financial data. Finds the intersection of a metric row and a time-period column.

**Workflow:**
1. Search the sheet for all cells containing `metric_name` → identifies candidate rows.
2. Search the sheet for all cells containing `time_period` → identifies candidate columns.
3. For each (row, column) intersection, reads the cell's value, formula, and visible text.
4. Returns all matches:

```json
{
  "Error": "",
  "Cells": [
    {
      "Cell": "D15",
      "Value": 500000,
      "Formula": "=SUM(D10:D14)",
      "VisibleText": "$500,000",
      "Row": 15,
      "Column": "D"
    }
  ]
}
```

---

## MCP Search Engine

**MCP tool:** `excel.search`
**Engine file:** [`search.py`](file:///d:/work/ExcelTamer/ExcelTamer/mcp/engine/search.py)

A more flexible search with configurable matching modes.

### Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `workbook_id` | `str` | — | ID of the open workbook |
| `query` | `str` | — | Search term |
| `sheet` | `str` | `None` | Specific sheet, or all sheets if omitted |
| `scope` | `str` | `"both"` | `"values"`, `"formulas"`, or `"both"` |
| `match_mode` | `str` | `"contains"` | `"contains"`, `"exact"`, or `"regex"` |
| `max_hits` | `int` | `200` | Maximum number of results |

### Match Modes

| Mode | Behaviour |
|---|---|
| `contains` | `query in str(cell_value)` |
| `exact` | `str(cell_value) == query` |
| `regex` | `re.search(query, str(cell_value))` |

### Response Format

```json
{
  "query": "Revenue",
  "hits": [
    {
      "sheet": "Income Statement",
      "cell": "A5",
      "value": "Total Revenue",
      "match_type": "value"
    }
  ],
  "truncated": false
}
```

### Implementation Notes

- Reads the entire used range of each sheet into a DataFrame for local iteration.
- Value-based search is fast (pandas operations).
- **Formula search** in scope `"both"` or `"formulas"` is currently limited — bulk formula reads are not yet optimised due to the cost of per-cell COM calls. Value matching is always performed; formula matching is deferred for future implementation.
- Stops early once `max_hits` is reached, setting `truncated: true`.

---

## LangChain Agent Search

| Tool | Description |
|---|---|
| `excel_search_cell` | Wraps `find_all_cells_by_value()` — exact-match search with optional whole-workbook scope |
| `excel_find_metric_value` | Wraps `find_metric_value()` — intersection search for financial metrics |
