# Search

`excel.search` scans values in one sheet or every sheet in an open workbook.

Inputs:

- `workbook_id`: identifier returned by `excel.open_workbook`
- `query`: text to match
- `sheet`: optional sheet restriction
- `scope`: `values`, `formulas`, or `both`
- `match_mode`: `contains`, `exact`, or `regex`
- `max_hits`: response limit

Each hit contains the sheet, A1 cell address, value, and match type. The
response sets `truncated` when `max_hits` is reached.

The search engine reads used ranges through the internal pandas-backed range
conversion and reconstructs Excel addresses from source row numbers and column
letters.

Formula scanning is not yet implemented in the bulk search path. `values` is
the fully supported scope; `formulas` does not currently produce formula
matches, and `both` currently behaves as value search.
