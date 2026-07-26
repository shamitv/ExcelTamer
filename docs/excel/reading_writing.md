# Reading and Writing

Every operation after `excel.open_workbook` uses its returned `workbook_id`.
Workbook handles remain in the server session until `excel.close`.

## Reading

`excel.query_cell` returns:

- `value`
- `formula`
- `visible_text`
- `entry_type`

`excel.read_range` accepts a sheet and optional A1 range. If the range is
omitted, the sheet's used range is read. Responses contain headers, a
two-dimensional value matrix, shape, truncation state, and warnings.

The internal backend records Excel row numbers while converting a range to a
pandas DataFrame. The engine removes that helper column before returning the
MCP value matrix and converts missing values to JSON-compatible `null`.

Reads are bounded by both request limits (`max_rows`, `max_cols`) and
`EXCELTAMER_MCP_MAX_CELLS_READ`. `excel.read_sheet_preview` applies smaller
preview defaults to the same read path.

Bulk formula reads are not currently implemented. When `include_formulas` is
requested for a range, values are returned with a warning.

## Writing

The server provides three write forms:

- `excel.change_cell_value` for one cell
- `excel.batch_update_cells` for non-contiguous cells
- `excel.write_range` for a rectangular matrix

Batch and range writes enforce `EXCELTAMER_MCP_MAX_CELLS_WRITE`. Successful
writes are appended to the configured audit log. Use `excel.save` or
`excel.save_as` to persist workbook changes.

For safer edits, create a checkpoint before writing, read the affected cells
back for verification, and then save.
