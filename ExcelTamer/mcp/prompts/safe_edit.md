
You are tasked with safely editing an Excel workbook using the `ExcelTamer` tools.
Follow this workflow strictly to ensure data integrity:

1. **Open & Inspect**: 
   - Open the workbook (`excel.open_workbook`).
   - Read the structure (`excel.get_structure`) and preview relevant sheets (`excel.read_sheet_preview`).

2. **Plan**: 
   - Identify the specific cells or ranges that need modification.
   - If searching for data, use `excel.search`.

3. **Checkpoint (Critical)**:
   - Before making ANY changes, create a checkpoint (`excel.checkpoint_create` with name "pre_edit").

4. **Edit**:
   - Apply changes using `excel.change_cell_value` or `excel.batch_update_cells` or `excel.write_range`.

5. **Verify**:
   - Read back the changed cells (`excel.read_range` or `excel.query_cell`) to confirm they match expectations.
   - Check the operation history (`excel.preview_diff`).

6. **Finalize**:
   - If successful, save the workbook (`excel.save` or `excel.save_as`).
   - If something went wrong, rollback immediately (`excel.checkpoint_rollback` name="pre_edit").
   - Close the workbook (`excel.close`).
