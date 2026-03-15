# Checkpoints, Rollback & Audit Trail

ExcelTamer provides data-safety mechanisms through named checkpoints (snapshot/rollback) and an append-only audit log for all write operations.

---

## Checkpoints

**Engine file:** [`diff.py`](file:///d:/work/ExcelTamer/ExcelTamer/mcp/engine/diff.py)

Checkpoints allow you to snapshot the current state of a workbook and restore it later if something goes wrong.

### Creating a Checkpoint

**MCP tool:** `excel.checkpoint_create`

```json
{"workbook_id": "abc-123", "name": "pre_edit"}
```

**What happens:**

1. The current workbook is saved to disk (`automation.save()`).
2. The saved file is **copied** to a temporary directory (`%TEMP%/exceltamer_checkpoints/`) with a unique filename.
3. The checkpoint path is stored in an in-memory registry keyed by `(workbook_id, name)`.

```
Original file (saved) ──copy──▶ %TEMP%/exceltamer_checkpoints/abc-123_pre_edit_a1b2c3d4.xlsx
```

> [!IMPORTANT]
> Creating a checkpoint **saves the workbook first**. Any unsaved changes in Excel will be persisted before the snapshot is taken.

### Rolling Back

**MCP tool:** `excel.checkpoint_rollback`

```json
{"workbook_id": "abc-123", "name": "pre_edit"}
```

**What happens:**

1. The current workbook is **closed** (without saving — changes are discarded).
2. The checkpoint file **overwrites** the original file on disk.
3. The original file is **re-opened** in Excel.
4. The session mapping is updated so the same `workbook_id` now points to the restored workbook.

```
Checkpoint file ──overwrite──▶ Original file path
                                    │
                              re-opened in Excel
                              (same workbook_id)
```

### Limitations

- Checkpoints are stored in the **temp directory** and will be cleaned up on system restart.
- Checkpoint metadata is held **in memory** only — if the MCP server restarts, checkpoint references are lost (the files may still exist in temp).
- Creating a checkpoint triggers a save, which may have side effects if the workbook contains volatile formulas (e.g., `NOW()`).

---

## Change History / Diff Preview

**MCP tool:** `excel.preview_diff`

Shows the recent write operations performed on a workbook by reading from the audit log.

```json
{
  "summary": "Found 3 recorded actions for this workbook.",
  "recent_actions": [
    {
      "timestamp": "2026-03-15T10:30:00",
      "tool": "change_cell_value",
      "workbook_id": "abc-123",
      "details": {"sheet": "Sheet1", "cell": "B5", "value_preview": "50000"}
    }
  ],
  "truncated": false
}
```

- Returns up to `max_changes` (default: 20) most recent actions.
- This is a **history** of operations, not a cell-by-cell diff against a baseline.

---

## Audit Log

**File:** [`audit.py`](file:///d:/work/ExcelTamer/ExcelTamer/mcp/audit.py)

All write operations are logged to an append-only JSONL file.

### Log Location

Configured via `EXCELTAMER_MCP_AUDIT_LOG_DIR` (default: `./.exceltamer_mcp_logs`).

Audit entries are written to `<AUDIT_LOG_DIR>/audit.jsonl`.

### Log Format

Each line is a JSON object:

```json
{
  "timestamp": "2026-03-15T10:30:00.123456",
  "tool": "change_cell_value",
  "workbook_id": "abc-123",
  "details": {
    "sheet": "Sheet1",
    "cell": "B5",
    "value_preview": "50000"
  }
}
```

### What Is Logged

| Write Tool | Logged Details |
|---|---|
| `excel.change_cell_value` | Sheet, cell, value preview (first 50 chars) |
| `excel.batch_update_cells` | Count of updated cells |
| `excel.write_range` | Sheet, start cell, rows, cols |

### Logging Implementation

- Uses a dedicated `logging.Logger` (`"ExcelTamerAudit"`) with a `FileHandler`.
- The formatter outputs raw messages (no timestamp prefix from the logger — timestamp is in the JSON payload).
- The log directory is created at module import time.

---

## Safe Edit Workflow

ExcelTamer ships a **prompt template** ([`safe_edit.md`](file:///d:/work/ExcelTamer/ExcelTamer/mcp/prompts/safe_edit.md)) that guides LLMs through a safe editing workflow:

1. **Open & Inspect** — open workbook, read structure, preview sheets
2. **Plan** — identify cells to modify, use search if needed
3. **Checkpoint** — create a checkpoint named `"pre_edit"` before any changes
4. **Edit** — apply changes
5. **Verify** — read back changed cells, check operation history
6. **Finalise** — save if successful, rollback if not, then close
