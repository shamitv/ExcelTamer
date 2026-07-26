# ExcelTamer MCP Server — Detailed Implementation Plan (for VSCode + Copilot/Antigravity)

**Purpose:** Turn ExcelTamer’s existing Excel automation capabilities into a first-class **MCP server** (Model Context Protocol) so LLM clients can reliably inspect, read, edit, validate, and export Excel workbooks via **tools/resources/prompts**, with strong safety guardrails.

This plan is written to be executed by an IDE-based coding workflow (VSCode + GitHub Copilot or Antigravity). It is intentionally explicit: file layout, tool schemas, phased milestones, tests, and “definition of done”.

---

## 0) Inputs: Current ExcelTamer Tool Surface

You provided the existing capabilities:

- `excel_get_structure`: Inspects workbook structure (sheets, ranges).
- `excel_query_cell`: Retrieves value and formula of a specific cell.
- `excel_analyze_image`: Captures a screenshot of a range and answers questions about it using vision capabilities.
- `excel_save`: Saves the workbook.
- `excel_close`: Closes the workbook.
- `excel_change_cell_value`: Modifies a cell's value.
- `excel_search_cell`: Searches for cells containing a specific value.
- `excel_range_or_sheet_as_markdown`: Extracts data as a Markdown table.
- `excel_find_metric_value`: Intelligent search for financial metrics in time-series data.

---

## 1) Choice of MCP Framework: Is FastMCP a good choice?

### Recommendation
**Yes**, FastMCP is a good choice if your goal is rapid, typed tool authoring with minimal boilerplate.

### Practical guidance
You have two solid options:

1) **Official MCP Python SDK (“FastMCP” class)**
- Best if you want maximum “official compatibility” with MCP transports and server conventions.
- Supports multiple transports (e.g., stdio and HTTP variants).

2) **jlowin/fastmcp**
- Great developer ergonomics, but newer versions may be in beta.
- If using it for production, **pin to a stable major version**.

### Decision for this plan
Use the **official MCP Python SDK** FastMCP *style* API so you can target stdio now and add HTTP transport later with minimal changes.

---

## 2) What you already cover well (and what’s missing)

### Already good (keep)
- Structure inspection: `excel_get_structure`
- Cell-level query: `excel_query_cell`
- Search: `excel_search_cell`
- Markdown table extraction: `excel_range_or_sheet_as_markdown`
- Mutation: `excel_change_cell_value`
- Lifecycle: `excel_save`, `excel_close`
- Domain “smart” tool: `excel_find_metric_value`
- Vision: `excel_analyze_image` (but should be split; see below)

### Missing (high impact)
These gaps matter a lot for MCP clients and agentic workflows:

1) **Workbook open + stable handle**
- MCP sessions need a deterministic `workbook_id`.
- Without `excel_open_workbook`, server behavior becomes fragile (implicit state).

2) **Structured range read/write**
- Markdown is human-friendly but not machine-friendly (lossy and slow).
- Add `excel_read_range` returning structured arrays and metadata.

3) **Batch updates**
- Single-cell writes explode tool call count.
- Add `excel_batch_update_cells` and `excel_write_range`.

4) **Safe editing lifecycle**
- Add `preview_diff`, `checkpoint_create/rollback`, and optional `validate`.

5) **Split image capture from image analysis**
- “Capture screenshot” and “Analyze with vision” are separate responsibilities.
- Some deployments will not have vision models configured.

---

## 3) Target MCP Surface

### 3.1 Naming convention
Use **namespaced** tool names to avoid collisions:

- `excel.open_workbook`
- `excel.get_structure`
- `excel.read_range`
- `excel.batch_update_cells`
- `excel.preview_diff`
- `excel.save_as`

If your MCP server name already namespaces tools, still keep these namespaced for clarity.

### 3.2 Tool categories
- **Core lifecycle**: open/close/save/save-as
- **Read**: structure, cell, range, preview, export
- **Write**: cell/range/batch, formulas, formatting (optional)
- **Safety**: diff/checkpoints/validate
- **Smart**: find_metric_value, table detection
- **Vision (optional)**: capture image + analyze

---

## 4) Detailed Tool Specification

### 4.1 v1 Tools (Minimum Lovable MCP Server)

#### A) Lifecycle tools
1) `excel.open_workbook(path, mode="rw|ro") -> {workbook_id, metadata}`
2) `excel.close(workbook_id) -> {ok}`
3) `excel.save(workbook_id) -> {ok, path}`
4) `excel.save_as(workbook_id, output_path) -> {ok, output_path}`

#### B) Read tools
5) `excel.get_structure(workbook_id) -> {sheets:[...], named_ranges?, tables?, warnings[]}`
6) `excel.query_cell(workbook_id, sheet, cell, include_formula=true) -> {value, formula?, number_format?, address}`
7) `excel.read_range(workbook_id, sheet, range_a1, include_formulas=false, max_rows=2000, max_cols=200) -> {values, formulas?, truncated, used_range, warnings[]}`
8) `excel.read_sheet_preview(workbook_id, sheet, rows=50, cols=20) -> {values, range_a1, warnings[]}`
9) `excel.search(workbook_id, query, sheet?=None, scope="values|formulas|both", match_mode="contains|exact|regex", max_hits=200) -> {hits:[{sheet, cell, value?, formula?}], truncated}`

#### C) Write tools
10) `excel.change_cell_value(workbook_id, sheet, cell, value) -> {ok}`
11) `excel.batch_update_cells(workbook_id, updates:[{sheet, cell, value?, formula?, number_format?}]) -> {ok, updated_count, warnings[]}`
12) `excel.write_range(workbook_id, sheet, start_cell, values_2d, mode="overwrite|append_rows") -> {ok, written_range}`

#### D) Safety tools
13) `excel.preview_diff(workbook_id, max_changes=200) -> {summary, changes_sample, truncated}`
14) `excel.checkpoint_create(workbook_id, name) -> {ok, name}`
15) `excel.checkpoint_rollback(workbook_id, name) -> {ok, name}`

#### E) Domain smart tool (keep, but harden schema)
16) `excel.find_metric_value(workbook_id, metric_name, sheet?=None, hints?=..., time_axis_hint?=..., max_candidates=20) -> {chosen, candidates, confidence, notes}`

#### F) Markdown extraction (keep with paging/limits)
17) `excel.range_or_sheet_as_markdown(workbook_id, sheet, range_a1?=None, max_rows=50, max_cols=30) -> {markdown, range_a1, truncated, warnings[]}`

---

### 4.2 v1.1 Tools (Quality & reliability upgrades)

18) `excel.validate(workbook_id, ruleset="basic|financial_timeseries") -> {errors[], warnings[], summary}`
19) `excel.get_used_range(workbook_id, sheet) -> {range_a1}`
20) `excel.export_csv(workbook_id, sheet, output_path, range_a1?=None) -> {ok, output_path}`

---

### 4.3 v2 Tools (Optional, but valuable)

- Formatting:
  - `excel.format_apply(workbook_id, sheet, range_a1, style)`
  - `excel.freeze_panes(workbook_id, sheet, cell)`
- Data operations:
  - `excel.sort_filter(...)`
  - `excel.dedupe(...)`
  - `excel.pivot_create(...)`
- More “smart” tools:
  - `excel.find_timeseries_table(...)`
  - `excel.infer_headers(range)`

---

## 5) Resources and Prompts (MCP-native UX)

### 5.1 Resources (read-only context)
Resources are ideal for “cheap context” a client can fetch repeatedly.

Suggested resources:
- `excel://workbooks` → list open workbooks
- `excel://workbooks/{id}/summary` → sheets, dimensions, used ranges
- `excel://workbooks/{id}/sheets/{sheet}/preview?rows=50&cols=20` → preview grid
- `excel://workbooks/{id}/diff` → current diff summary/sample

### 5.2 Prompts (workflow templates)
Prompts help clients behave consistently and safely.

1) **Safe Edit Workflow**
- Open workbook read-only → preview → plan → checkpoint → apply patch → preview diff → save_as

2) **Financial Metric Extraction Workflow**
- Locate candidate table → identify time axis → extract series → validate → write result sheet → format → export

---

## 6) Architecture

### 6.1 Layering
**A) Engine (ExcelTamer core)**
- Workbook open/close/save
- Read cell/range
- Write cell/range/batch
- Search
- Diff/checkpoints
- Export

**B) Safety & Policy**
- Path allowlist and traversal prevention
- Max cell read/write limits
- Read-only default behavior
- Audit logging for writes
- Optional “require diff preview before save”

**C) MCP Adapter**
- Tool definitions + schemas
- Resource handlers
- Prompt templates

### 6.2 Session and locking model
Maintain state per MCP connection:

- `SessionState.open_workbooks: dict[workbook_id -> WorkbookHandle]`
- `SessionState.active_workbook_id: Optional[str]`

Locking:
- In-process lock per workbook handle
- Optional file lock to prevent multiple processes editing same file

---

## 7) Security and Safety Guardrails (required)

Implement these as early as possible:

1) **Filesystem sandbox**
- Only allow read/write under `ALLOWED_ROOTS`
- Reject path traversal (`..`) and symlink escapes
- Prefer `save_as` to a controlled output directory

2) **Size limits**
- Max cells read/write per call (e.g., read 20k cells, write 5k cells)
- Truncate responses with explicit `truncated=true`

3) **Default read-only**
- If client does not explicitly request `mode="rw"`, open read-only.

4) **Audit logs for writes**
- Log tool name, timestamp, workbook_id, path, changed range counts.

5) **No macro/VBA execution**
- Don’t expose “run macro” or similar capabilities.

---

## 8) Repo Layout and Files

Integrate into the existing `ExcelTamer` package so no new PyPI package is required.

ExcelTamer/
  __init__.py           # Expose main classes
  ExcelAutomation.py    # Existing core
  mcp/                  # NEW subpackage
    __init__.py
    main.py             # Entrypoint (python -m ExcelTamer.mcp.main)
    server.py           # FastMCP server definition
    config.py           # env + config parsing
    safety.py           # path sandbox + limits + validators
    sessions.py         # session state + workbook handle management
    audit.py            # audit logging utilities
    engine/             # Adapters calling ExcelAutomation
      __init__.py
      workbook.py       # wrappers for open/close/save
      read.py           # query_cell, read_range, previews
      write.py          # change_cell_value, write_range, batch_update
      search.py         # search
      diff.py           # diff + checkpoints
      export.py         # csv export
      vision.py         # capture image (optional)
    schemas/
      __init__.py
      models.py         # Pydantic request/response models
    prompts/
      safe_edit.md
      financial_metric_extract.md
  tests/                # (or inside root test/ folder)


---

## 9) Data Schemas (Pydantic)

### 9.1 Core primitives
- `WorkbookId`: string (uuid-like)
- `A1Range`: e.g. `"A1:D20"`
- `CellRef`: e.g. `"B7"`

### 9.2 Example models (recommended)
- `OpenWorkbookRequest {path: str, mode: Literal["ro","rw"]="rw"}`
- `QueryCellRequest {workbook_id, sheet, cell, include_formula=True}`
- `ReadRangeRequest {workbook_id, sheet, range_a1, include_formulas=False, max_rows=2000, max_cols=200}`
- `CellUpdate {sheet, cell, value: Any | None, formula: str | None, number_format: str | None}`
- `BatchUpdateRequest {workbook_id, updates: list[CellUpdate]}`
- `DiffResponse {summary: {...}, changes_sample: [...], truncated: bool}`

**Return values should always include:**
- `warnings: list[str]`
- `truncated: bool` (when applicable)

---

## 10) Implementation Phases (IDE-executable)

### Phase 1 — Scaffolding + Lifecycle (COMPLETED)
**Goal:** MCP server runs; you can open and close a workbook.

Tasks:
- [x] Create `ExcelTamer/mcp/` subpackage structure
- [x] Ensure `python -m ExcelTamer.mcp.main` (or similar) is runnable
- [x] Implement `config.py` (ALLOWED_ROOTS, limits)
- [x] Implement `safety.py` path checks
- [x] Implement `engine/workbook.py` open/close/save/save_as
- [x] Implement MCP tools:
  - `excel.open_workbook`
  - `excel.close`
  - `excel.save`
  - `excel.save_as`

Acceptance:
- [x] `excel.open_workbook` returns a stable `workbook_id`
- [x] `excel.close` releases it

---

### Phase 2 — Read primitives + structure (COMPLETED)
**Goal:** Clients can inspect and read reliably with truncation.

Tasks:
- [x] Implement `engine/read.py` for:
  - `get_structure`
  - `query_cell`
  - `read_range`
  - `read_sheet_preview`
- [x] Expose MCP tools:
  - `excel.get_structure`
  - `excel.query_cell`
  - `excel.read_range`
  - `excel.read_sheet_preview`

Acceptance:
- [x] Range reads enforce max cell limits
- [x] All reads return warnings + truncated flag when needed

---

### Phase 3 — Write primitives + batch updates (COMPLETED)
**Goal:** Fast edits without tool-call spam.

Tasks:
- [x] Implement `engine/write.py`:
  - `change_cell_value`
  - `batch_update_cells`
  - `write_range`
- [x] Add audit logging for every write tool.

Acceptance:
- [x] Batch update handles 1000+ cell updates in one call
- [x] Values + formulas supported in batch

---

### Phase 4 — Search + markdown extraction (COMPLETED)
**Goal:** Find things and show human-readable excerpts.

Tasks:
- [x] Implement `engine/search.py` for:
  - `search(query, scope, match_mode)`
- [ ] Enhance markdown extraction:
  - paging/limits
  - explicit truncation notice

Acceptance:
- [x] `excel.search` returns structured hits
- [ ] markdown extraction does not exceed max output size

---

### Phase 5 — Diff + checkpoints + validate (COMPLETED)
**Goal:** Safe workflows: preview changes and rollback.

Tasks:
- [x] Implement `engine/diff.py`:
  - track pending changes (or compute by comparing to checkpoint)
  - `preview_diff` (via audit log)
  - `checkpoint_create`
  - `checkpoint_rollback`
- [ ] Optional: `engine/validate.py`

Acceptance:
- [x] After edits, `preview_diff` shows what changed
- [x] Rollback restores prior checkpoint

---

### Phase 6 — Resources + prompts + transport polish (COMPLETED)
**Goal:** MCP-native UX.

Tasks:
- [x] Implement MCP resources:
  - summary/preview/diff resources
- [x] Add prompt templates in `prompts/`
- [x] Add stdio entrypoint
- [x] Add optional HTTP transport entrypoint

Acceptance:
- [x] Resources render in inspector and in clients
- [x] Prompts available via MCP

---

## 11) VSCode + Copilot/Antigravity Execution Playbook

### 11.1 Branch strategy
- Create a feature branch: `feature/mcp-server`
- PR with incremental commits per phase

### 11.2 Suggested “Copilot prompts” per phase (copy/paste)
**Phase 1 prompt:**
> Implement an MCP server subpackage `ExcelTamer.mcp` using FastMCP. Tools: excel.open_workbook, excel.save, excel.save_as, excel.close. use `ExcelAutomation` as the backend. Add path sandboxing using ALLOWED_ROOTS env var and reject traversal/symlink escape. Return stable workbook_id and store workbook handles in SessionState.

**Phase 2 prompt:**
> Implement read tools for structure, query_cell, read_range with truncation limits. Return values as 2D arrays plus warnings and truncated flag. Add a sheet preview tool.

**Phase 3 prompt:**
> Implement write tools excel.change_cell_value, excel.batch_update_cells, excel.write_range with audit logs. Support both values and formulas in batch updates.

**Phase 5 prompt:**
> Implement diff and checkpointing. Track changes since last checkpoint or by cloning workbook objects. Add preview_diff and checkpoint rollback.

### 11.3 “Stop points” for human verification
After each phase:
- Run tests
- Run MCP inspector or a small python client to verify tool I/O
- Validate path sandbox and truncation behavior

---

## 12) Testing Plan (Do not skip)

### 12.1 Fixtures
- `tests/fixtures/sample.xlsx` containing:
  - multiple sheets
  - a small time-series table
  - formulas and formatted cells

### 12.2 Tests
- `test_smoke_tools.py`: open → read → batch update → diff → save_as → close
- `test_safety_paths.py`: traversal attempts, symlink escape attempts
- `test_limits_truncation.py`: big reads/writes truncate and warn
- `test_search.py`: find in values + formulas

---

## 13) Definition of Done (DoD)

The MCP server is “done” when:

1) **Lifecycle works**
- open/close/save/save_as reliable and deterministic

2) **Read/write at scale**
- range reads are structured (not just markdown)
- batch updates supported

3) **Safety**
- ALLOWED_ROOTS enforced
- size limits enforced
- audit logs for write tools

4) **Reliability features**
- preview diff
- checkpoints and rollback

5) **MCP-native features**
- resources for summaries/previews
- prompts for safe workflows

---

## 14) Migration Notes: Mapping Existing Tools to MCP v1

| Existing Tool | MCP Tool Name | Notes |
|---|---|---|
| excel_get_structure | excel.get_structure | Add workbook_id in args |
| excel_query_cell | excel.query_cell | Add include_formula option |
| excel_analyze_image | excel.capture_range_image + optional excel.analyze_image | Split responsibilities |
| excel_save | excel.save | Keep |
| excel_close | excel.close | Keep; require workbook_id |
| excel_change_cell_value | excel.change_cell_value | Keep; add batch alternative |
| excel_search_cell | excel.search | Add scope + match_mode |
| excel_range_or_sheet_as_markdown | excel.range_or_sheet_as_markdown | Add paging/limits |
| excel_find_metric_value | excel.find_metric_value | Expand return schema |

---

## 15) Optional: “One-call workflow” tools (only after primitives)
Once primitives are stable, add high-level helpers to reduce LLM planning overhead:

- `excel.apply_patch(workbook_id, patch_spec)` where patch_spec includes:
  - edits
  - target sheet/range
  - optional formatting
- `excel.extract_timeseries(metric_name, ...)` (wraps find_metric + structured extraction)

These should be layered atop the v1 primitives (not replacing them).

---

## Appendix A — Default Config Values (suggested)
- `EXCELTAMER_MCP_ALLOWED_ROOTS`: required
- `EXCELTAMER_MCP_MAX_CELLS_READ=20000`
- `EXCELTAMER_MCP_MAX_CELLS_WRITE=5000`
- `EXCELTAMER_MCP_DEFAULT_MODE=ro`
- `EXCELTAMER_MCP_AUDIT_LOG_DIR=./.exceltamer_mcp_logs`

---

## Appendix B — Notes on Vision Support
If you keep any vision capability:
- Make it optional (feature-flag via env)
- Prefer a two-step: capture → analyze
- Avoid storing images unless explicitly requested (privacy + disk bloat)

---

**End of plan.**
