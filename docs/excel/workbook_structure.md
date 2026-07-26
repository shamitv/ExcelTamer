# Workbook Structure and Lifecycle

## Lifecycle

There are two ways to register a workbook in the MCP session:

- `excel.open_workbook` validates a requested path against
  `EXCELTAMER_MCP_ALLOWED_ROOTS`, opens it through xlwings, and owns the
  resulting workbook handle.
- `excel.list_open_workbooks` discovers Excel workbooks already open in the
  current Windows user session. After the user focuses the intended workbook,
  `excel.attach_workbook` registers the active workbook as a non-owning
  attachment without reopening it.

Both paths return a UUID `workbook_id`.

Use that identifier for all later calls:

1. Inspect with `excel.get_structure`, `excel.query_cell`, or range tools.
2. Make changes only when the workflow allows writing.
3. Persist with `excel.save` or `excel.save_as`.
4. Release the handle with `excel.close`.

For an MCP-opened workbook, `excel.close` closes the workbook. For an attached
workbook, it only removes the MCP session entry, returns
`status: "detached"`, and leaves the workbook and Excel open. Repeated
attachment of the same live workbook returns the existing `workbook_id`.

The normal attachment sequence is:

1. List with `excel.list_open_workbooks`.
2. Focus the desired Excel workbook and call `excel.attach_workbook`.
3. Operate with the returned `workbook_id`.
4. Detach with `excel.close`.

Attached workbooks permit explicit writes, saves, Save As, and checkpoint
creation. Checkpoint rollback is rejected before any file or Excel operation
because rollback requires closing and reopening the workbook.

The session mapping is in memory and is lost when the server exits.

## Structure response

`excel.get_structure` returns each sheet's:

- name
- used-range dimensions
- used-range A1 address
- named ranges and their target addresses

This metadata is intended for discovery before requesting cell data.

## Resources

`excel://workbooks` lists registered workbook IDs, names, saved paths, Excel
application PIDs, attachment status, MCP access mode, physical read-only state,
and unsaved-change state.
`excel://workbooks/{id}/summary` returns structure for one open workbook.

## Prompts

The MCP server packages two native workflow prompts:

- `safe-edit`
- `financial-extract`

They are shipped as package data and loaded relative to the installed MCP
package.

## Path and size controls

`safety.py` resolves requested paths and rejects access outside configured
roots. Read and write engines additionally enforce the configured cell limits.
Write operations are recorded by `audit.py`.

Discovery and attachment intentionally do not call `safety.py` and do not
apply `EXCELTAMER_MCP_ALLOWED_ROOTS`. They expose every xlwings-visible
workbook in the same Windows user session, including unsaved workbooks and
files outside configured roots. This explicit user-selected bypass should be
used only with a trusted local MCP client.
