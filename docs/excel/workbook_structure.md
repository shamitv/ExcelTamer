# Workbook Structure and Lifecycle

## Lifecycle

`excel.open_workbook` validates the requested path against
`EXCELTAMER_MCP_ALLOWED_ROOTS`, opens it through xlwings, stores the internal
handle in the server session, and returns a UUID `workbook_id`.

Use that identifier for all later calls:

1. Inspect with `excel.get_structure`, `excel.query_cell`, or range tools.
2. Make changes only when the workflow allows writing.
3. Persist with `excel.save` or `excel.save_as`.
4. Release the handle with `excel.close`.

The session mapping is in memory and is lost when the server exits.

## Structure response

`excel.get_structure` returns each sheet's:

- name
- used-range dimensions
- used-range A1 address
- named ranges and their target addresses

This metadata is intended for discovery before requesting cell data.

## Resources

`excel://workbooks` lists open workbook IDs and names.
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
