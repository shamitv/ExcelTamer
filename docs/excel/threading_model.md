# Threading Model

ExcelTamer controls Microsoft Excel through xlwings, which uses Windows COM.
COM objects are tied to the apartment and thread that created them, so workbook
objects must not be passed to arbitrary worker threads.

## Current server behavior

The MCP transport is asynchronous, but Excel engine calls are synchronous.
Each registered workbook is represented by an internal `ExcelAutomation`
instance stored in the process-wide session. Entries record whether the
workbook was opened by MCP or attached as a non-owning handle. Engine functions
retrieve that instance and call xlwings on the server execution thread.

Stdio is the recommended local transport. SSE uses the same in-process session
and backend.

## Development rules

1. Keep all operations for a workbook on the thread that created its xlwings
   objects.
2. Do not move individual COM calls into generic thread pools.
3. Release workbooks through `excel.close` so handles are removed from the
   session. For attachments this detaches without closing Excel.
4. If concurrent workbook execution is introduced, add explicit per-workbook
   serialization and create/use each COM owner on its dedicated thread.
5. Test transport changes against real Excel on Windows; mock-only tests cannot
   validate COM apartment behavior.
