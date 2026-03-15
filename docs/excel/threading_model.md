# Threading Model for Excel Integration

ExcelTamer interacts with Microsoft Excel via COM automation (through [xlwings](https://www.xlwings.org/)). This imposes strict threading constraints that the codebase carefully manages. This document explains the problem, the solution, and how it is applied across the two integration surfaces.

---

## The COM / STA Constraint

Microsoft Excel is a COM server that operates under the **Single-Threaded Apartment (STA)** model. This means:

1. **COM objects are thread-bound** — an Excel COM object (workbook, sheet, range, etc.) can only be safely accessed from the thread that created it.
2. **Cross-thread access is dangerous** — calling COM methods from a different thread can cause crashes, deadlocks, or silent data corruption.
3. **xlwings wraps COM** — every xlwings call (`xw.App`, `xw.Book`, `xw.Range.value`, etc.) translates to one or more COM calls under the hood.

> [!CAUTION]
> Any architecture that creates Excel COM objects on one thread and then accesses them from another will produce unpredictable, hard-to-debug failures.

**References:**
- [STA Threading (Raymond Chen / The Old New Thing)](https://devblogs.microsoft.com/oldnewthing/20191125-00/?p=103135)
- [COM Threading Models (Microsoft Docs)](https://docs.microsoft.com/en-us/windows/win32/com/using-the-threading)

---

## The ExcelTamer Solution: Single-Worker ThreadPoolExecutor

ExcelTamer uses a `ThreadPoolExecutor(max_workers=1)` to funnel **all** Excel COM calls onto a single dedicated thread. This guarantees that:

- The `ExcelAutomation` instance is **created** on the executor thread.
- All subsequent xlwings/COM calls are **submitted** to the same executor thread.
- No COM object is ever touched from the main thread, async event loop, or any other thread.

```
Main / Async Thread(s)                   Executor Thread (single)
──────────────────────────               ────────────────────────
                                          ┌──────────────────┐
  executor.submit(ExcelAutomation, path) ─▶│ Create COM objects│
                                          └──────────────────┘
                                          ┌──────────────────┐
  executor.submit(automation.get_structure)▶│ COM: read sheets │
                                          └──────────────────┘
                                          ┌──────────────────┐
  executor.submit(automation.write_cell)  ─▶│ COM: write value │
                                          └──────────────────┘
```

---

## Implementation in the LangChain Agent

**File:** [`AgentBuilder.py`](file:///d:/work/ExcelTamer/ExcelTamer/ExcelTamerAgent/AgentBuilder.py)

```python
# Global single-threaded executor (created once, shared across tool calls)
executor = None

def create_agent(excel_path, llm, ...):
    global executor
    if executor is None:
        executor = ThreadPoolExecutor(max_workers=1)

    # ExcelAutomation is created ON the executor thread
    future = executor.submit(ExcelAutomation, file_path=excel_path)
    excel = future.result()

    # All tools receive the same executor
    tools = [
        ExcelGetStructureTool(excel_automation=excel, executor=executor),
        ExcelCellValueTool(excel_automation=excel, executor=executor),
        ExcelAnalyzeImageTool(excel_automation=excel, executor=executor, llm=llm),
        # ... all other tools
    ]
```

### How Each Tool Uses the Executor

Every LangChain tool follows the same pattern:

```python
class ExcelGetStructureTool(BaseTool):
    _excel_automation: ExcelAutomation = PrivateAttr()
    _executor: ThreadPoolExecutor = PrivateAttr()

    def _run(self, *args, **kwargs):
        # Submit the COM call to the executor thread
        future = self._executor.submit(self._excel_automation.get_structure)
        return future.result()   # Block until complete

    async def _arun(self, *args, **kwargs):
        # Same pattern — sync submission, blocking wait
        return self._run(*args, **kwargs)
```

**Key points:**
- `_run()` (sync) and `_arun()` (async) both submit work to the executor and block for the result.
- The `async` variant does **not** use `asyncio.run_in_executor()` — it calls the sync method directly. This is safe because LangChain agents typically await tool results sequentially.
- The executor is a **module-level global**, ensuring the same thread is reused across agent invocations within the same process.

---

## Implementation in the MCP Server

**File:** [`server.py`](file:///d:/work/ExcelTamer/ExcelTamer/mcp/server.py)

The MCP server takes a different approach: it uses `asyncio` for the server transport layer, but the engine functions call `ExcelAutomation` methods **directly** (not via an executor).

```python
@server.call_tool()
async def handle_call_tool(name, arguments):
    if name == "excel.open_workbook":
        result = workbook_engine.open_workbook(path, mode)
        return [TextContent(type="text", text=str(result))]
```

The engine functions (e.g., `workbook_engine.open_workbook()`) call `ExcelAutomation` synchronously within the async handler.

### Why This Works

The MCP server uses stdio-based transport (`stdio_server`), which runs in a single `asyncio` event loop. Since the server is single-threaded by design and all tool calls are awaited sequentially (not run in parallel), there is no concurrency conflict on the COM objects.

> [!NOTE]
> The MCP server model works because it is a single-process, single-client, stdio server. If the server were converted to handle concurrent requests (e.g., via SSE or WebSocket), an executor pattern similar to the LangChain agent would be required.

---

## Comparison of Threading Approaches

| Aspect | LangChain Agent | MCP Server |
|---|---|---|
| **Executor** | `ThreadPoolExecutor(max_workers=1)` | None (direct calls) |
| **COM thread** | Dedicated executor thread | Main asyncio thread |
| **Concurrency model** | Agent may use async tools | Sequential tool calls via stdio |
| **ExcelAutomation creation** | On the executor thread | In the engine function (main thread) |
| **Thread safety** | ✅ Explicit via executor | ✅ Implicit via single-threaded server |

---

## Guidelines for Contributors

1. **Never call xlwings from a thread that didn't create the COM objects.** If adding new functionality that might run concurrently, wrap all COM calls in `executor.submit()`.

2. **The executor must have `max_workers=1`.** Using more workers would create COM objects on multiple threads, breaking the STA guarantee.

3. **Creating `ExcelAutomation` must happen on the executor thread too.** The constructor creates `xw.App` and `xw.Book` COM objects — these must live on the same thread as all subsequent calls.

4. **Avoid `asyncio.run_in_executor` with COM.** The default asyncio executor creates threads indiscriminately. Always use the dedicated single-worker pool.

5. **If extending the MCP server to support concurrent clients**, introduce a `ThreadPoolExecutor(max_workers=1)` and route all engine calls through it, similar to the LangChain agent pattern.
