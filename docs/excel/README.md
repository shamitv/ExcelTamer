# Excel Integration Documentation

Technical documentation for how ExcelTamer interacts with Microsoft Excel.

---

## Documents

| Document | Description |
|---|---|
| [HOW-TO: Custom Python Integration](HOWTO_custom_python_integration.md) | **Start Here:** Definitive guide for developers integrating Excel automation (`xlwings`) directly into their own asynchronous or multithreaded Python codebases (includes screenshot example). |
| [Workbook Structure & Lifecycle](workbook_structure.md) | `ExcelAutomation` class, sheet management, workbook inspection, open/save/close lifecycle, MCP Resources/Prompts, session management, path safety, and config |
| [Reading & Writing](reading_writing.md) | Single-cell queries, DataFrame and Markdown range reads, 2D array reads, cell writes, batch updates, and range writes |
| [Screenshot & Image Analysis](screenshot_image_analysis.md) | PNG screenshot capture via xlwings and vision-LLM analysis of spreadsheet images |
| [Search](search.md) | Value search, regex/contains/exact matching, financial metric intersection search |
| [Checkpoints & Audit](checkpoints_and_audit.md) | Named checkpoints, rollback, JSONL audit trail, and the safe-edit workflow |
| [Threading Model](threading_model.md) | COM/STA constraints, single-worker `ThreadPoolExecutor` pattern, and threading differences between the LangChain Agent and MCP Server |

---

## Architecture Overview

ExcelTamer exposes Excel automation through two interfaces:

```
                       ┌─────────────────────┐
                       │   ExcelAutomation    │    Core layer (xlwings / COM)
                       └─────────┬───────────┘
                                 │
                ┌────────────────┴─────────────────┐
                │                                  │
    ┌───────────▼────────────┐      ┌──────────────▼────────────┐
    │   LangChain Agent      │      │      MCP Server           │
    │   (ExcelTamerTools)    │      │   (server.py + engine/)   │
    │                        │      │                           │
    │  • ThreadPoolExecutor  │      │  • asyncio stdio server   │
    │  • 9 LangChain tools   │      │  • 14 MCP tools           │
    │  • Vision LLM analysis │      │  • Session management     │
    └────────────────────────┘      │  • Audit logging          │
                                    │  • Checkpoints & rollback │
                                    │  • Path sandboxing        │
                                    └───────────────────────────┘
```

## Key Source Files

| File | Role |
|---|---|
| `ExcelTamer/ExcelAutomation.py` | Core xlwings wrapper |
| `ExcelTamer/ExcelTamerAgent/ExcelTamerTools.py` | LangChain tool definitions |
| `ExcelTamer/ExcelTamerAgent/AgentBuilder.py` | Agent factory with executor setup |
| `ExcelTamer/mcp/server.py` | MCP server (tool definitions + dispatch) |
| `ExcelTamer/mcp/engine/` | Engine modules (workbook, read, write, search, diff) |
| `ExcelTamer/mcp/sessions.py` | Singleton session state |
| `ExcelTamer/mcp/safety.py` | Path validation |
| `ExcelTamer/mcp/audit.py` | Write audit logging |
| `ExcelTamer/mcp/config.py` | Environment-based configuration |
