# Excel MCP Internals

These documents describe the behavior behind ExcelTamer's MCP protocol:

| Guide | Topic |
| --- | --- |
| [Workbook structure](workbook_structure.md) | Workbook lifecycle, sessions, resources, prompts, and path controls |
| [Reading and writing](reading_writing.md) | Cell/range I/O, limits, normalization, and audit behavior |
| [Search](search.md) | Value search modes, addressing, and current formula limitations |
| [Checkpoints and audit](checkpoints_and_audit.md) | Checkpoint creation, rollback, and write history |
| [Threading model](threading_model.md) | xlwings/COM ownership and transport considerations |

The public integration boundary is MCP. The xlwings backend and engine modules
are internal implementation details.
