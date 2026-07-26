
import asyncio
import json
import logging
from pathlib import Path

from mcp.server import Server
from mcp.types import (
    Tool,
    TextContent,
    ImageContent,
    EmbeddedResource
)
import mcp.types as types

from .engine import workbook as workbook_engine
from .engine import read as read_engine
from .engine import write as write_engine
from .engine import search as search_engine
from .engine import diff as diff_engine
from .sessions import session
from mcp.types import Resource, Prompt, PromptMessage, PromptArgument

# Configure logging (stderr so it doesn't break json-rpc on stdout)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger("ExcelTamerMCP")

server = Server("exceltamer-mcp")
PROMPT_DIR = Path(__file__).with_name("prompts")

@server.list_tools()
async def handle_list_tools() -> list[Tool]:
    return [
        Tool(
            name="excel.open_workbook",
            description="Open an Excel workbook and get a workbook_id. required for other operations.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Absolute path to the Excel file"},
                    "mode": {"type": "string", "enum": ["ro", "rw"], "default": "ro", "description": "Open mode: ro=read-only, rw=read-write"}
                },
                "required": ["path"]
            }
        ),
        Tool(
            name="excel.list_open_workbooks",
            description=(
                "List all Excel workbooks already open in the current Windows "
                "desktop session. This deliberately bypasses allowed-root filtering."
            ),
            inputSchema={
                "type": "object",
                "properties": {}
            }
        ),
        Tool(
            name="excel.attach_workbook",
            description=(
                "Attach the active Excel workbook without reopening or taking "
                "ownership of it."
            ),
            inputSchema={
                "type": "object",
                "properties": {}
            }
        ),
        Tool(
            name="excel.close",
            description=(
                "Close an MCP-opened workbook, or detach a workbook that was "
                "already open in Excel."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook_id": {"type": "string"}
                },
                "required": ["workbook_id"]
            }
        ),
        Tool(
            name="excel.save",
            description="Save the workbook (overwriting file).",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook_id": {"type": "string"}
                },
                "required": ["workbook_id"]
            }
        ),
        Tool(
            name="excel.save_as",
            description="Save the workbook to a new path.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook_id": {"type": "string"},
                    "output_path": {"type": "string"}
                },
                "required": ["workbook_id", "output_path"]
            }
        ),
        Tool(
            name="excel.get_structure",
            description="Get structure of the workbook (sheets, named ranges).",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook_id": {"type": "string"}
                },
                "required": ["workbook_id"]
            }
        ),
        Tool(
            name="excel.query_cell",
            description="Get value, formula, and text of a specific cell.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook_id": {"type": "string"},
                    "sheet": {"type": "string"},
                    "cell": {"type": "string"}
                },
                "required": ["workbook_id", "sheet", "cell"]
            }
        ),
        Tool(
            name="excel.read_range",
            description="Read a range of cells as a 2D array.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook_id": {"type": "string"},
                    "sheet": {"type": "string"},
                    "range_a1": {"type": "string", "description": "Range address (e.g. A1:C10). If omitted, uses used range."},
                    "max_rows": {"type": "integer", "default": 1000},
                    "max_cols": {"type": "integer", "default": 100}
                },
                "required": ["workbook_id", "sheet"]
            }
        ),
        Tool(
            name="excel.read_sheet_preview",
            description="Quick preview of a sheet's content (top-left).",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook_id": {"type": "string"},
                    "sheet": {"type": "string"},
                    "rows": {"type": "integer", "default": 50},
                    "cols": {"type": "integer", "default": 20}
                },
                "required": ["workbook_id", "sheet"]
            }
        ),
        Tool(
            name="excel.change_cell_value",
            description="Write a value/formula to a single cell.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook_id": {"type": "string"},
                    "sheet": {"type": "string"},
                    "cell": {"type": "string"},
                    "value": {"type": ["string", "number", "boolean", "null"]}
                },
                "required": ["workbook_id", "sheet", "cell", "value"]
            }
        ),
        Tool(
            name="excel.batch_update_cells",
            description="Write multiple non-contiguous cells in one go.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook_id": {"type": "string"},
                    "updates": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "sheet": {"type": "string"},
                                "cell": {"type": "string"},
                                "value": {"type": ["string", "number", "boolean", "null"]}
                            },
                            "required": ["sheet", "cell", "value"]
                        }
                    }
                },
                "required": ["workbook_id", "updates"]
            }
        ),
        Tool(
            name="excel.write_range",
            description="Write a 2D array of values starting at a specific cell.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook_id": {"type": "string"},
                    "sheet": {"type": "string"},
                    "start_cell": {"type": "string"},
                    "values": {
                        "type": "array", 
                        "items": {"type": "array", "items": {"type": ["string", "number", "boolean", "null"]}}
                    }
                },
                "required": ["workbook_id", "sheet", "start_cell", "values"]
            }
        ),
        Tool(
            name="excel.search",
            description="Search for a value across a sheet or the entire workbook.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook_id": {"type": "string"},
                    "query": {"type": "string"},
                    "sheet": {"type": "string", "description": "Optional. If omitted, searches all sheets."},
                    "scope": {"type": "string", "enum": ["values", "formulas", "both"], "default": "both"},
                    "match_mode": {"type": "string", "enum": ["contains", "exact", "regex"], "default": "contains"},
                    "max_hits": {"type": "integer", "default": 200}
                },
                "required": ["workbook_id", "query"]
            }
        ),
        Tool(
            name="excel.checkpoint_create",
            description="Create a named checkpoint of the current workbook state.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook_id": {"type": "string"},
                    "name": {"type": "string"}
                },
                "required": ["workbook_id", "name"]
            }
        ),
        Tool(
            name="excel.checkpoint_rollback",
            description=(
                "Restore a named checkpoint for an MCP-opened workbook. "
                "Rollback is not available for attached workbooks."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook_id": {"type": "string"},
                    "name": {"type": "string"}
                },
                "required": ["workbook_id", "name"]
            }
        ),
        Tool(
            name="excel.preview_diff",
            description="Show recent changes/audit history for this workbook.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook_id": {"type": "string"},
                    "max_changes": {"type": "integer", "default": 20}
                },
                "required": ["workbook_id"]
            }
        )
    ]

@server.list_resources()
async def handle_list_resources() -> list[Resource]:
    # Resource: List of open workbooks
    # URI: excel://workbooks
    wb_list_resource = Resource(
        uri="excel://workbooks",
        name="Open Workbooks",
        description="Metadata for workbooks registered in the MCP session",
        mimeType="application/json"
    )
    return [wb_list_resource]

@server.read_resource()
async def handle_read_resource(uri: str) -> str | bytes:
    if uri == "excel://workbooks":
        workbooks = []
        for wb_id, automation in session.open_workbooks.items():
            workbooks.append(
                {
                    "id": wb_id,
                    "workbook_id": wb_id,
                    **workbook_engine.workbook_metadata(automation),
                }
            )
        return json.dumps(workbooks)
    
    # Pattern: excel://workbooks/{id}/summary
    # Simple manual parsing since we don't have pattern matching in this basic skeleton
    import re
    match = re.match(r"excel://workbooks/([^/]+)/summary", uri)
    if match:
        wb_id = match.group(1)
        automation = session.get_workbook(wb_id)
        if automation:
            return str(read_engine.get_structure(wb_id))
        else:
            raise ValueError(f"Workbook {wb_id} not found")

    raise ValueError(f"Resource not found: {uri}")

@server.list_prompts()
async def handle_list_prompts() -> list[Prompt]:
    return [
        Prompt(
            name="safe-edit",
            description="Workflow for safely editing an Excel file with checkpoints.",
            arguments=[]
        ),
        Prompt(
            name="financial-extract",
            description="Workflow for extracting financial metrics.",
            arguments=[]
        )
    ]

@server.get_prompt()
async def handle_get_prompt(name: str, arguments: dict | None) -> types.GetPromptResult:
    if name == "safe-edit":
        content = (PROMPT_DIR / "safe_edit.md").read_text(encoding="utf-8")
        return types.GetPromptResult(
            messages=[
                PromptMessage(
                    role="user",
                    content=TextContent(type="text", text=content)
                )
            ]
        )
        
    elif name == "financial-extract":
        content = (PROMPT_DIR / "financial_metric_extract.md").read_text(
            encoding="utf-8"
        )
        return types.GetPromptResult(
            messages=[
                PromptMessage(
                    role="user",
                    content=TextContent(type="text", text=content)
                )
            ]
        )
        
    raise ValueError(f"Prompt not found: {name}")

@server.call_tool()
async def handle_call_tool(
    name: str, arguments: dict | None
) -> list[TextContent | ImageContent | EmbeddedResource]:
    if not arguments:
        arguments = {}
        
    try:
        if name == "excel.open_workbook":
            path = arguments.get("path")
            mode = arguments.get("mode", "ro")
            if not path:
                raise ValueError("path is required")
            result = workbook_engine.open_workbook(path, mode)
            return [TextContent(type="text", text=str(result))]

        elif name == "excel.list_open_workbooks":
            result = workbook_engine.list_open_workbooks()
            return [TextContent(type="text", text=str(result))]

        elif name == "excel.attach_workbook":
            result = workbook_engine.attach_workbook()
            return [TextContent(type="text", text=str(result))]
            
        elif name == "excel.close":
            workbook_id = arguments.get("workbook_id")
            if not workbook_id:
                raise ValueError("workbook_id is required")
            result = workbook_engine.close_workbook(workbook_id)
            return [TextContent(type="text", text=str(result))]
            
        elif name == "excel.save":
            workbook_id = arguments.get("workbook_id")
            if not workbook_id:
                raise ValueError("workbook_id is required")
            result = workbook_engine.save_workbook(workbook_id)
            return [TextContent(type="text", text=str(result))]
            
        elif name == "excel.save_as":
            workbook_id = arguments.get("workbook_id")
            output_path = arguments.get("output_path")
            if not workbook_id or not output_path:
                raise ValueError("workbook_id and output_path are required")
            result = workbook_engine.save_as_workbook(workbook_id, output_path)
            return [TextContent(type="text", text=str(result))]

        elif name == "excel.get_structure":
            workbook_id = arguments.get("workbook_id")
            if not workbook_id:
                raise ValueError("workbook_id is required")
            result = read_engine.get_structure(workbook_id)
            return [TextContent(type="text", text=str(result))]

        elif name == "excel.query_cell":
            workbook_id = arguments.get("workbook_id")
            sheet = arguments.get("sheet")
            cell = arguments.get("cell")
            if not workbook_id or not sheet or not cell:
                raise ValueError("workbook_id, sheet, and cell are required")
            result = read_engine.query_cell(workbook_id, sheet, cell)
            return [TextContent(type="text", text=str(result))]

        elif name == "excel.read_range":
            workbook_id = arguments.get("workbook_id")
            sheet = arguments.get("sheet")
            range_a1 = arguments.get("range_a1")
            max_rows = arguments.get("max_rows", 1000)
            max_cols = arguments.get("max_cols", 100)
            
            if not workbook_id or not sheet:
                raise ValueError("workbook_id and sheet are required")
            
            result = read_engine.read_range(
                workbook_id, 
                sheet, 
                range_a1=range_a1, 
                max_rows=max_rows, 
                max_cols=max_cols
            )
            return [TextContent(type="text", text=str(result))]
            
        elif name == "excel.read_sheet_preview":
            workbook_id = arguments.get("workbook_id")
            sheet = arguments.get("sheet")
            rows = arguments.get("rows", 50)
            cols = arguments.get("cols", 20)
            
            if not workbook_id or not sheet:
                raise ValueError("workbook_id and sheet are required")
            
            result = read_engine.read_sheet_preview(workbook_id, sheet, rows=rows, cols=cols)
            return [TextContent(type="text", text=str(result))]

        elif name == "excel.change_cell_value":
            workbook_id = arguments.get("workbook_id")
            sheet = arguments.get("sheet")
            cell = arguments.get("cell")
            value = arguments.get("value")
            
            if not workbook_id or not sheet or not cell:
                raise ValueError("workbook_id, sheet, and cell are required")
                
            result = write_engine.change_cell_value(workbook_id, sheet, cell, value)
            return [TextContent(type="text", text=str(result))]
            
        elif name == "excel.batch_update_cells":
            workbook_id = arguments.get("workbook_id")
            updates = arguments.get("updates")
            
            if not workbook_id or not updates:
                raise ValueError("workbook_id and updates are required")
                
            result = write_engine.batch_update_cells(workbook_id, updates)
            return [TextContent(type="text", text=str(result))]
            
        elif name == "excel.write_range":
            workbook_id = arguments.get("workbook_id")
            sheet = arguments.get("sheet")
            start_cell = arguments.get("start_cell")
            values = arguments.get("values")
            
            if not workbook_id or not sheet or not start_cell or values is None:
                raise ValueError("workbook_id, sheet, start_cell, and values are required")
                
            result = write_engine.write_range(workbook_id, sheet, start_cell, values)
            return [TextContent(type="text", text=str(result))]

        elif name == "excel.search":
            workbook_id = arguments.get("workbook_id")
            query = arguments.get("query")
            sheet = arguments.get("sheet")
            scope = arguments.get("scope", "both")
            match_mode = arguments.get("match_mode", "contains")
            max_hits = arguments.get("max_hits", 200)
            
            if not workbook_id or not query:
                raise ValueError("workbook_id and query are required")
                
            result = search_engine.search(
                workbook_id, query, 
                sheet=sheet, 
                scope=scope, 
                match_mode=match_mode, 
                max_hits=max_hits
            )
            return [TextContent(type="text", text=str(result))]

        elif name == "excel.checkpoint_create":
            workbook_id = arguments.get("workbook_id")
            name = arguments.get("name")
            if not workbook_id or not name:
                raise ValueError("workbook_id and name required")
            result = diff_engine.checkpoint_create(workbook_id, name)
            return [TextContent(type="text", text=str(result))]

        elif name == "excel.checkpoint_rollback":
            workbook_id = arguments.get("workbook_id")
            name = arguments.get("name")
            if not workbook_id or not name:
                raise ValueError("workbook_id and name required")
            result = diff_engine.checkpoint_rollback(workbook_id, name)
            return [TextContent(type="text", text=str(result))]

        elif name == "excel.preview_diff":
            workbook_id = arguments.get("workbook_id")
            max_changes = arguments.get("max_changes", 20)
            if not workbook_id:
                raise ValueError("workbook_id required")
            result = diff_engine.preview_diff(workbook_id, max_changes)
            return [TextContent(type="text", text=str(result))]
            
        else:
            raise ValueError(f"Unknown tool: {name}")

    except Exception as e:
        logger.error(f"Error executing {name}: {e}")
        return [TextContent(type="text", text=f"Error: {str(e)}")]


async def run_stdio():
    from mcp.server.stdio import stdio_server
    
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options()
        )

async def run_sse(port: int):
    from mcp.server.sse import SseServerTransport
    from starlette.applications import Starlette
    from starlette.responses import Response
    from starlette.routing import Mount, Route
    import uvicorn
    
    sse = SseServerTransport("/messages/")
    
    async def handle_sse(request):
        async with sse.connect_sse(
            request.scope, 
            request.receive, 
            request._send
        ) as streams:
            await server.run(
                streams[0], 
                streams[1], 
                server.create_initialization_options()
            )
        return Response()
        
    app = Starlette(
        debug=True,
        routes=[
            Route("/sse", endpoint=handle_sse, methods=["GET"]),
            Mount("/messages/", app=sse.handle_post_message),
        ]
    )
    
    config = uvicorn.Config(app, host="0.0.0.0", port=port, log_level="info")
    server_instance = uvicorn.Server(config)
    await server_instance.serve()

async def run(port: int | None = None):
    if port is not None:
        await run_sse(port)
    else:
        await run_stdio()

