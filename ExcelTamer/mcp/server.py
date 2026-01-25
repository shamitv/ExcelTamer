
import asyncio
import logging
from typing import Any, Sequence

from mcp.server import Server
from mcp.types import (
    Tool,
    TextContent,
    ImageContent,
    EmbeddedResource
)
import mcp.types as types

# Import engine functions
from .engine import workbook as workbook_engine
from .engine import read as read_engine

# Configure logging (stderr so it doesn't break json-rpc on stdout)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger("ExcelTamerMCP")

server = Server("exceltamer-mcp")

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
            name="excel.close",
            description="Close an open workbook by ID.",
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
        )
    ]

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
            
        else:
            raise ValueError(f"Unknown tool: {name}")

    except Exception as e:
        logger.error(f"Error executing {name}: {e}")
        return [TextContent(type="text", text=f"Error: {str(e)}")]

async def run():
    from mcp.server.stdio import stdio_server
    
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options()
        )
