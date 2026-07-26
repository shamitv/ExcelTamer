"""Pure MCP validation client for the ExcelTamer server."""

import argparse
import ast
import asyncio
import os
import sys
from contextlib import AsyncExitStack

from mcp import ClientSession, StdioServerParameters
from mcp.client.sse import sse_client
from mcp.client.stdio import stdio_client


async def run_client(
    file_path: str | None = None,
    transport: str = "stdio",
    port: int = 8123,
) -> None:
    """Validate discovery and, optionally, workbook operations."""
    async with AsyncExitStack() as stack:
        if transport == "stdio":
            env = os.environ.copy()
            env["PYTHONPATH"] = os.pathsep.join(
                path
                for path in (os.getcwd(), env.get("PYTHONPATH"))
                if path
            )
            params = StdioServerParameters(
                command=sys.executable,
                args=["-m", "ExcelTamer.mcp.main"],
                env=env,
            )
            read, write = await stack.enter_async_context(stdio_client(params))
        else:
            read, write = await stack.enter_async_context(
                sse_client(f"http://localhost:{port}/sse")
            )

        session = await stack.enter_async_context(ClientSession(read, write))
        await session.initialize()

        tools = await session.list_tools()
        resources = await session.list_resources()
        prompts = await session.list_prompts()
        print(
            f"Connected: {len(tools.tools)} tools, "
            f"{len(resources.resources)} resources, "
            f"{len(prompts.prompts)} prompts"
        )

        for prompt in prompts.prompts:
            result = await session.get_prompt(prompt.name)
            print(f"Prompt '{prompt.name}': {len(result.messages)} message(s)")

        if not file_path:
            return

        absolute_path = os.path.abspath(file_path)
        opened = await session.call_tool(
            "excel.open_workbook",
            arguments={"path": absolute_path, "mode": "ro"},
        )
        open_payload = ast.literal_eval(opened.content[0].text)
        workbook_id = open_payload["workbook_id"]
        try:
            structure = await session.call_tool(
                "excel.get_structure",
                arguments={"workbook_id": workbook_id},
            )
            print(f"Workbook structure: {structure.content[0].text}")
        finally:
            await session.call_tool(
                "excel.close",
                arguments={"workbook_id": workbook_id},
            )


def main() -> None:
    parser = argparse.ArgumentParser(description="Validate the ExcelTamer MCP server")
    parser.add_argument(
        "--file",
        help="Optional workbook path for an open/get-structure/close validation",
    )
    parser.add_argument(
        "--transport",
        choices=["stdio", "sse"],
        default="stdio",
    )
    parser.add_argument("--port", type=int, default=8123)
    args = parser.parse_args()
    asyncio.run(
        run_client(
            file_path=args.file,
            transport=args.transport,
            port=args.port,
        )
    )


if __name__ == "__main__":
    main()
