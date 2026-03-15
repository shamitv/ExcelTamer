
import asyncio
import os
import argparse
import sys
import json
from dotenv import load_dotenv
from typing import Optional
from contextlib import AsyncExitStack

# Load environment variables
load_dotenv()

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from mcp.client.sse import sse_client
from openai import AsyncOpenAI

async def run_client(
    file_path: str,
    transport: str = "stdio",
    port: int = 8123,
    model: str = "gpt-5-nano",
    base_url: Optional[str] = None
):
    print(f"Starting MCP Client...")
    print(f"File: {file_path}")
    print(f"Transport: {transport}")
    print(f"Model: {model}")

    # OpenAI Client Setup
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        print("Error: OPENAI_API_KEY not found in environment.")
        return

    openai_client = AsyncOpenAI(
        api_key=api_key,
        base_url=base_url or os.getenv("OPENAI_BASE_URL")
    )

    async with AsyncExitStack() as stack:
        # Connect to MCP Server
        if transport == "stdio":
            # Command to run the server
            command = sys.executable
            args = ["-m", "ExcelTamer.mcp.main"]
            env = os.environ.copy()
             # Update PYTHONPATH to include the current directory so ExcelTamer module is found
            current_dir = os.getcwd()
            if "PYTHONPATH" in env:
                env["PYTHONPATH"] += os.pathsep + current_dir
            else:
                 env["PYTHONPATH"] = current_dir

            server_params = StdioServerParameters(
                command=command,
                args=args,
                env=env
            )
            
            read, write = await stack.enter_async_context(stdio_client(server_params))
        
        elif transport == "sse":
            url = f"http://localhost:{port}/sse"
            read, write = await stack.enter_async_context(sse_client(url))
        
        else:
            print(f"Unknown transport: {transport}")
            return

        session = await stack.enter_async_context(ClientSession(read, write))
        await session.initialize()
        
        # List Tools
        tools_result = await session.list_tools()
        tools = tools_result.tools
        print(f"Connected to MCP Server. Found {len(tools)} tools.")
        
        # Open Workbook
        print(f"Opening workbook: {file_path}")
        wb_open_result = await session.call_tool("excel.open_workbook", arguments={"path": os.path.abspath(file_path)})
        wb_open_content = wb_open_result.content[0].text
        print(f"Open result: {wb_open_content}")
        
        # Extract workbook_id (simple parse for now, assuming result is string representation of dict or similar)
        # Ideally we parse the JSON response from the tool if it returns structured data.
        # But based on server.py, it returns str(result). 
        # For this test, let's assume we can get it or just ask LLM to use it.
        # To make it robust for LLM, LLM needs to see the output.
        

        # Chat Loop
        messages = [
            {
                "role": "system",
                "content": """You are an AI assistant that helps users with Excel files using the available tools. 
When you receive a tool call response, use it to output the user answer.
Warning: The 'excel.open_workbook' tool returns a result that contains the 'workbook_id'. You MUST use this 'workbook_id' for all subsequent tool calls.
"""
            },
            {
                "role": "user",
                "content": f"I have opened the workbook at {os.path.abspath(file_path)}. The output was: {wb_open_content}. Please tell me what sheets are in this workbook and read the first 5 rows of the first sheet."
            }
        ]

        # Convert MCP tools to OpenAI tools format
        openai_tools = []
        tool_map = {}
        for tool in tools:
            sanitized_name = tool.name.replace(".", "_")
            tool_map[sanitized_name] = tool.name
            openai_tools.append({
                "type": "function",
                "function": {
                    "name": sanitized_name,
                    "description": tool.description,
                    "parameters": tool.inputSchema
                }
            })

        print("\n--- Sending request to LLM ---")
        response = await openai_client.chat.completions.create(
            model=model,
            messages=messages,
            tools=openai_tools
        )

        message = response.choices[0].message
        print(f"LLM Response: {message.content or 'Tool Call'}")

        # Handle Tool Calls
        if message.tool_calls:
            messages.append(message) # Add assistant message with tool calls
            
            for tool_call in message.tool_calls:
                sanitized_fn_name = tool_call.function.name
                original_fn_name = tool_map.get(sanitized_fn_name, sanitized_fn_name)
                fn_args = json.loads(tool_call.function.arguments)
                
                print(f"Executing tool: {original_fn_name} (sanitized: {sanitized_fn_name}) with args: {fn_args}")
                
                try:
                    result = await session.call_tool(original_fn_name, arguments=fn_args)
                    result_text = "\n".join([c.text for c in result.content if c.type == "text"])
                except Exception as e:
                    result_text = f"Error executing tool: {e}"
                
                print(f"Result: {result_text[:200]}...") # Print preview
                
                messages.append({
                    "role": "tool",
                    "tool_call_id": tool_call.id,
                    "content": result_text
                })
            
            # Follow up with LLM
            print("\n--- Sending follow-up to LLM ---")
            response2 = await openai_client.chat.completions.create(
                model=model,
                messages=messages,
                tools=openai_tools
            )
            print(f"LLM Final Response: {response2.choices[0].message.content}")

        # Cleanup: Close workbook (best effort)
        # In a real app we'd track the ID properly. 
        # Here we rely on process termination to clean up unless we parse the ID.

def main():
    parser = argparse.ArgumentParser(description="MCP Validation Client")
    parser.add_argument("--file", required=True, help="Path to Excel file")
    parser.add_argument("--transport", default="stdio", choices=["stdio", "sse"], help="Transport mode")
    parser.add_argument("--port", type=int, default=8123, help="Port for SSE")
    parser.add_argument("--model", default="gpt-5-nano", help="OpenAI model")
    parser.add_argument("--url", help="OpenAI Base URL")
    
    args = parser.parse_args()
    
    try:
        asyncio.run(run_client(
            file_path=args.file,
            transport=args.transport,
            port=args.port,
            model=args.model,
            base_url=args.url
        ))
    except KeyboardInterrupt:
        pass

if __name__ == "__main__":
    main()
