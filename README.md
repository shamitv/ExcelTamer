# ExcelTamer

ExcelTamer is an AI Agent designed to work with Excel files. It can automate various tasks, making it easier to manage and manipulate Excel data programmatically.

## Features

- Automate Excel tasks
- Read and write Excel files
- Perform data analysis and manipulation

## Installation

To install ExcelTamer, clone the repo

## Documentation
For detailed usage instructions, including API examples and available tools, please refer to the [User Guide](docs/USER_GUIDE.md).

## Usage

test/invoke_agent.py is a sample script that demonstrates how to use ExcelTamer to automate Excel tasks.

1. Create an LLM 
2. Provide path to Excel File
3. Create an instance of Agent
4. Provide the task you want to perform
5. Run the agent

## ChatBot

test/ChainlitTest.py is a sample script that demonstrates how to use ExcelTamer as a ChatBot.


## MCP Server

ExcelTamer includes a full-featured [Model Context Protocol (MCP)](https://modelcontextprotocol.io) server. This allows you to use ExcelTamer capabilities directly within AI interfaces like **Claude Desktop**, **Cursor**, or any MCP-compliant client. It supports both **Stdio** and **HTTP SSE** transports.

Features:
*   Safe file access (sandboxing, read-only modes)
*   Structured reading and writing (batch updates, range reads)
*   Search and inspection
*   Checkpoints and rollback for safe editing

See the [MCP Server Guide](docs/mcp_server.md) for installation and configuration details.
