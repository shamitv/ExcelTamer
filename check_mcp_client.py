
try:
    from mcp import ClientSession, StdioServerParameters
    print("mcp.ClientSession found")
except ImportError as e:
    print(f"mcp.ClientSession NOT found: {e}")

try:
    from mcp.client.stdio import stdio_client
    print("mcp.client.stdio.stdio_client found")
except ImportError as e:
    print(f"mcp.client.stdio.stdio_client NOT found: {e}")

try:
    from mcp.client.sse import sse_client
    print("mcp.client.sse.sse_client found")
except ImportError as e:
    print(f"mcp.client.sse.sse_client NOT found: {e}")
