
import asyncio
import sys
import argparse
from .server import run

def main():
    parser = argparse.ArgumentParser(description="ExcelTamer MCP Server")
    parser.add_argument("--port", type=int, help="Port to run the SSE server on (default: stdio mode)")
    args = parser.parse_args()
    
    try:
        asyncio.run(run(port=args.port))
    except KeyboardInterrupt:
        sys.exit(0)

if __name__ == "__main__":
    main()
