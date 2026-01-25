
import asyncio
import sys
from .server import run

def main():
    try:
        asyncio.run(run())
    except KeyboardInterrupt:
        sys.exit(0)

if __name__ == "__main__":
    main()
