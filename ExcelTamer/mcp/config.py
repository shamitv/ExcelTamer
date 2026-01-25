
import os
from pathlib import Path

def get_env_list(key, default=None):
    val = os.environ.get(key, default)
    if not val:
        return []
    return [s.strip() for s in val.split(',')]

# Security: Allowed roots for file operations
# If empty, defaults to current working directory for safety in basic usage, 
# but in production should be explicit.
_default_root = str(Path.cwd())
ALLOWED_ROOTS = get_env_list("EXCELTAMER_MCP_ALLOWED_ROOTS", _default_root)

# Limits
MAX_CELLS_READ = int(os.environ.get("EXCELTAMER_MCP_MAX_CELLS_READ", "20000"))
MAX_CELLS_WRITE = int(os.environ.get("EXCELTAMER_MCP_MAX_CELLS_WRITE", "5000"))

# Defaults
DEFAULT_MODE = os.environ.get("EXCELTAMER_MCP_DEFAULT_MODE", "ro")

# Audit
AUDIT_LOG_DIR = os.environ.get("EXCELTAMER_MCP_AUDIT_LOG_DIR", "./.exceltamer_mcp_logs")
