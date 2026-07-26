
import logging
import json
import os
from datetime import datetime
from pathlib import Path
from .config import AUDIT_LOG_DIR

# Ensure audit directory exists
try:
    Path(AUDIT_LOG_DIR).mkdir(parents=True, exist_ok=True)
except Exception as e:
    # Fall back to the current directory if the configured location is unavailable.
    print(f"Warning: Could not create audit log dir {AUDIT_LOG_DIR}: {e}")

# Configure a specific logger for audit
audit_logger = logging.getLogger("ExcelTamerAudit")
audit_logger.setLevel(logging.INFO)

# File handler for audit logs (one log file per session/day or single rotating file)
# For simplicity, single file appened.
audit_file = Path(AUDIT_LOG_DIR) / "audit.jsonl"
handler = logging.FileHandler(audit_file)
handler.setFormatter(logging.Formatter('%(message)s'))
audit_logger.addHandler(handler)

def log_write(tool_name: str, workbook_id: str, details: dict):
    """
    Log a write operation.
    """
    entry = {
        "timestamp": datetime.now().isoformat(),
        "tool": tool_name,
        "workbook_id": workbook_id,
        "details": details
    }
    audit_logger.info(json.dumps(entry))
