
import os
from pathlib import Path
from .config import ALLOWED_ROOTS

class SecurityError(Exception):
    pass

def validate_path(request_path: str, allow_write: bool = False) -> Path:
    """
    Validates that the given path is within ALLOWED_ROOTS.
    Resolves symlinks and checks for traversal.
    """
    try:
        # Resolve user path to absolute
        abs_path = Path(request_path).resolve()
        
        # Check if it matches any allowed root
        is_allowed = False
        for root in ALLOWED_ROOTS:
            allowed_root = Path(root).resolve()
            # method 1: check if allowed_root is a parent of abs_path
            # calculate relative path; if it starts with '..', it's outside
            try:
                abs_path.relative_to(allowed_root)
                is_allowed = True
                break
            except ValueError:
                continue
        
        if not is_allowed:
            raise SecurityError(f"Path '{request_path}' is not within allowed roots: {ALLOWED_ROOTS}")
            
        return abs_path

    except Exception as e:
        if isinstance(e, SecurityError):
            raise
        raise SecurityError(f"Invalid path: {e}")
