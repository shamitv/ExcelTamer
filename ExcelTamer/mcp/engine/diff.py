
import os
import shutil
import uuid
from pathlib import Path
from tempfile import gettempdir
from typing import Dict, List, Optional
from ..sessions import session, ExcelAutomation

# We'll store checkpoint paths in session memory
# workbook_id -> {checkpoint_name -> file_path}
checkpoints: Dict[str, Dict[str, str]] = {}

def _get_checkpoint_dir():
    # Store checkpoints in a temp dir or hidden local dir
    base = Path(gettempdir()) / "exceltamer_checkpoints"
    base.mkdir(parents=True, exist_ok=True)
    return base

def checkpoint_create(workbook_id: str, name: str) -> dict:
    automation = session.get_workbook(workbook_id)
    if not automation:
        raise ValueError(f"Workbook {workbook_id} not found")
        
    # We need to save the current state to a separate file.
    # Xlwings .save() overwrites the current file.
    # To create a checkpoint without moving the user's active file pointer,
    # we can save a copy.
    
    cp_dir = _get_checkpoint_dir()
    cp_filename = f"{workbook_id}_{name}_{uuid.uuid4().hex[:8]}.xlsx"
    cp_path = cp_dir / cp_filename
    
    # Save a copy
    # automation.wb.save(path) changes the active workbook to that path in Excel UI usually?
    # Let's check xlwings docs/behavior. 
    # wb.save(path) "Saves the Workbook to the specified filename." -> Effectively Save As.
    # If we do that, our session is now pointing to the checkpoint file? Yes.
    # We want to stay on the main file but snapshot it.
    
    # Workaround: 
    # 1. Save current workbook to disk (ensure it's up to date).
    automation.save()
    
    # 2. Copy the file on disk to checkpoint path.
    # Access underlying full path
    current_path = automation.wb.fullname
    shutil.copy2(current_path, cp_path)
    
    if workbook_id not in checkpoints:
        checkpoints[workbook_id] = {}
    checkpoints[workbook_id][name] = str(cp_path)
    
    return {"status": "created", "name": name, "path": str(cp_path)}

def checkpoint_rollback(workbook_id: str, name: str) -> dict:
    automation = session.get_workbook(workbook_id)
    if not automation:
        raise ValueError(f"Workbook {workbook_id} not found")
        
    if workbook_id not in checkpoints or name not in checkpoints[workbook_id]:
        raise ValueError(f"Checkpoint '{name}' not found for this workbook")
        
    cp_path = checkpoints[workbook_id][name]
    
    # Rollback strategy:
    # 1. Close current workbook (discard changes? well we are rolling back)
    # 2. Overwrite current workbook file with checkpoint file
    # 3. Re-open
    
    original_path = automation.wb.fullname
    
    # Close
    automation.close(quit_app=False)
    session.remove_workbook(workbook_id)
    
    # Overwrite
    shutil.copy2(cp_path, original_path)
    
    # Re-open (reuse ID?)
    # ideally we keep the same ID for the client's sake
    # But we need to re-init automation
    new_automation = ExcelAutomation(file_path=original_path)
    
    # We need to hack session to restore the ID mapping
    session.open_workbooks[workbook_id] = new_automation
    
    return {"status": "rolled_back", "name": name}

def preview_diff(workbook_id: str, max_changes: int = 200) -> dict:
    """
    Shows a summary of changes.
    Since we don't track cell-by-cell diffs in memory yet (complex), 
    we will rely on:
    1. If we have a 'base' checkpoint, maybe compare? (Hard to do nicely in MVP)
    2. Or return the Audit Log entries for this session/workbook?
    
    Let's return the last N actions from the in-memory audit trail/session tracker?
    We didn't implement in-memory audit trail, only file log. 
    
    Let's read the audit log file and filter by workbook_id.
    """
    from ..config import AUDIT_LOG_DIR
    import json
    
    audit_file = Path(AUDIT_LOG_DIR) / "audit.jsonl"
    changes = []
    
    if audit_file.exists():
        # Read backward optimization could be done, but for MVP read all is fine
        with open(audit_file, "r") as f:
            for line in f:
                try:
                    entry = json.loads(line)
                    if entry.get("workbook_id") == workbook_id:
                        changes.append(entry)
                except:
                    pass
                    
    # Sort by timestamp desc?
    # Usually we want chronological.
    
    # This is "History" rather than "Diff". 
    # True Diff requires comparing values. 
    # For MVP, History is a good proxy for "What have I done?".
    
    return {
        "summary": f"Found {len(changes)} recorded actions for this workbook.",
        "recent_actions": changes[-max_changes:],
        "truncated": len(changes) > max_changes
    }
