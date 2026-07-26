
from ..config import DEFAULT_MODE
from ..excel import ExcelAutomation
from ..safety import validate_path
from ..sessions import session

def open_workbook(path: str, mode: str = DEFAULT_MODE) -> dict:
    """
    Opens a workbook and returns its ID and metadata.
    mode: 'ro' (read-only) or 'rw' (read-write)
    """
    # 1. Validate path
    abs_path = validate_path(path)
    
    # 2. Open workbook
    # Note: Xlwings doesn't strictly support 'ro' at open() level easily without API flags,
    # but we will enforce it at the Write tool level.
    # We pass the validated absolute path string.
    automation = ExcelAutomation(file_path=str(abs_path))
    
    # 3. Store in session
    wb_id = session.add_workbook(automation)
    
    # 4. Gather metadata
    sheets = automation.list_sheets()
    
    return {
        "workbook_id": wb_id,
        "filename": abs_path.name,
        "sheets": sheets,
        "mode": mode
    }

def close_workbook(workbook_id: str) -> dict:
    automation = session.get_workbook(workbook_id)
    if not automation:
        raise ValueError(f"Workbook {workbook_id} not found")
        
    # Close without quitting app to support multiple workbooks
    automation.close(quit_app=False)
    session.remove_workbook(workbook_id)
    
    return {"status": "closed", "workbook_id": workbook_id}

def save_workbook(workbook_id: str) -> dict:
    automation = session.get_workbook(workbook_id)
    if not automation:
        raise ValueError(f"Workbook {workbook_id} not found")
        
    automation.save()
    return {"status": "saved", "workbook_id": workbook_id}

def save_as_workbook(workbook_id: str, output_path: str) -> dict:
    # 1. Validate output path
    abs_path = validate_path(output_path, allow_write=True)
    
    automation = session.get_workbook(workbook_id)
    if not automation:
        raise ValueError(f"Workbook {workbook_id} not found")
        
    automation.save(str(abs_path))
    
    return {"status": "saved_as", "path": str(abs_path), "workbook_id": workbook_id}
