
from typing import Any, List
from ..sessions import session
from ..config import MAX_CELLS_WRITE
from ..audit import log_write

def change_cell_value(workbook_id: str, sheet: str, cell: str, value: Any) -> dict:
    automation = session.get_workbook(workbook_id)
    if not automation:
        raise ValueError(f"Workbook {workbook_id} not found")
        
    automation.write_cell(sheet, cell, value)
    
    # Audit logic
    log_write("change_cell_value", workbook_id, {
        "sheet": sheet,
        "cell": cell,
        "value_preview": str(value)[:50]
    })
    
    return {"status": "ok", "workbook_id": workbook_id}

def batch_update_cells(workbook_id: str, updates: List[dict]) -> dict:
    """
    updates: list of dicts with keys: sheet, cell, value
    """
    automation = session.get_workbook(workbook_id)
    if not automation:
        raise ValueError(f"Workbook {workbook_id} not found")
        
    if len(updates) > MAX_CELLS_WRITE:
        raise ValueError(f"Batch update exceeds limit of {MAX_CELLS_WRITE} cells")
        
    count = 0
    # Xlwings/ExcelAutomation doesn't have a native 'batch disjoint write' yet
    # so we iterate. Ideally, we optimize contiguous ranges later.
    # But this is still better than 100 separate tool calls over MCP.
    
    # Group by sheet to minimize sheet switching overhead if any (xlwings handles this well though)
    for update in updates:
        sheet = update.get("sheet")
        cell = update.get("cell")
        value = update.get("value")
        
        # We allow "formula" key as an alias for value if client separates them, 
        # but ExcelAutomation.write_cell handles both via .value assignment normally.
        # If specific formula assignment is needed, we treat it same as value.
        if "formula" in update and value is None:
            value = update["formula"]
            
        if sheet and cell:
            automation.write_cell(sheet, cell, value)
            count += 1
            
    log_write("batch_update_cells", workbook_id, {
        "count": count
    })
            
    return {"status": "ok", "updated_count": count}

def write_range(
    workbook_id: str, 
    sheet: str, 
    start_cell: str, 
    values: List[List[Any]]
) -> dict:
    automation = session.get_workbook(workbook_id)
    if not automation:
        raise ValueError(f"Workbook {workbook_id} not found")
        
    # Check dimensions
    rows = len(values)
    cols = len(values[0]) if rows > 0 else 0
    total_cells = rows * cols
    
    if total_cells > MAX_CELLS_WRITE:
        raise ValueError(f"Write range exceeds limit of {MAX_CELLS_WRITE} cells ({total_cells} requested)")
        
    ws = automation.wb.sheets[sheet]
    ws.range(start_cell).value = values
    
    # Calculate end range for return info
    # (Simplified assumption: contiguous block)
    # xlwings handles the size automatically.
    
    log_write("write_range", workbook_id, {
        "sheet": sheet,
        "start_cell": start_cell,
        "rows": rows,
        "cols": cols
    })
    
    return {
        "status": "ok", 
        "written_cells": total_cells, 
        "shape": [rows, cols]
    }
