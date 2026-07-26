
from typing import Optional
import pandas as pd
from ..sessions import session
from ..config import MAX_CELLS_READ

def get_structure(workbook_id: str) -> dict:
    automation = session.get_workbook(workbook_id)
    if not automation:
        raise ValueError(f"Workbook {workbook_id} not found")
        
    structure = automation.get_structure()
    return {
        "workbook_id": workbook_id,
        "structure": structure
    }

def query_cell(workbook_id: str, sheet: str, cell: str) -> dict:
    automation = session.get_workbook(workbook_id)
    if not automation:
        raise ValueError(f"Workbook {workbook_id} not found")
        
    # ExcelAutomation.query_cell returns {'Value', 'Formula', 'VisibleText'}
    # We normalized keys to lowercase for MCP consistency if preferred, but keeping explicit mapping is safer.
    data = automation.query_cell(sheet, cell)
    
    return {
        "sheet": sheet,
        "cell": cell,
        "value": data.get("Value"),
        "formula": data.get("Formula"),
        "visible_text": data.get("VisibleText"),
        "entry_type": str(type(data.get("Value")))
    }

def read_range(
    workbook_id: str, 
    sheet: str, 
    range_a1: Optional[str] = None, 
    include_formulas: bool = False,
    max_rows: int = 1000,
    max_cols: int = 100
) -> dict:
    automation = session.get_workbook(workbook_id)
    if not automation:
        raise ValueError(f"Workbook {workbook_id} not found")

    # If range_a1 is None, ExcelAutomation defaults to used_range
    # We fetch it as a dataframe
    df = automation.get_range_as_dataframe(sheet, range_a1)
    
    # Check size limits
    row_count, col_count = df.shape
    total_cells = row_count * col_count
    
    truncated = False
    warnings = []
    
    if total_cells > MAX_CELLS_READ:
        warnings.append(f"Range exceeds global max cell limit ({MAX_CELLS_READ}). Truncating.")
        truncated = True
    
    # Also check per-request limits
    if row_count > max_rows:
        warnings.append(f"Row count ({row_count}) exceeds request limit ({max_rows}). Truncating rows.")
        df = df.head(max_rows)
        truncated = True
        
    if col_count > max_cols:
        warnings.append(f"Column count ({col_count}) exceeds request limit ({max_cols}). Truncating cols.")
        df = df.iloc[:, :max_cols]
        truncated = True
        
    # Convert to 2D list (values)
    # Note: ExcelAutomation.get_range_as_dataframe adds a 'RowNumber' column at index 0.
    # We should probably exclude that if we want raw data, or keep it if useful.
    # The plan asked for "values_2d". Let's keep it clean and remove the artificial 'RowNumber' if present.
    
    if "RowNumber" in df.columns:
        # Keep RowNumber might be useful for context, but usually 'read_range' expects the raw grid.
        # Let's drop it to match the requested matrix shape of the range.
        df = df.drop(columns=["RowNumber"])
        
    # Handling NaNs: replace with None or "" for JSON serialization
    df = df.where(pd.notnull(df), None)
    
    values = df.values.tolist()
    headers = list(df.columns)
    
    # If include_formulas is True, we can't easily get them from the DF alone 
    # because xlwings/pandas fetch values. 
    # ExcelAutomation doesn't have a bulk "get formulas" for a range yet.
    # For MVP, we will warn if requested but not supported efficiently, 
    # or loop (expensive). Plan says "formulas?". 
    # Let's defer formula-in-range for now or implement a specific bulk fetch in ExcelAutomation later.
    if include_formulas:
        warnings.append("include_formulas=True is not yet optimized/supported for bulk ranges. Returning values only.")

    return {
        "range_address": range_a1 or "UsedRange",
        "shape": [len(values), len(headers)],
        "headers": headers,
        "values": values,
        "truncated": truncated,
        "warnings": warnings
    }

def read_sheet_preview(workbook_id: str, sheet: str, rows: int = 50, cols: int = 20) -> dict:
    return read_range(workbook_id, sheet, range_a1=None, max_rows=rows, max_cols=cols)
