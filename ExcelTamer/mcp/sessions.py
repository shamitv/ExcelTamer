
import uuid
from typing import Dict, Optional
from .excel import ExcelAutomation

class SessionState:
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(SessionState, cls).__new__(cls)
            cls._instance.open_workbooks = {}
        return cls._instance

    def __init__(self):
        # type hints
        self.open_workbooks: Dict[str, ExcelAutomation]

    def add_workbook(self, automation: ExcelAutomation) -> str:
        wb_id = str(uuid.uuid4())
        self.open_workbooks[wb_id] = automation
        return wb_id

    def get_workbook(self, wb_id: str) -> Optional[ExcelAutomation]:
        return self.open_workbooks.get(wb_id)

    def find_workbook_id(self, app_pid: int, workbook_name: str) -> Optional[str]:
        """Find an existing session entry by its live Excel identity."""
        for wb_id, automation in self.open_workbooks.items():
            try:
                if (
                    int(automation.app.pid) == int(app_pid)
                    and automation.wb.name == workbook_name
                ):
                    return wb_id
            except Exception:
                # A stale or inaccessible Excel handle must not prevent discovery
                # of other workbooks.
                continue
        return None

    def remove_workbook(self, wb_id: str):
        if wb_id in self.open_workbooks:
            del self.open_workbooks[wb_id]

    def clear(self):
        self.open_workbooks.clear()

# Global session singleton
session = SessionState()
