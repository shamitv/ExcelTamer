
import uuid
from typing import Dict, Optional
from ..ExcelAutomation import ExcelAutomation

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

    def remove_workbook(self, wb_id: str):
        if wb_id in self.open_workbooks:
            del self.open_workbooks[wb_id]

    def clear(self):
        self.open_workbooks.clear()

# Global session singleton
session = SessionState()
