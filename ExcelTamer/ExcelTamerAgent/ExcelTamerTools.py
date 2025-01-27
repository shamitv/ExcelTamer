import asyncio
from typing import ClassVar, Any, List
from concurrent.futures import ThreadPoolExecutor
from pydantic import PrivateAttr
from langchain.tools import BaseTool
from ExcelTamer.ExcelAutomation import ExcelAutomation

class ExcelGetStructureTool(BaseTool):
    """Tool to inspect the structure of an Excel workbook."""

    tool_name: ClassVar[str] = "excel_get_structure"
    tool_description: ClassVar[str] = """Inspect Excel workbook structure. It lists all the sheets. For each sheet, it provides :
    - Sheet Name
    - Number of Rows
    - Number of Columns
    - Range of the used cells
    - Named ranges in the sheet (name and reference)

    Useful for inspecting workbook for data operations. Does not access raw cell data."""

    _excel_automation: ExcelAutomation = PrivateAttr()
    _executor: ThreadPoolExecutor = PrivateAttr()

    def __init__(self, excel_automation: ExcelAutomation, executor: ThreadPoolExecutor):
        """Constructor accepts an ExcelAutomation instance and a ThreadPoolExecutor."""
        super().__init__(name=self.tool_name, description=self.tool_description)
        self._excel_automation = excel_automation
        self._executor = executor

    def _get_structure_sync(self) -> List[dict]:
        """Sync wrapper for the get_structure method."""
        # Use the ThreadPoolExecutor to ensure that xlwings interacts with Excel in a separate thread
        future = self._executor.submit(self._excel_automation.get_structure)
        return future.result()

    async def _get_structure_async(self) -> List[dict]:
        """Async wrapper for the get_structure method using ThreadPoolExecutor."""
        loop = asyncio.get_event_loop()
        return await loop.run_in_executor(self._executor, self._get_structure_sync)

    def _run(self, *args: Any, **kwargs: Any) -> Any:
        """Sync entry point for the tool."""
        return self._get_structure_sync()

    async def _arun(self, *args: Any, **kwargs: Any) -> Any:
        """Async entry point for the tool."""
        return await self._get_structure_async()

    @property
    def name(self) -> str:
        """The name of the tool."""
        return self.tool_name

    @property
    def description(self) -> str:
        """A brief description of the tool's functionality."""
        return self.tool_description