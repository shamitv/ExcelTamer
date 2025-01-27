import asyncio
import base64
import os
import tempfile
from typing import ClassVar, Any, List
from concurrent.futures import ThreadPoolExecutor

from langchain_core.language_models import BaseChatModel
from langchain_core.messages import HumanMessage
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
        #loop = asyncio.get_event_loop()
        #return await loop.run_in_executor(self._executor, self._get_structure_sync)
        return self._get_structure_sync()

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


class ExcelCellValueTool(BaseTool):
        """Tool to query Formula / Value of a cell."""

        tool_name: ClassVar[str] = "excel_query_cell"
        tool_description: ClassVar[str] =  """Retrieve the value and formula of a specific cell.  Ensure that sheet name and cell are provided as two parameters."""

        _excel_automation: ExcelAutomation = PrivateAttr()
        _executor: ThreadPoolExecutor = PrivateAttr()

        def __init__(self, excel_automation: ExcelAutomation, executor: ThreadPoolExecutor):
            """Constructor accepts an ExcelAutomation instance and a ThreadPoolExecutor."""
            super().__init__(name=self.tool_name, description=self.tool_description)
            self._excel_automation = excel_automation
            self._executor = executor

        def _impl(self, sheet_name: str, cell: str) -> dict:
            """Sync wrapper for the get_structure method."""
            # Use the ThreadPoolExecutor to ensure that xlwings interacts with Excel in a separate thread
            future = self._executor.submit(self._excel_automation.query_cell, sheet_name, cell)
            return future.result()

        def _run(self, sheet_name: str, cell: str) -> Any:
            """Sync entry point for the tool."""
            return self._impl(sheet_name, cell)

        async def _arun(self, sheet_name: str, cell: str) -> Any:
            """Async entry point for the tool."""
            return self._impl(sheet_name, cell)

        @property
        def name(self) -> str:
            """The name of the tool."""
            return self.tool_name

        @property
        def description(self) -> str:
            """A brief description of the tool's functionality."""
            return self.tool_description



class ExcelAnalyzeImageTool(BaseTool):
    """Tool to analyze an image of an Excel sheet."""

    tool_name: ClassVar[str] = "excel_analyze_image"
    tool_description: ClassVar[str] ="""Capture a screenshot of the specified sheet or range
    and answers a related question.

    :param question: The question related to the image.
    :param sheet_name: The name of the sheet to capture the screenshot.
    :param cell_range: (optional) The range of cells to
                capture the screenshot. Screenshot of whole sheet is captured if this parameter is not provided.
    :return: Answer to the question.
    """

    _excel_automation: ExcelAutomation = PrivateAttr()
    _llm: BaseChatModel = PrivateAttr()
    _executor: ThreadPoolExecutor = PrivateAttr()

    def __init__(self, llm: BaseChatModel, excel_automation: ExcelAutomation, executor: ThreadPoolExecutor):
        """Constructor accepts an image path and a ThreadPoolExecutor."""
        super().__init__(name=self.tool_name, description=self.tool_description)
        self._llm = llm
        self._executor = executor
        self._excel_automation = excel_automation

    def take_screenshot(self,sheet_name: str, cell_range: str) -> str:
        """Capture a screenshot of the specified sheet or range,
         and return the base64 string of the image data.
         If cell_range is not provided, the entire sheet will be captured.
         """

        # Create a temporary file path
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_file:
            output_path = temp_file.name

        # Capture the screenshot
        self._excel_automation.capture_screenshot_png(sheet_name, output_path, cell_range)

        # Read the image and convert to base64
        with open(output_path, "rb") as image_file:
            base64_image = base64.b64encode(image_file.read()).decode('utf-8')

        # Clean up the temporary file
        os.remove(output_path)

        # Create data URL
        data_url = f"data:image/png;base64,{base64_image}"

        return data_url

    def ask_question_about_image_base64(self,encoded_image_url, question):
        """
        Submits an image (in form of Data URL) and a related question to LLM and returns the concise response.

        :param encoded_image_url: Base64 encoded image string (data URL with header).
        :param question: The question related to the image.
        :return: The response from LLM.
        """
        # Prepare the message content
        messages = [
            HumanMessage(content=[
                {"type": "text", "text": f"Please provide a "
                                         f"concise response to following question \n\n##Question\n\n{question} ."},
                {
                    "type": "image_url",
                    "image_url": {
                        "url": encoded_image_url
                    }
                }
            ])
        ]

        # Get the response from OpenAI using the global llm
        response = self._llm.invoke(messages)
        return response.content


    def _impl(self, question: str, sheet_name: str, cell_range: str = None) -> str:
        """Sync wrapper for the analyze_image method."""
        # Use the ThreadPoolExecutor to ensure that image processing is done in a separate thread
        future = self._executor.submit(self.take_screenshot,  sheet_name, cell_range)
        image_data_url =  future.result()

        response = self.ask_question_about_image_base64(image_data_url, question)

        return response


    def _run(self,question: str, sheet_name: str, cell_range: str = None) -> str:
        """Sync entry point for the tool."""
        return self._impl(question, sheet_name, cell_range)

    async def _arun(self, question: str, sheet_name: str, cell_range: str = None) -> str:
        """Async entry point for the tool."""
        return self._impl(question, sheet_name, cell_range)

    @property
    def name(self) -> str:
        """The name of the tool."""
        return self.tool_name

    @property
    def description(self) -> str:
        """A brief description of the tool's functionality."""
        return self.tool_description

class ExcelSaveTool(BaseTool):
    """Tool to save the Excel workbook."""

    tool_name: ClassVar[str] = "excel_save"
    tool_description: ClassVar[str] = """Save the Excel workbook to the specified file path. If no file path is provided, the workbook will be saved in its current location."""

    _excel_automation: ExcelAutomation = PrivateAttr()
    _executor: ThreadPoolExecutor = PrivateAttr()

    def __init__(self, excel_automation: ExcelAutomation, executor: ThreadPoolExecutor):
        """Constructor accepts an ExcelAutomation instance and a ThreadPoolExecutor."""
        super().__init__(name=self.tool_name, description=self.tool_description)
        self._excel_automation = excel_automation
        self._executor = executor

    def _impl(self, file_path: str = None) -> None:
        """Sync wrapper for the save method."""
        # Use the ThreadPoolExecutor to ensure that xlwings interacts with Excel in a separate thread
        future = self._executor.submit(self._excel_automation.save, file_path)
        future.result()

    def _run(self, file_path: str = None) -> None:
        """Sync entry point for the tool."""
        return self._impl(file_path)

    async def _arun(self, file_path: str = None) -> None:
        """Async entry point for the tool."""
        return self._impl(file_path)

    @property
    def name(self) -> str:
        """The name of the tool."""
        return self.tool_name

    @property
    def description(self) -> str:
        """A brief description of the tool's functionality."""
        return self.tool_description

class ExcelCloseTool(BaseTool):
    """Tool to close the Excel workbook."""

    tool_name: ClassVar[str] = "excel_close"
    tool_description: ClassVar[str] = """Close the Excel workbook."""

    _excel_automation: ExcelAutomation = PrivateAttr()
    _executor: ThreadPoolExecutor = PrivateAttr()

    def __init__(self, excel_automation: ExcelAutomation, executor: ThreadPoolExecutor):
        """Constructor accepts an ExcelAutomation instance and a ThreadPoolExecutor."""
        super().__init__(name=self.tool_name, description=self.tool_description)
        self._excel_automation = excel_automation
        self._executor = executor

    def _impl(self) -> None:
        """Sync wrapper for the close method."""
        # Use the ThreadPoolExecutor to ensure that xlwings interacts with Excel in a separate thread
        future = self._executor.submit(self._excel_automation.close)
        future.result()

    def _run(self) -> None:
        """Sync entry point for the tool."""
        return self._impl()

    async def _arun(self) -> None:
        """Async entry point for the tool."""
        return self._impl()

    @property
    def name(self) -> str:
        """The name of the tool."""
        return self.tool_name

    @property
    def description(self) -> str:
        """A brief description of the tool's functionality."""
        return self.tool_description

class ExcelWriteCellTool(BaseTool):
    """Tool to modify Value of a cell."""

    tool_name: ClassVar[str] = "excel_change_cell_value"
    tool_description: ClassVar[
        str] = """Modify value of a specific cell.  Ensure that sheet name and cell are provided as two parameters.
                    :param sheet_name: The name of the sheet to capture the screenshot.
                    :param cell: Cell address.
                    :param value: New value to be written to the cell.
        """

    _excel_automation: ExcelAutomation = PrivateAttr()
    _executor: ThreadPoolExecutor = PrivateAttr()

    def __init__(self, excel_automation: ExcelAutomation, executor: ThreadPoolExecutor):
        """Constructor accepts an ExcelAutomation instance and a ThreadPoolExecutor."""
        super().__init__(name=self.tool_name, description=self.tool_description)
        self._excel_automation = excel_automation
        self._executor = executor

    def _impl(self, sheet_name: str, cell: str, value: str) -> dict:
        """Sync wrapper for the get_structure method."""
        # Use the ThreadPoolExecutor to ensure that xlwings interacts with Excel in a separate thread
        future = self._executor.submit(self._excel_automation.write_cell,sheet_name, cell,value)
        return future.result()

    def _run(self, sheet_name: str, cell: str,value:str) -> Any:
        """Sync entry point for the tool."""
        return self._impl(sheet_name, cell,value)

    async def _arun(self, sheet_name: str, cell: str, value:str) -> Any:
        """Async entry point for the tool."""
        return self._impl(sheet_name, cell,value)

    @property
    def name(self) -> str:
        """The name of the tool."""
        return self.tool_name

    @property
    def description(self) -> str:
        """A brief description of the tool's functionality."""
        return self.tool_description