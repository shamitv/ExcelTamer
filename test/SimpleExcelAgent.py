import os

from langchain.agents import initialize_agent, AgentType
from langchain_core.tools import tool
from langchain_openai import ChatOpenAI

from dotenv import load_dotenv

from ExcelManager import ExcelManager

load_dotenv()


def get_absolute_path(file_path: str) -> str:
    return os.path.abspath(file_path)

excel:ExcelManager = ExcelManager(file_path="example.xlsx")

@tool
def query_cell(sheet_name: str, cell: str) -> dict:
    """Retrieve the value and formula of a specific cell.  Ensure that sheet name and cell are provided as two parameters."""
    contents = excel.query_cell(sheet_name, cell)
    return contents


@tool
def take_screenshot(sheet_name: str, cell_range: str, output_path: str) -> str:
    """Capture a screenshot of the specified sheet or range. Outfile path must be provided, this file will be a png."""
    excel.capture_screenshot(sheet_name, cell_range, output_path)
    abs_path:str = get_absolute_path(output_path)
    return f"Screenshot saved to {abs_path}"


@tool
def get_structure() -> list[dict]:
    """Inspect Excel workbook structure (sheets, dimensions, named ranges).

    Useful for inspecting workbook organization before data operations. Does not access raw cell data."""

    structure = excel.get_structure()
    return structure


llm = ChatOpenAI(temperature=0)
agent = initialize_agent(
    tools=[query_cell, take_screenshot, get_structure],
    llm=llm,
    agent=AgentType.OPENAI_FUNCTIONS,
    verbose=True
)



res = agent.invoke(input="Query cell B15 in Expenses , also take screenshot of A1:E28", context=excel)

print("Agent gave this final response:", res)

res = agent.invoke(input="What is the most expensive thing in Feb. Note, first look at excel's structure to figure how where this data might be", context=excel)

print("Agent gave this final response:", res)