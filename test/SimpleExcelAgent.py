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
    take_screenshot_impl(sheet_name, cell_range, output_path)


def take_screenshot_impl(sheet_name: str, cell_range: str, output_path: str) -> str:
    excel.capture_screenshot(sheet_name, cell_range, output_path)
    abs_path:str = get_absolute_path(output_path)
    return f"Screenshot saved to {abs_path}"

llm = ChatOpenAI(temperature=0)
agent = initialize_agent(
    tools=[query_cell, take_screenshot],
    llm=llm,
    agent=AgentType.OPENAI_FUNCTIONS,
    verbose=True
)

take_screenshot_impl(sheet_name="Expenses", cell_range="A1:E28", output_path="screenshot_2.pdf")

res = agent.invoke(input="Query cell B15 in Expenses , also take screenshot of A1:E28", context=excel)

print("Agent gave this final response:", res)