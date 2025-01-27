from concurrent.futures import ThreadPoolExecutor

from dotenv import load_dotenv
load_dotenv()

import base64
import os
import tempfile

from ExcelTamer.ExcelAutomation import ExcelManager

from langchain.agents import initialize_agent, AgentType
from langchain_core.tools import tool
from langchain_openai import ChatOpenAI
from langchain.schema import HumanMessage

executor = ThreadPoolExecutor()

def get_absolute_path(file_path: str) -> str:
    return os.path.abspath(file_path)

# Initialize the ExcelManager
excel:ExcelManager
# Function to initialize ExcelManager
def init_excel_manager(file_path: str) -> ExcelManager:
    return ExcelManager(file_path=file_path)

# Submit the task to the executor
future_excel = executor.submit(init_excel_manager, "example.xlsx")

# Get the result (ExcelManager instance)
excel = future_excel.result()





@tool
def query_cell(sheet_name: str, cell: str) -> dict:
    """Retrieve the value and formula of a specific cell.  Ensure that sheet name and cell are provided as two parameters."""
    future = executor.submit(excel.query_cell, sheet_name, cell)
    contents = future.result()
    return contents


def take_screenshot(sheet_name: str, cell_range: str) -> str:
    """Capture a screenshot of the specified sheet or range,
     and return the base64 string of the image data.
     If cell_range is not provided, the entire sheet will be captured.
     """

    # Create a temporary file path
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_file:
        output_path = temp_file.name

    # Capture the screenshot
    excel.capture_screenshot(sheet_name, output_path, cell_range)

    # Read the image and convert to base64
    with open(output_path, "rb") as image_file:
        base64_image = base64.b64encode(image_file.read()).decode('utf-8')

    # Clean up the temporary file
    os.remove(output_path)

    # Create data URL
    data_url = f"data:image/png;base64,{base64_image}"

    return data_url


@tool
def get_structure() -> list[dict]:
    """Inspect Excel workbook structure. It lists all the sheets. For each sheet, it provides :
    - Sheet Name
    - Number of Rows
    - Number of Columns
    - Range of the used cells
    - Named ranges in the sheet (name and reference)


    Useful for inspecting workbook for data operations. Does not access raw cell data."""

    future = executor.submit(excel.get_structure)
    structure = future.result()
    return structure


@tool
def analyze_excel_image(question:str, sheet_name: str, cell_range: str = None) -> str:
    """
    Capture a screenshot of the specified sheet or range
    and answers a related question.

    :param question: The question related to the image.
    :param sheet_name: The name of the sheet to capture the screenshot.
    :param cell_range: (optional) The range of cells to
                capture the screenshot. Screenshot of whole sheet is captured if this parameter is not provided.
    :return: Answer to the question.
    """
    # Capture the screenshot and get the base64 encoded image
    future = executor.submit(take_screenshot, sheet_name, cell_range)
    base64_image = future.result()

    # Ask the question about the image
    response = ask_question_about_image_base64(base64_image, question)

    return response


def ask_question_about_image_base64(encoded_image, question):
    """
    Submits an image (in form of Data URL) and a related question to LLM and returns the concise response.

    :param encoded_image: Base64 encoded image string (data URL with header).
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
                    "url": encoded_image
                }
            }
        ])
    ]

    # Get the response from OpenAI using the global llm
    response = llm.invoke(messages)
    return response.content


llm = ChatOpenAI(model_name="gpt-4o-mini", temperature=0)
agent = initialize_agent(
    tools=[query_cell, analyze_excel_image, get_structure],
    llm=llm,
    agent=AgentType.OPENAI_FUNCTIONS,
    verbose=True
)

"""###

res = agent.invoke(input="Query cell B15 in Expenses", context=excel)

print("Agent gave this final response:", res)

res = agent.invoke(input="Which sheet might have total expenses for Feb ? "
                         "What are the expenses for Jan ? "
                         "Note: that you can request a screenshot to examine the sheet or range"
                         "If a screenshot is available, examine the screenshot to get values", context=excel)

print("Agent gave this final response:", res)

"""