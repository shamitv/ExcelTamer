# ExcelTamer User Guide

`ExcelTamer` is a Python package that provides agentic tools for Excel automation using LLMs. It exposes a low-level automation wrapper (`ExcelAutomation`) and a set of LangChain-compatible tools (`ExcelTamerTools`).

## Installation

```bash
pip install ExcelTamer
```

## Usage

### 1. Basic Automation with `ExcelAutomation`

The `ExcelAutomation` class provides a direct wrapper around `xlwings` for common Excel operations.

```python
from ExcelTamer import ExcelAutomation

# Open a workbook (or create a new one if path is None)
excel = ExcelAutomation(file_path="path/to/workbook.xlsx")

# List sheets
sheets = excel.list_sheets()
print(f"Sheets: {sheets}")

# Read a cell
value = excel.read_cell(sheet_name="Sheet1", cell="A1")
print(f"Value in A1: {value}")

# Write to a cell
excel.write_cell(sheet_name="Sheet1", cell="B2", value="Hello World")

# Get a range as a DataFrame
df = excel.get_range_as_dataframe(sheet_name="Sheet1", cell_range="A1:C10")
print(df)

# Save and Close
excel.save()
excel.close()
```

### 2. Agentic Tools with `ExcelTamerTools`

`ExcelTamerTools` provides a collection of tools ready to be used with LangChain agents.

```python
from langchain_openai import ChatOpenAI
from ExcelTamer import ExcelAutomation, create_agent

# Initialize the agent
llm = ChatOpenAI(model="gpt-4o", temperature=0)
agent_executor = create_agent(
    excel_path="path/to/workbook.xlsx",
    llm=llm
)

# Run a query
response = agent_executor.invoke({"input": "What is the total generated in 2023?"})
print(response["output"])
```

### Available Tools

The agent has access to the following tools:

-   `excel_get_structure`: Inspects workbook structure (sheets, ranges).
-   `excel_query_cell`: Retrieves value and formula of a specific cell.
-   `excel_analyze_image`: Captures a screenshot of a range and answers questions about it using vision capabilities.
-   `excel_save`: Saves the workbook.
-   `excel_close`: Closes the workbook.
-   `excel_change_cell_value`: Modifies a cell's value.
-   `excel_search_cell`: Searches for cells containing a specific value.
-   `excel_range_or_sheet_as_markdown`: Extracts data as a Markdown table.
-   `excel_find_metric_value`: Intelligent search for financial metrics in time-series data.
