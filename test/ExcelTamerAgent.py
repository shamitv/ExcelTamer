from langchain.agents import AgentType, initialize_agent
from langchain.tools import Tool
from langchain.chat_models import ChatOpenAI
import xlwings as xw
import pandas as pd



class ExcelTamerAgent:
    def __init__(self, file_path=None):
        if file_path:
            self.file_path = file_path
            self.app = xw.App(visible=True)
            self.wb = self.app.books.open(file_path)
        else:
            self.app = xw.apps.active
            self.wb = self.app.books.active

    def search_value(self, metric, time_period):
        """Search for a value in the workbook based on metric and time period."""
        results = []
        for sheet in self.wb.sheets:
            metric_cell = sheet.api.Cells.Find(What=metric, LookAt=xw.constants.LookAt.xlPart)
            time_period_cell = sheet.api.Cells.Find(What=time_period, LookAt=xw.constants.LookAt.xlPart)
            if metric_cell and time_period_cell:
                value_cell = sheet.range((metric_cell.Row, time_period_cell.Column)).value
                results.append({
                    'Sheet': sheet.name,
                    'Metric': metric,
                    'Time Period': time_period,
                    'Address': sheet.range((metric_cell.Row, time_period_cell.Column)).address,
                    'Value': value_cell
                })
        return results

    def identify_business_drivers(self, metric):
        """Identify business drivers for a given metric by analyzing formulas and using LLM to interpret business context."""
        query_type = self.identify_query_type(metric)

        if query_type == "Temporal Metric Search":
            results = self.search_value(metric, "")  # Assuming no time period given
        else:
            results = []
            for sheet in self.wb.sheets:
                metric_cell = sheet.api.Cells.Find(What=metric, LookAt=xw.constants.LookAt.xlPart)
                if metric_cell:
                    results.append(metric_cell)

        if results:
            for result in results:
                formula = self.wb.sheets[result['Sheet']].range(result['Address']).formula
                if formula:
                    referenced_cells = self.wb.sheets[result['Sheet']].range(result['Address']).precedents
                    cell_addresses = [ref_cell.Address for ref_cell in referenced_cells]
                    driver_prompt = f"Identify business drivers for the metric '{metric}' based on the following formula: {formula} and cell references: {cell_addresses}"
                    response = llm.predict(driver_prompt)
                    return response
        return "No drivers found."

    def list_open_workbooks(self):
        """List all currently open Excel workbooks."""
        return [wb.name for wb in xw.books]

    def identify_query_type(self, search_term):
        """Identify if a search term refers to a temporal metric search, atemporal metric search, or generic text search using LLM."""
        query_prompt = {
            "input": search_term,
            "schema": {
                "query_type": {
                    "type": "string",
                    "enum": ["Temporal Metric Search", "Atemporal Metric Search", "Generic Text Search"]
                }
            }
        }
        response = llm.predict_messages([query_prompt])
        return response['query_type']

    def get_structure(self):
        """Return the structure of the workbook."""
        structure_info = []
        for sheet in self.wb.sheets:
            used_range = sheet.used_range
            structure_info.append({
                'Sheet Name': sheet.name,
                'Rows': used_range.rows.count,
                'Columns': used_range.columns.count
            })
        return structure_info

    def query_cell(self, sheet_name, cell):
        """Retrieve the value and formula of a specific cell."""
        sheet = self.wb.sheets[sheet_name]
        value = sheet.range(cell).value
        formula = sheet.range(cell).formula
        return {'Value': value, 'Formula': formula}

    def get_named_ranges(self):
        """Retrieve all named ranges in the workbook."""
        return {name.name: name.refers_to_range.address for name in self.wb.names}

    def capture_screenshot(self, sheet_name, cell_range):
        """Capture a screenshot of the specified sheet or range."""
        sheet = self.wb.sheets[sheet_name]
        sheet.range(cell_range).api.Show()
        screenshot_path = f"{sheet_name}_{cell_range}.png"
        sheet.api.ExportAsFixedFormat(0, screenshot_path)
        return screenshot_path

    def change_value(self, sheet_name, cell, new_value):
        """Change the value of a specific cell."""
        sheet = self.wb.sheets[sheet_name]
        sheet.range(cell).value = new_value
        self.wb.save()
        return f"Value in {sheet_name} at {cell} changed to {new_value}"


    def close_workbook(self):
        """Close the workbook and clean up resources."""
        self.wb.close()


if __name__ == "__main__":
    """agent = ExcelTamerAgent()
    res = agent.identify_business_drivers("Revenue")
    print(res)
    agent.close_workbook()"""

    excel_agent = ExcelTamerAgent("example.xlsx")

    llm = ChatOpenAI(temperature=0, model_name="gpt-4")

    tools = [
        Tool(name="Search Function",
             func=lambda args: excel_agent.search_value(args["metric"], args["time_period"]),
             description="Search for specific values based on metric and time period in the Excel workbook."),
        Tool(name="Identify Query Type",
             func=lambda query: excel_agent.identify_query_type(query),
             description="Identify if a search term refers to a temporal metric search, atemporal metric search, or generic text search."),
        Tool(name="Structure Function",
             func=excel_agent.get_structure,
             description="Get the structure of the workbook including sheets, rows, and columns."),
        Tool(name="Query Function",
             func=lambda args: excel_agent.query_cell(args["sheet_name"], args["cell"]),
             description="Query a cell to get its value and formula."),
        Tool(name="Names Function",
             func=excel_agent.get_named_ranges,
             description="Retrieve all named ranges from the workbook."),
        Tool(name="Screenshot Function",
             func=lambda args: excel_agent.capture_screenshot(args["sheet_name"], args["cell_range"]),
             description="Capture a screenshot of a specific sheet or range."),
        Tool(name="Change Value Function",
             func=lambda args: excel_agent.change_value(args["sheet_name"], args["cell"], args["new_value"]),
             description="Change the value of a specific cell in the workbook."),
        Tool(name="Identify Business Drivers",
             func=lambda metric: excel_agent.identify_business_drivers(metric),
             description="Identify business drivers for a given metric by analyzing formulas and using LLM to interpret business context."),
        Tool(name="List Open Workbooks",
             func=excel_agent.list_open_workbooks,
             description="List all currently open Excel workbooks.")
    ]

    agent = initialize_agent(tools=tools, llm=llm, agent=AgentType.ZERO_SHOT_REACT_DESCRIPTION, verbose=True)

    # Example usage
    # agent.run("Get structure of workbook")
    # agent.run("Query cell A1 in Sheet1")
    # agent.run("Identify business drivers for Total Net Sales")
    res= agent.run("Query cell B15 in Expenses")
    print(res)

    excel_agent.close_workbook()