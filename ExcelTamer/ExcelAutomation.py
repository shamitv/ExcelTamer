import pandas as pd
import xlwings as xw

import logging
from datetime import datetime

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(message)s',
    datefmt='%d:%m:%Y %H:%M:%S'
)


class ExcelAutomation:
    def __init__(self, file_path: str = None):
        self.app = xw.apps.active if xw.apps else xw.App(visible=True)

        if file_path:
            self.wb = self.app.books.open(file_path)
        else:
            self.wb = self.app.books.active if self.app.books else self.app.books.add()


    def list_open_workbooks(self) -> list[str]:
        return [wb.fullname for wb in self.app.books]

    def save(self, file_path: str = None) -> None:
        if file_path:
            self.wb.save(file_path)
        else:
            self.wb.save()

    def close(self) -> None:
        self.wb.close()
        self.app.quit()

    def list_sheets(self) -> list[str]:
        return [sheet.tool_name for sheet in self.wb.sheets]

    def add_sheet(self, sheet_name: str) -> None:
        self.wb.sheets.add(sheet_name)

    def remove_sheet(self, sheet_name: str) -> None:
        sheet = self.wb.sheets[sheet_name]
        sheet.delete()

    def read_cell(self, sheet_name: str, cell: str) -> any:
        sheet = self.wb.sheets[sheet_name]
        return sheet.range(cell).value

    def query_cell(self, sheet_name:str, cell:str) ->dict:
        """Retrieve the value and formula of a specific cell."""
        sheet = self.wb.sheets[sheet_name]
        value = sheet.range(cell).value
        formula = sheet.range(cell).formula
        return {'Value': value, 'Formula': formula}

    def get_range_as_markdown(self, sheet_name: str, cell_range: str=None) -> str:
        df = self.get_range_as_dataframe(sheet_name, cell_range)

        return df.to_markdown(index=True)

    def get_range_as_dataframe(self, sheet_name, cell_range=None):
        """
        Returns a pandas DataFrame from the specified sheet and range.
        The DataFrame columns are the Excel column letters (I, J, K, etc.)
        rather than the first row of data.

        Also adds a 'RowNumber' column with the actual Excel row indices.
        """
        logging.debug(f"Getting range as DataFrame for sheet: {sheet_name}, cell_range: {cell_range}")

        sheet = self.wb.sheets[sheet_name]

        # If cell_range is empty or blank, treat it as None
        if not cell_range or cell_range.strip() == "":
            cell_range = None

        # If no specific range is given, default to entire used range.
        if cell_range is None:
            cell_range = sheet.used_range.address
            range = sheet.range(cell_range)
        else:
            range = sheet.range(cell_range)

        df = self.get_dataframe_with_excel_headers_impl(sheet,range)

        return df

    def write_cell(self, sheet_name: str, cell: str, value: any) -> None:
        sheet = self.wb.sheets[sheet_name]
        sheet.range(cell).value = value

    def list_named_ranges(self) -> dict[str, str]:
        return {name.tool_name: name.refers_to_range.address for name in self.wb.names}

    def capture_screenshot_png(self, sheet_name: str, output_path: str, cell_range: str = None) -> bool:
        try:
            #cell_range = None
            sheet = self.wb.sheets[sheet_name]
            if not cell_range:
                cell_range = sheet.used_range.address
            sheet.range(cell_range).api.Show()
            sheet.range(cell_range).to_png(output_path)
            return True
        except Exception as e:
            print(f"Failed to capture screenshot: {e}")
            return False

    def get_dataframe_with_excel_headers_impl(self, sheet: xw.Sheet, cell_range:xw.Range):
        """
        Returns a DataFrame from the given xlwings sheet and cell_range.
        The columns of the DataFrame are the actual Excel column letters
        (e.g. I, J, K, ... AH). A 'RowNumber' column is added to reflect
        actual Excel row indices.

        :param sheet: An xlwings Sheet object.
        :param cell_range: A string like 'I3:AH10', or any valid Excel range.
        :return: pandas DataFrame
        """
        # TO-DO: Remove this var and use cell_range directly
        rng:xw.Range = cell_range  # e.g. "I3:AH10"

        # Read the raw 2D list of values
        data_2d = rng.value
        if not data_2d:
            return pd.DataFrame()  # Empty range => empty DataFrame

        # Dimensions: number of rows/columns in the 2D data
        row_count = rng.rows.count
        col_count = rng.columns.count

        # Excel's top-left row/col for the specified range
        start_row = rng.row
        start_col = rng.column

        # 1) Build column labels from the top row of each column's address
        #    For column c in [0..col_count-1], we parse e.g. "$I$3" -> "I".
        columns_letters = []
        for c_offset in range(col_count):
            col_address = sheet.range((start_row, start_col + c_offset)).address
            # col_address might look like "$I$3" => extract just the letters
            letters = ''.join(ch for ch in col_address if ch.isalpha())
            columns_letters.append(letters)

        # 2) Build the 'RowNumber' list from start_row -> (start_row + row_count - 1)
        row_numbers = list(range(start_row, start_row + row_count))

        # 3) Create the DataFrame
        df = pd.DataFrame(data_2d, columns=columns_letters)

        # 4) Insert 'RowNumber' at the beginning
        df.insert(0, "RowNumber", row_numbers)

        return df

    def find_all_cells_by_value(self, value: str, sheet_name: str = None, search_whole_workbook: bool = False):
        logging.debug(
            f"Searching for cells with value '{value}' in sheet '{sheet_name}' (search whole workbook: {search_whole_workbook})")
        # If search_whole_workbook is True, search all sheets
        if search_whole_workbook:
            found_cells = []
            for sheet in self.wb.sheets:
                found_cells += self.find_all_cells_in_sheet(sheet, value)
            return found_cells

        # If sheet_name is provided, search within that sheet
        if sheet_name:
            sheet = self.wb.sheets[sheet_name]
        else:
            # Use the active sheet if no sheet_name is provided
            sheet = self.wb.sheets.active

        return self.find_all_cells_in_sheet(sheet, value)

    def find_all_cells_in_sheet(self, sheet, value: str) -> list[tuple[str, str, int]]:
        logging.debug(f"Searching for value '{value}' in sheet '{sheet.name}'")

        search_range = sheet.used_range

        # Get DataFrame from the specified range
        df = self.get_dataframe_with_excel_headers_impl(sheet, search_range)

        # Find all cells with the specified value
        found_cells = df[df.isin([value])].stack().index.tolist()

        # Convert the DataFrame index to Excel row and column indices
        found_cells = [(sheet.name, df.columns.get_loc(col) + 1, int(df.at[row, 'RowNumber'])) for row, col in
                       found_cells]

        # Convert column indices to Excel column letters
        found_cells = [(sheet_name, df.columns[col - 1], row) for sheet_name, col, row in found_cells]

        logging.debug(f"Found {len(found_cells)} cells with value '{value}' in sheet '{sheet.name}'")
        return found_cells


    def get_structure(self):
        """Return the structure of the workbook."""
        structure_info = []
        for sheet in self.wb.sheets:
            used_range = sheet.used_range
            # Access all named ranges in the sheet
            named_ranges = sheet.names
            named_range_info = []
            for name in named_ranges:
                named_range_info.append({
                    'Name': name.name,
                    'Refers To': name.refers_to_range.address
                })
            structure_info.append({
                'Sheet Name': sheet.name,
                'Rows': used_range.rows.count,
                'Columns': used_range.columns.count,
                'Range': used_range.address,
                'Named Ranges': named_range_info
            })
        return structure_info