import xlwings as xw

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

    def find_cells_by_value(self, value: str, sheet_name: str = None, search_whole_workbook: bool = False) -> list[str]:
        """
        Searches for all cells containing the specified value.

        :param value: The value to search for.
        :param sheet_name: The name of the sheet to search in (optional, searches active sheet if not provided).
        :param search_whole_workbook: Boolean flag to search across all sheets or just one.
        :return: A list of cell addresses containing the value.
        """
        found_cells = []
        sheets = self.wb.sheets if search_whole_workbook else [
            self.wb.sheets[sheet_name] if sheet_name else self.wb.sheets.active]
        for sheet in sheets:
            for cell in sheet.used_range:
                if cell.value == value:
                    found_cells.append(f"{sheet.name}!{cell.address}")
        return found_cells

    def find_cells_by_partial_value(self, value: str, sheet_name: str = None, search_whole_workbook: bool = False) -> \
    list[str]:
        """
        Searches for all cells containing the specified partial value.

        :param value: The substring to search for within cells.
        :param sheet_name: The name of the sheet to search in (optional, searches active sheet if not provided).
        :param search_whole_workbook: Boolean flag to search across all sheets or just one.
        :return: A list of cell addresses containing the partial value.
        """
        found_cells = []
        sheets = self.wb.sheets if search_whole_workbook else [
            self.wb.sheets[sheet_name] if sheet_name else self.wb.sheets.active]
        for sheet in sheets:
            for cell in sheet.used_range:
                if isinstance(cell.value, str) and value in cell.value:
                    found_cells.append(f"{sheet.name}!{cell.address}")
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