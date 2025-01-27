import xlwings as xw

class ExcelAutomation:
    def __init__(self, file_path: str = None):
        self.app = xw.App(visible=False)
        self.wb = self.app.books.open(file_path) if file_path else self.app.books.add()

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

    def capture_screenshot(self, sheet_name: str, output_path: str, cell_range: str = None) -> bool:
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