"""Internal xlwings backend used by the ExcelTamer MCP engines."""

import logging
from typing import Any

import pandas as pd
import xlwings as xw

logger = logging.getLogger(__name__)


class ExcelAutomation:
    """Manage an Excel application/workbook pair for one MCP session entry."""

    def __init__(
        self,
        file_path: str | None = None,
        *,
        app: Any | None = None,
        workbook: Any | None = None,
        attached: bool = False,
        access_mode: str = "rw",
    ):
        if workbook is not None:
            self.app = app if app is not None else workbook.app
            self.wb = workbook
        else:
            self.app = xw.apps.active if xw.apps else xw.App(visible=True)
            self.wb = (
                self.app.books.open(file_path)
                if file_path
                else self.app.books.active
                if self.app.books
                else self.app.books.add()
            )
        self.attached = attached
        self.access_mode = access_mode

    @classmethod
    def attach(cls, app: Any, workbook: Any) -> "ExcelAutomation":
        """Create a non-owning handle for a workbook already open in Excel."""
        return cls(
            app=app,
            workbook=workbook,
            attached=True,
            access_mode="rw",
        )

    def save(self, file_path: str | None = None) -> None:
        if file_path:
            self.wb.save(file_path)
        else:
            self.wb.save()

    def close(self, quit_app: bool = True) -> None:
        self.wb.close()
        if quit_app:
            self.app.quit()

    def list_sheets(self) -> list[str]:
        return [sheet.name for sheet in self.wb.sheets]

    def capture_screenshot_png(
        self,
        sheet_name: str,
        output_path: str,
        cell_range: str | None = None,
    ) -> bool:
        """Render a worksheet range to a PNG file."""
        try:
            sheet = self.wb.sheets[sheet_name]
            target = (
                sheet.range(cell_range)
                if cell_range and cell_range.strip()
                else sheet.used_range
            )
            target.api.Show()
            target.to_png(output_path)
            return True
        except Exception:
            logger.exception(
                "Failed to capture screenshot for sheet=%s range=%s",
                sheet_name,
                cell_range,
            )
            return False

    def query_cell(self, sheet_name: str, cell: str) -> dict[str, Any]:
        """Return the value, formula, and rendered text for one cell."""
        target = self.wb.sheets[sheet_name].range(cell)
        return {
            "Value": target.value,
            "Formula": target.formula,
            "VisibleText": target.api.Text,
        }

    def get_range_as_dataframe(
        self, sheet_name: str, cell_range: str | None = None
    ) -> pd.DataFrame:
        """Read a range with Excel column letters and source row numbers."""
        logger.debug(
            "Getting range as DataFrame for sheet=%s range=%s",
            sheet_name,
            cell_range,
        )
        sheet = self.wb.sheets[sheet_name]
        if not cell_range or not cell_range.strip():
            cell_range = sheet.used_range.address
        return self._range_to_dataframe(sheet, sheet.range(cell_range))

    def _range_to_dataframe(
        self, sheet: xw.Sheet, cell_range: xw.Range
    ) -> pd.DataFrame:
        data = cell_range.value
        if data is None:
            return pd.DataFrame()

        row_count = cell_range.rows.count
        col_count = cell_range.columns.count
        if not isinstance(data, list):
            data = [[data]]
        elif row_count == 1 and (not data or not isinstance(data[0], list)):
            data = [data]
        elif col_count == 1 and data and not isinstance(data[0], list):
            data = [[value] for value in data]

        start_row = cell_range.row
        start_col = cell_range.column
        columns = []
        for offset in range(col_count):
            address = sheet.range((start_row, start_col + offset)).address
            columns.append("".join(char for char in address if char.isalpha()))

        frame = pd.DataFrame(data, columns=columns)
        frame.insert(0, "RowNumber", range(start_row, start_row + row_count))
        return frame

    def write_cell(self, sheet_name: str, cell: str, value: Any) -> None:
        self.wb.sheets[sheet_name].range(cell).value = value

    def get_structure(self) -> list[dict[str, Any]]:
        """Return sheet dimensions, used ranges, and named ranges."""
        structure = []
        for sheet in self.wb.sheets:
            used_range = sheet.used_range
            structure.append(
                {
                    "Sheet Name": sheet.name,
                    "Rows": used_range.rows.count,
                    "Columns": used_range.columns.count,
                    "Range": used_range.address,
                    "Named Ranges": [
                        {
                            "Name": name.name,
                            "Refers To": name.refers_to_range.address,
                        }
                        for name in sheet.names
                    ],
                }
            )
        return structure
