"""MCP-only smoke tests that do not require Microsoft Excel."""

import asyncio
import os
import sys
import unittest
from types import SimpleNamespace
from unittest.mock import patch

import pandas as pd
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

from ExcelTamer.mcp.engine import read, search, write
from ExcelTamer.mcp.excel import ExcelAutomation
from ExcelTamer.mcp.server import (
    handle_get_prompt,
    handle_list_prompts,
    handle_list_resources,
    handle_list_tools,
)
from ExcelTamer.mcp.sessions import session


class ProtocolSurfaceTests(unittest.TestCase):
    def test_protocol_surface_is_unchanged(self):
        tools = asyncio.run(handle_list_tools())
        resources = asyncio.run(handle_list_resources())
        prompts = asyncio.run(handle_list_prompts())

        self.assertEqual(len(tools), 15)
        self.assertEqual(len(resources), 1)
        self.assertEqual(
            {prompt.name for prompt in prompts},
            {"safe-edit", "financial-extract"},
        )

    def test_packaged_prompt_bodies_are_available(self):
        for name in ("safe-edit", "financial-extract"):
            prompt = asyncio.run(handle_get_prompt(name, None))
            self.assertEqual(len(prompt.messages), 1)
            self.assertTrue(prompt.messages[0].content.text.strip())

    def test_single_cell_range_normalization(self):
        backend = object.__new__(ExcelAutomation)
        cell_range = SimpleNamespace(
            value="value",
            rows=SimpleNamespace(count=1),
            columns=SimpleNamespace(count=1),
            row=3,
            column=2,
        )

        class Sheet:
            @staticmethod
            def range(position):
                row, column = position
                return SimpleNamespace(address=f"${chr(64 + column)}${row}")

        frame = backend._range_to_dataframe(Sheet(), cell_range)
        self.assertEqual(
            frame.to_dict(orient="records"),
            [{"RowNumber": 3, "B": "value"}],
        )


class EngineTests(unittest.TestCase):
    class TargetRange:
        value = None

    class Sheet:
        def __init__(self):
            self.target = EngineTests.TargetRange()

        def range(self, _address):
            return self.target

    class Backend:
        def __init__(self):
            self.frame = pd.DataFrame(
                {
                    "RowNumber": [4, 5],
                    "A": ["Revenue", "Costs"],
                    "B": [100, 40],
                }
            )
            self.writes = []
            self.sheet = EngineTests.Sheet()
            self.wb = SimpleNamespace(sheets={"Sheet1": self.sheet})

        @staticmethod
        def list_sheets():
            return ["Sheet1"]

        def get_range_as_dataframe(self, _sheet, _cell_range=None):
            return self.frame.copy()

        def write_cell(self, sheet, cell, value):
            self.writes.append((sheet, cell, value))

    def setUp(self):
        self.workbook_id = "test-workbook"
        self.backend = self.Backend()
        session.open_workbooks[self.workbook_id] = self.backend

    def tearDown(self):
        session.remove_workbook(self.workbook_id)

    def test_read_range_returns_source_values_without_helper_column(self):
        result = read.read_range(
            self.workbook_id,
            "Sheet1",
            range_a1="A4:B5",
        )
        self.assertEqual(result["headers"], ["A", "B"])
        self.assertEqual(result["values"], [["Revenue", 100], ["Costs", 40]])
        self.assertEqual(result["shape"], [2, 2])

    def test_search_reconstructs_excel_addresses(self):
        result = search.search(
            self.workbook_id,
            "Revenue",
            match_mode="exact",
            scope="values",
        )
        self.assertEqual(
            result["hits"],
            [
                {
                    "sheet": "Sheet1",
                    "cell": "A4",
                    "value": "Revenue",
                    "match_type": "value",
                }
            ],
        )

    @patch("ExcelTamer.mcp.engine.write.log_write")
    def test_write_paths_update_cells_and_ranges(self, _log_write):
        write.change_cell_value(self.workbook_id, "Sheet1", "B4", 120)
        batch = write.batch_update_cells(
            self.workbook_id,
            [
                {"sheet": "Sheet1", "cell": "B5", "value": 45},
                {"sheet": "Sheet1", "cell": "C5", "formula": "=B5*2"},
            ],
        )
        block = write.write_range(
            self.workbook_id,
            "Sheet1",
            "D1",
            [[1, 2], [3, 4]],
        )

        self.assertEqual(
            self.backend.writes,
            [
                ("Sheet1", "B4", 120),
                ("Sheet1", "B5", 45),
                ("Sheet1", "C5", "=B5*2"),
            ],
        )
        self.assertEqual(batch["updated_count"], 2)
        self.assertEqual(block["written_cells"], 4)
        self.assertEqual(self.backend.sheet.target.value, [[1, 2], [3, 4]])


class StdioHandshakeTests(unittest.IsolatedAsyncioTestCase):
    async def test_stdio_discovery_and_prompt_retrieval(self):
        env = os.environ.copy()
        env["PYTHONPATH"] = os.pathsep.join(
            path for path in (os.getcwd(), env.get("PYTHONPATH")) if path
        )
        params = StdioServerParameters(
            command=sys.executable,
            args=["-m", "ExcelTamer.mcp.main"],
            env=env,
        )

        async with stdio_client(params) as streams:
            async with ClientSession(*streams) as session:
                await session.initialize()
                tools = await session.list_tools()
                resources = await session.list_resources()
                prompts = await session.list_prompts()
                safe_edit = await session.get_prompt("safe-edit")

                self.assertEqual(len(tools.tools), 15)
                self.assertEqual(len(resources.resources), 1)
                self.assertEqual(len(prompts.prompts), 2)
                self.assertTrue(safe_edit.messages[0].content.text.strip())


if __name__ == "__main__":
    unittest.main()
