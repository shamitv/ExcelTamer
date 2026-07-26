"""MCP-only smoke tests that do not require Microsoft Excel."""

import asyncio
import json
import os
import socket
import sys
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

import pandas as pd
from mcp import ClientSession, StdioServerParameters
from mcp.client.sse import sse_client
from mcp.client.stdio import stdio_client

from ExcelTamer.mcp.engine import diff, read, search, workbook, write
from ExcelTamer.mcp.excel import ExcelAutomation
from ExcelTamer.mcp.server import (
    handle_get_prompt,
    handle_list_prompts,
    handle_list_resources,
    handle_list_tools,
    handle_read_resource,
)
from ExcelTamer.mcp.sessions import session


class ProtocolSurfaceTests(unittest.TestCase):
    def test_protocol_surface(self):
        tools = asyncio.run(handle_list_tools())
        resources = asyncio.run(handle_list_resources())
        prompts = asyncio.run(handle_list_prompts())

        self.assertEqual(len(tools), 17)
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


class AttachmentTests(unittest.TestCase):
    class Book:
        def __init__(
            self,
            name,
            path=None,
            *,
            saved=True,
            read_only=False,
        ):
            self.name = name
            self.fullname = path or name
            directory = path.rsplit("\\", 1)[0] if path else ""
            self.api = SimpleNamespace(
                Path=directory,
                Saved=saved,
                ReadOnly=read_only,
            )
            self.sheets = [SimpleNamespace(name="Sheet1")]
            self.close = Mock()

    class Books(list):
        def __init__(self, books, active=None):
            super().__init__(books)
            self.active = active

    class App:
        def __init__(self, pid, books):
            self.pid = pid
            self.books = books

    class Apps(list):
        def __init__(self, apps, active=None):
            super().__init__(apps)
            self.active = active

    def setUp(self):
        session.clear()
        diff.checkpoints.clear()

    def tearDown(self):
        session.clear()
        diff.checkpoints.clear()

    def test_lists_apps_saved_unsaved_and_out_of_root_workbooks(self):
        saved = self.Book(
            "Budget.xlsx",
            r"C:\Users\you\Documents\Excel\Budget.xlsx",
            read_only=True,
        )
        unsaved = self.Book("Book1", saved=False)
        outside_root = self.Book(
            "Private.xlsx",
            r"D:\OutsideAllowedRoots\Private.xlsx",
        )
        first_app = self.App(101, self.Books([saved, unsaved], active=saved))
        second_app = self.App(
            202,
            self.Books([outside_root], active=outside_root),
        )

        class InaccessibleApp:
            pid = 303

            @property
            def books(self):
                raise RuntimeError("Excel instance is inaccessible")

        apps = self.Apps(
            [first_app, second_app, InaccessibleApp()],
            active=first_app,
        )
        registered_id = session.add_workbook(
            ExcelAutomation.attach(first_app, saved)
        )

        with patch.object(workbook.xw, "apps", apps):
            result = workbook.list_open_workbooks()

        self.assertEqual(result["count"], 3)
        by_name = {item["name"]: item for item in result["workbooks"]}
        self.assertEqual(by_name["Budget.xlsx"]["workbook_id"], registered_id)
        self.assertTrue(by_name["Budget.xlsx"]["active"])
        self.assertTrue(by_name["Budget.xlsx"]["read_only"])
        self.assertIsNone(by_name["Book1"]["path"])
        self.assertTrue(by_name["Book1"]["has_unsaved_changes"])
        self.assertEqual(
            by_name["Private.xlsx"]["path"],
            r"D:\OutsideAllowedRoots\Private.xlsx",
        )
        self.assertFalse(by_name["Private.xlsx"]["active"])
        self.assertEqual(result["warnings"][0]["app_pid"], 303)
        self.assertIn("inaccessible", result["warnings"][0]["error"])

    def test_attaches_active_workbook_idempotently(self):
        active_book = self.Book(
            "Active.xlsx",
            r"C:\Work\Active.xlsx",
        )
        app = self.App(404, self.Books([active_book], active=active_book))
        apps = self.Apps([app], active=app)

        with patch.object(workbook.xw, "apps", apps):
            first = workbook.attach_workbook()
            second = workbook.attach_workbook()

        self.assertEqual(first["workbook_id"], second["workbook_id"])
        self.assertEqual(first["app_pid"], 404)
        self.assertEqual(first["sheets"], ["Sheet1"])
        self.assertTrue(first["attached"])
        self.assertFalse(first["already_attached"])
        self.assertTrue(second["already_attached"])
        self.assertEqual(len(session.open_workbooks), 1)

    def test_attach_reports_no_excel_and_no_active_workbook(self):
        with patch.object(workbook.xw, "apps", self.Apps([])):
            with self.assertRaisesRegex(
                ValueError,
                "No running Excel application",
            ):
                workbook.attach_workbook()

        app = self.App(505, self.Books([]))
        with patch.object(workbook.xw, "apps", self.Apps([app], active=app)):
            with self.assertRaisesRegex(ValueError, "no open workbook"):
                workbook.attach_workbook()

    def test_detaches_without_closing_and_rejects_rollback(self):
        active_book = self.Book(
            "Attached.xlsx",
            r"C:\Work\Attached.xlsx",
        )
        app = self.App(606, self.Books([active_book], active=active_book))
        apps = self.Apps([app], active=app)

        with patch.object(workbook.xw, "apps", apps):
            attached = workbook.attach_workbook()
        workbook_id = attached["workbook_id"]

        resource = json.loads(asyncio.run(handle_read_resource(
            "excel://workbooks"
        )))
        self.assertEqual(resource[0]["app_pid"], 606)
        self.assertEqual(resource[0]["path"], r"C:\Work\Attached.xlsx")
        self.assertTrue(resource[0]["attached"])
        self.assertEqual(resource[0]["access_mode"], "rw")

        with patch.object(diff.shutil, "copy2") as copy2:
            with self.assertRaisesRegex(
                ValueError,
                "not supported for attached workbooks",
            ):
                diff.checkpoint_rollback(workbook_id, "before")
            copy2.assert_not_called()
        active_book.close.assert_not_called()
        self.assertIsNotNone(session.get_workbook(workbook_id))

        result = workbook.close_workbook(workbook_id)
        self.assertEqual(result["status"], "detached")
        active_book.close.assert_not_called()
        self.assertIsNone(session.get_workbook(workbook_id))


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

                self.assertEqual(len(tools.tools), 17)
                self.assertEqual(len(resources.resources), 1)
                self.assertEqual(len(prompts.prompts), 2)
                self.assertTrue(safe_edit.messages[0].content.text.strip())


class SseHandshakeTests(unittest.IsolatedAsyncioTestCase):
    async def test_sse_discovery_and_prompt_retrieval(self):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as port_socket:
            port_socket.bind(("127.0.0.1", 0))
            port = port_socket.getsockname()[1]

        env = os.environ.copy()
        env["PYTHONPATH"] = os.pathsep.join(
            path for path in (os.getcwd(), env.get("PYTHONPATH")) if path
        )
        process = await asyncio.create_subprocess_exec(
            sys.executable,
            "-m",
            "ExcelTamer.mcp.main",
            "--port",
            str(port),
            env=env,
            stdout=asyncio.subprocess.DEVNULL,
            stderr=asyncio.subprocess.PIPE,
        )
        stderr_text = ""
        try:
            await self._wait_for_listener(process, port)
            async with asyncio.timeout(10):
                async with sse_client(
                    f"http://127.0.0.1:{port}/sse",
                    timeout=5,
                    sse_read_timeout=5,
                ) as streams:
                    async with ClientSession(*streams) as session:
                        await session.initialize()
                        tools = await session.list_tools()
                        resources = await session.list_resources()
                        prompts = await session.list_prompts()
                        safe_edit = await session.get_prompt("safe-edit")

                        self.assertEqual(len(tools.tools), 17)
                        self.assertEqual(len(resources.resources), 1)
                        self.assertEqual(len(prompts.prompts), 2)
                        self.assertTrue(safe_edit.messages[0].content.text.strip())

            # Give the request handler time to finish after the SSE client closes.
            await asyncio.sleep(0.1)
        finally:
            if process.returncode is None:
                process.terminate()
                try:
                    await asyncio.wait_for(process.wait(), timeout=5)
                except TimeoutError:
                    process.kill()
                    await process.wait()
            if process.stderr is not None:
                stderr_text = (await process.stderr.read()).decode(
                    errors="replace"
                )

        self.assertNotIn("Exception in ASGI application", stderr_text)
        self.assertNotIn(
            "TypeError: 'NoneType' object is not callable",
            stderr_text,
        )

    async def _wait_for_listener(self, process, port):
        loop = asyncio.get_running_loop()
        deadline = loop.time() + 10
        while loop.time() < deadline:
            if process.returncode is not None:
                stderr = ""
                if process.stderr is not None:
                    stderr = (await process.stderr.read()).decode(
                        errors="replace"
                    )
                self.fail(
                    f"SSE server exited before accepting connections:\n{stderr}"
                )
            try:
                _reader, writer = await asyncio.open_connection(
                    "127.0.0.1",
                    port,
                )
            except OSError:
                await asyncio.sleep(0.05)
                continue
            writer.close()
            await writer.wait_closed()
            return
        self.fail(f"SSE server did not listen on port {port} within 10 seconds")


if __name__ == "__main__":
    unittest.main()
