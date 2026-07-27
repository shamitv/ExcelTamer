"""Tests for MCP worksheet and range screenshot capture."""

import asyncio
import base64
import json
import os
import sys
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

import mcp.types as types
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from mcp.types import ImageContent, TextContent

from ExcelTamer.mcp.engine import image
from ExcelTamer.mcp.excel import ExcelAutomation
from ExcelTamer.mcp.server import handle_call_tool, handle_list_tools
from ExcelTamer.mcp.sessions import session


PNG_BYTES = b"\x89PNG\r\n\x1a\nExcelTamer screenshot test"


class ScreenshotBackend:
    def __init__(self, capture_result=True, write_bytes=PNG_BYTES):
        self.capture_result = capture_result
        self.write_bytes = write_bytes
        self.calls = []

    @staticmethod
    def list_sheets():
        return ["Dashboard"]

    def capture_screenshot_png(self, sheet, output_path, cell_range=None):
        self.calls.append((sheet, output_path, cell_range))
        if self.write_bytes is not None:
            Path(output_path).write_bytes(self.write_bytes)
        return self.capture_result


class ExcelScreenshotBackendTests(unittest.TestCase):
    def test_used_range_is_rendered_when_range_is_omitted(self):
        target = SimpleNamespace(
            api=SimpleNamespace(Show=Mock()),
            to_png=Mock(),
        )
        sheet = SimpleNamespace(
            used_range=target,
            range=Mock(),
        )
        backend = object.__new__(ExcelAutomation)
        backend.wb = SimpleNamespace(sheets={"Dashboard": sheet})

        captured = backend.capture_screenshot_png(
            "Dashboard",
            r"C:\Temp\dashboard.png",
        )

        self.assertTrue(captured)
        sheet.range.assert_not_called()
        target.api.Show.assert_called_once_with()
        target.to_png.assert_called_once_with(r"C:\Temp\dashboard.png")

    def test_explicit_range_is_rendered(self):
        target = SimpleNamespace(
            api=SimpleNamespace(Show=Mock()),
            to_png=Mock(),
        )
        sheet = SimpleNamespace(
            used_range=object(),
            range=Mock(return_value=target),
        )
        backend = object.__new__(ExcelAutomation)
        backend.wb = SimpleNamespace(sheets={"Dashboard": sheet})

        captured = backend.capture_screenshot_png(
            "Dashboard",
            r"C:\Temp\dashboard.png",
            "A1:H20",
        )

        self.assertTrue(captured)
        sheet.range.assert_called_once_with("A1:H20")
        target.api.Show.assert_called_once_with()
        target.to_png.assert_called_once_with(r"C:\Temp\dashboard.png")


class ScreenshotEngineTests(unittest.TestCase):
    def setUp(self):
        session.clear()
        self.workbook_id = "screenshot-workbook"
        self.backend = ScreenshotBackend()
        session.open_workbooks[self.workbook_id] = self.backend
        self.retained_paths = []

    def tearDown(self):
        session.clear()
        for path in self.retained_paths:
            path.unlink(missing_ok=True)

    def test_image_mode_returns_base64_and_removes_temporary_file(self):
        result = image.capture_range_image(
            self.workbook_id,
            "Dashboard",
            range_a1="A1:H20",
        )

        self.assertEqual(result["status"], "success")
        self.assertTrue(result["image"])
        self.assertFalse(result["file"])
        self.assertEqual(base64.b64decode(result["image_data"]), PNG_BYTES)
        self.assertIsNone(result["file_path"])
        self.assertEqual(result["image_mime_type"], "image/png")
        self.assertIsNone(result["error"])

        self.assertEqual(self.backend.calls[0][0], "Dashboard")
        self.assertEqual(self.backend.calls[0][2], "A1:H20")
        self.assertFalse(Path(self.backend.calls[0][1]).exists())

    def test_file_mode_uses_used_range_and_retains_temporary_file(self):
        result = image.capture_range_image(
            self.workbook_id,
            "Dashboard",
            range_a1="   ",
            return_image=False,
        )

        self.assertEqual(result["status"], "success")
        self.assertFalse(result["image"])
        self.assertTrue(result["file"])
        self.assertIsNone(result["image_data"])
        self.assertEqual(result["image_mime_type"], "image/png")
        self.assertIsNone(result["error"])
        self.assertIsNone(self.backend.calls[0][2])

        output_path = Path(result["file_path"])
        self.retained_paths.append(output_path)
        self.assertTrue(output_path.is_absolute())
        self.assertEqual(output_path.read_bytes(), PNG_BYTES)

    def test_unknown_workbook_and_sheet_return_structured_errors(self):
        unknown_workbook = image.capture_range_image(
            "missing-workbook",
            "Dashboard",
        )
        missing_sheet = image.capture_range_image(
            self.workbook_id,
            "Missing",
        )

        for result in (unknown_workbook, missing_sheet):
            self.assertEqual(result["status"], "error")
            self.assertFalse(result["image"])
            self.assertFalse(result["file"])
            self.assertIsNone(result["image_data"])
            self.assertIsNone(result["file_path"])
            self.assertIsNone(result["image_mime_type"])
            self.assertTrue(result["error"])

        self.assertEqual(self.backend.calls, [])

    def test_capture_failure_and_empty_output_remove_partial_files(self):
        for capture_result, write_bytes in (
            (False, b"partial"),
            (True, None),
        ):
            with self.subTest(
                capture_result=capture_result,
                write_bytes=write_bytes,
            ):
                backend = ScreenshotBackend(
                    capture_result=capture_result,
                    write_bytes=write_bytes,
                )
                session.open_workbooks[self.workbook_id] = backend

                result = image.capture_range_image(
                    self.workbook_id,
                    "Dashboard",
                )

                self.assertEqual(result["status"], "error")
                self.assertFalse(result["image"])
                self.assertFalse(result["file"])
                self.assertTrue(result["error"])
                self.assertFalse(Path(backend.calls[0][1]).exists())

    def test_temporary_file_creation_failure_is_structured(self):
        with patch.object(
            image.tempfile,
            "NamedTemporaryFile",
            side_effect=OSError("temporary directory unavailable"),
        ):
            result = image.capture_range_image(
                self.workbook_id,
                "Dashboard",
            )

        self.assertEqual(result["status"], "error")
        self.assertFalse(result["image"])
        self.assertFalse(result["file"])
        self.assertIn("temporary directory unavailable", result["error"])
        self.assertEqual(self.backend.calls, [])


class ScreenshotMcpSurfaceTests(unittest.TestCase):
    def setUp(self):
        session.clear()
        self.workbook_id = "mcp-screenshot-workbook"
        self.backend = ScreenshotBackend()
        session.open_workbooks[self.workbook_id] = self.backend
        self.retained_paths = []

    def tearDown(self):
        session.clear()
        for path in self.retained_paths:
            path.unlink(missing_ok=True)

    def test_tool_discovery_exposes_input_and_output_schemas(self):
        tools = asyncio.run(handle_list_tools())
        screenshot_tool = next(
            tool for tool in tools if tool.name == "excel.capture_range_image"
        )

        self.assertEqual(
            screenshot_tool.inputSchema["required"],
            ["workbook_id", "sheet"],
        )
        self.assertTrue(
            screenshot_tool.inputSchema["properties"]["return_image"]["default"]
        )
        self.assertEqual(
            set(screenshot_tool.outputSchema["required"]),
            {
                "status",
                "image",
                "file",
                "image_data",
                "file_path",
                "image_mime_type",
                "error",
            },
        )

    def test_image_mode_returns_native_content_and_structured_data(self):
        content, structured = asyncio.run(
            handle_call_tool(
                "excel.capture_range_image",
                {
                    "workbook_id": self.workbook_id,
                    "sheet": "Dashboard",
                    "range_a1": "B2:D8",
                },
            )
        )

        self.assertEqual(len(content), 1)
        self.assertIsInstance(content[0], ImageContent)
        self.assertEqual(content[0].data, structured["image_data"])
        self.assertEqual(content[0].mimeType, "image/png")
        self.assertTrue(structured["image"])
        self.assertFalse(structured["file"])

    def test_file_mode_returns_text_content_and_structured_data(self):
        content, structured = asyncio.run(
            handle_call_tool(
                "excel.capture_range_image",
                {
                    "workbook_id": self.workbook_id,
                    "sheet": "Dashboard",
                    "return_image": False,
                },
            )
        )

        self.assertEqual(len(content), 1)
        self.assertIsInstance(content[0], TextContent)
        self.assertEqual(json.loads(content[0].text), structured)
        self.assertFalse(structured["image"])
        self.assertTrue(structured["file"])

        output_path = Path(structured["file_path"])
        self.retained_paths.append(output_path)
        self.assertTrue(output_path.exists())

    def test_runtime_error_is_an_mcp_error_with_structured_content(self):
        result = asyncio.run(
            handle_call_tool(
                "excel.capture_range_image",
                {
                    "workbook_id": "missing-workbook",
                    "sheet": "Dashboard",
                },
            )
        )

        self.assertIsInstance(result, types.CallToolResult)
        self.assertTrue(result.isError)
        self.assertEqual(result.structuredContent["status"], "error")
        self.assertFalse(result.structuredContent["image"])
        self.assertFalse(result.structuredContent["file"])
        self.assertIsInstance(result.content[0], TextContent)


class ScreenshotProtocolTests(unittest.IsolatedAsyncioTestCase):
    async def test_malformed_input_uses_standard_mcp_validation_error(self):
        environment = os.environ.copy()
        environment["PYTHONPATH"] = os.pathsep.join(
            path
            for path in (os.getcwd(), environment.get("PYTHONPATH"))
            if path
        )
        parameters = StdioServerParameters(
            command=sys.executable,
            args=["-m", "ExcelTamer.mcp.main"],
            env=environment,
        )

        async with stdio_client(parameters) as streams:
            async with ClientSession(*streams) as client:
                await client.initialize()
                result = await client.call_tool(
                    "excel.capture_range_image",
                    {"sheet": "Dashboard"},
                )

        self.assertTrue(result.isError)
        self.assertIsNone(result.structuredContent)
        self.assertIn("Input validation error", result.content[0].text)


if __name__ == "__main__":
    unittest.main()
