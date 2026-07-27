"""Screenshot capture support for the ExcelTamer MCP server."""

from __future__ import annotations

import base64
import tempfile
from pathlib import Path
from typing import Any

from ..sessions import session

PNG_MIME_TYPE = "image/png"


def _error_response(error: str) -> dict[str, Any]:
    return {
        "status": "error",
        "image": False,
        "file": False,
        "image_data": None,
        "file_path": None,
        "image_mime_type": None,
        "error": error,
    }


def capture_range_image(
    workbook_id: str,
    sheet: str,
    range_a1: str | None = None,
    return_image: bool = True,
) -> dict[str, Any]:
    """Capture a sheet range as base64 PNG data or a temporary PNG file."""
    automation = session.get_workbook(workbook_id)
    if not automation:
        return _error_response(f"Workbook {workbook_id} not found")

    try:
        if sheet not in automation.list_sheets():
            return _error_response(f"Sheet {sheet!r} not found")
    except Exception as exc:
        return _error_response(f"Unable to inspect workbook sheets: {exc}")

    normalized_range = range_a1.strip() if range_a1 and range_a1.strip() else None
    temporary_path: Path | None = None
    keep_file = False

    try:
        with tempfile.NamedTemporaryFile(
            prefix="exceltamer_",
            suffix=".png",
            delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name).resolve()

        captured = automation.capture_screenshot_png(
            sheet,
            str(temporary_path),
            normalized_range,
        )
        if not captured:
            raise RuntimeError("Excel could not capture the requested sheet or range")

        if not temporary_path.is_file() or temporary_path.stat().st_size == 0:
            raise RuntimeError("Excel produced an empty screenshot")

        if return_image:
            image_data = base64.b64encode(temporary_path.read_bytes()).decode("ascii")
            return {
                "status": "success",
                "image": True,
                "file": False,
                "image_data": image_data,
                "file_path": None,
                "image_mime_type": PNG_MIME_TYPE,
                "error": None,
            }

        keep_file = True
        return {
            "status": "success",
            "image": False,
            "file": True,
            "image_data": None,
            "file_path": str(temporary_path),
            "image_mime_type": PNG_MIME_TYPE,
            "error": None,
        }
    except Exception as exc:
        return _error_response(str(exc))
    finally:
        if temporary_path is not None and not keep_file:
            try:
                temporary_path.unlink(missing_ok=True)
            except OSError:
                pass
