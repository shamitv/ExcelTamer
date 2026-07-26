
from typing import Any

import xlwings as xw

from ..config import DEFAULT_MODE
from ..excel import ExcelAutomation
from ..safety import validate_path
from ..sessions import session


def _workbook_path(workbook: Any) -> str | None:
    """Return a saved workbook path, or None for an unsaved workbook."""
    try:
        if not workbook.api.Path:
            return None
        return str(workbook.fullname)
    except Exception:
        return None


def _api_boolean(workbook: Any, property_name: str) -> bool | None:
    try:
        return bool(getattr(workbook.api, property_name))
    except Exception:
        return None


def workbook_metadata(automation: ExcelAutomation) -> dict:
    """Return live metadata for a registered workbook."""
    saved = _api_boolean(automation.wb, "Saved")
    return {
        "name": automation.wb.name,
        "path": _workbook_path(automation.wb),
        "app_pid": int(automation.app.pid),
        "attached": bool(getattr(automation, "attached", False)),
        "access_mode": getattr(automation, "access_mode", DEFAULT_MODE),
        "read_only": _api_boolean(automation.wb, "ReadOnly"),
        "has_unsaved_changes": None if saved is None else not saved,
    }


def list_open_workbooks() -> dict:
    """Enumerate workbooks visible to xlwings in the current desktop session."""
    workbooks = []
    warnings = []

    try:
        apps = list(xw.apps)
    except Exception as exc:
        return {
            "count": 0,
            "workbooks": [],
            "warnings": [{"app_pid": None, "error": str(exc)}],
        }

    try:
        active_app = xw.apps.active
        active_pid = int(active_app.pid) if active_app is not None else None
    except Exception:
        active_pid = None

    for app in apps:
        try:
            app_pid = int(app.pid)
        except Exception as exc:
            warnings.append({"app_pid": None, "error": str(exc)})
            continue

        try:
            books = list(app.books)
            active_book = app.books.active if app.books else None
            active_name = active_book.name if active_book is not None else None
        except Exception as exc:
            warnings.append({"app_pid": app_pid, "error": str(exc)})
            continue

        for book in books:
            try:
                name = book.name
                saved = _api_boolean(book, "Saved")
                workbooks.append(
                    {
                        "app_pid": app_pid,
                        "name": name,
                        "path": _workbook_path(book),
                        "active": (
                            app_pid == active_pid and name == active_name
                        ),
                        "read_only": _api_boolean(book, "ReadOnly"),
                        "has_unsaved_changes": (
                            None if saved is None else not saved
                        ),
                        "workbook_id": session.find_workbook_id(
                            app_pid,
                            name,
                        ),
                    }
                )
            except Exception as exc:
                try:
                    workbook_name = book.name
                except Exception:
                    workbook_name = None
                warnings.append(
                    {
                        "app_pid": app_pid,
                        "workbook": workbook_name,
                        "error": str(exc),
                    }
                )

    return {
        "count": len(workbooks),
        "workbooks": workbooks,
        "warnings": warnings,
    }


def attach_workbook() -> dict:
    """Attach the active workbook without opening or taking ownership of it."""
    try:
        apps = list(xw.apps)
    except Exception as exc:
        raise ValueError("No running Excel application found") from exc
    if not apps:
        raise ValueError("No running Excel application found")

    try:
        app = xw.apps.active
    except Exception as exc:
        raise ValueError("No active Excel application found") from exc
    if app is None:
        raise ValueError("No active Excel application found")

    try:
        if not app.books:
            raise ValueError("The active Excel application has no open workbook")
        workbook = app.books.active
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError(
            "The active Excel application has no active workbook"
        ) from exc
    if workbook is None:
        raise ValueError("The active Excel application has no active workbook")

    app_pid = int(app.pid)
    name = workbook.name
    existing_id = session.find_workbook_id(app_pid, name)
    if existing_id:
        automation = session.get_workbook(existing_id)
        return {
            "workbook_id": existing_id,
            "app_pid": app_pid,
            "name": name,
            "path": _workbook_path(workbook),
            "sheets": automation.list_sheets(),
            "attached": bool(getattr(automation, "attached", False)),
            "already_attached": True,
        }

    automation = ExcelAutomation.attach(app, workbook)
    workbook_id = session.add_workbook(automation)
    return {
        "workbook_id": workbook_id,
        "app_pid": app_pid,
        "name": name,
        "path": _workbook_path(workbook),
        "sheets": automation.list_sheets(),
        "attached": True,
        "already_attached": False,
    }


def open_workbook(path: str, mode: str = DEFAULT_MODE) -> dict:
    """
    Opens a workbook and returns its ID and metadata.
    mode: 'ro' (read-only) or 'rw' (read-write)
    """
    # 1. Validate path
    abs_path = validate_path(path)
    
    # 2. Open workbook
    # Note: Xlwings doesn't strictly support 'ro' at open() level easily without API flags,
    # but we will enforce it at the Write tool level.
    # We pass the validated absolute path string.
    automation = ExcelAutomation(
        file_path=str(abs_path),
        access_mode=mode,
    )
    
    # 3. Store in session
    wb_id = session.add_workbook(automation)
    
    # 4. Gather metadata
    sheets = automation.list_sheets()
    
    return {
        "workbook_id": wb_id,
        "filename": abs_path.name,
        "sheets": sheets,
        "mode": mode
    }

def close_workbook(workbook_id: str) -> dict:
    automation = session.get_workbook(workbook_id)
    if not automation:
        raise ValueError(f"Workbook {workbook_id} not found")
        
    if getattr(automation, "attached", False):
        session.remove_workbook(workbook_id)
        return {"status": "detached", "workbook_id": workbook_id}

    # Close without quitting app to support multiple workbooks
    automation.close(quit_app=False)
    session.remove_workbook(workbook_id)
    
    return {"status": "closed", "workbook_id": workbook_id}

def save_workbook(workbook_id: str) -> dict:
    automation = session.get_workbook(workbook_id)
    if not automation:
        raise ValueError(f"Workbook {workbook_id} not found")
        
    automation.save()
    return {"status": "saved", "workbook_id": workbook_id}

def save_as_workbook(workbook_id: str, output_path: str) -> dict:
    # 1. Validate output path
    abs_path = validate_path(output_path, allow_write=True)
    
    automation = session.get_workbook(workbook_id)
    if not automation:
        raise ValueError(f"Workbook {workbook_id} not found")
        
    automation.save(str(abs_path))
    
    return {"status": "saved_as", "path": str(abs_path), "workbook_id": workbook_id}
