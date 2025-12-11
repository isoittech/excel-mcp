"""Thin Python wrappers for the Java Excel tools.

Each function here corresponds to a single Java main class under
``java/src/jp/isoittech`` and is responsible for:

* building the correct Java command line
* executing it via :mod:`subprocess`
* translating results to Python data structures where appropriate

These functions are intended to be called from an MCP server
implementation, but they can also be imported and used directly.
"""

from __future__ import annotations

import json
import os
import subprocess
from pathlib import Path
from typing import Any, Iterable, List, Optional

# Base paths relative to repository root (this file lives in py/src/excel).
# ``wrapper.py`` is located at ``<repo>/py/src/excel/wrapper.py`` so the
# repository root is the 4th parent of this file (index 3).
_REPO_ROOT = Path(__file__).resolve().parents[3]
_JAVA_DIR = _REPO_ROOT / "java"
_JAVA_DIST = _JAVA_DIR / "dist"
_JAVA_JARS = _JAVA_DIR / "jars"


def _java_classpath() -> str:
    """Return the classpath used to invoke the Java tools.

    It combines the compiled classes under ``java/dist`` and all JARs under
    ``java/jars``. A colon (``:``) is used as the separator which works on
    Unix-like environments (the typical target for MCP servers).
    """

    jars_pattern = str(_JAVA_JARS / "*")
    return f"{jars_pattern}:{_JAVA_DIST}"


def _run_java(class_name: str, args: Iterable[str]) -> subprocess.CompletedProcess:
    """Run a Java tool and return the completed process.

    Parameters
    ----------
    class_name:
        Fully-qualified Java class name (for example
        ``"jp.isoittech.ReadExcelTool"``).
    args:
        Iterable of additional command-line arguments to pass to ``main``.
    """

    cmd = [
        "java",
        "-cp",
        _java_classpath(),
        class_name,
        *list(args),
    ]

    return subprocess.run(cmd, check=False, capture_output=True, text=True)


# ---------------------------------------------------------------------------
# High-level wrappers
# ---------------------------------------------------------------------------


def create_excel(file_path: str, sheet_name: str = "Sheet1") -> None:
    """Create a new Excel file with a single worksheet.

    This is a thin wrapper around ``jp.isoittech.CreateExcelTool``.
    """

    result = _run_java("jp.isoittech.CreateExcelTool", [file_path, sheet_name])
    if result.returncode != 0:
        raise RuntimeError(result.stderr or f"CreateExcelTool failed: {result.returncode}")


def read_excel(file_path: str, sheet_name: str, range_str: str) -> List[List[Any]]:
    """Read a rectangular range from an Excel sheet.

    The Java tool prints a JSON matrix to stdout which is parsed and
    returned as a list of lists.
    """

    result = _run_java("jp.isoittech.ReadExcelTool", [file_path, sheet_name, range_str])
    if result.returncode != 0:
        raise RuntimeError(result.stderr or f"ReadExcelTool failed: {result.returncode}")
    return json.loads(result.stdout.strip() or "[]")


def write_excel(file_path: str, sheet_name: str, data: List[List[Any]]) -> None:
    """Write a matrix of values into an Excel sheet starting at A1.

    ``data`` must be JSON-serializable.
    """

    json_data = json.dumps(data, ensure_ascii=False)
    result = _run_java("jp.isoittech.WriteExcelTool", [file_path, sheet_name, json_data])
    if result.returncode != 0:
        raise RuntimeError(result.stderr or f"WriteExcelTool failed: {result.returncode}")


def create_sheet(file_path: str, sheet_name: str) -> None:
    result = _run_java("jp.isoittech.CreateSheetTool", [file_path, sheet_name])
    if result.returncode != 0:
        raise RuntimeError(result.stderr or f"CreateSheetTool failed: {result.returncode}")


def rename_worksheet(file_path: str, old_name: str, new_name: str) -> None:
    result = _run_java("jp.isoittech.RenameWorksheetTool", [file_path, old_name, new_name])
    if result.returncode != 0:
        raise RuntimeError(result.stderr or f"RenameWorksheetTool failed: {result.returncode}")


def delete_worksheet(file_path: str, sheet_name: str) -> None:
    result = _run_java("jp.isoittech.DeleteWorksheetTool", [file_path, sheet_name])
    if result.returncode != 0:
        raise RuntimeError(result.stderr or f"DeleteWorksheetTool failed: {result.returncode}")


def copy_worksheet(file_path: str, source_sheet: str, target_sheet: str) -> None:
    result = _run_java("jp.isoittech.CopyWorksheetTool", [file_path, source_sheet, target_sheet])
    if result.returncode != 0:
        raise RuntimeError(result.stderr or f"CopyWorksheetTool failed: {result.returncode}")


def apply_formula(file_path: str, sheet_name: str, cell: str, formula: str) -> None:
    result = _run_java("jp.isoittech.ApplyFormulaTool", [file_path, sheet_name, cell, formula])
    if result.returncode != 0:
        raise RuntimeError(result.stderr or f"ApplyFormulaTool failed: {result.returncode}")


def validate_formula_syntax(file_path: str, sheet_name: str, formula: str) -> bool:
    result = _run_java("jp.isoittech.ValidateFormulaSyntaxTool", [file_path, sheet_name, formula])
    if result.returncode != 0:
        raise RuntimeError(result.stderr or f"ValidateFormulaSyntaxTool failed: {result.returncode}")
    # Java side prints "OK" when the formula is syntactically valid.
    return "OK" in result.stdout


def format_range(
    file_path: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    bold: bool = False,
    italic: bool = False,
    font_size: int = 0,
    font_color: str = "",
    bg_color: str = "",
) -> None:
    args = [
        file_path,
        sheet_name,
        start_cell,
        end_cell,
        str(bold).lower(),
        str(italic).lower(),
        str(font_size),
        font_color,
        bg_color,
    ]
    result = _run_java("jp.isoittech.FormatRangeTool", args)
    if result.returncode != 0:
        raise RuntimeError(result.stderr or f"FormatRangeTool failed: {result.returncode}")


def merge_cells(file_path: str, sheet_name: str, start_cell: str, end_cell: str) -> None:
    result = _run_java("jp.isoittech.MergeCellsTool", [file_path, sheet_name, start_cell, end_cell])
    if result.returncode != 0:
        raise RuntimeError(result.stderr or f"MergeCellsTool failed: {result.returncode}")


def unmerge_cells(file_path: str, sheet_name: str, start_cell: str, end_cell: str) -> None:
    result = _run_java("jp.isoittech.UnmergeCellsTool", [file_path, sheet_name, start_cell, end_cell])
    if result.returncode != 0:
        raise RuntimeError(result.stderr or f"UnmergeCellsTool failed: {result.returncode}")


def copy_range(
    file_path: str,
    sheet_name: str,
    source_start: str,
    source_end: str,
    target_start: str,
    target_sheet: Optional[str] = None,
) -> None:
    args = [file_path, sheet_name, source_start, source_end, target_start]
    if target_sheet is not None:
        args.append(target_sheet)
    result = _run_java("jp.isoittech.CopyRangeTool", args)
    if result.returncode != 0:
        raise RuntimeError(result.stderr or f"CopyRangeTool failed: {result.returncode}")


def delete_range(
    file_path: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    shift_direction: str = "up",
) -> None:
    result = _run_java(
        "jp.isoittech.DeleteRangeTool",
        [file_path, sheet_name, start_cell, end_cell, shift_direction],
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr or f"DeleteRangeTool failed: {result.returncode}")


def validate_excel_range(
    file_path: str,
    sheet_name: str,
    start_cell: str,
    end_cell: Optional[str] = None,
) -> bool:
    args = [file_path, sheet_name, start_cell]
    if end_cell is not None:
        args.append(end_cell)
    result = _run_java("jp.isoittech.ValidateExcelRangeTool", args)
    if result.returncode != 0:
        raise RuntimeError(result.stderr or f"ValidateExcelRangeTool failed: {result.returncode}")
    return "OK" in result.stdout


def create_chart(
    file_path: str,
    sheet_name: str,
    data_range: str,
    chart_type: str,
    target_cell: str,
    title: Optional[str] = None,
    x_axis: Optional[str] = None,
    y_axis: Optional[str] = None,
) -> None:
    args = [file_path, sheet_name, data_range, chart_type, target_cell]
    if title is not None:
        args.append(title)
    if x_axis is not None:
        args.append(x_axis)
    if y_axis is not None:
        args.append(y_axis)
    result = _run_java("jp.isoittech.CreateChartTool", args)
    if result.returncode != 0:
        raise RuntimeError(result.stderr or f"CreateChartTool failed: {result.returncode}")


def create_pivot_table(
    file_path: str,
    sheet_name: str,
    data_range: str,
    rows: Iterable[str],
    values: Iterable[str],
    columns: Optional[Iterable[str]] = None,
    agg_func: str = "sum",
) -> str:
    """Simulate creation of a pivot table.

    The underlying Java tool currently does not modify the Excel file; it
    simply prints a confirmation message. That message is returned from
    this function.
    """

    rows_str = ",".join(rows)
    values_str = ",".join(values)
    columns_str = ",".join(columns) if columns is not None else ""

    args = [file_path, sheet_name, data_range, rows_str, values_str, columns_str, agg_func]
    result = _run_java("jp.isoittech.CreatePivotTableTool", args)
    if result.returncode != 0:
        raise RuntimeError(result.stderr or f"CreatePivotTableTool failed: {result.returncode}")
    return result.stdout.strip()
