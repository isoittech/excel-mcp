#!/usr/bin/env python
"""MCP server for Excel operations backed by Java tools.

This server exposes the Excel functionality implemented in the Java
classes under ``java/src/jp/isoittech`` via the ``py/src/excel``
wrappers. The design mirrors the PowerPoint MCP server under
``pptx-mcp`` but is focused on workbook and worksheet manipulation.
"""

from __future__ import annotations

import argparse
from typing import Any, Dict

from mcp.server.fastmcp import FastMCP

from excel import (
    create_excel,
    read_excel,
    write_excel,
    create_sheet,
    rename_worksheet,
    delete_worksheet,
    copy_worksheet,
    apply_formula,
    validate_formula_syntax,
    format_range,
    merge_cells,
    unmerge_cells,
    copy_range,
    delete_range,
    validate_excel_range,
    create_chart,
    create_pivot_table,
)


app = FastMCP(name="excel-mcp-server")


# ---------------------------------------------------------------------------
# Tool registrations
# ---------------------------------------------------------------------------


@app.tool()
async def tool_create_excel(path: str, sheet_name: str = "Sheet1") -> str:
    """Create a new Excel workbook on disk.

    Args:
        path: Target file path to create.
        sheet_name: Name of the initial worksheet.
    """

    create_excel(path, sheet_name)
    return f"Created Excel workbook at {path} with sheet '{sheet_name}'"


@app.tool()
async def tool_read_excel(path: str, sheet_name: str, range_str: str) -> Dict[str, Any]:
    """Read a rectangular range from an Excel sheet.

    Returns a JSON-serializable object with a ``data`` field containing
    the 2D array of cell values.
    """

    data = read_excel(path, sheet_name, range_str)
    return {"path": path, "+sheet": sheet_name, "range": range_str, "data": data}


@app.tool()
async def tool_write_excel(path: str, sheet_name: str, data: Any) -> str:
    """Write a 2D array of values into an Excel sheet starting at A1."""

    write_excel(path, sheet_name, data)
    return f"Wrote {len(data)} rows to {path}:{sheet_name}!A1"


@app.tool()
async def tool_create_sheet(path: str, sheet_name: str) -> str:
    create_sheet(path, sheet_name)
    return f"Created sheet '{sheet_name}' in {path}"


@app.tool()
async def tool_rename_worksheet(path: str, old_name: str, new_name: str) -> str:
    rename_worksheet(path, old_name, new_name)
    return f"Renamed sheet '{old_name}' to '{new_name}' in {path}"


@app.tool()
async def tool_delete_worksheet(path: str, sheet_name: str) -> str:
    delete_worksheet(path, sheet_name)
    return f"Deleted sheet '{sheet_name}' in {path}"


@app.tool()
async def tool_copy_worksheet(path: str, source_sheet: str, target_sheet: str) -> str:
    copy_worksheet(path, source_sheet, target_sheet)
    return f"Copied sheet '{source_sheet}' to '{target_sheet}' in {path}"


@app.tool()
async def tool_apply_formula(path: str, sheet_name: str, cell: str, formula: str) -> str:
    apply_formula(path, sheet_name, cell, formula)
    return f"Applied formula '{formula}' to {path}:{sheet_name}!{cell}"


@app.tool()
async def tool_validate_formula_syntax(path: str, sheet_name: str, formula: str) -> Dict[str, Any]:
    ok = validate_formula_syntax(path, sheet_name, formula)
    return {"valid": ok, "formula": formula}


@app.tool()
async def tool_format_range(
    path: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    bold: bool = False,
    italic: bool = False,
    font_size: int = 0,
    font_color: str = "",
    bg_color: str = "",
) -> str:
    format_range(path, sheet_name, start_cell, end_cell, bold, italic, font_size, font_color, bg_color)
    return f"Formatted range {sheet_name}!{start_cell}:{end_cell} in {path}"


@app.tool()
async def tool_merge_cells(path: str, sheet_name: str, start_cell: str, end_cell: str) -> str:
    merge_cells(path, sheet_name, start_cell, end_cell)
    return f"Merged cells {sheet_name}!{start_cell}:{end_cell} in {path}"


@app.tool()
async def tool_unmerge_cells(path: str, sheet_name: str, start_cell: str, end_cell: str) -> str:
    unmerge_cells(path, sheet_name, start_cell, end_cell)
    return f"Unmerged cells {sheet_name}!{start_cell}:{end_cell} in {path}"


@app.tool()
async def tool_copy_range(
    path: str,
    sheet_name: str,
    source_start: str,
    source_end: str,
    target_start: str,
    target_sheet: str | None = None,
) -> str:
    copy_range(path, sheet_name, source_start, source_end, target_start, target_sheet)
    return f"Copied range {sheet_name}!{source_start}:{source_end} to {target_sheet or sheet_name}!{target_start} in {path}"


@app.tool()
async def tool_delete_range(
    path: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    shift_direction: str = "up",
) -> str:
    delete_range(path, sheet_name, start_cell, end_cell, shift_direction)
    return f"Deleted range {sheet_name}!{start_cell}:{end_cell} in {path} (shift={shift_direction})"


@app.tool()
async def tool_validate_excel_range(
    path: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str | None = None,
) -> Dict[str, Any]:
    ok = validate_excel_range(path, sheet_name, start_cell, end_cell)
    return {"valid": ok, "sheet": sheet_name, "start": start_cell, "end": end_cell}


@app.tool()
async def tool_create_chart(
    path: str,
    sheet_name: str,
    data_range: str,
    chart_type: str,
    target_cell: str,
    title: str | None = None,
    x_axis: str | None = None,
    y_axis: str | None = None,
) -> str:
    create_chart(path, sheet_name, data_range, chart_type, target_cell, title, x_axis, y_axis)
    return f"Created chart of type '{chart_type}' at {sheet_name}!{target_cell} in {path}"


@app.tool()
async def tool_create_pivot_table(
    path: str,
    sheet_name: str,
    data_range: str,
    rows: list[str],
    values: list[str],
    columns: list[str] | None = None,
    agg_func: str = "sum",
) -> Dict[str, Any]:
    message = create_pivot_table(path, sheet_name, data_range, rows, values, columns, agg_func)
    return {"message": message}


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


# ---- Main Function ----
def main(transport: str = "stdio", port: int = 8000) -> None:
    """Entry point used by the ``excel-mcp-server`` console script.

    Mirrors the PowerPoint MCP server interface, allowing the transport
    to be selected from the command line while keeping stdio as the
    default for MCP runtimes.
    """

    if transport == "http":
        import asyncio

        # Set the port for HTTP transport if supported by FastMCP
        app.settings.port = port
        try:
            app.run(transport="streamable-http")
        except asyncio.exceptions.CancelledError:
            print("Server stopped by user.")
        except KeyboardInterrupt:
            print("Server stopped by user.")
        except Exception as e:  # pragma: no cover - defensive logging
            print(f"Error starting server: {e}")

    elif transport == "sse":
        # Run the FastMCP server in SSE (Server Side Events) mode
        app.run(transport="sse")

    else:
        # Default: stdio transport for MCP tooling
        app.run(transport="stdio")


if __name__ == "__main__":  # pragma: no cover
    # Parse command line arguments
    parser = argparse.ArgumentParser(description="MCP Server for Excel manipulation")

    parser.add_argument(
        "-t",
        "--transport",
        type=str,
        default="stdio",
        choices=["stdio", "http", "sse"],
        help="Transport method for the MCP server (default: stdio)",
    )

    parser.add_argument(
        "-p",
        "--port",
        type=int,
        default=8000,
        help="Port to run the MCP server on (default: 8000)",
    )

    args = parser.parse_args()
    main(args.transport, args.port)
