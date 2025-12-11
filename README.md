# Python-based MCP Server for Excel Operations

This repository provides an MCP server that exposes Excel manipulation tools to MCP clients
(such as Microsoft Copilot or Claude) by wrapping Java-based tools from Python.

The core Excel operations are implemented in Java using Apache POI.
On top of that, we build the MCP server in Python, which gives us the following advantages:

- Apache POI is a very mature library for Excel operations, with rich features and good compatibility
- It is easier to implement an MCP server in Python and to leverage libraries like `fastmcp`

By adding a Python wrapper around the Java logic, we can:

- Let Java/Apache POI handle the heavy lifting for Excel processing
- Keep the MCP server itself lightweight and easy to extend in Python

## Table of Contents

1. [Requirements](#requirements)
2. [Installation](#installation)
3. [Using with MCP clients](#using-with-mcp-clients)
4. [Available tools](#available-tools)
5. [Notes](#notes)

## Requirements

- Java 11 or later
- Python 3.11 or later
- Excel file format: .xlsx
- Supported OS: Windows 10/11, macOS 12+, Linux (Ubuntu 20.04+)

## Installation

### 1. Clone the repository

```bash
git clone git@github.com:isoittech/excel-mcp.git
cd excel-mcp
```

### 2. Build the Java tools

You can compile the Java-side Excel tools using `java/tools/compile.sh`:

```bash
cd java
./tools/compile.sh
cd ..
```

### 3. Set up the Python MCP server

The Python MCP server lives under the `py` directory.
It is designed to be run locally using `uv`.

```bash
cd py
uv run excel-mcp-server --help
```

On the first run, dependencies defined in `pyproject.toml` will be installed automatically.

## Using with MCP clients

### Example: Claude Desktop

1. Point your MCP configuration to the `excel-mcp/py` directory.
2. The file `py/mcp-config.json` contains the definition for this MCP server.

Example `mcp-config.json`:

```json
{
  "mcpServers": {
    "excel-mcp-server": {
      "command": "uv",
      "args": [
        "run",
        "excel-mcp-server"
      ],
      "env": {
        "PYTHONPATH": "./src"
      }
    }
  }
}
```

After restarting your MCP client, you should see a server named `excel-mcp-server`
with a set of tools for working with Excel files.

### Using Docker

The `py` directory also includes a `Dockerfile` and `docker-compose.yml`.
You can use them if you prefer to connect via HTTP/SSE from your MCP client.

Build and run the container manually:

```bash
cd py
docker build -t excel-mcp-server .
docker run --rm -p 8585:8585 excel-mcp-server excel-mcp-server -t sse -p 8585
```

Or use docker compose:

```bash
cd py
docker compose up -d
```

From the MCP client side, configure an SSE endpoint such as
`http://host.docker.internal:8585/sse`.

## Available tools

Below are some examples of the tools exposed to MCP clients.

### Read from an Excel file

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "read_excel",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "range": "A1:C10"
  }
}
```

### Write to an Excel file

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "write_excel",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "data": [
      ["A1", "B1", "C1"],
      ["A2", "B2", "C2"]
    ]
  }
}
```

### Create a new sheet

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "create_sheet",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "NewSheet"
  }
}
```

### Create a new Excel file

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "create_excel",
  "arguments": {
    "filePath": "/path/to/new_file.xlsx",
    "sheetName": "Sheet1"  // optional, default is "Sheet1"
  }
}
```

### Get workbook metadata

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "get_workbook_metadata",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "includeRanges": false  // optional, whether to include range info
  }
}
```

### Rename a worksheet

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "rename_worksheet",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "oldName": "Sheet1",
    "newName": "NewName"
  }
}
```

### Delete a worksheet

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "delete_worksheet",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1"
  }
}
```

### Copy a worksheet

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "copy_worksheet",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sourceSheet": "Sheet1",
    "targetSheet": "Sheet1Copy"
  }
}
```

### Apply a formula to a cell

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "apply_formula",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "cell": "C1",
    "formula": "=SUM(A1:B1)"
  }
}
```

### Validate formula syntax

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "validate_formula_syntax",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "cell": "C1",
    "formula": "=SUM(A1:B1)"
  }
}
```

### Format a range

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "format_range",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C3",
    "bold": true,
    "italic": false,
    "fontSize": 12,
    "fontColor": "#FF0000",
    "bgColor": "#FFFF00"
  }
}
```

### Merge cells

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "merge_cells",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C1"
  }
}
```

### Unmerge cells

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "unmerge_cells",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C1"
  }
}
```

### Copy a range

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "copy_range",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "sourceStart": "A1",
    "sourceEnd": "C3",
    "targetStart": "D1",
    "targetSheet": "Sheet2"  // optional, defaults to the same sheet
  }
}
```

### Delete a range

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "delete_range",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C3",
    "shiftDirection": "up"  // "up" or "left", default is "up"
  }
}
```

### Validate an Excel range

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "validate_excel_range",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C3"  // optional
  }
}
```

### Create a chart

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "create_chart",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "dataRange": "A1:C10",
    "chartType": "column",  // "column", "line", "bar", "area", "scatter", "pie"
    "targetCell": "E1",
    "title": "Sample Chart",  // optional
    "xAxis": "X Axis",        // optional
    "yAxis": "Y Axis"         // optional
  }
}
```

### Create a pivot table

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "create_pivot_table",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "dataRange": "A1:D100",
    "rows": ["Category"],
    "values": ["Sales"],
    "columns": ["Region"],  // optional
    "aggFunc": "sum"  // "sum", "count", "average", "max", "min", etc.
  }
}
```

## Notes

- Always use absolute paths for file paths.
- If you omit the sheet name, the first sheet will be used by default.
- Ranges should be specified in A1 notation such as `"A1:C10"`.
- `create_excel` will fail if the target file already exists.
- Depending on the current implementation, pivot table support may only build metadata
  and may not create a full Excel pivot table object in the file.

## Author

- Author: isoittech

## License

This MCP server is provided under the MIT License.
You are free to use, modify, and redistribute it.
See the LICENSE file in this repository for details.