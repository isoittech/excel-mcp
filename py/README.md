# excel-mcp-server

Local MCP server for Excel operations backed by Java tools in the excel-mcp repository.

This package is not intended for publishing; it exists only to provide an MCP
server entry point using `uv run` as configured in `mcp-config.json`.

## Running the server with uv

The Excel MCP server is designed to be run with `uv` from the `py` directory:

```bash
cd /path/to/excel-mcp/py
uv run excel-mcp-server -t stdio
```

You can also use HTTP or SSE transports for local testing:

```bash
# HTTP (streamable-http) on port 8000
uv run excel-mcp-server -t http -p 8000

# Server-Sent Events (SSE) on port 8001
uv run excel-mcp-server -t sse -p 8001
```

> Note: MCP runtimes like Claude Desktop will typically use the `stdio`
> transport and will manage the server process lifecycle themselves.

## Using the server from an MCP client (e.g. Claude)

1. Make sure Java is available and the Excel Java tools are built under
   `../java` (`java/dist` and `java/jars` as in this repository).
2. From `excel-mcp/py`, install dependencies and build the virtual
   environment (the first `uv run` will do this automatically).
3. Configure your MCP client to use the `mcp-config.json` file in this
   directory. For example, Claude Desktop can be pointed at this
   directory so that it discovers the `excel-mcp-server` and uses:

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

4. Restart the MCP client so it picks up the new server configuration.
5. In the client UI, you should now see an `excel-mcp-server` toolset
   that exposes operations such as creating workbooks, reading/writing
   ranges, managing sheets, listing sheet names, and creating charts or pivot tables.

### List sheet names

A tool is available to list worksheet names in a workbook:

- Tool: `tool_list_sheets`
- Args: `path` (e.g. `/mnt/data/book.xlsx`)
- Returns: `{ "path": "...", "sheets": ["Sheet1", "Data", ...] }`

If the server fails to start, check the client logs for errors about
Java, the classpath, or missing Python dependencies, and verify that you
can run:

```bash
cd /path/to/excel-mcp/py
uv run excel-mcp-server --help
```

successfully from a terminal.

