FROM python:3.11-alpine

# Install system dependencies needed by Java and Python wrappers
RUN apk add --no-cache openjdk11-jre-headless gcc musl-dev libffi-dev

# Set work directory
WORKDIR /app

# Copy the Excel MCP server package (this directory corresponds to excel-mcp/py)
COPY . .

# Install uv and project dependencies
RUN pip install --no-cache-dir uv && \
    uv pip install --system .

# Default entrypoint: run the Excel MCP server in stdio mode
ENTRYPOINT ["uv", "run", "excel-mcp-server"]
