from mcp.server.fastmcp import FastMCP
import win32com.client
from typing import Optional
from datetime import datetime
from fastapi import FastAPI
from fastapi.responses import JSONResponse, StreamingResponse

app = FastAPI()
mcp = FastMCP("local-file-search")

# MCPエンドポイントをFastAPIにマウント
app.mount("/mcp", mcp)

# /mcpにアクセスされた時の簡易レスポンス
@app.get("/mcp")
async def mcp_root_info():
    def info_generator():
        yield "retry: 1000\n"  # Retry field for reconnection
        yield "id: 1\n"  # Add unique ID for the message
        yield "data: {\n"
        yield 'data: "message": "This is an MCP-compatible server.",\n'
        yield 'data: "endpoints": "/mcp/tools/search_files/prompt",\n'
        yield 'data: "spec": "https://modelcontextprotocol.io/specification/2025-03-26/server/prompts"\n'
        yield "data: }\n\n"

    return StreamingResponse(info_generator(), media_type="text/event-stream")

@mcp.tool()
def search_files(
    query: str,
    extension: Optional[str] = None,
    modified_after: Optional[str] = None,
    min_size_kb: Optional[int] = None,
    max_size_kb: Optional[int] = None
) -> StreamingResponse:
    """Search indexed files on Windows using Windows Search. Optionally filter by file extension, modified date, and file size range (KB)."""
    conn = win32com.client.Dispatch("ADODB.Connection")
    conn.Open("Provider=Search.CollatorDSO;Extended Properties='Application=Windows';")

    conditions = [f"CONTAINS('{query}')"]
    base_path = "C:\\Users\\%USERNAME%\\Documents"
    conditions.append(f"System.ItemPathDisplay LIKE '{base_path}%'")

    if extension:
        conditions.append(f"System.FileExtension = '{extension}'")

    if modified_after:
        try:
            dt = datetime.fromisoformat(modified_after)
            iso_time = dt.strftime('%Y-%m-%dT%H:%M:%S')
            conditions.append(f"System.DateModified >= '{iso_time}'")
        except ValueError:
            def error_generator():
                yield "data: Invalid date format. Use ISO 8601 (e.g., 2024-01-01T00:00:00)\n\n"
            return StreamingResponse(error_generator(), media_type="text/event-stream")

    if min_size_kb is not None:
        conditions.append(f"System.Size >= {min_size_kb * 1024}")
    if max_size_kb is not None:
        conditions.append(f"System.Size <= {max_size_kb * 1024}")

    where_clause = " AND ".join(conditions)

    rs = conn.Execute(
        f"SELECT System.ItemName, System.ItemPathDisplay FROM SYSTEMINDEX WHERE {where_clause}"
    )[0]

    results = []
    while not rs.EOF:
        name = rs.Fields.Item("System.ItemName").Value
        path = rs.Fields.Item("System.ItemPathDisplay").Value
        results.append(f"{name} - {path}")
        rs.MoveNext()

    rs.Close()
    conn.Close()

    def result_generator():
        yield "retry: 1000\n"  # Retry field for reconnection
        message_id = 1  # Initialize message ID
        if results:
            for result in results:
                yield f"id: {message_id}\n"  # Add unique ID for each message
                yield f"data: {result}\n\n"
                message_id += 1
        else:
            yield f"id: {message_id}\n"
            yield "data: No matching files found.\n\n"

    return StreamingResponse(result_generator(), media_type="text/event-stream")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
