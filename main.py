from mcp.server.fastmcp import FastMCP
import win32com.client
from typing import Optional
from datetime import datetime
from fastapi import FastAPI
from fastapi.responses import JSONResponse

mcp = FastMCP("local-file-search")

@mcp.tool()
def search_local_files(
    query: str,
    extension: Optional[str] = None,
    modified_after: Optional[str] = None,
    min_size_kb: Optional[int] = None,
    max_size_kb: Optional[int] = None
) -> str:
    """Search indexed files on Windows using Windows Search. Optionally filter by file extension, modified date, and file size range (KB)."""
    conn = win32com.client.Dispatch("ADODB.Connection")
    conn.Open("Provider=Search.CollatorDSO;Extended Properties='Application=Windows';")

    conditions = [f"CONTAINS('{query}')"]
    # base_path = "C:\\Users\\%USERNAME%\\Documents"
    # conditions.append(f"System.ItemPathDisplay LIKE '{base_path}%'")

    if extension:
        conditions.append(f"System.FileExtension = '{extension}'")

    if modified_after:
        try:
            dt = datetime.fromisoformat(modified_after)
            iso_time = dt.strftime('%Y-%m-%dT%H:%M:%S')
            conditions.append(f"System.DateModified >= '{iso_time}'")
        except ValueError:
            return "Invalid date format. Use ISO 8601 (e.g., 2024-01-01T00:00:00)"

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

    return "\n".join(results) if results else "No matching files found."

if __name__ == "__main__":
    mcp.run(transport="sse")
