from mcp.server.fastmcp import FastMCP
import win32com.client
from typing import Optional
from datetime import datetime

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
    # base_path = 'C:/Users/' + win32com.client.Dispatch("WScript.Network").UserName + '/Documents'
    # conditions.append(f"SCOPE='file:///{base_path}'")

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
        f"SELECT System.ItemName, System.ItemUrl FROM SYSTEMINDEX WHERE {where_clause}"
    )[0]

    results = []
    while not rs.EOF:
        name = rs.Fields.Item("System.ItemName").Value
        path = rs.Fields.Item("System.ItemUrl").Value
        results.append(f"{name} - {path}")
        rs.MoveNext()

    rs.Close()
    conn.Close()

    return "\n".join(results) if results else "No matching files found."


import os
import win32com.client


def read_word_com(file_path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(file_path, ReadOnly=True)
        text = doc.Content.Text
        doc.Close(False)
        return text
    finally:
        word.Quit()


def read_excel_com(file_path):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(file_path)
        text = []
        for sheet in wb.Sheets:
            for row in sheet.UsedRange.Rows:
                row_values = [str(cell.Value) if cell.Value is not None else '' for cell in row.Columns]
                text.append('\t'.join(row_values))
        wb.Close(False)
        return '\n'.join(text)
    finally:
        excel.Quit()


def read_ppt_com(file_path):
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    ppt.Visible = False
    try:
        presentation = ppt.Presentations.Open(file_path, WithWindow=False)
        text = []
        for slide in presentation.Slides:
            for shape in slide.Shapes:
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    text.append(shape.TextFrame.TextRange.Text)
        presentation.Close()
        return '\n'.join(text)
    finally:
        ppt.Quit()


def read_office_file_com(file_path):
    ext = os.path.splitext(file_path)[1].lower()

    # まずテキストファイルとして試みる
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()
    except:
        pass  # テキストとして読めなかったらCOMに切り替え

    # COMで読み取りを試行
    try:
        if ext in [".doc", ".docx"]:
            return read_word_com(file_path)
        elif ext in [".xls", ".xlsx"]:
            return read_excel_com(file_path)
        elif ext in [".ppt", ".pptx"]:
            return read_ppt_com(file_path)
        else:
            return f"[SKIP] Unsupported file type: {ext}"
    except Exception as e:
        return f"[ERROR] Failed to read {file_path}: {e}"

@mcp.tool()
def local_read_file(path: str)-> str:
    """Read a file."""
    return read_office_file_com(path)

if __name__ == "__main__":
    mcp.run(transport="sse")
