from mcp.server.fastmcp import FastMCP
from typing import Optional
from datetime import datetime
import os
import win32com.client
from Foundation import NSDate
from CoreServices import (
    MDQueryCreate,
    MDQueryExecute,
    MDQueryGetAttributeValueAtIndex,
    kMDItemPath,    
)
import time
import platform
# オフィスファイル用のライブラリ
from docx import Document
from openpyxl import load_workbook
from pptx import Presentation

mcp = FastMCP("local-file-search")

@mcp.tool()
def search_local_files(
    query: str,
    extension: Optional[str] = None,
    modified_after: Optional[str] = None,
    min_size_kb: Optional[int] = None,
    max_size_kb: Optional[int] = None
) -> str:
    """Search indexed files on Windows or Mac. Optionally filter by file extension, modified date, and file size range (KB)."""
    system = platform.system()
    
    if system == 'Windows':
        return search_local_files_windows(query, extension, modified_after, min_size_kb, max_size_kb)
    elif system == 'Darwin':  # macOS
        return search_local_files_mac(query, extension, modified_after, min_size_kb, max_size_kb)
    else:
        return "Unsupported operating system"

def search_local_files_mac(
    query: str,
    extension: Optional[str] = None,
    modified_after: Optional[str] = None,
    min_size_kb: Optional[int] = None,
    max_size_kb: Optional[int] = None
) -> str:
    """Search files on macOS using MDQuery. Optionally filter by file extension, modified date, and file size range."""
    # MDQueryのクエリ文字列を構築
    query_parts = [f'kMDItemTextContent == "*{query}*"wc']
    
    if extension:
        query_parts.append(f'kMDItemFSName == "*.{extension}"')
    
    if modified_after:
        try:
            dt = datetime.fromisoformat(modified_after)
            timestamp = time.mktime(dt.timetuple())
            date = NSDate.dateWithTimeIntervalSince1970_(timestamp)
            query_parts.append(f'kMDItemFSContentChangeDate >= $time')
        except ValueError:
            return "Invalid date format. Use ISO 8601 (e.g., 2024-01-01T00:00:00)"

    # クエリの作成と実行
    query_string = ' && '.join(query_parts)
    mdquery = MDQueryCreate(None, query_string, None, None)
    
    if not mdquery:
        return "Failed to create search query"

    MDQueryExecute(mdquery, 0)
    
    # 結果の取得とフィルタリング
    filtered_files = []
    count = mdquery.resultCount()
    
    for i in range(count):
        path = MDQueryGetAttributeValueAtIndex(mdquery, kMDItemPath, i)
        if not path:
            continue
            
        try:
            size_kb = os.path.getsize(path) / 1024
            if min_size_kb and size_kb < min_size_kb:
                continue
            if max_size_kb and size_kb > max_size_kb:
                continue
                
            name = os.path.basename(path)
            filtered_files.append(f"{name} - file://{path}")
        except OSError:
            continue
    
    return '\n'.join(filtered_files) if filtered_files else "No matching files found."


def search_local_files_windows(
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


def read_word_file(file_path: str) -> str:
    """Wordファイルをpython-docxで読み取る"""
    try:
        doc = Document(file_path)
        return '\n'.join([paragraph.text for paragraph in doc.paragraphs])
    except ModuleNotFoundError:
        return "[ERROR] python-docx is not installed. Run: pip install python-docx"
    except Exception as e:
        return f"[ERROR] Failed to read Word file: {e}"

def read_excel_file(file_path: str) -> str:
    """Excelファイルをopenpyxlで読み取る"""
    try:
        wb = load_workbook(file_path, read_only=True, data_only=True)
        text = []
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            for row in ws.rows:
                row_values = [str(cell.value) if cell.value is not None else '' for cell in row]
                text.append('\t'.join(row_values))
        return '\n'.join(text)
    except ModuleNotFoundError:
        return "[ERROR] openpyxl is not installed. Run: pip install openpyxl"
    except Exception as e:
        return f"[ERROR] Failed to read Excel file: {e}"

def read_ppt_file(file_path: str) -> str:
    """PowerPointファイルをpython-pptxで読み取る"""
    try:
        prs = Presentation(file_path)
        text = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    text.append(shape.text)
        return '\n'.join(text)
    except ModuleNotFoundError:
        return "[ERROR] python-pptx is not installed. Run: pip install python-pptx"
    except Exception as e:
        return f"[ERROR] Failed to read PowerPoint file: {e}"

def read_office_file_com(file_path: str) -> str:
    """オフィスファイルをPythonライブラリで読み取る"""
    ext = os.path.splitext(file_path)[1].lower()

    # まずテキストファイルとして試みる
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()
    except:
        pass

    # Pythonライブラリで読み取りを試行
    try:
        if ext in [".doc", ".docx"]:
            return read_word_file(file_path)
        elif ext in [".xls", ".xlsx"]:
            return read_excel_file(file_path)
        elif ext in [".ppt", ".pptx"]:
            return read_ppt_file(file_path)
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
