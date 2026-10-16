# excel_mcp_server.py
import os
import io
import contextlib
import traceback
import openpyxl
from openpyxl.cell import Cell
from openpyxl.utils import get_column_letter
import pandas as pd
from fastmcp import FastMCP

# Khởi tạo FastMCP Server
mcp = FastMCP("ExcelUniversalQueryServer")


def _resolve_path(file_path: str) -> str:
    """Hỗ trợ tìm file linh hoạt trong thư mục hiện tại hoặc thư mục xlsx_files."""
    if os.path.exists(file_path):
        return file_path
    alt_path = os.path.join("xlsx_files", os.path.basename(file_path))
    if os.path.exists(alt_path):
        return alt_path
    raise FileNotFoundError(
        f"Không tìm thấy file: {file_path} (hoặc {alt_path})")


@mcp.tool()
def inspect_workbook(file_path: str) -> str:
    """
    Kiểm tra cấu trúc tổng quan của file Excel (.xlsx):
    Trả về danh sách các sheet, số dòng, số cột, các vùng gộp ô (merged cells) và 5 dòng đầu tiên.
    Luôn gọi hàm này đầu tiên khi làm việc với một file mới.
    """
    try:
        path = _resolve_path(file_path)
        wb = openpyxl.load_workbook(path, data_only=True)
        info = [f"=== THÔNG TIN FILE: {os.path.basename(path)} ==="]

        for name in wb.sheetnames:
            sheet = wb[name]
            merged_Count = len(sheet.merged_cells.ranges)
            merged_sample = [str(r)
                             for r in list(sheet.merged_cells.ranges)[:8]]
            info.append(
                f"\n- Sheet: '{name}' | Kích thước: {sheet.max_row} dòng x {sheet.max_column} cột "
                f"| Số vùng gộp ô: {merged_Count} (Ví dụ: {merged_sample})"
            )
            # Xem trước 8 dòng đầu có dữ liệu
            preview = []
            for r in range(1, min(9, sheet.max_row + 1)):
                row_vals = [sheet.cell(r, c).value for c in range(
                    1, min(12, sheet.max_column + 1))]
                if any(v is not None for v in row_vals):
                    preview.append(f"  Dòng {r}: {row_vals}")
            info.extend(preview)

        wb.close()
        return "\n".join(info)
    except Exception as e:
        return f"Lỗi khi kiểm tra workbook: {str(e)}"


@mcp.tool()
def read_sheet_grid(
    file_path: str,
    sheet_name: str,
    start_row: int = 1,
    end_row: int = 100,
    start_col: int = 1,
    end_col: int = 25,
    fill_merged: bool = True
) -> str:
    """
    Đọc nội dung chi tiết của một sheet dưới dạng bảng có gắn chỉ số Dòng (R) và Cột (C).
    Nếu fill_merged=True, các ô gộp (merged cells) sẽ được tự động điền đầy giá trị giúp dễ dàng
    nhận diện tiêu đề cột/nhóm phân cấp.
    """
    try:
        path = _resolve_path(file_path)
        wb = openpyxl.load_workbook(path, data_only=True)
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet không tồn tại: {sheet_name}")
        sheet = wb[sheet_name]

        # Xử lý gộp ô: sao chép giá trị từ ô gốc sang toàn bộ vùng gộp
        if fill_merged:
            for m_range in list(sheet.merged_cells.ranges):
                top_left_val = sheet.cell(
                    m_range.min_row, m_range.min_col).value
                sheet.unmerge_cells(str(m_range))
                for r in range(m_range.min_row, m_range.max_row + 1):
                    for c in range(m_range.min_col, m_range.max_col + 1):
                        cell = sheet.cell(r, c)
                        if isinstance(cell, Cell):
                            cell.value = top_left_val

        max_r = min(end_row, sheet.max_row)
        max_c = min(end_col, sheet.max_column)

        rows_data = []
        headers = [
            "Row"] + [f"C{c}({get_column_letter(c)})" for c in range(start_col, max_c + 1)]

        for r in range(start_row, max_r + 1):
            row_vals = [sheet.cell(r, c).value for c in range(
                start_col, max_c + 1)]
            if any(v is not None for v in row_vals):
                formatted_vals = [
                    str(v).replace("\n", " ").strip() if v is not None else ""
                    for v in row_vals
                ]
                rows_data.append([f"R{r}"] + formatted_vals)

        wb.close()
        if not rows_data:
            return f"Sheet '{sheet_name}' không có dữ liệu trong vùng dòng {start_row}-{max_r}."

        df = pd.DataFrame(rows_data, columns=headers)
        return f"Sheet: '{sheet_name}' (Hiển thị dòng {start_row} đến {max_r}):\n" + df.to_markdown(index=False)
    except Exception as e:
        return f"Lỗi khi đọc sheet: {str(e)}"


@mcp.tool()
def search_in_workbook(file_path: str, keyword: str) -> str:
    """
    Tìm kiếm một từ khóa (không phân biệt hoa thường) trên tất cả các sheet của file Excel.
    Trả về tọa độ (Sheet, Dòng, Cột) và giá trị của các ô khớp.
    """
    try:
        path = _resolve_path(file_path)
        wb = openpyxl.load_workbook(path, data_only=True)
        matches = []
        kw_lower = keyword.lower()

        for name in wb.sheetnames:
            sheet = wb[name]
            for r in range(1, sheet.max_row + 1):
                for c in range(1, sheet.max_column + 1):
                    val = sheet.cell(r, c).value
                    if val is not None and kw_lower in str(val).lower():
                        col_letter = get_column_letter(c)
                        matches.append(
                            f"[{name}] R{r}C{c} ({col_letter}{r}): {repr(val)}")
                        if len(matches) >= 50:
                            wb.close()
                            return "Tìm thấy quá nhiều kết quả (hiển thị 50 kết quả đầu):\n" + "\n".join(matches)
        wb.close()
        return "\n".join(matches) if matches else f"Không tìm thấy từ khóa '{keyword}'."
    except Exception as e:
        return f"Lỗi khi tìm kiếm: {str(e)}"


@mcp.tool()
def run_python_analysis(file_path: str, python_code: str) -> str:
    """
    Thực thi đoạn mã Python tùy chỉnh để phân tích và tính toán số liệu chính xác từ file Excel.
    Biến `file_path` (đường dẫn thực tế), `pd` (pandas), và `openpyxl` đã được cung cấp sẵn trong môi trường chạy.
    BẮT BUỘC dùng hàm `print()` trong `python_code` để in kết quả tính toán ra màn hình.
    """
    try:
        resolved_path = _resolve_path(file_path)
        stdout_capture = io.StringIO()

        # Môi trường biến cục bộ cung cấp sẵn cho đoạn code của AI
        exec_globals = {
            "__builtins__": __builtins__,
            "file_path": resolved_path,
            "pd": pd,
            "openpyxl": openpyxl,
        }

        with contextlib.redirect_stdout(stdout_capture):
            exec(python_code, exec_globals)

        output = stdout_capture.getvalue()
        return output if output.strip() else "Code đã chạy thành công nhưng không có lệnh print() nào xuất dữ liệu."
    except Exception:
        return f"Lỗi khi thực thi Python code:\n{traceback.format_exc()}"


if __name__ == "__main__":
    # Chạy MCP server qua giao thức stdio chuẩn
    mcp.run(transport="stdio")
