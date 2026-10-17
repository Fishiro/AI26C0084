# main.py
import os
import sys
import io
import re
import csv
import json
import argparse
import traceback
import contextlib
import threading
import concurrent.futures
import time
from typing import List, Dict, Any
import openpyxl
from openpyxl.cell import Cell
from openpyxl.utils import get_column_letter
import pandas as pd
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

PRICE_1M_PROMPT = 0.10
PRICE_1M_COMPLETION = 0.20

SYSTEM_PROMPT = """Bạn là hệ thống AI chuyên phân tích dữ liệu bảng tính Excel (.xlsx) với độ chính xác tuyệt đối.
Câu hỏi có thể bằng tiếng Anh, tiếng Việt, tiếng Pháp hoặc ngôn ngữ khác. Hãy trả lời đúng định dạng mà câu hỏi yêu cầu.

QUY TẮC BẮT BUỘC (CHAIN OF THOUGHT):
1. Suy nghĩ ngắn gọn trong thẻ <THOUGHT>...</THOUGHT>.
2. Nếu bảng nhỏ và bạn đã thấy rõ đáp án chính xác trên bảng xem trước (Overview), hãy xuất ngay <FINAL_ANSWER>.
3. Nếu cần chạy Python tool, hãy SỬ DỤNG CÁC HÀM TRỢ THỦ SAU đây thay vì tự viết logic pandas phức tạp để tránh lỗi:
   - `df = load_df()`: Lấy DataFrame an toàn (đã tự động gộp header).
   - `df = extract_table(df, "keyword")`: TỐI QUAN TRỌNG VÀ BẮT BUỘC! File thống kê luôn có 3-4 dòng title rác trên cùng. Ngay dòng code đầu tiên, BẠN PHẢI GỌI `df = extract_table(df, "từ_khóa")` (ví dụ từ khóa là 'Year', 'Occupation', 'Age') để cắt bỏ title rác và gán lại Header. Nếu không làm bước này, code của bạn sẽ bị lỗi KeyError 100%.- `df = extract_table(df, "keyword")`: RẤT QUAN TRỌNG! Nếu DataFrame bị dính các dòng tiêu đề báo cáo dài (title) làm hỏng Header, hãy gọi hàm này với `keyword` là tên của một cột thực sự để cắt bỏ title và thiết lập lại Header chuẩn.
   - `find_row(df, "keyword")`: Trả về Index của dòng ĐẦU TIÊN chứa keyword.
   - `find_cells(df, "keyword")`: Trả về DANH SÁCH TẤT CẢ các Index dòng chứa keyword.
   - `find_col(df, "keyword")`: Trả về tên cột chứa keyword.
   - `get_val(df, row_idx, col_name_or_idx)`: Lấy giá trị an toàn.
   - `clean_num(val)`: Ép kiểu 1 giá trị thành số thực (float) an toàn (xóa $, %, dấu phẩy, xử lý ký tự '-').
   - `clean_col(df, "col_name")`: Ép kiểu toàn bộ 1 cột thành số thực an toàn. Trả về Pandas Series.

   * VÍ DỤ AN TOÀN 1 (Tìm và lấy 1 ô dữ liệu):
     ```python
     df = load_df()
     r = find_row(df, "Mexican")
     c = find_col(df, "2023")
     if r != -1 and c:
         val = get_val(df, r, c)
         print(clean_num(val)) # Luôn clean_num trước khi in/tính toán
     ```
   
   * VÍ DỤ AN TOÀN 2 (Cắt bảng, làm sạch cột và tính toán):
     ```python
     df = load_df()
     df = extract_table(df, "Occupation") 
     c = find_col(df, "Total")
     if c:
         df[c] = clean_col(df, c) 
         # BẮT BUỘC DÙNG round(val, 4) ĐỂ TRÁNH LỖI SỐ THẬP PHÂN PYTHON (VD: 30.299999997)
         result = round(df[c].sum(), 4)
         print(result)
     ```

   - NẾU TỰ VIẾT PANDAS: 
     - KHÔNG BAO GIỜ dùng `.loc['Tên dòng']` hoặc `df['Tên cột đầy đủ']` vì tên cột do gộp ô rất dài và có ký tự ẩn. Ưu tiên dùng `.iloc[row_idx, col_idx]`.
     - KHÔNG DÙNG `.astype(float)` hay `pd.to_numeric` thủ công vì sẽ lỗi với ký tự `,` hoặc `$`. Luôn dùng `clean_col()`.
     - Luôn in kết quả ra màn hình bằng lệnh `print()`. Nếu code báo lỗi, hãy đọc log lỗi để sửa lại code chứ không lặp lại.
4. CẢNH BÁO QUAN TRỌNG VỀ SUY LUẬN DỮ LIỆU ĐỂ ĐẠT EXACT MATCH:
   - QUY TẮC TÍNH TRUNG BÌNH/TỔNG: Theo quy ước BTC, nếu câu hỏi yêu cầu tính trung bình/tổng "across all..." (ví dụ: across all education levels) mà KHÔNG yêu cầu loại trừ, bạn BẮT BUỘC PHẢI TÍNH TRÊN TOÀN BỘ CÁC CỘT (BAO GỒM CẢ CỘT "TOTAL"). Không tự ý loại bỏ dòng/cột Total trong trường hợp này.
   - QUY TẮC TÍNH TUỔI: Nếu bảng có cột Date of Birth và Date of Admission, hãy tính tuổi bằng cách lấy Date of Admission trừ đi Date of Birth (Tuổi nhập học thấp = Trẻ nhất). BẮT BUỘC dùng hàm trợ thủ df['col'] = parse_date(df, 'col') để ép kiểu cả 2 cột sang ngày tháng trước khi trừ để tránh lỗi TypeError.
   - GIỮ NGUYÊN GIÁ TRỊ GỐC: Nếu tiêu đề bảng có ghi chú đơn vị (ví dụ: "[In thousands]", "[In millions]"), TUYỆT ĐỐI KHÔNG tự ý nhân thêm 1000 hay 1000000. Trả về con số y hệt như nó hiển thị trong ô.
5. Khi đã có kết quả cuối cùng, đặt ĐÁP ÁN DUY NHẤT vào thẻ <FINAL_ANSWER>...</FINAL_ANSWER>:
   - Chỉ ghi đáp án, tuyệt đối không kèm lời giải thích hay đơn vị thừa.
   - Tuân thủ chính xác số chữ số thập phân được yêu cầu (ví dụ: `50.0`, `14.20`, `40.60, 37.60`). Nếu câu hỏi có yêu cầu định dạng thập phân, bạn PHẢI bảo toàn các số 0 ở cuối.
"""

TOOLS_SCHEMA = [
    {
        "type": "function",
        "function": {
            "name": "run_python_analysis",
            "description": "Thực thi code Python (pandas/openpyxl) để tính toán chính xác trên dữ liệu Excel. Bắt buộc dùng print() để xuất kết quả.",
            "parameters": {
                "type": "object",
                "properties": {
                    "python_code": {
                        "type": "string",
                        "description": "Code Python cần thực thi. Có sẵn các biến: `pd`, `sheet_names`, `sheets_raw` và các hàm trợ thủ `load_df()`, `find_row(df, kw)`, `find_col(df, kw)`, `get_val(df, r, c)`."
                    }
                },
                "required": ["python_code"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "read_sheet_grid",
            "description": "Đọc chi tiết một vùng dòng/cột của sheet cụ thể dưới dạng bảng có tọa độ R (dòng) và C (cột).",
            "parameters": {
                "type": "object",
                "properties": {
                    "sheet_name": {"type": "string", "description": "Tên sheet cần đọc."},
                    "start_row": {"type": "integer", "default": 1},
                    "end_row": {"type": "integer", "default": 120},
                },
                "required": [],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "search_keyword",
            "description": "Tìm kiếm nhanh vị trí (Sheet, Dòng, Cột) của một từ khóa trên toàn bộ các sheets của file Excel. Dùng khi không biết dữ liệu nằm ở sheet nào hoặc bảng quá lớn.",
            "parameters": {
                "type": "object",
                "properties": {
                    "keyword": {"type": "string", "description": "Từ khóa cần tìm (Ví dụ: 'Total', 'Revenue', 'Mexican')"}
                },
                "required": ["keyword"],
            },
        },
    },
]


class ExcelWorkspace:
    """Đọc, tách ô gộp và lưu cache toàn bộ các sheets của file Excel trên RAM."""

    def __init__(self, xlsx_dir: str, file_name: str):
        self.file_path = self._resolve_path(xlsx_dir, file_name)
        self.sheets_raw: Dict[str, List[List[Any]]] = {}
        self.merged_counts: Dict[str, int] = {}
        self._load_workbook()

    @staticmethod
    def _resolve_path(xlsx_dir: str, file_name: str) -> str:
        base = os.path.basename(file_name)
        candidates = [
            os.path.join(xlsx_dir, file_name),
            os.path.join(xlsx_dir, base),
            os.path.join(xlsx_dir, "xlsx_files", base),
            os.path.join("input", base),
            os.path.join("input", "xlsx_files", base),
            os.path.join("/input/xlsx_files", base),
        ]
        for c in candidates:
            if os.path.isfile(c):
                return c
        for root, _, files in os.walk(xlsx_dir if os.path.exists(xlsx_dir) else "."):
            if "venv" not in root and ".venv" not in root and base in files:
                return os.path.join(root, base)
        raise FileNotFoundError(
            f"Không tìm thấy tệp '{file_name}' trong '{xlsx_dir}'.")

    def _load_workbook(self):
        wb = openpyxl.load_workbook(self.file_path, data_only=True)
        for name in wb.sheetnames:
            sheet = wb[name]
            self.merged_counts[name] = len(sheet.merged_cells.ranges)
            for m_range in list(sheet.merged_cells.ranges):
                top_val = sheet.cell(m_range.min_row, m_range.min_col).value
                sheet.unmerge_cells(str(m_range))
                for r in range(m_range.min_row, m_range.max_row + 1):
                    for c in range(m_range.min_col, m_range.max_col + 1):
                        cell = sheet.cell(r, c)
                        if isinstance(cell, Cell):
                            cell.value = top_val

            data = []
            for r in range(1, sheet.max_row + 1):
                data.append(
                    [sheet.cell(r, c).value for c in range(1, sheet.max_column + 1)])
            self.sheets_raw[name] = data
        wb.close()

    def get_overview(self, global_max_rows: int = 150, max_cols: int = 20) -> str:
        out = [
            f"=== WORKBOOK: {os.path.basename(self.file_path)} | Danh sách Sheets: {list(self.sheets_raw.keys())} ==="]
        total_sheets = len(self.sheets_raw)

        if total_sheets == 0:
            return "\n".join(out)

        base_budget_per_sheet = max(15, global_max_rows // total_sheets)

        for name, data in self.sheets_raw.items():
            total_r = len(data)
            total_c = len(data[0]) if total_r > 0 else 0

            max_r = min(total_r, max(base_budget_per_sheet,
                        35 if total_sheets == 1 else 20))

            out.append(
                f"\n--- SHEET: '{name}' (Kích thước: {total_r} dòng x {total_c} cột | Ô gộp đã tách: {self.merged_counts.get(name, 0)}) ---")
            if total_r == 0 or total_c == 0:
                out.append("(Sheet trống)")
                continue

            max_c = min(max_cols, total_c)
            headers = [
                "Row(0-idx)"] + [f"C{c}({get_column_letter(c)})" for c in range(1, max_c + 1)]
            rows_str = []

            for r_idx in range(max_r):
                row_vals = data[r_idx][:max_c]
                if any(v is not None for v in row_vals):
                    clean = [str(v).replace("\n", " ").strip()[:65]
                             if v is not None else "" for v in row_vals]
                    rows_str.append([f"R{r_idx+1}({r_idx})"] + clean)

            if rows_str:
                out.append(pd.DataFrame(
                    rows_str, columns=headers).to_markdown(index=False))
                if total_r > max_r:
                    out.append(
                        f"... (Còn {total_r - max_r} dòng bên dưới. Dùng tool `search_keyword` hoặc `run_python_analysis` để xử lý)")

        return "\n".join(out)

    def read_grid(self, sheet_name: str = "", start_row: int = 1, end_row: int = 120) -> str:
        if not sheet_name or sheet_name not in self.sheets_raw:
            sheet_name = next(iter(self.sheets_raw.keys()))
        data = self.sheets_raw[sheet_name]
        total_r = len(data)
        total_c = len(data[0]) if total_r > 0 else 0
        s_idx = max(0, start_row - 1)
        e_idx = min(total_r, end_row)

        headers = [
            "Row(0-idx)"] + [f"C{c}({get_column_letter(c)})" for c in range(1, min(25, total_c) + 1)]
        rows_str = []
        for r_idx in range(s_idx, e_idx):
            row_vals = data[r_idx][:25]
            if any(v is not None for v in row_vals):
                clean = [str(v).replace("\n", " ").strip()
                         if v is not None else "" for v in row_vals]
                rows_str.append([f"R{r_idx+1}({r_idx})"] + clean)

        if not rows_str:
            return f"Không có dữ liệu trong dòng {start_row}-{end_row}."
        return pd.DataFrame(rows_str, columns=headers).to_markdown(index=False)

    def search_keyword(self, keyword: str) -> str:
        """Tìm kiếm từ khóa trên toàn bộ WorkSpace và trả về toạ độ cụ thể."""
        kw = keyword.lower().strip()
        results = []
        for sheet_name, data in self.sheets_raw.items():
            for r_idx, row in enumerate(data):
                for c_idx, val in enumerate(row):
                    if val is not None and kw in str(val).lower():
                        results.append(
                            f"- Sheet: '{sheet_name}' | Row: {r_idx} | Col_idx: {c_idx} | Value: '{str(val).strip()}'"
                        )
                        if len(results) >= 15:
                            return "\n".join(results) + "\n... (Nhiều hơn 15 kết quả, vui lòng dùng keyword cụ thể hơn)."

        if not results:
            return f"Không tìm thấy '{keyword}' trong toàn bộ file."
        return "\n".join(results)

    def load_df(self, sheet_name: Any = None, header_rows: Any = 1) -> pd.DataFrame:
        if sheet_name is None or sheet_name == "":
            target_sheet = next(iter(self.sheets_raw.keys()))
        elif isinstance(sheet_name, int):
            target_sheet = list(self.sheets_raw.keys())[sheet_name]
        elif sheet_name in self.sheets_raw:
            target_sheet = sheet_name
        else:
            matches = [k for k in self.sheets_raw.keys() if k.lower()
                       == str(sheet_name).lower()]
            target_sheet = matches[0] if matches else next(
                iter(self.sheets_raw.keys()))

        raw_df = pd.DataFrame(self.sheets_raw[target_sheet])
        raw_df.dropna(how="all", inplace=True)
        raw_df.dropna(how="all", axis=1, inplace=True)
        raw_df.reset_index(drop=True, inplace=True)

        if raw_df.empty or header_rows is None or header_rows == 0 or header_rows is False:
            raw_df.columns = list(range(raw_df.shape[1]))
            return raw_df

        # TÌM DÒNG HEADER THỰC SỰ: Dòng đầu tiên có nhiều hơn 2 giá trị khác rỗng
        first_valid_row_idx = 0
        for idx in range(min(5, len(raw_df))):
            valid_count = raw_df.iloc[idx].count()
            if valid_count >= 2:  # Khả năng cao là Header
                first_valid_row_idx = idx
                break

        # TẠO DANH SÁCH CÁC DÒNG LÀM HEADER (h_indices)
        if header_rows == 1 and len(raw_df) > first_valid_row_idx + 1:
            r0 = [str(x).strip() for x in raw_df.iloc[first_valid_row_idx].tolist(
            ) if pd.notna(x) and str(x).strip() != ""]
            r1 = [str(x).strip() for x in raw_df.iloc[first_valid_row_idx +
                                                      1].tolist() if pd.notna(x) and str(x).strip() != ""]
            # Tự động nhận diện header 2 tầng
            if len(r0) != len(set(r0)) and all(not x.replace(".", "", 1).lstrip("-").isdigit() for x in r1):
                h_indices = [first_valid_row_idx, first_valid_row_idx + 1]
            else:
                h_indices = [first_valid_row_idx]

        elif isinstance(header_rows, int):
            # Nếu truyền vào 1 số cụ thể (VD: 2), lấy từ first_valid_row_idx tiến xuống
            h_indices = list(
                range(first_valid_row_idx, first_valid_row_idx + header_rows))

        else:
            # Nếu truyền vào list cụ thể (VD: [2, 3])
            h_indices = list(header_rows)

        max_h = max(h_indices) if h_indices else -1
        if max_h >= len(raw_df):
            return raw_df

        combined_cols, sub_cols = [], []
        for col_idx in range(raw_df.shape[1]):
            parts = []
            for r_idx in h_indices:
                val = raw_df.iloc[r_idx, col_idx]
                if pd.notna(val):
                    s_val = str(val).replace("\n", " ").strip()
                    if s_val and (not parts or parts[-1].lower() != s_val.lower()):
                        parts.append(s_val)
            full_name = " ".join(parts) if parts else f"Col_{col_idx}"
            combined_cols.append(full_name)
            sub_cols.append(parts[-1] if parts else full_name)

        data_df = raw_df.iloc[max_h + 1:].copy().reset_index(drop=True)
        data_df.columns = combined_cols

        for full_c, short_c in zip(combined_cols, sub_cols):
            if short_c != full_c and short_c not in data_df.columns and sub_cols.count(short_c) == 1:
                data_df[short_c] = data_df[full_c]

        for col in data_df.columns:  # type: ignore
            if isinstance(data_df[col], pd.Series):
                try:
                    data_df[col] = pd.to_numeric(data_df[col])
                except (ValueError, TypeError):
                    pass

        return data_df

    def run_python(self, code: str) -> str:
        buf = io.StringIO()

        # --- Các hàm trợ thủ an toàn (Helper Functions) ---
        def helper_extract_table(df, keyword_in_header):
            kw = str(keyword_in_header).lower().strip()
            for idx, row in df.iterrows():
                if any(kw in str(val).lower() for val in row.values if pd.notna(val)):
                    new_df = df.iloc[idx+1:].copy()
                    new_df.columns = [str(x).replace(
                        '\n', ' ').strip() for x in row.values]
                    new_df.reset_index(drop=True, inplace=True)
                    return new_df
            return df

        def helper_find_row(df, keyword):
            kw = str(keyword).lower().strip()
            for idx, row in df.iterrows():
                if any(kw in str(val).lower() for val in row.values if pd.notna(val)):
                    return idx
            return -1

        def helper_find_cells(df, keyword):
            kw = str(keyword).lower().strip()
            results = []
            for idx, row in df.iterrows():
                if any(kw in str(val).lower() for val in row.values if pd.notna(val)):
                    results.append(idx)
            return results

        def helper_find_col(df, keyword):
            kw = str(keyword).lower().strip()
            for col in df.columns:
                if kw in str(col).lower():
                    return col
            return None

        def helper_get_val(df, row_idx, col):
            try:
                if isinstance(col, int):
                    return df.iloc[row_idx, col]
                return df.loc[row_idx, col]
            except Exception:
                return None

        def helper_clean_num(val):
            """Gọt sạch $, %, phẩy ngàn, và xử lý dấu gạch ngang '-' thành 0.0"""
            try:
                if pd.isna(val):
                    return 0.0
                clean = str(val).replace(',', '').replace(
                    '$', '').replace('%', '').strip()
                # Thường file thống kê dùng '-' hoặc 'N/A' để chỉ số 0 hoặc không có dữ liệu
                if clean in ['-', 'N/A', 'NA', '']:
                    return 0.0
                return float(clean)
            except:
                return 0.0

        def helper_clean_col(df, col_identifier):
            """Trả về một Pandas Series chứa các số thực đã được làm sạch an toàn từ cột chỉ định"""
            try:
                if isinstance(col_identifier, int):
                    col_name = df.columns[col_identifier]
                else:
                    col_name = col_identifier
                return df[col_name].apply(helper_clean_num)
            except Exception:
                return pd.Series(dtype=float)

        def helper_parse_date(df, col_identifier):
            """Ép kiểu cột chuỗi ngày tháng sang datetime an toàn để trừ tuổi"""
            try:
                if isinstance(col_identifier, int):
                    col_name = df.columns[col_identifier]
                else:
                    col_name = col_identifier
                # Coerce để những ô lỗi ngày tháng tự động biến thành NaT
                return pd.to_datetime(df[col_name], errors='coerce')
            except Exception:
                return df[col_identifier] if not isinstance(col_identifier, int) else df.iloc[:, col_identifier]

        # --------------------------------------------------

        exec_globals = {
            "__builtins__": __builtins__,
            "pd": pd,
            "openpyxl": openpyxl,
            "file_path": self.file_path,
            "sheet_names": list(self.sheets_raw.keys()),
            "sheets_raw": self.sheets_raw,
            "load_df": self.load_df,
            "load_unmerged_df": self.load_df,
            "extract_table": helper_extract_table,
            "find_row": helper_find_row,
            "find_cells": helper_find_cells,
            "find_col": helper_find_col,
            "get_val": helper_get_val,
            "clean_num": helper_clean_num,
            "clean_col": helper_clean_col,
            "parse_date": helper_parse_date,
        }

        try:
            with contextlib.redirect_stdout(buf):
                exec(code, exec_globals)
            output = buf.getvalue()
            return output if output.strip() else "(Code chạy thành công nhưng chưa có lệnh print(). Hãy dùng print() để in kết quả!)"
        except Exception as e:
            tb_lines = traceback.format_exc().strip().splitlines()
            error_msg = "Lỗi khi thực thi Python code:\n" + \
                "\n".join(tb_lines[-6:])

            # Ý TƯỞNG MỚI: Bơm "thuốc giải" trực tiếp vào thông báo lỗi để AI tỉnh ngộ!
            if isinstance(e, KeyError):
                error_msg += "\n[HINT]: Lỗi KeyError! Tên cột bạn dùng không tồn tại. Đừng cố đoán tên cột. Hãy in df.columns.tolist() ra để xem tên chính xác, HOẶC dùng extract_table(df, 'từ_khóa') để cắt lại bảng."
            elif isinstance(e, ValueError):
                error_msg += "\n[HINT]: Lỗi ValueError! Có thể bạn đang tính toán toán học trên chuỗi chứa '$', '%', ','. BẮT BUỘC dùng clean_col(df, 'tên_cột') trước khi dùng .sum() hoặc .mean()."
            elif isinstance(e, IndexError):
                error_msg += "\n[HINT]: Lỗi IndexError! Tọa độ dòng/cột bạn trích xuất nằm ngoài phạm vi bảng."

            return error_msg


# ---------------------------------------------------------------------------
# Nhận diện số chữ số thập phân đa ngôn ngữ (đặt NGOÀI class, cấp module)
# ---------------------------------------------------------------------------
_DEC_KW = re.compile(
    r"decimal|d[ée]cimal|thập phân|nachkommastelle|dezimalstelle|casas? decimais|小数|小數|소수",
    re.IGNORECASE,
)

_NUM_WORDS = {
    1: ["1", "one", "một", "une", "un", "uno", "um", "eine", "ein",
        "一", "壹", "첫째", "第一"],
    2: ["2", "two", "hai", "deux", "dos", "dois", "zwei",
        "二", "两", "兩", "둘째", "第二"],
    3: ["3", "three", "ba", "trois", "tres", "três", "drei",
        "三", "셋째", "第三"],
    4: ["4", "four", "bốn", "quatre", "cuatro", "quatro", "vier",
        "四", "넷째", "第四"],
}


def _tok_pattern(tok: str) -> str:
    # CJK/Hangul không có khoảng trắng giữa từ -> không dùng ranh giới từ
    if any(ord(ch) >= 0x1100 for ch in tok):
        return re.escape(tok)
    return rf"(?<!\w){re.escape(tok)}(?!\w)"


_NUM_RE = re.compile(
    "|".join(
        f"(?P<n{n}_{i}>{_tok_pattern(t)})"
        for n, toks in _NUM_WORDS.items()
        for i, t in enumerate(sorted(toks, key=len, reverse=True))
    ),
    re.IGNORECASE,
)


def _num_from_match(m: "re.Match") -> int:
    return int(m.lastgroup.split("_")[0][1:])  # type: ignore


def detect_decimals(q: str):
    """Trả về số chữ số thập phân được yêu cầu trong câu hỏi, hoặc None."""
    q = q.lower()
    for kw in _DEC_KW.finditer(q):
        before = q[max(0, kw.start() - 16):kw.start()]
        after = q[kw.end():kw.end() + 16]
        ms = list(_NUM_RE.finditer(before))
        if ms:
            return _num_from_match(ms[-1])
        m = _NUM_RE.search(after)
        if m:
            return _num_from_match(m)
    return None


def normalize_final_answer(query: str, raw_text: str) -> str:
    """Chuẩn hóa đáp án hỗ trợ đa ngôn ngữ và bảo toàn định dạng thập phân."""
    if not raw_text:
        return ""

    matches = re.findall(r"<FINAL_ANSWER>(.*?)</FINAL_ANSWER>",
                         raw_text, re.DOTALL | re.IGNORECASE)
    if matches:
        ans = matches[-1].strip()
    else:
        clean_text = re.sub(r"<THOUGHT>.*?</THOUGHT>", "",
                            raw_text, flags=re.DOTALL).strip()
        lines = [line.strip()
                 for line in clean_text.splitlines() if line.strip()]
        ans = lines[-1] if lines else ""

    ans = re.sub(
        r"^(Answer|Final Answer|Đáp án|Kết quả|Output|Réponse)\s*[:=-]\s*", "", ans, flags=re.IGNORECASE).strip()
    if (ans.startswith('"') and ans.endswith('"')) or (ans.startswith("'") and ans.endswith("'")):
        ans = ans[1:-1].strip()
    if ans.startswith("`") and ans.endswith("`"):
        ans = ans[1:-1].strip()
    if ans.endswith(".") and not query.strip().endswith("."):
        ans = ans[:-1].strip()

    q_lower = query.lower()

    # 1. Xử lý câu hỏi Yes / No
    if "yes or no" in q_lower or "'yes' or 'no'" in q_lower:
        if re.search(r"\byes\b", ans, re.IGNORECASE):
            return "Yes"
        if re.search(r"\bno\b", ans, re.IGNORECASE):
            return "No"

    # 2. Nhận diện số chữ số thập phân (đa ngôn ngữ)
    target_decimals = detect_decimals(query)

    is_integer_req = any(
        kw in q_lower for kw in [
            "only the integer", "as an integer", "as integers", "nearest integer",
            "số nguyên", "nombre entier", "entier", "ganzzahl", "entero",
            "整数", "정수",
        ]
    )

    # 3. Xử lý câu trả lời có dấu phẩy ngăn cách nhiều giá trị
    if "," in ans and ("separated by a comma" in q_lower or "answer as `" in q_lower):
        parts = [p.strip() for p in ans.split(",")]

        # KIỂM TRA MỚI: Nếu đây là danh sách chữ (ví dụ rank categories), không được ép kiểu
        is_numeric_list = all(re.fullmatch(
            r"[-+]?\d*\.?\d+", p.replace("$", "").replace("%", "").strip()) for p in parts if p.strip())

        if is_numeric_list:
            formatted_parts = []
            both_decimals = target_decimals is not None and not any(
                k in q_lower for k in ["group,", "class,", "year,", "sheet name,"]
            )
            for idx, p in enumerate(parts):
                clean_p = p.replace("$", "").replace(
                    "%", "").strip()  # KHÔNG replace phẩy ở đây nữa
                try:
                    val = float(clean_p)
                    if target_decimals is not None and (both_decimals or "." in clean_p or idx == len(parts) - 1):
                        formatted_parts.append(f"{val:.{target_decimals}f}")
                    elif is_integer_req or clean_p.lstrip("-").isdigit():
                        formatted_parts.append(str(int(round(val))))
                    else:
                        formatted_parts.append(clean_p)
                except ValueError:
                    formatted_parts.append(p)
            return ", ".join(formatted_parts)
        else:
            # Nếu là mảng chữ, trả về nguyên bản chuỗi ban đầu để Exact Match không bị sai
            return ans

    # 4. Xử lý câu trả lời là một giá trị số đơn lẻ
    clean_single = ans.replace(",", "").replace(
        "$", "").replace("%", "").strip()
    if re.fullmatch(r"[-+]?\d+(\.\d+)?", clean_single):
        try:
            val = float(clean_single)
            if target_decimals is not None:
                return f"{val:.{target_decimals}f}"
            if is_integer_req:
                return str(int(round(val)))
            if "." in clean_single:
                return clean_single
            return str(int(round(val)))
        except ValueError:
            pass

    return ans


class QuerySolver:
    def __init__(self, api_key: str, model: str, base_url: str):
        self.client = OpenAI(
            base_url=base_url, api_key=api_key, timeout=60.0, max_retries=3)
        self.model = model
        self.workspaces: Dict[str, ExcelWorkspace] = {}
        # Thread-lock cho bộ nhớ đệm (Cache) để an toàn khi nhiều luồng truy cập RAM
        self._ws_lock = threading.Lock()

    def get_workspace(self, xlsx_dir: str, file_name: str) -> ExcelWorkspace:
        key = f"{xlsx_dir}::{file_name}"
        with self._ws_lock:
            if key not in self.workspaces:
                if len(self.workspaces) >= 5:
                    del self.workspaces[next(iter(self.workspaces.keys()))]
                self.workspaces[key] = ExcelWorkspace(xlsx_dir, file_name)
            return self.workspaces[key]

    def process_query(self, q_id: Any, xlsx_dir: str, file_name: str, query: str, log_cb=None) -> Dict[str, Any]:
        if log_cb is None:
            def log_cb(msg): return print(msg, file=sys.stderr)

        ws = self.get_workspace(xlsx_dir, file_name)
        overview = ws.get_overview()

        messages: List[Dict[str, Any]] = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": f"Question: {query}\n\n{overview}"},
        ]

        total_p_tokens = 0
        total_c_tokens = 0
        total_cost = 0.0
        final_answer = ""

        for step in range(6):
            response = self.client.chat.completions.create(
                model=self.model,
                messages=messages,  # type: ignore
                tools=TOOLS_SCHEMA,  # type: ignore
                tool_choice="auto",
                temperature=0.0,
            )

            msg = response.choices[0].message

            if response.usage:
                p_tok = response.usage.prompt_tokens
                c_tok = response.usage.completion_tokens
                total_p_tokens += p_tok
                total_c_tokens += c_tok
                step_cost = (p_tok / 1_000_000 * PRICE_1M_PROMPT) + \
                    (c_tok / 1_000_000 * PRICE_1M_COMPLETION)
                total_cost += step_cost
                log_cb(
                    f"  [Step {step+1}] Tokens: Prompt={p_tok}, Comp={c_tok} | Chi phí: ${step_cost:.6f}")

            if msg.content:
                for t in re.findall(r"<THOUGHT>(.*?)</THOUGHT>", msg.content, re.DOTALL | re.IGNORECASE):
                    log_cb(f"  [THOUGHT]: {t.strip()[:250]}")

            if msg.tool_calls:
                messages.append({
                    "role": "assistant",
                    "content": msg.content or "",
                    "tool_calls": [
                        {
                            "id": tc.id,
                            "type": "function",
                            "function": {"name": tc.function.name, "arguments": tc.function.arguments},
                        }
                        for tc in msg.tool_calls
                        if tc.type == "function"
                    ],
                })

                for tc in msg.tool_calls:
                    if tc.type != "function":
                        continue
                    t_name = tc.function.name
                    try:
                        t_args = json.loads(tc.function.arguments)
                    except Exception:
                        t_args = {}

                    log_cb(f"  [CALL]: {t_name}")
                    if t_name == "run_python_analysis":
                        tool_out = ws.run_python(t_args.get("python_code", ""))
                    elif t_name == "read_sheet_grid":
                        tool_out = ws.read_grid(
                            t_args.get("sheet_name", ""),
                            int(t_args.get("start_row", 1)),
                            int(t_args.get("end_row", 120)),
                        )
                    elif t_name == "search_keyword":
                        tool_out = ws.search_keyword(t_args.get("keyword", ""))
                    else:
                        tool_out = f"Tool '{t_name}' không tồn tại."

                    log_cb(f"  [OUTPUT]: {str(tool_out).strip()[:250]}")
                    messages.append({
                        "role": "tool",
                        "tool_call_id": tc.id,
                        "name": t_name,
                        "content": str(tool_out)[:6000],
                    })
            else:
                messages.append(
                    {"role": "assistant", "content": msg.content or ""})
                final_answer = normalize_final_answer(query, msg.content or "")
                if final_answer:
                    break
                messages.append({
                    "role": "user",
                    "content": "Hãy xuất đáp án cuối cùng vào trong cặp thẻ <FINAL_ANSWER>...</FINAL_ANSWER>.",
                })

        if not final_answer:
            messages.append({
                "role": "user",
                "content": "Dựa vào kết quả trên, xuất ngay đáp án duy nhất trong thẻ <FINAL_ANSWER>...</FINAL_ANSWER>.",
            })
            resp = self.client.chat.completions.create(
                model=self.model, messages=messages, temperature=0.0)  # type: ignore
            if resp.usage:
                total_p_tokens += resp.usage.prompt_tokens
                total_c_tokens += resp.usage.completion_tokens
            final_answer = normalize_final_answer(
                query, resp.choices[0].message.content or "")

        return {"answer": final_answer, "cost": total_cost, "tokens": total_p_tokens + total_c_tokens}


def main():
    start_time = time.time()  # Bắt đầu đếm giờ tổng thời gian xử lý

    parser = argparse.ArgumentParser(
        description="AI26C0084 - Giải pháp Hỏi đáp trên Bảng tính Excel")
    parser.add_argument("--xlsx-dir", type=str, default="input",
                        help="Thư mục chứa các tệp .xlsx")
    parser.add_argument("--queries", type=str,
                        default="input/queries.json", help="Đường dẫn tệp queries.json")
    parser.add_argument(
        "--output", type=str, default="output/submission.csv", help="Đường dẫn tệp kết quả")
    parser.add_argument("--openrouter-key", type=str,
                        default=os.getenv("OPENROUTER_API_KEY", ""), help="Khóa API")
    parser.add_argument("--openrouter-model", type=str,
                        default="qwen/qwen3.7-flash", help="Model ID")
    parser.add_argument("--openrouter-base-url", type=str,
                        default="https://openrouter.ai/api/v1", help="Base URL")
    args = parser.parse_args()

    if not args.openrouter_key:
        print("Lỗi: Thiếu --openrouter-key hoặc biến môi trường OPENROUTER_API_KEY.", file=sys.stderr)
        sys.exit(1)

    out_dir = os.path.dirname(os.path.abspath(args.output))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    with open(args.queries, "r", encoding="utf-8") as f:
        queries_list = json.load(f)

    solver = QuerySolver(api_key=args.openrouter_key,
                         model=args.openrouter_model, base_url=args.openrouter_base_url)
    results = []
    total_cost = 0.0
    total_tokens = 0

    # Khóa luồng (Lock) cho việc in ra màn hình
    print_lock = threading.Lock()

    # Hàm xử lý từng câu hỏi cho luồng (thread)
    def process_item(item):
        q_id = item.get("id")
        file_name = item.get("file_name") or item.get("file name")
        query_text = item.get("query", "")

        # Buffer: Lưu tạm toàn bộ dòng log của ID này vào một list
        logs_buffer = []
        logs_buffer.append(
            f"\n{'='*55}\n>>> Đang xử lý ID: {q_id} | File: {file_name}")

        try:
            # Truyền logs_buffer.append vào log_cb để Gom log thay vì In trực tiếp
            res = solver.process_query(
                q_id, args.xlsx_dir, file_name, query_text, log_cb=logs_buffer.append)

            answer = res["answer"]
            cost = res["cost"]
            tokens = res["tokens"]

            logs_buffer.append(
                f"[THỐNG KÊ ID {q_id}] Chi phí: ${cost:.6f} | Tokens: {tokens}")
            logs_buffer.append(f"<<< ĐÁP ÁN ID {q_id}: {answer}")

            # Xuất log ra màn hình: Đợi lấy khóa lock xong mới in "một cục" ra console
            with print_lock:
                print("\n".join(logs_buffer), file=sys.stderr, flush=True)

            return {"id": q_id, "answer": answer, "cost": cost, "tokens": tokens}
        except Exception as e:
            logs_buffer.append(f"[LỖI ID {q_id}]: {str(e)}")
            with print_lock:
                print("\n".join(logs_buffer), file=sys.stderr, flush=True)

            return {"id": q_id, "answer": "", "cost": 0.0, "tokens": 0}

    # Giới hạn 3 luồng chạy song song để đảm bảo không bị tràn RAM của BTC
    max_workers = 3
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(process_item, item)
                   for item in queries_list]

        for future in concurrent.futures.as_completed(futures):
            res_data = future.result()
            # 1. CẬP NHẬT: Lưu thêm cost và tokens vào mảng kết quả chung để xuất ra JSON dự phòng
            results.append({
                "id": res_data["id"],
                "answer": res_data["answer"]
            })
            total_cost += res_data["cost"]
            total_tokens += res_data["tokens"]

    # Sắp xếp lại thứ tự theo id tăng dần theo quy định BTC
    try:
        results = sorted(results, key=lambda x: int(x["id"]))
    except ValueError:
        results = sorted(results, key=lambda x: str(x["id"]))

    # 2. GHI FILE SUBMISSION.CSV CHUẨN ĐỂ NỘP BÀI
    # Lọc lại chỉ lấy 2 cột id và answer để ghi vào CSV (Bắt buộc theo luật BTC)
    csv_data = [{"id": r["id"], "answer": r["answer"]} for r in results]
    with open(args.output, "w", newline="", encoding="utf-8") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=["id", "answer"])
        writer.writeheader()
        writer.writerows(csv_data)

    # 3. GHI FILE ANSWERS.JSON DỰ PHÒNG
    # Đặt file json đồng cấp (cùng thư mục) với tệp submission.csv
    answers_json_path = os.path.join(
        out_dir, "answers.json") if out_dir else "answers.json"
    try:
        with open(answers_json_path, "w", encoding="utf-8") as jsonfile:
            # Ghi toàn bộ dữ liệu (id, answer, cost, tokens) vào file JSON
            json.dump(results, jsonfile, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Lỗi khi lưu tệp {answers_json_path}: {e}", file=sys.stderr)

    # Kết thúc đếm giờ và định dạng thời gian
    elapsed_time = time.time() - start_time
    mins = int(elapsed_time // 60)
    secs = elapsed_time % 60

    print(f"\n{'='*55}", file=sys.stderr)
    print(
        f"HOÀN TẤT! Đã xử lý: {len(results)} câu | Thời gian: {mins}p {secs:.2f}s", file=sys.stderr)
    print(
        f"Tổng Token: {total_tokens} | Tổng Chi phí: ${total_cost:.6f}", file=sys.stderr)
    print(
        f"-> File submission.csv : {os.path.abspath(args.output)}", file=sys.stderr)
    print(
        f"-> File answers.json  : {os.path.abspath(answers_json_path)}", file=sys.stderr)
    print(f"{'='*55}\n", file=sys.stderr)


if __name__ == "__main__":
    main()
