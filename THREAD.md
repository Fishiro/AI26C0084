# PHÂN TÍCH LUỒNG THỰC THI MAIN.PY

## 1. Tổng quan kiến trúc ứng dụng

Dự án AI26C0084 là hệ thống hỏi đáp tự động trên dữ liệu Excel. Người dùng cung cấp danh sách câu hỏi trong file JSON và các workbook Excel tương ứng. Chương trình sử dụng mô hình ngôn ngữ lớn (LLM) thông qua OpenRouter để phân tích câu hỏi, đọc dữ liệu Excel, thực hiện phép tính khi cần và trả về đáp án theo định dạng yêu cầu.

Mã nguồn chính nằm trong file `main.py`.

### Luồng tổng thể

```text
Bắt đầu chạy main.py
        |
        v
Import thư viện và nạp biến môi trường
        |
        v
Khởi tạo SYSTEM_PROMPT và TOOLS_SCHEMA
        |
        v
Định nghĩa ExcelWorkspace
        |
        v
Định nghĩa các hàm chuẩn hóa đáp án
        |
        v
Định nghĩa QuerySolver
        |
        v
Gọi hàm main()
        |
        v
Đọc tham số dòng lệnh
        |
        v
Kiểm tra API Key và tạo thư mục output
        |
        v
Đọc input/queries.json
        |
        v
Khởi tạo QuerySolver và các biến thống kê
        |
        v
Tạo ThreadPoolExecutor với tối đa 3 worker
        |
        v
process_item() xử lý từng câu hỏi
        |
        v
QuerySolver.process_query()
        |
        v
ExcelWorkspace đọc workbook và tạo Overview
        |
        v
Gửi câu hỏi + Overview đến LLM
        |
        v
LLM trả lời trực tiếp hoặc gọi Tool
        |
        +------------------------------+
        |                              |
        v                              v
run_python_analysis               read_sheet_grid
        |                              |
        +---------------+--------------+
                        |
                        v
               Nhận kết quả Tool
                        |
                        v
               LLM tiếp tục suy luận
                        |
                        v
               Chuẩn hóa đáp án
                        |
                        v
               Trả answer, cost, tokens
                        |
                        v
               Thu thập kết quả các worker
                        |
                        v
               Sắp xếp theo ID
                        |
                        v
               Ghi submission.csv
                        |
                        v
               Ghi answers.json
                        |
                        v
               In thời gian, token, chi phí
                        |
                        v
                     KẾT THÚC
```

Các hàm trong `main.py` được định nghĩa trước khi chương trình thực hiện `main()`. Khi chạy file trực tiếp, Python thực hiện các lệnh ở cấp module, đăng ký các lớp và hàm, sau đó mới gọi `main()` thông qua:

```python
if __name__ == "__main__":
    main()
```

Vì vậy, cần phân biệt thứ tự định nghĩa hàm trong mã nguồn với thứ tự các hàm thực sự được gọi khi chương trình chạy.

---

## 2. Giai đoạn 1: Import thư viện và khởi tạo biến toàn cục

### 2.1. Import thư viện

Ở đầu file, chương trình nạp các thư viện cần thiết:

```python
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
```

Vai trò của các thư viện:

| Thư viện             | Chức năng                                                                   |
| -------------------- | --------------------------------------------------------------------------- |
| `os`                 | Xử lý đường dẫn, tìm file, tạo thư mục                                      |
| `sys`                | Ghi log ra stderr và kết thúc chương trình                                  |
| `io`                 | Tạo bộ đệm để thu nhận kết quả `print()`                                    |
| `re`                 | Sử dụng Regex nhận diện và chuẩn hóa đáp án                                 |
| `csv`                | Ghi kết quả cuối cùng thành file CSV                                        |
| `json`               | Đọc câu hỏi và ghi kết quả JSON                                             |
| `argparse`           | Đọc tham số khi chạy chương trình                                           |
| `traceback`          | Thu thập thông tin lỗi Python                                               |
| `contextlib`         | Chuyển hướng stdout để thu kết quả của code thực thi                        |
| `threading`          | Khóa đồng bộ khi truy cập cache hoặc in log                                 |
| `concurrent.futures` | Chạy các câu hỏi song song                                                  |
| `time`               | Đo tổng thời gian chạy                                                      |
| `openpyxl`           | Đọc và xử lý các workbook Excel                                             |
| `pandas`             | Chuyển bảng Excel thành DataFrame để phân tích                              |
| `dotenv`             | Đọc các biến môi trường từ file `.env`                                      |
| `OpenAI`             | Gửi yêu cầu chat completion đến API tương thích OpenAI, ở đây là OpenRouter |

### 2.2. Nạp biến môi trường

```python
load_dotenv()
```

Lệnh trên nạp cấu hình từ file `.env` vào môi trường của tiến trình, nếu có.

Ví dụ, chương trình có thể sử dụng biến môi trường:

```text
OPENROUTER_API_KEY=your_api_key
```

Giá trị API Key sẽ được đọc khi khởi tạo tham số dòng lệnh trong hàm `main()`.

### 2.3. Các hằng số giá dự phòng

```python
PRICE_1M_PROMPT = 0.030
PRICE_1M_COMPLETION = 0.130
PRICE_1M_CACHED = 0.0060
```

Đơn vị USD / 1 triệu token, theo trang giá OpenRouter của `qwen/qwen3.7-flash`. Các hằng số này chỉ dùng khi API không trả về chi phí thực; giá có thể đổi theo provider nên cần cập nhật khi OpenRouter đổi giá. Cách tính xem mục 6.3.

---

## 3. Giai đoạn 2: Khởi tạo SYSTEM_PROMPT và TOOLS_SCHEMA

### 3.1. SYSTEM_PROMPT

`SYSTEM_PROMPT` là chuỗi hướng dẫn lớn dành cho LLM, được khai báo ngay sau các hằng số chi phí.

Nhiệm vụ của prompt là hướng dẫn mô hình cách giải quyết các câu hỏi liên quan đến Excel, đặc biệt là dữ liệu thống kê có cấu trúc phức tạp.

Một số yêu cầu chính:

1. Phân tích câu hỏi và dữ liệu bảng tính.
2. Sử dụng Overview khi dữ liệu đã đủ rõ để trả lời.
3. Gọi công cụ Python khi phải tính toán hoặc truy xuất dữ liệu chi tiết.
4. Dùng các hàm trợ thủ như `load_df()`, `extract_table()`, `find_row()`, `find_col()`, `get_val()`, `clean_num()` và `clean_col()`.
5. Xử lý các tình huống có tiêu đề nhiều tầng, ô gộp, dữ liệu dạng số được lưu dưới dạng chuỗi và ngày tháng.
6. Trả về đáp án trong thẻ `<FINAL_ANSWER>...</FINAL_ANSWER>`.

Ví dụ về định dạng kết quả được yêu cầu:

```text
<FINAL_ANSWER>2023</FINAL_ANSWER>
```

`SYSTEM_PROMPT` là biến toàn cục được truyền vào hội thoại mỗi khi `QuerySolver.process_query()` bắt đầu xử lý một câu hỏi.

### 3.2. TOOLS_SCHEMA

Chương trình khai báo danh sách các công cụ mà LLM có thể yêu cầu thực thi.

```python
TOOLS_SCHEMA = [
    # Tool 1: run_python_analysis
    # Tool 2: read_sheet_grid
    # Tool 3: search_keyword
]
```

Ba công cụ bao gồm:

| Tên Tool              | Chức năng                              |
| --------------------- | -------------------------------------- |
| `run_python_analysis` | Chạy Python để phân tích dữ liệu Excel |
| `read_sheet_grid`     | Đọc chi tiết một vùng dòng của sheet   |
| `search_keyword`      | Tìm từ khóa trên toàn bộ các sheet     |

Đối với `run_python_analysis`, tham số quan trọng là:

```python
{
    "python_code": "Đoạn code Python mà mô hình yêu cầu thực thi"
}
```

Đối với `read_sheet_grid`, mô hình có thể chỉ định:

```python
{
    "sheet_name": "Sheet1",
    "start_row": 1,
    "end_row": 120
}
```

Đối với `search_keyword`:

```python
{
    "keyword": "Total"
}
```

`TOOLS_SCHEMA` chỉ mô tả giao diện các công cụ với API. Bản thân nó không thực thi các hàm. Việc thực thi thật sự được điều phối bởi `QuerySolver.process_query()`.

---

## 4. Giai đoạn 3: Lớp ExcelWorkspace

`ExcelWorkspace` chịu trách nhiệm đọc workbook Excel, xử lý ô gộp, lưu dữ liệu vào RAM và cung cấp các phương thức để truy xuất bảng.

Đây là lớp nền tảng cho mọi thao tác dữ liệu.

### 4.1. Hàm khởi tạo `__init__()`

Mã nguồn:

```python
def __init__(self, xlsx_dir: str, file_name: str):
    self.file_path = self._resolve_path(xlsx_dir, file_name)
    self.sheets_raw: Dict[str, List[List[Any]]] = {}
    self.merged_counts: Dict[str, int] = {}
    self._load_workbook()
```

Các biến được tạo:

- `xlsx_dir`: thư mục chứa các file Excel.
- `file_name`: tên file cần xử lý.
- `self.file_path`: đường dẫn thật đến file sau khi được tìm thấy.
- `self.sheets_raw`: dictionary lưu dữ liệu thô của tất cả các sheet.
- `self.merged_counts`: dictionary lưu số vùng ô gộp ban đầu của từng sheet.

Ngay sau khi các biến này được khởi tạo, chương trình gọi:

```python
self._load_workbook()
```

Đây là bước thực sự đọc workbook và nạp dữ liệu vào bộ nhớ.

### 4.2. Hàm `_resolve_path()`

```python
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

    # Tiếp tục tìm trong các thư mục con.
    # Bỏ qua thư mục venv và .venv.

    raise FileNotFoundError(...)
```

Hàm nhận hai đầu vào:

- `xlsx_dir`: thư mục chứa dữ liệu.
- `file_name`: tên hoặc đường dẫn tương đối của workbook.

Biến `base` chứa tên file đã bỏ phần đường dẫn.

Thuật toán thực hiện:

1. Tạo danh sách những đường dẫn có thể chứa file.
2. Kiểm tra từng đường dẫn bằng `os.path.isfile()`.
3. Nếu chưa tìm thấy, tìm tiếp qua các thư mục con bằng `os.walk()`.
4. Bỏ qua các thư mục có tên `venv` hoặc `.venv`.
5. Nếu vẫn không tìm thấy thì phát sinh `FileNotFoundError`.

Kết quả trả về là một chuỗi đường dẫn được gán cho `self.file_path`.

### 4.3. Hàm `_load_workbook()`

```python
def _load_workbook(self):
    wb = openpyxl.load_workbook(
        self.file_path,
        data_only=True
    )

    for name in wb.sheetnames:
        sheet = wb[name]

        self.merged_counts[name] = len(
            sheet.merged_cells.ranges
        )

        for m_range in list(sheet.merged_cells.ranges):
            top_val = sheet.cell(
                m_range.min_row,
                m_range.min_col
            ).value

            sheet.unmerge_cells(str(m_range))

            for r in range(
                m_range.min_row,
                m_range.max_row + 1
            ):
                for c in range(
                    m_range.min_col,
                    m_range.max_col + 1
                ):
                    cell = sheet.cell(r, c)

                    if isinstance(cell, Cell):
                        cell.value = top_val

        data = []

        for r in range(1, sheet.max_row + 1):
            data.append([
                sheet.cell(r, c).value
                for c in range(1, sheet.max_column + 1)
            ])

        self.sheets_raw[name] = data

    wb.close()
```

Hàm này thực hiện ba nhiệm vụ chính.

**Nhiệm vụ 1: Đọc workbook**

```python
wb = openpyxl.load_workbook(
    self.file_path,
    data_only=True
)
```

`data_only=True` yêu cầu đọc giá trị được lưu của các ô công thức thay vì lấy chuỗi công thức làm giá trị trả về.

**Nhiệm vụ 2: Tách các ô gộp**

Các bảng thống kê thường gộp một ô trên nhiều cột hoặc dòng. Nếu giữ nguyên cấu trúc đó, nhiều ô có thể mang giá trị `None`.

Chương trình lấy giá trị của ô trên cùng, bên trái:

```python
top_val = sheet.cell(
    m_range.min_row,
    m_range.min_col
).value
```

Sau đó tách vùng gộp và điền `top_val` vào các ô thành phần.

Nhờ đó, những ô vốn nằm trong cùng một vùng gộp được chuyển thành các ô có giá trị nhất quán.

**Nhiệm vụ 3: Lưu nội dung sheet vào RAM**

```python
data = []

for r in range(1, sheet.max_row + 1):
    data.append([
        sheet.cell(r, c).value
        for c in range(1, sheet.max_column + 1)
    ])

self.sheets_raw[name] = data
```

Mỗi sheet được chuyển thành danh sách hai chiều:

```text
sheets_raw
    |
    +-- "Sheet1"
    |      +-- [giá trị dòng 1]
    |      +-- [giá trị dòng 2]
    |
    +-- "Sheet2"
           +-- [giá trị dòng 1]
           +-- [giá trị dòng 2]
```

Sau khi hoàn thành toàn bộ sheet, workbook được đóng bằng `wb.close()`.

### 4.4. Hàm `get_overview()`

```python
def get_overview(
    self,
    global_max_rows: int = 150,
    max_cols: int = 20
) -> str:
```

Mục đích: tạo bản xem trước của workbook để gửi cho LLM.

Các biến chính:

- `global_max_rows = 150`: ngân sách dòng tham chiếu.
- `max_cols = 20`: giới hạn số cột được hiển thị trên mỗi sheet.
- `out`: danh sách chuỗi đầu ra.
- `total_sheets`: số sheet của workbook.
- `base_budget_per_sheet`: ngân sách dòng cơ sở cho từng sheet.
- `total_r`, `total_c`: số dòng và số cột.
- `max_r`, `max_c`: giới hạn thực tế của vùng hiển thị.
- `rows_str`: dữ liệu đã định dạng để tạo bảng Markdown.

Cách tính ngân sách:

```python
base_budget_per_sheet = max(
    15,
    global_max_rows // total_sheets
)

max_r = min(
    total_r,
    max(
        base_budget_per_sheet,
        35 if total_sheets == 1 else 20
    )
)
```

Ví dụ, nếu workbook có một sheet và dữ liệu không vượt quá giới hạn, Overview có thể hiển thị tối đa 150 dòng.

Nếu có nhiều sheet, số dòng hiển thị sẽ được phân bổ dựa trên tổng số sheet, đồng thời áp dụng giới hạn tối thiểu trong đoạn code.

Dữ liệu được chuyển thành bảng Markdown bằng:

```python
pd.DataFrame(
    rows_str,
    columns=headers
).to_markdown(index=False)
```

Giá trị trong các ô còn được rút gọn khi tạo Overview, chẳng hạn mỗi giá trị có thể bị giới hạn khoảng 65 ký tự.

Cuối cùng, hàm trả về một chuỗi:

```python
return "\n".join(out)
```

Chuỗi này được truyền cho LLM ở đầu `QuerySolver.process_query()`.

### 4.5. Hàm `read_grid()`

```python
def read_grid(
    self,
    sheet_name: str = "",
    start_row: int = 1,
    end_row: int = 120
) -> str:
```

Hàm được sử dụng khi mô hình cần đọc chi tiết một vùng dữ liệu.

Các biến:

- `sheet_name`: tên sheet cần đọc.
- `start_row`: dòng bắt đầu, đánh số từ 1.
- `end_row`: dòng kết thúc.
- `data`: danh sách dữ liệu của sheet.
- `s_idx`: chỉ số dòng bắt đầu theo quy ước Python, bắt đầu từ 0.
- `e_idx`: giới hạn dòng kết thúc.
- `rows_str`: dữ liệu được chuyển thành bảng văn bản.

Nếu tên sheet bị bỏ trống hoặc không hợp lệ, hàm chọn sheet đầu tiên.

Phần chuyển đổi chỉ số dòng:

```python
s_idx = max(0, start_row - 1)
e_idx = min(total_r, end_row)
```

Chương trình lấy dữ liệu trong vùng được yêu cầu, giới hạn tối đa 25 cột, bỏ qua các dòng hoàn toàn trống rồi trả về bảng Markdown.

Điểm khác với `get_overview()` là `read_grid()` tập trung đọc vùng dòng chi tiết thay vì tạo bản xem trước cho toàn bộ workbook.

### 4.6. Hàm `search_keyword()`

```python
def search_keyword(self, keyword: str) -> str:
    kw = keyword.lower().strip()
    results = []

    for sheet_name, data in self.sheets_raw.items():
        for r_idx, row in enumerate(data):
            for c_idx, val in enumerate(row):
                if val is not None and kw in str(val).lower():
                    results.append(
                        f"- Sheet: '{sheet_name}' | "
                        f"Row: {r_idx} | "
                        f"Col_idx: {c_idx} | "
                        f"Value: '{str(val).strip()}'"
                    )

                if len(results) >= 15:
                    return "\n".join(results) + \
                        "\n... (Nhiều hơn 15 kết quả...)"

    if not results:
        return f"Không tìm thấy '{keyword}' trong toàn bộ file."

    return "\n".join(results)
```

Hàm tìm chuỗi con không phân biệt chữ hoa và chữ thường trên toàn bộ workbook.

Các biến:

- `kw`: từ khóa đã chuẩn hóa.
- `sheet_name`: tên sheet hiện tại.
- `r_idx`: chỉ số dòng, bắt đầu từ 0.
- `c_idx`: chỉ số cột, bắt đầu từ 0.
- `val`: giá trị ô hiện tại.
- `results`: danh sách kết quả tìm được.

Hàm dừng khi có đủ 15 kết quả.

Kết quả trả về cho mô hình biết giá trị nằm ở sheet nào, dòng nào, cột nào.

### 4.7. Hàm `load_df()`

```python
def load_df(
    self,
    sheet_name: Any = None,
    header_rows: Any = 1
) -> pd.DataFrame:
```

Hàm chuyển dữ liệu thô trong `self.sheets_raw` thành Pandas DataFrame để thuận tiện cho việc tính toán.

Đầu tiên, hàm xác định sheet cần sử dụng:

- Nếu bỏ trống, chọn sheet đầu tiên.
- Nếu truyền số nguyên, chọn sheet theo vị trí trong danh sách sheet.
- Nếu truyền tên hợp lệ, chọn đúng sheet đó.
- Nếu tên không trùng chính xác, thử so sánh không phân biệt hoa thường; nếu vẫn không tìm thấy thì dùng sheet đầu tiên.

Tiếp theo:

```python
raw_df = pd.DataFrame(
    self.sheets_raw[target_sheet]
)

raw_df.dropna(how="all", inplace=True)
raw_df.dropna(how="all", axis=1, inplace=True)
raw_df.reset_index(drop=True, inplace=True)
```

Ý nghĩa:

- Chuyển dữ liệu thô thành DataFrame.
- Loại bỏ các dòng hoàn toàn rỗng.
- Loại bỏ các cột hoàn toàn rỗng.
- Đánh lại chỉ số dòng.

Sau đó, hàm tìm dòng header thực tế bằng cách xem xét tối đa năm dòng đầu tiên. Dòng đầu tiên có từ hai giá trị khác rỗng trở lên được xem là ứng viên header.

```python
for idx in range(min(5, len(raw_df))):
    valid_count = raw_df.iloc[idx].count()

    if valid_count >= 2:
        first_valid_row_idx = idx
        break
```

Tiếp theo, chương trình xác định header gồm một hay nhiều dòng dựa trên tham số `header_rows` và một số điều kiện nhận diện header hai tầng.

Tên cột được ghép từ các dòng header:

```python
full_name = " ".join(parts) if parts else f"Col_{col_idx}"
```

Ví dụ:

```text
Header tầng 1: Revenue
Header tầng 2: 2023

Tên cột sau khi ghép: Revenue 2023
```

Phần dữ liệu nằm dưới header được tách ra:

```python
data_df = raw_df.iloc[max_h + 1:].copy()
data_df.reset_index(drop=True, inplace=True)
data_df.columns = combined_cols
```

Cuối cùng, hàm cố gắng chuyển các cột có dữ liệu số sang dạng số bằng `pd.to_numeric()`. Những cột không thể chuyển đổi được giữ nguyên.

Kết quả trả về là DataFrame đã được làm sạch cơ bản.

### 4.8. Hàm `run_python()`

Đây là một trong những phương thức quan trọng nhất của `ExcelWorkspace`.

```python
def run_python(self, code: str) -> str:
    buf = io.StringIO()
```

Đầu vào `code` là đoạn Python do mô hình yêu cầu thực thi.

Hàm có ba nhóm nhiệm vụ:

1. Tạo các hàm trợ thủ để phân tích Excel.
2. Thực thi đoạn code và thu kết quả in ra.
3. Bắt lỗi và trả thông báo để mô hình có thể sửa code.

**Các hàm trợ thủ được định nghĩa bên trong `run_python()`:**

| Hàm                                           | Chức năng                                                                     |
| --------------------------------------------- | ----------------------------------------------------------------------------- |
| `helper_extract_table(df, keyword_in_header)` | Tìm dòng có từ khóa rồi dùng dòng đó làm header mới                           |
| `helper_find_row(df, keyword)`                | Trả về chỉ số dòng đầu tiên chứa từ khóa                                      |
| `helper_find_cells(df, keyword)`              | Trả về danh sách chỉ số các dòng chứa từ khóa                                 |
| `helper_find_col(df, keyword)`                | Tìm tên cột có chứa từ khóa                                                   |
| `helper_get_val(df, row_idx, col)`            | Lấy giá trị theo chỉ số dòng và tên hoặc chỉ số cột                           |
| `helper_clean_num(val)`                       | Chuyển giá trị thành số thực sau khi loại bỏ các ký tự như `$`, `%`, dấu phẩy |
| `helper_clean_col(df, col_identifier)`        | Làm sạch toàn bộ một cột thành các giá trị số                                 |
| `helper_parse_date(df, col_identifier)`       | Chuyển đổi một cột thành kiểu ngày tháng                                      |

Ví dụ hàm làm sạch số:

```python
def helper_clean_num(val):
    try:
        if pd.isna(val):
            return 0.0

        clean = str(val).replace(
            ',', ''
        ).replace(
            '$', ''
        ).replace(
            '%', ''
        ).strip()

        if clean in ['-', 'N/A', 'NA', '']:
            return 0.0

        return float(clean)

    except:
        return 0.0
```

Ví dụ, chuỗi `"1,250"` được chuyển thành `1250.0`, còn `"-"` được chuyển thành `0.0`.

**Tạo môi trường thực thi**

```python
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
```

Đối tượng `exec_globals` cung cấp các biến và hàm mà đoạn code do mô hình sinh ra có thể truy cập.

Ví dụ:

```python
df = load_df()
r = find_row(df, "Mexican")
c = find_col(df, "2023")

if r != -1 and c:
    val = get_val(df, r, c)
    print(clean_num(val))
```

Đoạn code trên sử dụng các hàm trợ thủ đã được đăng ký trong `exec_globals`.

**Thực thi code và thu kết quả**

```python
try:
    with contextlib.redirect_stdout(buf):
        exec(code, exec_globals)

    output = buf.getvalue()

    return output if output.strip() else (
        "(Code chạy thành công nhưng chưa có lệnh print().)"
    )
```

Ở đây:

- `exec()` thực thi đoạn code.
- `contextlib.redirect_stdout(buf)` chuyển hướng kết quả `print()` vào bộ đệm.
- `buf.getvalue()` lấy toàn bộ đầu ra thành chuỗi `output`.
- Hàm trả về chuỗi kết quả để LLM tiếp tục xử lý.

Nếu code xảy ra lỗi, hàm lấy traceback và bổ sung hướng dẫn sửa lỗi.

Ví dụ, nếu phát sinh `KeyError`, chương trình hướng dẫn mô hình kiểm tra lại tên cột bằng `df.columns.tolist()` hoặc dùng `extract_table()` để thiết lập lại header.

**Lưu ý:** `exec()` chạy code ngay trong tiến trình Python hiện tại với các built-in được cung cấp, không phải một sandbox biệt lập hoàn toàn.

---

## 5. Giai đoạn 4: Các hàm chuẩn hóa đáp án ở cấp module

Sau lớp `ExcelWorkspace`, mã nguồn định nghĩa các biểu thức Regex và các hàm xử lý định dạng đáp án.

Các hàm này không trực tiếp đọc workbook; chúng xử lý yêu cầu về định dạng và đáp án trả về từ LLM.

### 5.1. `_DEC_KW`

```python
_DEC_KW = re.compile(
    r"decimal|d[ée]cimal|thập phân|nachkommastelle|dezimalstelle|casas? decimais|小数|小數|소수",
    re.IGNORECASE,
)
```

Biểu thức này nhận diện các từ liên quan đến số chữ số thập phân trong nhiều ngôn ngữ.

Ví dụ:

- `two decimal places`
- `hai chữ số thập phân`
- `deux décimales`

### 5.2. `_NUM_WORDS`

```python
_NUM_WORDS = {
    1: [...],
    2: [...],
    3: [...],
    4: [...]
}
```

Dictionary này ánh xạ các cách viết số từ 1 đến 4 trong nhiều ngôn ngữ sang số nguyên tương ứng.

### 5.3. `_tok_pattern()`

```python
def _tok_pattern(tok: str) -> str:
    if any(ord(ch) >= 0x1100 for ch in tok):
        return re.escape(tok)

    return rf"(?<!\w){re.escape(tok)}(?!\w)"
```

Hàm xây dựng mẫu Regex cho từng cách viết số.

Đối với ký tự thuộc nhóm CJK hoặc Hangul, hàm không sử dụng ranh giới từ thông thường, giúp nhận diện các chuỗi không có dấu cách giữa các từ.

### 5.4. `_NUM_RE`

```python
_NUM_RE = re.compile(
    "|".join(
        f"(?P<n{n}_{i}>{_tok_pattern(t)})"
        for n, toks in _NUM_WORDS.items()
        for i, t in enumerate(
            sorted(toks, key=len, reverse=True)
        )
    ),
    re.IGNORECASE,
)
```

Biến `_NUM_RE` tổng hợp các mẫu số thành một Regex lớn để tìm số trong câu hỏi.

Các nhóm có tên như `n2_0` giúp xác định số vừa được nhận diện là 1, 2, 3 hay 4.

### 5.5. `_num_from_match()`

```python
def _num_from_match(m: "re.Match") -> int:
    return int(m.lastgroup.split("_")[0][1:])
```

Hàm nhận kết quả khớp Regex và chuyển tên nhóm thành số nguyên.

Ví dụ: nhóm `n2_0` tương ứng với kết quả `2`.

### 5.6. `detect_decimals()`

```python
def detect_decimals(q: str):
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
```

Hàm tìm số lượng chữ số thập phân được yêu cầu trong câu hỏi.

Thuật toán:

1. Chuyển câu hỏi thành chữ thường.
2. Tìm từ khóa liên quan đến chữ số thập phân.
3. Tìm số đứng trước hoặc sau từ khóa trong phạm vi ký tự được xét.
4. Trả về số chữ số thập phân.
5. Trả về `None` nếu không nhận diện được.

Ví dụ:

```python
detect_decimals("Round the answer to two decimal places")
```

Kết quả:

```text
2
```

### 5.7. `normalize_final_answer()`

```python
def normalize_final_answer(
    query: str,
    raw_text: str
) -> str:
```

Đây là hàm chuẩn hóa câu trả lời cuối cùng của mô hình.

Hai tham số:

- `query`: câu hỏi ban đầu.
- `raw_text`: nội dung phản hồi thô từ LLM.

**Bước 1: Tách đáp án khỏi thẻ FINAL_ANSWER**

```python
matches = re.findall(
    r"<FINAL_ANSWER>(.*?)</FINAL_ANSWER>",
    raw_text,
    re.DOTALL | re.IGNORECASE
)
```

Nếu có thẻ đáp án, hàm lấy nội dung của thẻ khớp cuối cùng.

Nếu không có thẻ, hàm loại bỏ nội dung trong `<THOUGHT>...</THOUGHT>` và lấy dòng không rỗng cuối cùng.

**Bước 2: Loại bỏ phần định dạng thừa**

Hàm xóa các tiền tố như `Answer:`, `Final Answer:`, `Đáp án:` hoặc `Kết quả:`.

Hàm cũng loại bỏ một số dấu ngoặc kép, dấu backtick và dấu chấm câu cuối chuỗi theo điều kiện trong code.

**Bước 3: Xử lý câu hỏi Yes/No**

Nếu câu hỏi yêu cầu trả lời Yes hoặc No, hàm chuẩn hóa đáp án về đúng một trong hai chuỗi `Yes` hoặc `No` khi nhận diện được từ tương ứng.

**Bước 4: Chuẩn hóa số thập phân**

Hàm gọi:

```python
target_decimals = detect_decimals(query)
```

Nếu câu hỏi yêu cầu hai chữ số thập phân, giá trị như `14.2` có thể được chuẩn hóa thành:

```text
14.20
```

Nếu câu hỏi yêu cầu số nguyên, hàm làm tròn số và chuyển kết quả thành chuỗi số nguyên.

**Bước 5: Xử lý danh sách nhiều giá trị**

Nếu đáp án có dấu phẩy và câu hỏi yêu cầu danh sách phân tách bằng dấu phẩy, hàm phân biệt:

- Danh sách số: có thể chuẩn hóa từng phần tử.
- Danh sách văn bản: giữ nguyên để tránh làm biến dạng tên danh mục hoặc tên sheet.

Cuối cùng, hàm trả về một chuỗi đáp án duy nhất.

---

## 6. Giai đoạn 5: Lớp QuerySolver

`QuerySolver` kết nối với API LLM, lưu cache các workbook và điều phối toàn bộ quá trình trả lời một câu hỏi.

### 6.1. Hàm `__init__()`

```python
class QuerySolver:
    def __init__(
        self,
        api_key: str,
        model: str,
        base_url: str
    ):
        self.client = OpenAI(
            base_url=base_url,
            api_key=api_key,
            timeout=60.0,
            max_retries=3
        )

        self.model = model
        self.workspaces: Dict[str, ExcelWorkspace] = {}
        self._ws_lock = threading.Lock()
```

Các biến:

| Biến              | Ý nghĩa                               |
| ----------------- | ------------------------------------- |
| `api_key`         | API Key dùng để xác thực              |
| `model`           | Tên mô hình                           |
| `base_url`        | Địa chỉ API endpoint                  |
| `self.client`     | Client để gửi yêu cầu chat completion |
| `self.model`      | Tên model được sử dụng                |
| `self.workspaces` | Cache các đối tượng `ExcelWorkspace`  |
| `self._ws_lock`   | Khóa đồng bộ khi truy cập cache       |

Các tham số kết nối được truyền từ `main()`.

### 6.2. Hàm `get_workspace()`

```python
def get_workspace(
    self,
    xlsx_dir: str,
    file_name: str
) -> ExcelWorkspace:
    key = f"{xlsx_dir}::{file_name}"

    with self._ws_lock:
        if key not in self.workspaces:
            if len(self.workspaces) >= 5:
                del self.workspaces[
                    next(iter(self.workspaces.keys()))
                ]

            self.workspaces[key] = ExcelWorkspace(
                xlsx_dir,
                file_name
            )

        return self.workspaces[key]
```

Hàm này giúp tái sử dụng workbook đã tải lên RAM.

Biến `key` được tạo từ đường dẫn thư mục và tên file. Nếu khóa đã tồn tại, chương trình trả lại workspace đang có.

Nếu chưa tồn tại, hàm:

1. Kiểm tra số workspace đã lưu.
2. Nếu đã có ít nhất 5 workspace, xóa workspace đầu tiên trong thứ tự duyệt dictionary.
3. Khởi tạo workspace mới.
4. Lưu vào cache và trả về.

`self._ws_lock` bảo đảm các thao tác truy cập cache không diễn ra đồng thời theo cách thiếu đồng bộ giữa các worker.

Cơ chế này giảm việc đọc lại cùng một file Excel nhiều lần.

### 6.3. Lớp `UsageTracker` và hàm `_chat()`

`UsageTracker` cộng dồn token và chi phí của mọi lần gọi API trong một câu hỏi. Mỗi câu hỏi có một tracker riêng nên an toàn khi chạy đa luồng.

Mỗi lần gọi `add(usage)`:

- Đọc `prompt_tokens`, `completion_tokens`, `cached_tokens`, `reasoning_tokens` từ phản hồi API.
- Nếu API trả `usage.cost`: dùng chi phí thực này.
- Nếu không: ước tính = token chưa cache × giá prompt + token cache × giá cached + token completion × giá completion.

`_chat()` bọc lời gọi API: gửi kèm `extra_body={"usage": {"include": True}}` để OpenRouter trả chi phí, rồi gọi `tracker.add()`. Vòng lặp chính và lần gọi dự phòng đều đi qua hàm này nên không bị sót.

Thuộc tính `cost_source` cho biết chi phí lấy từ "API", "ước tính" hoặc "hỗn hợp"; giá trị được in ở dòng log `[USAGE]`.

### 6.4. Hàm `process_query()`

```python
def process_query(
    self,
    q_id: Any,
    xlsx_dir: str,
    file_name: str,
    query: str,
    log_cb=None
) -> Dict[str, Any]:
```

Đây là hàm trung tâm giải quyết một câu hỏi.

Đầu vào:

- `q_id`: ID của câu hỏi.
- `xlsx_dir`: thư mục Excel.
- `file_name`: tên workbook.
- `query`: nội dung câu hỏi.
- `log_cb`: hàm ghi log được truyền từ `process_item()`.

Nếu không truyền `log_cb`, hàm mặc định in log ra stderr.

#### Bước 1: Lấy workbook và tạo Overview

```python
ws = self.get_workspace(
    xlsx_dir,
    file_name
)

overview = ws.get_overview()
```

`ws` là đối tượng `ExcelWorkspace` dùng để truy vấn dữ liệu.

`overview` là bản xem trước workbook được chuyển thành chuỗi Markdown.

#### Bước 2: Tạo messages

```python
messages: List[Dict[str, Any]] = [
    {
        "role": "system",
        "content": SYSTEM_PROMPT
    },
    {
        "role": "user",
        "content": f"Question: {query}\n\n{overview}"
    },
]
```

Danh sách `messages` lưu lịch sử hội thoại cho lần giải câu hỏi hiện tại.

Hai thông điệp ban đầu:

- `system`: chỉ dẫn cách phân tích dữ liệu và sử dụng Tool.
- `user`: câu hỏi ban đầu kết hợp với Overview.

Mỗi lượt gọi API hoặc gọi Tool sẽ bổ sung thông tin vào danh sách này.

#### Bước 3: Khởi tạo bộ đếm token

```python
tracker = UsageTracker()
final_answer = ""
```

Toàn bộ xử lý (bước 1 đến 9) nằm trong khối `try`. Nếu lỗi giữa chừng, `tracker` vẫn giữ token và chi phí đã tiêu để trả về ở bước 10.

#### Bước 4: Bắt đầu vòng lặp Agent

```python
for step in range(6):
    response, step_cost = self._chat(
        tracker,
        model=self.model,
        messages=messages,
        tools=TOOLS_SCHEMA,
        tool_choice="auto",
        temperature=0.0,
    )
    msg = response.choices[0].message
```

Tối đa sáu lượt gọi. `tool_choice="auto"` cho phép mô hình trả lời trực tiếp hoặc gọi công cụ; `temperature=0.0` giảm tính ngẫu nhiên. Token và chi phí của mỗi lượt được `_chat()` ghi vào `tracker`, `log_cb` in ra phần token và chi phí của lượt đó.

#### Bước 5: Ghi nhận token và chi phí

Đã được `_chat()` thực hiện ở bước 4 (xem mục 6.3), không cần tính lại trong vòng lặp.

#### Bước 6: Ghi nhận phần THOUGHT

```python
if msg.content:
    for t in re.findall(
        r"<THOUGHT>(.*?)</THOUGHT>",
        msg.content,
        re.DOTALL | re.IGNORECASE
    ):
        log_cb(f"  [THOUGHT]: {t.strip()[:250]}")
```

Đoạn code tìm nội dung nằm trong thẻ `<THOUGHT>` và ghi tối đa 250 ký tự cho mỗi phần tìm thấy vào log.

Mục đích là lưu thông tin theo dõi đầu ra mô hình.

#### Bước 7: Nếu mô hình yêu cầu gọi Tool

```python
if msg.tool_calls:
    messages.append({
        "role": "assistant",
        "content": msg.content or "",
        "tool_calls": [
            {
                "id": tc.id,
                "type": "function",
                "function": {
                    "name": tc.function.name,
                    "arguments": tc.function.arguments
                },
            }
            for tc in msg.tool_calls
            if tc.type == "function"
        ],
    })
```

Các yêu cầu gọi công cụ được thêm vào lịch sử hội thoại dưới vai trò assistant.

Sau đó, chương trình lần lượt duyệt các Tool được yêu cầu.

```python
for tc in msg.tool_calls:
    if tc.type != "function":
        continue

    t_name = tc.function.name

    try:
        t_args = json.loads(
            tc.function.arguments
        )
    except Exception:
        t_args = {}
```

Trong đó:

- `tc`: một yêu cầu gọi Tool.
- `t_name`: tên Tool.
- `t_args`: dictionary chứa các tham số của Tool.

Việc điều phối được thực hiện theo tên công cụ:

```python
if t_name == "run_python_analysis":
    tool_out = ws.run_python(
        t_args.get("python_code", "")
    )

elif t_name == "read_sheet_grid":
    tool_out = ws.read_grid(
        t_args.get("sheet_name", ""),
        int(t_args.get("start_row", 1)),
        int(t_args.get("end_row", 120)),
    )

elif t_name == "search_keyword":
    tool_out = ws.search_keyword(
        t_args.get("keyword", "")
    )

else:
    tool_out = f"Tool '{t_name}' không tồn tại."
```

Ý nghĩa:

- `run_python_analysis`: thực thi code Python trên dữ liệu workbook.
- `read_sheet_grid`: đọc một vùng dòng của sheet.
- `search_keyword`: tìm từ khóa.
- Tên công cụ không hợp lệ: trả về thông báo lỗi.

Kết quả `tool_out` được thêm vào lịch sử hội thoại:

```python
messages.append({
    "role": "tool",
    "tool_call_id": tc.id,
    "name": t_name,
    "content": str(tool_out)[:6000],
})
```

Nội dung phản hồi Tool được giới hạn tối đa 6000 ký tự trước khi gửi lại cho mô hình.

Sau đó vòng lặp tiếp tục, và LLM có thể sử dụng kết quả Tool để gọi công cụ khác hoặc đưa ra đáp án.

#### Bước 8: Nếu mô hình trả lời trực tiếp

Khi `msg.tool_calls` không tồn tại hoặc rỗng, chương trình thực hiện nhánh `else`:

```python
messages.append({
    "role": "assistant",
    "content": msg.content or ""
})

final_answer = normalize_final_answer(
    query,
    msg.content or ""
)

if final_answer:
    break
```

Nội dung phản hồi được lưu vào hội thoại, sau đó đưa qua `normalize_final_answer()`.

Nếu nhận được đáp án không rỗng, vòng lặp Agent kết thúc bằng `break`.

Nếu chưa có đáp án, chương trình bổ sung yêu cầu mô hình trả về thẻ `<FINAL_ANSWER>` và tiếp tục vòng lặp.

#### Bước 9: Xử lý khi hết sáu lượt nhưng chưa có đáp án

Chương trình thêm yêu cầu xuất `<FINAL_ANSWER>` rồi gọi API thêm một lần qua `self._chat(tracker, ...)`. Lần gọi này được tính cả token lẫn chi phí vào `tracker`.

Nếu có ngoại lệ ở bất kỳ bước nào, khối `except` ghi log lỗi và hàm vẫn trả về kết quả với token và chi phí đã tiêu.

#### Bước 10: Trả kết quả

```python
return {
    "answer": final_answer,
    "cost": tracker.cost,
    "tokens": tracker.total_tokens,
}
```

Trước khi trả về, hàm ghi dòng log `[USAGE]` gồm prompt (kèm cached), completion (kèm reasoning) và nguồn chi phí. Dictionary này được chuyển về `process_item()` trong `main()`.

---

## 7. Giai đoạn 6: Hàm main() — Điểm bắt đầu điều phối ứng dụng

Sau khi các lớp và hàm hỗ trợ đã được định nghĩa, Python thực thi hàm `main()` khi file được chạy trực tiếp.

```python
def main():
    start_time = time.time()
```

`start_time` lưu thời điểm bắt đầu xử lý để cuối chương trình tính tổng thời gian chạy.

### 7.1. Tạo parser và đọc tham số dòng lệnh

```python
parser = argparse.ArgumentParser(
    description="AI26C0084 - Giải pháp Hỏi đáp trên Bảng tính Excel"
)

parser.add_argument(
    "--xlsx-dir",
    type=str,
    default="input",
    help="Thư mục chứa các tệp .xlsx"
)

parser.add_argument(
    "--queries",
    type=str,
    default="input/queries.json",
    help="Đường dẫn tệp queries.json"
)

parser.add_argument(
    "--output",
    type=str,
    default="output/submission.csv",
    help="Đường dẫn tệp kết quả"
)

parser.add_argument(
    "--openrouter-key",
    type=str,
    default=os.getenv("OPENROUTER_API_KEY", ""),
    help="Khóa API"
)

parser.add_argument(
    "--openrouter-model",
    type=str,
    default="qwen/qwen3.7-flash",
    help="Model ID"
)

parser.add_argument(
    "--openrouter-base-url",
    type=str,
    default="https://openrouter.ai/api/v1",
    help="Base URL"
)

args = parser.parse_args()
```

Các tham số:

| Tham số                 | Biến trong chương trình    | Giá trị mặc định               |
| ----------------------- | -------------------------- | ------------------------------ |
| `--xlsx-dir`            | `args.xlsx_dir`            | `input`                        |
| `--queries`             | `args.queries`             | `input/queries.json`           |
| `--output`              | `args.output`              | `output/submission.csv`        |
| `--openrouter-key`      | `args.openrouter_key`      | Đọc từ `OPENROUTER_API_KEY`    |
| `--openrouter-model`    | `args.openrouter_model`    | `qwen/qwen3.7-flash`           |
| `--openrouter-base-url` | `args.openrouter_base_url` | `https://openrouter.ai/api/v1` |

`args` là đối tượng chứa các tham số đã được phân tích.

Nếu người dùng truyền tham số khi chạy lệnh Python, các giá trị đó sẽ thay thế giá trị mặc định tương ứng.

### 7.2. Kiểm tra API Key

```python
if not args.openrouter_key:
    print(
        "Lỗi: Thiếu --openrouter-key hoặc "
        "biến môi trường OPENROUTER_API_KEY.",
        file=sys.stderr
    )

    sys.exit(1)
```

Nếu API Key trống, chương trình ghi lỗi ra stderr và kết thúc với mã thoát `1`.

Trong trường hợp hợp lệ, chương trình tiếp tục.

### 7.3. Tạo thư mục output

```python
out_dir = os.path.dirname(
    os.path.abspath(args.output)
)

if out_dir:
    os.makedirs(
        out_dir,
        exist_ok=True
    )
```

Các biến:

- `args.output`: đường dẫn file CSV đầu ra.
- `out_dir`: thư mục chứa file CSV đầu ra.

Ví dụ, nếu `args.output` là `output/submission.csv`, `out_dir` sẽ trỏ đến thư mục `output` theo đường dẫn tuyệt đối.

`os.makedirs(..., exist_ok=True)` tạo thư mục nếu chưa tồn tại, giúp tránh lỗi khi ghi file.

### 7.4. Đọc danh sách câu hỏi

```python
with open(
    args.queries,
    "r",
    encoding="utf-8"
) as f:
    queries_list = json.load(f)
```

Đây là bước đọc đầu vào quan trọng.

`queries_list` là danh sách các dictionary được đọc từ file JSON.

Ví dụ, nội dung `input/queries.json`:

```json
[
    {
        "id": "ext_q001",
        "file_name": "p2_t03.xlsx",
        "query": "In which year did the civilian labor force first cross the 70,000 mark (in thousands)? Answer with only the year as an integer."
    }
]
```

Mỗi phần tử chứa:

- `id`: mã câu hỏi.
- `file_name`: workbook cần sử dụng.
- `query`: nội dung câu hỏi.

Đoạn code thực tế cũng chấp nhận trường `file name` thay cho `file_name` ở bước xử lý từng câu hỏi.

### 7.5. Khởi tạo QuerySolver

```python
solver = QuerySolver(
    api_key=args.openrouter_key,
    model=args.openrouter_model,
    base_url=args.openrouter_base_url
)
```

Ba giá trị cấu hình được truyền vào hàm `QuerySolver.__init__()`.

Từ đây, biến `solver` trở thành đối tượng chịu trách nhiệm gọi API LLM và xử lý mọi câu hỏi.

### 7.6. Khởi tạo các biến kết quả và thống kê

```python
results = []
total_cost = 0.0
total_tokens = 0

print_lock = threading.Lock()
```

Ý nghĩa:

- `results`: danh sách đáp án cuối cùng.
- `total_cost`: tổng chi phí được cộng từ các tác vụ đã xử lý.
- `total_tokens`: tổng token được cộng từ các tác vụ đã xử lý.
- `print_lock`: khóa đồng bộ để các worker không in log xen kẽ nhau.

Biến `print_lock` được dùng trong hàm `process_item()` bên trong `main()`.

---

## 8. Giai đoạn 7: Hàm process_item() xử lý từng câu hỏi

`process_item()` là hàm lồng bên trong `main()`, được giao cho các worker xử lý song song.

```python
def process_item(item):
    q_id = item.get("id")
    file_name = item.get("file_name") or item.get("file name")
    query_text = item.get("query", "")
```

Các biến:

| Biến          | Ý nghĩa                                        |
| ------------- | ---------------------------------------------- |
| `item`        | Một dictionary câu hỏi từ `queries_list`       |
| `q_id`        | ID câu hỏi                                     |
| `file_name`   | Tên workbook cần đọc                           |
| `query_text`  | Nội dung câu hỏi                               |
| `logs_buffer` | Danh sách các dòng log của riêng câu hỏi       |
| `res`         | Dictionary kết quả từ `solver.process_query()` |
| `answer`      | Đáp án cuối cùng                               |
| `cost`        | Chi phí được trả về                            |
| `tokens`      | Token được trả về                              |

Đầu tiên, chương trình tạo bộ đệm log:

```python
logs_buffer = []

logs_buffer.append(
    f"\n{'='*55}\n>>> Đang xử lý ID: {q_id} | File: {file_name}"
)
```

Thay vì in từng thông báo ngay lập tức, chương trình lưu chúng vào `logs_buffer`.

### 8.1. Gọi QuerySolver.process_query()

```python
res = solver.process_query(
    q_id,
    args.xlsx_dir,
    file_name,
    query_text,
    log_cb=logs_buffer.append
)
```

Đây là điểm kết nối giữa `main()` và bộ xử lý LLM.

Tham số `log_cb=logs_buffer.append` khiến các thông báo bên trong `process_query()` được nối vào danh sách log thay vì in trực tiếp.

Khi `process_query()` hoàn thành, `res` có dạng:

```python
{
    "answer": "...",
    "cost": 0.000123,
    "tokens": 1000
}
```

Các giá trị trên chỉ là ví dụ minh họa cho cấu trúc, không phải kết quả của một lần chạy thực tế.

Sau đó, hàm lấy dữ liệu:

```python
answer = res["answer"]
cost = res["cost"]
tokens = res["tokens"]
```

Và bổ sung thông tin thống kê cùng đáp án vào log.

### 8.2. Đồng bộ in log

```python
with print_lock:
    print(
        "\n".join(logs_buffer),
        file=sys.stderr,
        flush=True
    )
```

Chương trình sử dụng khóa để in toàn bộ log của một tác vụ thành một nhóm.

Điều này hạn chế việc log của nhiều câu hỏi bị trộn lẫn khi chạy đa luồng.

### 8.3. Trả kết quả cho worker pool

```python
return {
    "id": q_id,
    "answer": answer,
    "cost": cost,
    "tokens": tokens
}
```

Kết quả trả về của worker chứa cả thông tin đáp án và thống kê.

Nếu có ngoại lệ trong quá trình xử lý:

```python
except Exception as e:
    logs_buffer.append(
        f"[LỖI ID {q_id}]: {str(e)}"
    )

    with print_lock:
        print(
            "\n".join(logs_buffer),
            file=sys.stderr,
            flush=True
        )

    return {
        "id": q_id,
        "answer": "",
        "cost": 0.0,
        "tokens": 0
    }
```

Lỗi phát sinh bên trong `process_query()` đã được hàm này tự bắt và vẫn giữ lại token đã tiêu. Khối `except` ở đây chỉ xử lý lỗi ngoài `process_query()` và trả về đáp án rỗng, chi phí và token bằng 0.

Nhờ đó, một lỗi trong câu hỏi không nhất thiết khiến worker phải trả về một kết quả không đúng cấu trúc mong đợi của `main()`.

---

## 9. Giai đoạn 8: Chạy song song bằng ThreadPoolExecutor

Sau khi định nghĩa `process_item()`, chương trình thiết lập worker pool:

```python
max_workers = 3

with concurrent.futures.ThreadPoolExecutor(
    max_workers=max_workers
) as executor:
    futures = [
        executor.submit(process_item, item)
        for item in queries_list
    ]
```

Các biến:

- `max_workers`: số worker tối đa, được đặt là 3.
- `executor`: đối tượng quản lý nhóm worker.
- `futures`: danh sách đối tượng Future đại diện cho các tác vụ.

Dòng:

```python
executor.submit(process_item, item)
```

đưa việc xử lý từng câu hỏi vào worker pool.

Nhờ vậy, tối đa ba câu hỏi có thể được xử lý đồng thời.

### 9.1. Thu thập kết quả hoàn thành

```python
for future in concurrent.futures.as_completed(futures):
    res_data = future.result()

    results.append({
        "id": res_data["id"],
        "answer": res_data["answer"],
        "cost": res_data["cost"],
        "tokens": res_data["tokens"]
    })

    total_cost += res_data["cost"]
    total_tokens += res_data["tokens"]
```

`as_completed(futures)` trả về các tác vụ theo thứ tự chúng hoàn thành, không nhất thiết theo thứ tự trong file JSON.

`future.result()` lấy giá trị được worker trả về.

Sau đó:

- `id`, `answer`, `cost`, `tokens` được thêm vào `results`.
- `cost` được cộng vào `total_cost`.
- `tokens` được cộng vào `total_tokens`.

Đây là nơi thống kê của từng câu hỏi được tổng hợp thành thống kê toàn chương trình.

---

## 10. Giai đoạn 9: Sắp xếp và ghi submission.csv

Khi tất cả Future đã hoàn thành, chương trình sắp xếp kết quả:

```python
try:
    results = sorted(
        results,
        key=lambda x: int(x["id"])
    )
except ValueError:
    results = sorted(
        results,
        key=lambda x: str(x["id"])
    )
```

Chương trình cố gắng sắp xếp ID dưới dạng số nguyên.

Nếu không thể chuyển đổi một ID nào đó thành số nguyên, toàn bộ danh sách sẽ được sắp xếp theo chuỗi.

Việc sắp xếp này giúp kết quả không bị phụ thuộc vào thứ tự hoàn thành các worker.

### 10.1. Tạo danh sách CSV

```python
csv_data = [
    {
        "id": r["id"],
        "answer": r["answer"]
    }
    for r in results
]
```

`csv_data` chứa chính xác hai trường:

- `id`
- `answer`

### 10.2. Ghi file CSV

```python
with open(
    args.output,
    "w",
    newline="",
    encoding="utf-8"
) as csvfile:

    writer = csv.DictWriter(
        csvfile,
        fieldnames=["id", "answer"]
    )

    writer.writeheader()
    writer.writerows(csv_data)
```

Chức năng:

1. Mở file đầu ra ở chế độ ghi.
2. Thiết lập tên hai cột.
3. Ghi header.
4. Ghi toàn bộ kết quả.

Ví dụ cấu trúc đầu ra:

```csv
id,answer
ext_q001,1972
ext_q002,1985
```

Các đáp án ở trên chỉ mang tính minh họa cấu trúc CSV.

Đường dẫn file phụ thuộc vào tham số `--output`. Với cấu hình mặc định, file sẽ là `output/submission.csv`.

---

## 11. Giai đoạn 10: Ghi answers.json

Sau khi ghi CSV, chương trình tạo đường dẫn file JSON:

```python
answers_json_path = os.path.join(
    out_dir,
    "answers.json"
) if out_dir else "answers.json"
```

Tiếp theo:

```python
try:
    with open(
        answers_json_path,
        "w",
        encoding="utf-8"
    ) as jsonfile:

        json.dump(
            results,
            jsonfile,
            ensure_ascii=False,
            indent=2
        )

except Exception as e:
    print(
        f"Lỗi khi lưu tệp {answers_json_path}: {e}",
        file=sys.stderr
    )
```

File `answers.json` được đặt cùng thư mục với `submission.csv`.

Tham số `ensure_ascii=False` giữ các ký tự Unicode, còn `indent=2` định dạng JSON dễ đọc.

File `answers.json` chứa `id`, `answer`, `cost`, `tokens` cho từng câu hỏi, dùng để đối chiếu chi phí. `submission.csv` vẫn chỉ có hai cột `id` và `answer` theo yêu cầu nộp bài.

---

## 12. Giai đoạn 11: Tính thời gian và in thống kê cuối chương trình

Sau khi ghi các file, chương trình tính thời gian chạy:

```python
elapsed_time = time.time() - start_time

mins = int(elapsed_time // 60)
secs = elapsed_time % 60
```

Các biến:

- `start_time`: thời điểm bắt đầu thực thi `main()`.
- `elapsed_time`: tổng thời gian đã trôi qua, tính bằng giây.
- `mins`: số phút nguyên.
- `secs`: số giây còn lại, có thể chứa phần thập phân.

Chương trình sau đó in thông tin tổng kết ra stderr:

```python
print(
    f"HOÀN TẤT! Đã xử lý: {len(results)} câu | "
    f"Thời gian: {mins}p {secs:.2f}s",
    file=sys.stderr
)

print(
    f"Tổng Token: {total_tokens} | "
    f"Tổng Chi phí: ${total_cost:.6f}",
    file=sys.stderr
)

print(
    f"-> File submission.csv : "
    f"{os.path.abspath(args.output)}",
    file=sys.stderr
)

print(
    f"-> File answers.json  : "
    f"{os.path.abspath(answers_json_path)}",
    file=sys.stderr
)
```

Các thông tin này cho biết:

- Tổng số câu hỏi có trong danh sách kết quả.
- Thời gian chạy.
- Tổng token được ghi nhận.
- Tổng chi phí (ưu tiên chi phí thực từ API).
- Đường dẫn tuyệt đối đến hai file đầu ra.

Khi hoàn thành đoạn này, hàm `main()` kết thúc và chương trình kết thúc nếu không còn tác vụ nào khác.

---

## 13. Theo dõi các biến quan trọng từ đầu đến cuối

Bảng dưới đây tổng hợp vòng đời của những biến chính trong ứng dụng.

| Biến              | Khởi tạo ở đâu                 | Được sử dụng như thế nào                 | Kết quả cuối                                  |
| ----------------- | ------------------------------ | ---------------------------------------- | --------------------------------------------- |
| `args`            | `main()`                       | Chứa tham số dòng lệnh                   | Quyết định đầu vào, model và đường dẫn đầu ra |
| `queries_list`    | `main()`                       | Danh sách câu hỏi đọc từ JSON            | Được đưa vào worker pool                      |
| `solver`          | `main()`                       | Đối tượng `QuerySolver`                  | Xử lý từng câu hỏi và gọi API                 |
| `self.workspaces` | `QuerySolver.__init__()`       | Cache workbook                           | Tái sử dụng dữ liệu Excel trong RAM           |
| `self.sheets_raw` | `ExcelWorkspace.__init__()`    | Lưu dữ liệu các sheet                    | Phục vụ Overview, tìm kiếm và tính toán       |
| `messages`        | `QuerySolver.process_query()`  | Lịch sử trao đổi với LLM                 | Chứa prompt, câu trả lời và kết quả Tool      |
| `msg`             | `QuerySolver.process_query()`  | Phản hồi hiện tại của model              | Quyết định gọi Tool hay chuẩn hóa đáp án      |
| `tool_out`        | `QuerySolver.process_query()`  | Kết quả của một Tool                     | Được gửi lại cho model                        |
| `final_answer`    | `QuerySolver.process_query()`  | Đáp án đang được hoàn thiện              | Trở thành `answer`                            |
| `logs_buffer`     | `process_item()`               | Lưu log riêng cho từng câu hỏi           | Được in ra stderr                             |
| `res_data`        | Vòng lặp Future trong `main()` | Kết quả của worker                       | Được dùng để tổng hợp đáp án và thống kê      |
| `results`         | `main()`                       | Danh sách kết quả đã thu thập            | Được sắp xếp rồi ghi CSV và JSON              |
| `total_cost`      | `main()`                       | Tổng chi phí thực/ước tính cộng dồn      | Được in ở cuối chương trình                   |
| `total_tokens`    | `main()`                       | Tổng token được cộng dồn                 | Được in ở cuối chương trình                   |
| `elapsed_time`    | Cuối `main()`                  | Chênh lệch thời gian bắt đầu và kết thúc | Được chuyển thành phút và giây                |

### Chuỗi truyền dữ liệu quan trọng nhất

```text
queries_list
    |
    v
process_item(item)
    |
    +--> q_id
    +--> file_name
    +--> query_text
    |
    v
solver.process_query(...)
    |
    +--> get_workspace(...)
    |        |
    |        v
    |    ExcelWorkspace
    |        |
    |        +--> sheets_raw
    |        +--> get_overview()
    |
    +--> messages
    |
    +--> LLM API
    |        |
    |        +--> msg
    |        +--> tool_calls
    |        +--> tool_out
    |        +--> final_answer
    |
    v
res = {
    answer,
    cost,
    tokens
}
    |
    v
process_item trả kết quả
    |
    v
future.result()
    |
    +--> results
    +--> total_cost
    +--> total_tokens
    |
    v
Sắp xếp results
    |
    +--> submission.csv
    +--> answers.json
    |
    v
In thống kê và kết thúc
```

---

## 14. Một số điểm đáng chú ý khi đọc và bảo trì main.py

### 14.1. Overview không thay thế việc đọc toàn bộ dữ liệu

`get_overview()` chỉ tạo bản xem trước có giới hạn. Nếu câu trả lời phụ thuộc vào các dòng hoặc cột chưa xuất hiện trong Overview, mô hình có thể cần gọi `search_keyword`, `read_sheet_grid` hoặc `run_python_analysis`.

### 14.2. Cơ chế tự sửa lỗi hoạt động qua kết quả Tool

Nếu code do mô hình sinh ra bị lỗi, `run_python()` trả lại thông báo lỗi và gợi ý phù hợp. Mô hình có thể đọc thông báo đó trong vòng tiếp theo để sửa truy vấn hoặc code phân tích.

Cơ chế này cho phép Agent thử lại trong giới hạn vòng lặp đã thiết lập.

### 14.3. Xử lý song song không đồng nghĩa với kết quả được lưu theo thứ tự hoàn thành

Các worker hoàn thành không theo thứ tự cố định. Chương trình chủ động sắp xếp lại `results` theo ID trước khi ghi CSV để giảm sự phụ thuộc vào thứ tự thực thi.

### 14.4. Nguồn của chi phí

Chi phí ưu tiên lấy từ `usage.cost` của OpenRouter; chỉ khi thiếu mới ước tính theo giá dự phòng ở mục 2.3. Sau mỗi câu hỏi, dòng log `[USAGE]` cho biết nguồn là "API", "ước tính" hay "hỗn hợp". Nếu thấy "ước tính", hãy kiểm tra lại tham số `usage.include` và giá trong code.

---

## 15. Kết luận

`main.py` được chia thành bốn lớp chức năng chính:

1. **ExcelWorkspace:** Đọc workbook, xử lý ô gộp, tạo Overview, đọc vùng dữ liệu, tìm từ khóa và cung cấp hàm hỗ trợ phân tích Pandas.
2. **Chuẩn hóa đáp án:** Dùng Regex để tách đáp án cuối cùng, xử lý Yes/No, danh sách và định dạng số thập phân.
3. **QuerySolver:** Quản lý cache, giao tiếp với LLM, điều phối Tool, ghi nhận token và tạo đáp án cuối cùng.
4. **main():** Đọc tham số CLI và JSON, tạo worker pool, thu thập kết quả, sắp xếp dữ liệu và ghi các file đầu ra.

Toàn bộ luồng thực thi bắt đầu tại `if __name__ == "__main__": main()`, đi qua các bước chuẩn bị dữ liệu, xử lý từng câu hỏi bằng Agent, rồi kết thúc tại phần tổng kết thống kê.

Tệp kết quả chính là `submission.csv`, chứa hai trường `id` và `answer`. File `answers.json` được tạo cùng thư mục, bổ sung `cost` và `tokens` của từng câu.

**Mã nguồn:** https://github.com/Fishiro/AI26C0084/blob/main/main.py

**README dự án:** https://github.com/Fishiro/AI26C0084/blob/main/README.md
