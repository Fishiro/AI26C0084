# AI26C0084 - Giải pháp Hỏi đáp Tự động trên Bảng tính Excel

## 1. Kiến trúc giải pháp

Giải pháp được thiết kế tinh gọn trong một tệp thực thi chính (`main.py`) nhằm tối ưu hóa tốc độ phản hồi, tiết kiệm bộ nhớ (< 1 GB RAM) và giảm thiểu số lượng token LLM sử dụng:

- **`ExcelWorkspace` (In-Memory Excel Engine)**:
    - Sử dụng `openpyxl` và `pandas` tự động tách và điền giá trị cho toàn bộ các ô bị gộp (`merged_cells`) ngay trên RAM.
    - Lưu bộ nhớ đệm (cache) cấu trúc bảng tính để dùng chung cho nhiều câu hỏi trên cùng một tệp Excel mà không cần đọc lại từ ổ cứng.
    - Hàm `load_df()` tự động nhận diện và ghép nối tiêu đề phân cấp nhiều tầng (multi-row headers), hỗ trợ truy vấn xuyên suốt nhiều trang tính (`sheets_raw`, `sheet_names`).
- **`QuerySolver` (Agentic Tool-Calling Loop)**:
    - Kết nối mô hình `qwen/qwen3.7-flash` qua OpenRouter API với `temperature=0.0` để đảm bảo tính nhất quán và chính xác.
    - Tự động sinh và thực thi mã phân tích Python/Pandas (`run_python_analysis`) hoặc đọc lưới tọa độ (`read_sheet_grid`) khi xử lý bảng dữ liệu phức tạp.
    - Bộ hậu xử lý `normalize_final_answer()` chuẩn hóa định dạng câu trả lời đa ngôn ngữ (số nguyên, độ chính xác thập phân, danh sách cách nhau bởi dấu phẩy, Yes/No) để khớp chính xác (Exact Match) và xuất ra tệp CSV chuẩn UTF-8 gồm hai cột `id,answer`.

## 2. Cấu trúc thư mục

```text
AI26C0084/
├── Dockerfile          # Cấu hình đóng gói image Linux amd64
├── main.py             # Mã nguồn thực thi chính (ENTRYPOINT)
├── requirements.txt    # Danh sách thư viện Python phụ thuộc
└── README.md           # Tài liệu hướng dẫn cài đặt và vận hành
```

## 3. Cài đặt môi trường trực tiếp (Local)

Yêu cầu hệ thống: Python 3.10 trở lên (khuyến nghị Python 3.11).

Cài đặt các thư viện phụ thuộc bằng lệnh:

```bash
pip install -r requirements.txt
```

## 4. Chạy chương trình bằng dòng lệnh (CLI)

### 4.1. Chạy thử nghiệm trên Windows PowerShell (Dữ liệu tại thư mục `./input`)

```pwsh
python main.py `
  --xlsx-dir ./input `
  --queries ./input/queries.json `
  --output ./output/submission.csv `
  --openrouter-key "<YOUR_OPENROUTER_KEY>" `
  --openrouter-model qwen/qwen3.7-flash `
  --openrouter-base-url https://openrouter.ai/api/v1
```

### 4.2. Chạy thử nghiệm trên Windows CMD

```cmd
python main.py ^
  --xlsx-dir ./input ^
  --queries ./input/queries.json ^
  --output ./output/submission.csv ^
  --openrouter-key "<YOUR_OPENROUTER_KEY>" ^
  --openrouter-model qwen/qwen3.7-flash ^
  --openrouter-base-url https://openrouter.ai/api/v1
```

### 4.3. Chạy trên Linux / macOS / Git Bash

```bash
python main.py \
  --xlsx-dir ./input \
  --queries ./input/queries.json \
  --output ./output/submission.csv \
  --openrouter-key "<YOUR_OPENROUTER_KEY>" \
  --openrouter-model qwen/qwen3.7-flash \
  --openrouter-base-url https://openrouter.ai/api/v1
```

## 5. Đóng gói và chạy kiểm thử qua Docker (Chuẩn môi trường chấm thi)

### Bước 1: Build Docker Image (Linux amd64)

```bash
docker build --platform linux/amd64 -t ai26c0084 .
```

### Bước 2: Chạy Container mô phỏng hệ thống chấm

- **Trên Windows PowerShell:**

```pwsh
docker run --rm `
  -v "${PWD}/input:/input/xlsx_files:ro" `
  -v "${PWD}/input/queries.json:/input/queries.json:ro" `
  -v "${PWD}/output:/output" `
  ai26c0084 `
  --xlsx-dir /input/xlsx_files `
  --queries /input/queries.json `
  --output /output/submission.csv `
  --openrouter-key "<YOUR_OPENROUTER_KEY>" `
  --openrouter-model qwen/qwen3.7-flash `
  --openrouter-base-url https://openrouter.ai/api/v1)
```

- **Trên Linux / macOS / Git Bash:**

```bash
docker run --rm \
  -v "$(pwd)/input:/input/xlsx_files:ro" \
  -v "$(pwd)/input/queries.json:/input/queries.json:ro" \
  -v "$(pwd)/output:/output" \
  ai26c0084 \
  --xlsx-dir /input/xlsx_files \
  --queries /input/queries.json \
  --output /output/submission.csv \
  --openrouter-key "<YOUR_OPENROUTER_KEY>" \
  --openrouter-model qwen/qwen3.7-flash \
  --openrouter-base-url https://openrouter.ai/api/v1
```
