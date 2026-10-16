# AI26C0084 - Giải pháp Hỏi đáp Tự động trên Bảng tính Excel (MCP + FastMCP)

## 1. Kiến trúc giải pháp

- **MCP Server (`helper/excel_mcp_server.py`)**: Sử dụng `FastMCP`, `openpyxl` và `pandas` để tự động làm phẳng các ô bị gộp (merged cells), đọc tiêu đề phân cấp nhiều tầng, tìm kiếm từ khóa đa trang tính và thực thi mã phân tích Python/Pandas trực tiếp.
- **MCP Client (`main.py`)**: Kết nối MCP Server qua giao thức `stdio`, sử dụng mô hình `qwen/qwen3.7-flash` qua OpenRouter API, kết hợp cơ chế bộ nhớ đệm cấu trúc bảng (`overview_cache`) và bộ chuẩn hóa định dạng đầu ra (`normalize_final_answer`) để xuất tệp `submission.csv` chuẩn UTF-8.

## 2. Cài đặt môi trường trực tiếp (Local)

```bash
pip install -r requirements.txt
```

## 3. Chạy chương trình bằng dòng lệnh (CLI)

```bash
python main.py \
 --xlsx-dir /input/xlsx_files \
 --queries /input/queries.json \
 --output /output/submission.csv \
 --openrouter-key <YOUR_OPENROUTER_KEY> \
 --openrouter-model qwen/qwen3.7-flash \
 --openrouter-base-url https://openrouter.ai/api/v1
```

```pwsh
python main.py `
  --xlsx-dir /input/xlsx_files `
  --queries /input/queries.json `
  --output /output/submission.csv `
  --openrouter-key <YOUR_OPENROUTER_KEY> `
  --openrouter-model qwen/qwen3.7-flash `
  --openrouter-base-url https://openrouter.ai/api/v1
```

```cmd
python main.py ^
  --xlsx-dir /input/xlsx_files ^
  --queries /input/queries.json ^
  --output /output/submission.csv ^
  --openrouter-key <YOUR_OPENROUTER_KEY> ^
  --openrouter-model qwen/qwen3.7-flash ^
  --openrouter-base-url https://openrouter.ai/api/v1
```

## 4. Đóng gói và chạy qua Docker

```bash
docker build --platform linux/amd64 -t ai26c0084 .

docker run --rm \
  -v $(pwd)/excel_input:/input/xlsx_files:ro \
  -v $(pwd)/queries.json:/input/queries.json:ro \
  -v $(pwd)/submission_output:/output \
  ai26c0084 \
  --xlsx-dir /input/xlsx_files \
  --queries /input/queries.json \
  --output /output/submission.csv \
  --openrouter-key <YOUR_OPENROUTER_KEY> \
  --openrouter-model qwen/qwen3.7-flash \
  --openrouter-base-url https://openrouter.ai/api/v1
```
