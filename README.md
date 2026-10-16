# AI26C0084 - Giải pháp Hỏi đáp Tự động trên Bảng tính Excel

## Cấu trúc dự án

```text
AI26C0084/
├── Dockerfile
├── main.py
├── requirements.txt
├── README.md
├── input/
│   ├── queries.json
│   └── xlsx_files/
│       ├── workbook_1.xlsx
│       └── workbook_2.xlsx
└── output/
```

Trong đó:

- `input/`: chứa tệp `queries.json` và thư mục con `xlsx_files/` chứa các tệp Excel cần truy vấn.
- `output/`: thư mục nhận kết quả, sẽ được tạo trước khi chạy nếu chưa tồn tại.
- `queries.json`: chứa danh sách câu hỏi theo định dạng mà hệ thống chấm yêu cầu.
- Các tệp Excel phải được đặt trong thư mục được truyền qua tham số `--xlsx-dir`.

### Build Docker image

Chạy lệnh tại thư mục gốc chứa `Dockerfile`:

```bash
docker build --platform linux/amd64 -t ai26c0084 .
```

Lệnh trên tạo Docker image có tên `ai26c0084`.

Nếu quá trình build thất bại, hãy kiểm tra `Dockerfile`, `requirements.txt`, kết nối mạng và khả năng tải các gói phụ thuộc.

### Chạy chương trình bằng Docker

Ví dụ lệnh chạy tiêu chuẩn với các tham số đúng định dạng theo yêu cầu của BTC:

```bash
docker run --rm \
  -v "$(pwd)/input:/input:ro" \
  -v "$(pwd)/output:/output" \
  ai26c0084 \
  --xlsx-dir /input/xlsx_files \
  --queries /input/queries.json \
  --output /output/submission.csv \
  --openrouter-key "$OPENROUTER_API_KEY" \
  --openrouter-model qwen/qwen3.7-flash \
  --openrouter-base-url https://openrouter.ai/api/v1
```

## Chạy trực tiếp bằng Python

Phương án này phù hợp để phát triển, gỡ lỗi và kiểm tra nhanh trong môi trường cục bộ.

### Yêu cầu môi trường

- Python 3.10 trở lên; khuyến nghị Python 3.11.
- Các thư viện trong `requirements.txt`.
- Kết nối Internet và API key hợp lệ của OpenRouter.

### Cài đặt thư viện

Tại thư mục gốc dự án `AI260084`, chạy:

```bash
python -m pip install -r requirements.txt
```

### Chạy chương trình

```bash
python main.py \
  --xlsx-dir ./input/xlsx_files \
  --queries ./input/queries.json \
  --output ./output/submission.csv \
  --openrouter-key "$OPENROUTER_API_KEY" \
  --openrouter-model qwen/qwen3.7-flash \
  --openrouter-base-url https://openrouter.ai/api/v1
```

---

## Tham số dòng lệnh

| Tham số                 | Ý nghĩa                                  |
| ----------------------- | ---------------------------------------- |
| `--xlsx-dir`            | Đường dẫn đến thư mục chứa các tệp Excel |
| `--queries`             | Đường dẫn đến tệp JSON chứa câu hỏi      |
| `--output`              | Đường dẫn lưu tệp CSV kết quả            |
| `--openrouter-key`      | API key dùng để gọi OpenRouter           |
| `--openrouter-model`    | Tên mô hình LLM sử dụng                  |
| `--openrouter-base-url` | Địa chỉ API endpoint của OpenRouter      |

Cấu hình mô hình mặc định trong hướng dẫn:

```text
Model:    qwen/qwen3.7-flash
Endpoint: https://openrouter.ai/api/v1
```

Các tham số trên đã tương thích với giao diện CLI được triển khai trong `main.py`.

---

## Kiến trúc và phương pháp xử lý

Giải pháp được tổ chức tập trung trong `main.py`, hướng đến giảm chi phí xử lý dữ liệu Excel, hạn chế sử dụng bộ nhớ và chuẩn hóa câu trả lời trước khi xuất kết quả.

### ExcelWorkspace - Xử lý bảng tính trong bộ nhớ

- Sử dụng `openpyxl` và `pandas` để đọc và xử lý dữ liệu Excel.
- Hỗ trợ xử lý ô gộp (`merged_cells`) bằng cách điền giá trị vào các ô liên quan trong bộ nhớ.
- Tái sử dụng dữ liệu và cấu trúc bảng tính đã nạp khi xử lý nhiều câu hỏi trên cùng dữ liệu.
- Hỗ trợ nhận diện tiêu đề nhiều tầng (`multi-row headers`) thông qua `load_df()`.
- Cung cấp thông tin về các trang tính để hỗ trợ truy vấn trên nhiều sheet.

### QuerySolver - Vòng lặp gọi công cụ phân tích

- Kết nối mô hình Qwen thông qua OpenRouter API.
- Sử dụng `temperature=0.0` theo cấu hình mô hình để hướng đến kết quả nhất quán.
- Có thể tạo và thực thi mã Python/Pandas thông qua `run_python_analysis` để phân tích dữ liệu có cấu trúc.
- Hỗ trợ đọc lưới ô bằng `read_sheet_grid` khi cần xác định vị trí hoặc quan hệ giữa các ô.
- Phối hợp giữa suy luận của LLM và thao tác trên dữ liệu bảng tính để tìm câu trả lời.

### Chuẩn hóa kết quả

Hàm `normalize_final_answer()` xử lý định dạng câu trả lời cuối cùng nhằm tăng khả năng tương thích với đánh giá tự động, bao gồm:

- Chuẩn hóa số nguyên và số thập phân.
- Chuẩn hóa danh sách câu trả lời theo định dạng phân tách bằng dấu phẩy khi phù hợp.
- Chuẩn hóa các câu trả lời dạng Yes/No.
- Xuất kết quả thành CSV với hai cột `id,answer`.

Chuẩn hóa định dạng không đảm bảo câu trả lời đúng về mặt ngữ nghĩa. Kết quả cuối cùng vẫn phụ thuộc vào dữ liệu đầu vào, khả năng suy luận, kết quả phân tích và quy tắc đánh giá của BTC.

---

## Cấu trúc thư mục dự án

```text
AI26C0084/
├── Dockerfile
├── main.py
├── requirements.txt
├── README.md
├── input/                 # Dữ liệu thử nghiệm cục bộ
│   ├── queries.json
│   └── xlsx_files/
│       └── *.xlsx
└── output/                # Kết quả sinh ra khi chạy
    └── submission.csv
```

`input/` và `output/` là các thư mục dùng trong ví dụ chạy thử; không nhất thiết phải được đóng gói vào image. Khi chấm thi, BTC có thể mount dữ liệu thực tế vào container theo đúng giao diện đầu vào đã quy định.

---

## Tóm tắt quy trình chạy

Quy trình thực thi chuẩn:

```text
Chuẩn bị Excel và queries.json
            |
            v
Build Docker image
            |
            v
Mount dữ liệu đầu vào
            |
            v
Chạy main.py với cấu hình OpenRouter
            |
            v
Xử lý câu hỏi và phân tích dữ liệu Excel
            |
            v
Ghi output/submission.csv
            |
            v
Kiểm tra schema và số lượng câu trả lời
```

**Tệp kết quả:** `output/submission.csv`.

## MINH CHỨNG

| STT | Nội dung                                                             | Link lịch sử đoạn chat                                                 |
| :-: | -------------------------------------------------------------------- | ---------------------------------------------------------------------- |
|  1  | Cải thiện sản phẩm                                                   | [CLAUDE](https://claude.ai/share/00a90e01-b257-4a10-bda1-a73c8be7ee96) |
|  2  | Phân tích đầu ra và cải thiện truy vấn đa ngôn ngữ                   | [GEMINI](https://share.gemini.google/kQXXzSQovcjK)                     |
|  3  | Phân tích Excel và cải thiện instruction truy vấn và sử dụng công cụ | [GEMINI](https://share.gemini.google/NyiM40zkniMh)                     |
|  4  | Fix bug lỗi                                                          | [GEMINI](https://share.gemini.google/1avvkxISVCel)                     |
|  5  | Cải thiện instruction prompt                                         | [GEMINI](https://share.gemini.google/kv7FioQPRpSc)                     |
|  6  | Tái cấu trúc dự án                                                   | [GEMINI](https://share.gemini.google/pZBdfOUsBobi)                     |
|  7  | Nâng cấp tính năng tự động ghi log và tính toán chi phí token        | [GEMINI](https://share.gemini.google/PMzh5CrVYxEV)                     |
|  8  | Giải pháp MCP (đã bị thay thế)                                       | [GEMINI](https://share.gemini.google/Srtz5QqUG9CO)                     |

## LINK MÃ NGUỒN DỰ ÁN: [GitHub | AI26C0084](https://github.com/Fishiro/AI26C0084)
