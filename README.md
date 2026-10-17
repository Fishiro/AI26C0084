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

Giải pháp được thiết kế gói gọn trong `main.py`, kết hợp giữa tiền xử lý bảng tính trên RAM, vòng lặp suy luận với công cụ (Tool Calling) và bộ chuẩn hóa đầu ra bằng biểu thức chính quy (Regex).

### 1. ExcelWorkspace - Tiền xử lý và lưu đệm bảng tính trên RAM

- **Đọc và làm phẳng ô gộp (`merged_cells`):** Sử dụng `openpyxl` (`data_only=True`) để lấy giá trị thực của ô thay vì công thức. Tự động tách toàn bộ các vùng ô gộp và điền lặp lại giá trị của ô gốc (top-left) vào tất cả các ô thành phần nhằm bảo toàn ngữ cảnh dòng/cột.
- **Tạo bảng xem trước (`get_overview`):** Trích xuất tối đa 35 dòng và 20 cột đầu tiên của mỗi trang tính dưới dạng bảng Markdown có gắn tọa độ dòng (`R1(0)`) và cột (`C1(A)`). Với các bảng dữ liệu nhỏ, mô hình có thể trả lời trực tiếp ngay từ bước xem trước mà không cần gọi thêm công cụ.
- **Chuẩn hóa DataFrame (`load_df`):** Tự động loại bỏ các dòng/cột trống hoàn toàn, nhận diện và gộp tiêu đề nhiều tầng (`multi-row headers`) thành tên cột hoàn chỉnh (kèm tên cột rút gọn nếu không trùng lặp), đồng thời tự động ép kiểu số (`pd.to_numeric`) cho các cột dữ liệu.

### 2. QuerySolver - Vòng lặp suy luận và thực thi công cụ

- **Cơ chế bộ nhớ đệm an toàn luồng (Thread-safe Cache):** Lưu giữ tối đa 5 đối tượng `ExcelWorkspace` gần nhất trên RAM và đồng bộ hóa truy cập bằng `threading.Lock`, giúp tái sử dụng dữ liệu khi có nhiều câu hỏi cùng truy vấn trên một tệp Excel.
- **Vòng lặp Agent (tối đa 6 bước):** Kết nối với mô hình ngôn ngữ (mặc định `qwen/qwen3.7-flash`) qua OpenRouter API với `temperature=0.0`. Mô hình suy luận từng bước trong thẻ `<THOUGHT>` và chủ động lựa chọn gọi công cụ trước khi đưa ra đáp án trong thẻ `<FINAL_ANSWER>`.
- **Hệ thống công cụ và hàm trợ thủ an toàn:**
    - `read_sheet_grid`: Đọc chi tiết một vùng dòng/cột cụ thể trên sheet khi bảng vượt quá phạm vi xem trước.
    - `run_python_analysis`: Thực thi mã Python/Pandas trực tiếp trên dữ liệu trong bộ nhớ và thu nhận kết quả qua `stdout`. Cung cấp sẵn các hàm trợ thủ (`load_df`, `find_row`, `find_col`, `get_val`) giúp tìm kiếm dòng/cột theo từ khóa không phân biệt hoa thường, hạn chế lỗi sai lệch tên cột do ký tự ẩn hoặc tiêu đề gộp quá dài.
- **Quy tắc suy luận miền dữ liệu (Domain Rules):** Tích hợp chỉ dẫn trực tiếp trong System Prompt để xử lý các trường hợp dễ mất điểm Exact Match như: giữ nguyên cột `Total` khi đề bài yêu cầu tính trên toàn bộ các nhóm (`across all...`), tính tuổi dựa trên khoảng cách giữa `Date of Admission` và `Date of Birth`, và giữ nguyên giá trị số gốc trên ô mà không tự ý nhân theo đơn vị ghi chú (`[In thousands]`, `[In millions]`).

### 3. Chuẩn hóa kết quả đa ngôn ngữ (`normalize_final_answer`)

- **Làm sạch chuỗi kết quả:** Trích xuất nội dung từ thẻ `<FINAL_ANSWER>` (hoặc dòng kết quả cuối cùng), loại bỏ các tiền tố giải thích thừa (`Answer:`, `Đáp án:`, `Réponse:`,...), dấu ngoặc kép và dấu chấm câu ở cuối.
- **Định dạng số học đa ngôn ngữ (`detect_decimals`):** Sử dụng Regex để nhận diện yêu cầu về số chữ số thập phân hoặc số nguyên bằng nhiều ngôn ngữ (Anh, Việt, Pháp, Đức, Tây Ban Nha, Bồ Đào Nha, Trung, Hàn), đảm bảo làm tròn chuẩn xác và giữ nguyên các chữ số `0` ở phần thập phân (ví dụ: `14.20`).
- **Xử lý các định dạng câu trả lời đặc thù:** Tự động chuẩn hóa câu hỏi dạng `Yes/No`, định dạng từng phần tử trong câu trả lời chứa nhiều giá trị ngăn cách bằng dấu phẩy, và lọc bỏ ký hiệu tiền tệ/phần trăm (`$`, `%`, `,`) khi chuẩn hóa giá trị số.

### 4. Xử lý song song, theo dõi chi phí và xuất kết quả

- **Thực thi đa luồng (`ThreadPoolExecutor`):** Chạy song song tối đa 3 luồng (`max_workers=3`) để tối ưu hóa thời gian chờ mạng khi gọi API mà vẫn kiểm soát tốt dung lượng RAM tiêu thụ.
- **Gom log và thống kê chi phí:** Mỗi câu hỏi được gom toàn bộ lịch sử xử lý vào bộ đệm riêng (`logs_buffer`) và in tuần tự qua khóa `print_lock` để tránh trộn lẫn log giữa các luồng, kèm theo thống kê chi tiết số lượng token và chi phí ước tính.
- **Đầu ra tiêu chuẩn và dự phòng:** Sắp xếp danh sách kết quả theo `id` tăng dần, xuất ra tệp `submission.csv` gồm 2 cột `id,answer` theo đúng quy định và tự động lưu thêm tệp `answers.json` cùng thư mục để đối chiếu khi cần.

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
