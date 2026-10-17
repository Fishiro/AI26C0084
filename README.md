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

Giải pháp được thiết kế gói gọn trong `main.py`, kết hợp giữa tiền xử lý bảng tính trên RAM, hệ thống Agent tự sửa lỗi qua công cụ (Tool Calling), và bộ chuẩn hóa đầu ra bằng biểu thức chính quy (Regex) hỗ trợ Exact Match đa ngôn ngữ.

### 1. ExcelWorkspace - Tiền xử lý thông minh và Tối ưu Token động

- **Làm phẳng không gian ô gộp (Unmerge & Context Preservation):** Các tệp thống kê thường lạm dụng ô gộp (`merged_cells`). Hệ thống sử dụng `openpyxl` với `data_only=True` để đọc giá trị thực, sau đó tự động unmerge toàn bộ và điền sao chép giá trị của ô gốc (top-left) vào tất cả các ô thành phần. Điều này giúp Pandas sau này không bị lỗi `NaN` khi truy xuất tọa độ.
- **Thuật toán "Ngân sách dòng toàn cục" (Global Token Budget):** Không giới hạn cố định số dòng cố định cho hàm `get_overview`. Hệ thống tính toán động dựa trên số lượng sheets: nếu file chỉ có 1 sheet, nó cấp phát tới 50 dòng để mường tượng tối đa; nhưng nếu file có 15 sheets, nó sẽ ép nén mỗi sheet xuống 15 dòng đầu. Đảm bảo Context Window của Prompt luôn an toàn dưới giới hạn Token và tiết kiệm chi phí (`< 0.10$`).
- **Chuẩn hóa Header & Loại bỏ Title rác (`load_df`):** Tự động bỏ qua các dòng rỗng. Phát hiện dòng Header đích thực bằng cách quét tìm dòng đầu tiên chứa nhiều hơn 2 cột có dữ liệu thực tế. Tránh hoàn toàn việc Pandas lấy nhầm dòng Title báo cáo (VD: _"Báo cáo tài chính quý 3"_) làm tên Header, đồng thời gộp an toàn các tiêu đề đa tầng (MultiIndex) thành chuỗi phẳng.

### 2. QuerySolver - Agent Tự phục hồi và Thực thi an toàn (Self-Healing Agent)

- **Cơ chế Cache luồng an toàn (Thread-safe Lock):** Cache tối đa 5 đối tượng `ExcelWorkspace` gần nhất trên RAM. Sử dụng `threading.Lock` để cho phép các câu hỏi song song truy cập cùng 1 file Excel mà không bị lỗi xung đột biến nhớ hay phải đọc lại file từ ổ cứng nhiều lần.
- **Vòng lặp Agent 6 bước & Tool Calling tự phục hồi:** Kết nối qua OpenRouter (`qwen/qwen3.7-flash`, `temperature=0.0`). Đặc biệt, hàm `run_python_analysis` tích hợp cơ chế **Auto-correction hint** trực tiếp vào Traceback. Ví dụ: nếu LLM dùng sai tên cột sinh ra `KeyError`, hệ thống không chỉ báo lỗi mà còn gợi ý ngay trong Console: _"[HINT]: Lỗi KeyError! Đừng cố đoán tên cột, hãy in df.columns.tolist() ra..."_. Mô hình sẽ tự học và sửa code đúng ở vòng lặp sau.
- **Bộ Công cụ & Helper Vượt Ải (Domain Tools):**
    - `extract_table`: Vũ khí hạng nặng xử lý các file quá dị biệt. Tự động cắt phăng toàn bộ các dòng title dư thừa phía trên và lấy đúng dòng chứa `keyword` làm Header mới.
    - `search_keyword`: Tầm soát toàn file. Báo cáo chính xác (Sheet, Dòng, Cột) khi dữ liệu nằm lọt thỏm giữa 15 sheet hoặc hàng ngàn dòng mà Overview không thể thấy.
    - `clean_num` & `clean_col`: Hóa giải mọi bẫy ký tự ẩn của BTC (xóa `$`, `%`, `,`) và biến đổi các chuỗi rỗng/vô định (như -, `N/A`) thành số `0.0` an toàn trước khi dùng phép toán `.sum()` hay `.mean()`.
    - `parse_date`: Ép kiểu dữ liệu chuỗi `dd/mm/yyyy` về chuẩn Datetime Pandas, khắc phục triệt để lỗi `TypeError` khi BTC yêu cầu trừ ngày tháng (VD: tính tuổi dựa trên `Date of Admission`).

### 3. Chuẩn hóa kết quả Khắc nghiệt (Exact Match Normalization)

- **Gọt dũa tiền tố/hậu tố:** Quét `Regex` sâu để bóc tách đáp án khỏi thẻ `<FINAL_ANSWER>`. Lọc sạch các từ dẫn dắt thừa (như `Answer:`, `Kết quả:`), loại bỏ dấu ngoặc kép thừa và dấu chấm câu ở cuối chuỗi.
- **Bảo toàn định dạng Thập phân đa ngôn ngữ (`detect_decimals`):** Dùng Regex quét câu hỏi (hỗ trợ Anh, Việt, Pháp, Đức, Trung, Hàn, Tây Ban Nha) để bắt yêu cầu về định dạng (VD: _"two decimal places"_). Ép kiểu và bảo toàn số `0` ở phần thập phân (VD: `14.20`) hoặc làm tròn nguyên số.
- **Chống sai lệch do Phép toán mảng (Numeric Array Fallback):** Khi đáp án trả về là một mảng xếp hạng tên chữ có chứa dấu phẩy (VD: _"Food, Energy, Services"_), hệ thống tự động nhận diện mảng String để không can thiệp biến dạng chuỗi. Nhưng nếu là mảng số (_"2020, 9034"_), hệ thống tự bóc tách, dọn dẹp `$ / %` và làm tròn số học trên từng phần tử.

### 4. Tối ưu chi phí và Quản lý đa luồng

- **Thực thi Đa luồng (`ThreadPoolExecutor`):** Chạy song song `max_workers=3`. Giúp tận dụng tối đa 6 vCPU để đánh bại thời gian trễ của API mạng, nhưng vẫn nằm an toàn dưới ngưỡng tràn RAM của BTC.
- **Quản lý Console Log bất đồng bộ:** Áp dụng mô hình "Gom log cục bộ" (`logs_buffer`) và khóa đồng bộ hóa in ấn (`print_lock`), giúp các đoạn log chi tiết (THOUGHT, CALL, OUTPUT, COST) của từng câu hỏi không bị trộn lẫn nát bét trên màn hình Terminal.
- **Xuất File & Đo lường chi phí:** Đáp ứng tuyệt đối cấu trúc nộp bài bằng cách lưu mảng cuối cùng (đã sort theo `id`) ra tệp `submission.csv` (2 cột: id, answer). Đồng thời ghi toàn bộ siêu dữ liệu vòng đời (Cost, Tokens, Log lỗi) vào tệp `answers.json` nội bộ để dễ dàng audit chi phí thực tế.

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
