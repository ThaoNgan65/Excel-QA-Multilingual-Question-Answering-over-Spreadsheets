# Excel QA – Hỏi đáp tự động trên bảng tính Excel đa ngôn ngữ

Giải pháp tự động trả lời câu hỏi ngôn ngữ tự nhiên dựa trên các tệp bảng tính `.xlsx` (cấu trúc phức tạp, đa ngôn ngữ EN/VI/ZH/JA/KO/FR). Với mỗi câu hỏi, hệ thống trả về **một đáp án ngắn gọn, đúng định dạng** để so khớp chính xác (Exact Match).

- **Đầu vào:** thư mục chứa các tệp `.xlsx` và tệp `queries.json`.
- **Đầu ra:** tệp `submission.csv` gồm 2 cột `id`, `answer` (mỗi `id` đúng một dòng).
- **Mô hình LLM:** `qwen/qwen3.7-flash` qua OpenRouter.

---

## 1. Ý tưởng chính

Thay vì để LLM "đọc rồi đoán" đáp án, hệ thống cho LLM **viết một đoạn mã Python ngắn** để tra cứu/tính toán trên dữ liệu bảng, sau đó **chạy mã trong sandbox**. Nhờ vậy các phép cộng, đếm, xếp hạng, làm tròn… đều do máy tính thực hiện thay vì ước lượng bằng ngôn ngữ.

### Luồng xử lý

```
xlsx ──► Đọc & chuẩn hoá (merge, header path, multi-table, công thức)
     ──► Dump văn bản (kèm địa chỉ ô)
     ──► (Tuỳ chọn) Routing: chọn sheet liên quan khi workbook lớn
     ──► LLM sinh code Python  ──► Sandbox chạy code ──► (lỗi → sửa 1 lần)
     ──► Bỏ phiếu nhiều lượt (VOTES) ──► Trọng tài (judge) khi phiếu lệch
     ──► Hậu xử lý theo định dạng yêu cầu trong câu hỏi (làm tròn, ngày, Yes/No, nhãn…)
     ──► submission.csv
```

### Các thành phần quan trọng

| Thành phần | Mô tả |
|---|---|
| **Đọc Excel → `Sheet`** | Xử lý ô gộp (merge), tiêu đề nhiều tầng (header path dạng `A > B`), tự tách nhiều bảng xếp chồng trong một sheet, làm sạch giá trị (marker thiếu như `-`, `n/a`, số kiểu kế toán, fullwidth…). |
| **Bộ tính công thức** | Với ô công thức không có giá trị tính sẵn, hệ thống tự tính (hỗ trợ `SUM`, `IF`, `VLOOKUP`, `INDEX/MATCH`, `SUMIF(S)`, `COUNTIF(S)`, `ROUND*`… kể cả tham chiếu xuyên sheet và cache liên kết ngoài). |
| **Tra cứu ô theo địa chỉ** | Câu hỏi dạng "ô D5 chứa gì / công thức nào" được đọc trực tiếp từ tệp gốc (`CELL LOOKUP`), có thể trả lời không cần LLM. |
| **Sandbox** | Chạy mã do LLM sinh với danh sách builtin và module bị giới hạn (`math`, `re`, `datetime`, `statistics`, …), có timeout. |
| **Bỏ phiếu + Judge** | Sinh nhiều phiếu (mặc định 3); nếu lệch, một lượt LLM làm trọng tài so sánh code và bảng để chọn đáp án đúng nhất. |
| **Hậu xử lý đa ngôn ngữ** | Nhận diện yêu cầu làm tròn/định dạng trong câu hỏi (`2 chữ số thập phân`, `保留2位小数`, `소수점 2자리`, `round up`, `làm tròn xuống`…), định dạng ngày, Yes/No, khôi phục đúng nhãn gốc (giữ dấu chấm/phẩy). |
| **Routing sheet** | Với workbook lớn (nhiều sheet, dump dài), LLM chọn sheet cần thiết để giảm token và tăng tốc. |

---

## 2. Cấu trúc thư mục nộp bài

```
.
├── excelqa.py          # mã nguồn chính
├── requirements.txt    # thư viện phụ thuộc
├── Dockerfile
└── README.md
```

Gợi ý nội dung `requirements.txt`:

```
openpyxl
requests
```

> Chỉ cần hai thư viện này. Hệ thống **không** dùng mô hình Embedding/Reranker chạy local nên bộ nhớ rất nhẹ.

---

## 3. Định dạng dữ liệu

**`queries.json`** – danh sách câu hỏi, mỗi phần tử gồm:

```json
[
  {
    "id": "q001",
    "file_name": "report.xlsx",
    "query": "Tổng doanh thu năm 2023 là bao nhiêu? Làm tròn 2 chữ số thập phân."
  }
]
```

**`submission.csv`** – kết quả:

```csv
id,answer
q001,12345.67
```

---

## 4. Cách chạy

### 4.1. Chạy bằng Docker (khuyến nghị)

Build image:

```bash
docker build -t excelqa .
```

Chạy:

```bash
docker run --rm \
  -e OPENROUTER_API_KEY="<khoá-api-của-bạn>" \
  -v /đường/dẫn/data:/data:ro \
  -v /đường/dẫn/out:/out \
  excelqa \
  --xlsx-dir /data \
  --queries /data/queries.json \
  --output /out/submission.csv
```

`ENTRYPOINT` của image là `python excelqa.py`, nên chỉ cần truyền tham số phía sau tên image.

### 4.2. Chạy trực tiếp bằng Python

```bash
pip install -r requirements.txt
export OPENROUTER_API_KEY="<khoá-api-của-bạn>"

python excelqa.py \
  --xlsx-dir ./data \
  --queries ./data/queries.json \
  --output ./submission.csv
```

### 4.3. Tham số dòng lệnh

| Tham số | Ý nghĩa |
|---|---|
| `--xlsx-dir` | Thư mục chứa các tệp `.xlsx` (tìm đệ quy). |
| `--queries` | Đường dẫn tới `queries.json`. Nếu bỏ qua, tự tìm `queries.json` trong thư mục dữ liệu. |
| `--output`, `-o` | Đường dẫn `submission.csv` (mặc định `submission.csv`). |
| `--openrouter-key` | API key (hoặc dùng biến môi trường `OPENROUTER_API_KEY`). |
| `--openrouter-model` | Tên model (mặc định `qwen/qwen3.7-flash`). |
| `--openrouter-base-url` | Base URL, ví dụ `https://openrouter.ai/api/v1`. |
| `--input`, `-i` | (Cũ) thư mục chứa cả `queries.json` và `.xlsx`. |
| `--api-key` | (Cũ) tương đương `--openrouter-key`. |
| `--dump <file.xlsx>` | In ra bản dump văn bản của một tệp xlsx rồi thoát (dùng để kiểm tra khâu đọc dữ liệu). |

> Không có đường dẫn nào được hardcode; chương trình chạy lại bình thường trên tập kiểm thử ẩn.

---

## 5. Cấu hình bằng biến môi trường

Mọi tham số đều có giá trị mặc định hợp lý, chỉ đặt khi cần tinh chỉnh.

**Kết nối & tài nguyên**

| Biến | Mặc định | Ý nghĩa |
|---|---|---|
| `OPENROUTER_API_KEY` | – | **Bắt buộc.** Khoá API OpenRouter (không ghi vào mã nguồn). |
| `OPENROUTER_MODEL` | `qwen/qwen3.7-flash` | Model sử dụng. |
| `OPENROUTER_API_URL` | `https://openrouter.ai/api/v1/chat/completions` | Endpoint. |
| `WORKERS` | `6` | Số câu hỏi xử lý song song. |
| `TIME_BUDGET_SEC` | `2100` | Ngân sách thời gian toàn bộ (giây), nằm dưới giới hạn 40 phút. |
| `MAX_COST_PER_Q` | `0.08` | Trần chi phí mỗi câu (USD), thấp hơn mức 0.10 USD cho phép. |
| `LLM_MAX_TOKENS` | `3000` | Số token đầu ra tối đa mỗi lượt gọi. |
| `PRICE_IN_PER_M` / `PRICE_OUT_PER_M` | `0.03` / `0.13` | Đơn giá (USD/1M token) để ước tính chi phí. |
| `HTTP_TIMEOUT`, `LLM_RETRIES` | `90`, `4` | Timeout và số lần thử lại khi gọi API. |
| `REASONING` | `off` | `off` = tắt chế độ suy luận của model để tiết kiệm token/thời gian. |

**Bỏ phiếu & trọng tài**

| Biến | Mặc định | Ý nghĩa |
|---|---|---|
| `VOTES` | `3` | Số phiếu mỗi câu. |
| `VOTE_TEMP` | `0.3` | Nhiệt độ cho các phiếu sau phiếu đầu (phiếu đầu dùng 0). |
| `PARALLEL_VOTES` | `1` | Chạy song song 2 phiếu đầu. |
| `JUDGE` | `1` | Bật/tắt trọng tài khi phiếu lệch. |
| `JUDGE_OVERRIDE` | `1` | `0` = chỉ gọi judge khi chưa có đa số. |
| `REJECT_ERROR_ANSWERS` | `1` | Không chấp nhận chuỗi kiểu "Error…" làm đáp án. |

**Đọc dữ liệu & routing**

| Biến | Mặc định | Ý nghĩa |
|---|---|---|
| `MULTI_TABLE` | `1` | Tự tách nhiều bảng xếp chồng trong một sheet. |
| `ROUTING` | `1` | Bật chọn sheet liên quan. |
| `ROUTE_MIN_CHARS` | `12000` | Chỉ routing khi dump dài hơn ngưỡng này. |
| `ROUTE_SHRINK_RATIO` | `0.8` | Chỉ dùng kết quả routing nếu dump giảm dưới tỉ lệ này. |
| `MAX_HEADER_ROWS` | `12` | Số dòng tối đa được coi là tiêu đề. |
| `CELL_FALLBACK` | `1` | Tra cứu ô theo địa chỉ, đọc công thức thô. |
| `SANDBOX_TIMEOUT` | `15` | Timeout (giây) khi chạy mã do LLM sinh. |

**Ghi log**

| Biến | Mặc định | Ý nghĩa |
|---|---|---|
| `DEBUG_LOG` | `debug_log.jsonl` | Tệp log chi tiết từng câu (phiếu, code, judge, token, chi phí, thời gian). |
| `LOG_CANDS` | `1` | Ghi cả mã và đáp án của từng phiếu vào log. |

---

## 6. Đáp ứng ràng buộc của đề bài

| Ràng buộc | Cách đáp ứng |
|---|---|
| Mô hình `qwen/qwen3.7-flash` qua OpenRouter | Mặc định dùng đúng model này, gọi qua HTTP bằng `requests`. |
| Không hardcode API key | Key lấy từ biến môi trường `OPENROUTER_API_KEY` hoặc tham số dòng lệnh. |
| RAM ≤ 1 GB | Không dùng Embedding/Reranker local; mỗi workbook được cache một lần và chỉ chứa dữ liệu văn bản. |
| Gói nộp ≤ 100 MB | Chỉ gồm `excelqa.py`, `requirements.txt`, `Dockerfile`, `README.md`. |
| ≤ 0.10 USD / câu | Có bộ đếm chi phí theo từng câu, ngắt khi vượt `MAX_COST_PER_Q` (mặc định 0.08). |
| ≤ 40 phút / 20 câu | Chạy song song nhiều câu, có `TIME_BUDGET_SEC` làm hạn chót cho toàn bộ pipeline. |
| Mỗi `id` đúng một dòng | `submission.csv` được ghi theo đúng thứ tự `id` trong `queries.json`, ghi lại sau mỗi câu (an toàn nếu bị ngắt giữa chừng). |
| Chỉ ghi đáp án cuối | Hậu xử lý loại bỏ tiền tố ("Answer: …"), chú thích, dấu chấm thừa… |

### Tối ưu theo tiêu chí xếp hạng

1. **Độ chính xác (Exact Match):** sinh mã + chạy sandbox, bỏ phiếu, trọng tài, hậu xử lý định dạng đa ngôn ngữ.
2. **Thời gian trung bình:** xử lý song song, dừng bỏ phiếu sớm khi đã có đa số, trả lời thẳng các câu hỏi tra cứu ô mà không cần LLM.
3. **Token trung bình:** tắt reasoning, routing sheet cho workbook lớn, mã sinh ra ngắn gọn.

---

## 7. Gỡ lỗi

- Xem dump của một tệp xlsx để kiểm tra khâu đọc dữ liệu:
  ```bash
  python excelqa.py --dump data/report.xlsx
  ```
- Mở `debug_log.jsonl` (mỗi dòng một câu) để xem các phiếu, mã đã sinh, quyết định của judge, số token, chi phí và thời gian.
- Log tiến trình được in ra `stderr` theo từng câu, ví dụ: `[q001] vote3  12.3s tok=8200/900 $0.00036 ...`.

---

## 8. Lưu ý

- Mã do LLM sinh chỉ chạy trong sandbox giới hạn (không `open`, không `os`, không mạng).
- Không đổi đơn vị của số (ví dụ "đơn vị: nghìn") trừ khi câu hỏi yêu cầu rõ.
- Đáp án dạng nhãn được giữ nguyên chuỗi gốc trong sheet, không dịch.
