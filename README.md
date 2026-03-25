# Using-Python-to-process-soil-physical-mechanical-property-data
Using Python to process soil physical–mechanical property data from spreadsheets for AI training data preparation
# Xlsx2CSV.py (Excel → CSV and TCVN3 → Unicode normalization)

`Xlsx2CSV.py` is a Python script that converts legacy Excel spreadsheets (`.xls`, `.xlsx`) into **CSV files** for ML training data preparation. It scans an input directory in batch mode, reads each workbook, and **splits data by worksheet** (skipping empty worksheets).

In addition to exporting CSV, the script detects Vietnamese text stored in **TCVN3 / VietKey**-style encoding (often found in older Excel files using fonts like `.Vn...`) and converts it into standard **Unicode** using the mapping embedded in the script. CSV files are written as **UTF-8 with BOM** to improve compatibility with Excel and common data tools.

---

## Key features (English)

- Batch scan all Excel files in the input directory.
- For each workbook:
  - Export **1 CSV per worksheet** with actual data.
  - Skip empty worksheets to avoid polluting the training set.
- CSV output:
  - Encoding: **UTF-8 with BOM** (`utf-8-sig`)
  - Naming: `file_stem_sheetname.csv` (sheet name is sanitized for Windows)
- Text normalization:
  - Detect potentially TCVN3-encoded strings using **font metadata** and/or **TCVN3 marker characters**.
  - Convert TCVN3 → Unicode with a VietKey-compatible mapping.
- Logging:
  - `OK: <filename>` for successful processing
  - `FAIL: <filename> -> <reason>` for unreadable/encrypted files

## Installation requirements

- `.xlsx`: `openpyxl`
- `.xls`: `xlrd==1.2.0` (newer `xlrd` versions removed `.xls` support)

Install with:

```bash
pip install openpyxl xlrd==1.2.0
```

## Usage

From the project folder:

```bash
python "Xlsx2CSV.py"
```

Or specify input/output folders:

```bash
python "Xlsx2CSV.py" --input "d:\...\dataXLSX" --output "d:\...\dataCSV_UTF8"
```

## Common errors (English)

- `File is not a zip file`  
  Typical for corrupted/invalid `.xlsx` files (not a valid OOXML ZIP container).
- `Workbook is encrypted`  
  Typical for password-protected `.xls` files.

---

# Xlsx2CSV.py (chuyển Excel → CSV và chuẩn hóa TCVN3 → Unicode)

`Xlsx2CSV.py` là script Python dùng để quét toàn bộ file bảng tính trong một thư mục, đọc từng *workbook* và **tách dữ liệu theo từng sheet** ra file **CSV UTF-8**. Đồng thời script có bước nhận diện và **chuyển chuỗi tiếng Việt mã TCVN3/VietKey sang Unicode chuẩn** để tránh lỗi font khi đưa dữ liệu vào pipeline học máy.

## Tính năng chính

- Quét hàng loạt trong thư mục đầu vào.
- Với mỗi workbook:
  - Tách **từng worksheet** thành **1 file CSV**.
  - **Bỏ qua sheet rỗng** (không có dữ liệu đáng kể).
- Xuất CSV:
  - Mã hóa **UTF-8 có BOM** (`utf-8-sig`) để mở đúng trong Excel và các công cụ phổ biến.
  - Quy ước tên: `tenFileGoc_tenSheet.csv`.
- Chuẩn hóa tiếng Việt:
  - Nhận diện chuỗi nghi ngờ TCVN3 dựa trên **metadata font** (font kiểu `.Vn...`/chứa `TCVN3`) và/hoặc **mẫu ký tự đặc trưng** TCVN3.
  - Chuyển TCVN3 → Unicode theo bảng VietKey (cài đặt dạng thay thế theo **hai pha**).

## Yêu cầu cài đặt

### Với `.xlsx`

- `openpyxl`

### Với `.xls`

- `xlrd==1.2.0` (nhất thiết cần bản hỗ trợ `.xls` cũ)

Cài đặt:

```bash
pip install openpyxl xlrd==1.2.0
```

## Cách chạy

Trong thư mục dự án (hoặc gọi theo đường dẫn đầy đủ):

### Mặc định (theo cấu trúc dự án)

- `--input`: thư mục `dataXLSX` (nằm cùng cấp với `PythonCode`)
- `--output`: thư mục `dataCSV`

```bash
python "Xlsx2CSV.py"
```

### Chỉ định thư mục input/output

```bash
python "Xlsx2CSV.py" --input "d:\...\dataXLSX" --output "d:\...\dataCSV_UTF8"
```

## Đầu ra (CSV)

- Mỗi **worksheet có dữ liệu** tạo ra **một file CSV**.
- Tên file:
  - `file_stem_sheetname.csv` (ví dụ `okVinh.xls_Goc.csv`)
  - Trong đó `sheetname` được “làm sạch” ký tự không hợp lệ trên Windows.
- CSV dùng `utf-8-sig` (có BOM).

## Nhận diện và chuyển mã TCVN3 → Unicode

Script áp dụng điều kiện (là chuỗi `str` và thỏa ít nhất một trong hai tín hiệu):

- `font_name`:
  - chứa chuỗi `TCVN3` (không phân biệt hoa thường), hoặc
  - bắt đầu bằng `.vn` (ví dụ `.VnTime`, `.VnArial`, `.VnCourier`, …)
- nội dung chuỗi:
  - chứa các ký tự “marker” đặc trưng TCVN3 (ví dụ `µ¸©·¡§ÌÐÝÓ¨®Ü¬­¹¶Ê»¼½Æ`, …)

Sau đó chuỗi được chuyển theo bảng ánh xạ VietKey đã tích hợp sẵn trong script.

## Log & lỗi thường gặp

Script ghi:

- `OK: <tên file>` khi xử lý được workbook và xuất ra CSV.
- `FAIL: <tên file> -> <lý do>` khi không đọc được workbook (tệp hỏng / sai định dạng / mã hóa / …).
- Cuối cùng: `Đã xuất N file CSV ...`

Các lỗi hay gặp:

- `File is not a zip file` (thường với `.xlsx`)
  - file bị hỏng, hoặc bị đổi đuôi, hoặc không phải định dạng Excel OOXML chuẩn.
- `Workbook is encrypted` (thường với `.xls`)
  - file bị mật khẩu/mã hóa; không trích xuất được nếu không có mật khẩu.

## Gợi ý mở rộng (nếu cần cho nghiên cứu tiếp theo)

- Xuất thêm `failures.txt` / `failures.json` để lưu log lỗi có cấu trúc.
- Thêm bước chuẩn hóa cột/đơn vị (schema master) trước khi đưa vào huấn luyện.
- Tích hợp kiểm tra chất lượng số (phát hiện ngoại lệ, giá trị thiếu, định dạng số thập phân).
- Chuyển từ CSV phẳng sang format phù hợp hơn (Parquet/JSON/TFRecord…) tùy pipeline ML.

