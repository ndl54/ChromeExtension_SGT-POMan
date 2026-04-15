# SGT POMan Chrome Extension

Tiện ích Chrome hỗ trợ tạo nhanh nội dung ghi chú để sao chép cho 2 nghiệp vụ:
- `Ghi chú Phiếu Nhập`
- `Ghi chú Trả Hàng`

## Tính năng chính

- Chuyển nhanh giữa 2 chế độ ghi chú bằng dropdown trên cùng.
- Gợi ý dữ liệu từ Google Sheet cho:
  - đối tác giao/trả hàng
  - nhà cung cấp
  - mã sản phẩm
- Hỗ trợ nhiều dòng sản phẩm.
- Sao chép nhanh vào clipboard.
- Có thể `Điền lại dữ liệu cũ` từ note đã copy trước đó.
- Có thể tải mới dữ liệu và cache vào `chrome.storage`.

## 1. Ghi chú Phiếu Nhập

### Trường nhập

- `Trường hợp` bắt buộc
- `Đối tác giao hàng` tùy chọn
- `Nhà cung cấp` tùy chọn
- Ít nhất 1 dòng `mã sản phẩm + tình trạng kích/quà`
- `Tiền cước` tùy chọn
- `Ghi chú thêm` tùy chọn

### Trường hợp khi copy

- `case1`: `Lấy NCC giao khách`
- `case2`: `Lấy NCC về kho`
- `case3`: `NCC giao về kho`
- `case4`: `NCC giao khách`

### Format copy

Với `case1` hoặc `case2`:

```text
[Đối tác giao hàng] | [Nhà cung cấp] | [Trường hợp]
[Mã SP] - [Tình trạng] |
[Mã SP] - [Tình trạng] |
Cước: [số tiền] | [Ghi chú]
```

Với `case3` hoặc `case4`:

```text
[Trường hợp] | [Đối tác giao hàng] | [Nhà cung cấp]
[Mã SP] - [Tình trạng] |
[Mã SP] - [Tình trạng] |
Cước: [số tiền] | [Ghi chú]
```

Ghi chú:
- Nếu không có cước hoặc ghi chú thì tự bỏ qua.
- Đối tác được chuẩn hóa theo data.
- Có hỗ trợ `Urbox` cho từng dòng sản phẩm.

## 2. Ghi chú Trả Hàng

### Trường nhập

- Ít nhất 1 dòng sản phẩm, mỗi dòng gồm:
  - `Mã sản phẩm`
  - `Tình trạng hàng`
  - `Tình trạng vỏ, đai`
- `Đối tác trả hàng` bắt buộc
- `Ghi chú thêm` tùy chọn

### Format copy

```text
[Mã SP] - [Tình trạng hàng] - [Tình trạng vỏ, đai] |
[Mã SP] - [Tình trạng hàng] - [Tình trạng vỏ, đai] |
Đối tác/GV: [Tên đối tác]
Ghi chú: [Nội dung ghi chú]
```

Ví dụ:

```text
UA65DU7700 - Hàng không lỗi - Vỏ đẹp |
WT-85NG1 - Hàng lỗi - Vỏ không đẹp |
Đối tác/GV: Lộc BM
Ghi chú: Khách lùi giờ giao
```

Ghi chú:
- Có thể thêm nhiều dòng sản phẩm.
- Nếu không nhập `Ghi chú thêm` thì tự bỏ qua dòng `Ghi chú: ...`.
- `Điền lại dữ liệu cũ` cũng hỗ trợ format này.

## Dữ liệu nguồn

Nguồn dữ liệu là Google Sheet:
- `Doi_Tac_Giao_Hang`
- `Nha_Cung_Cap`
- `San_Pham`

File CSV trong repo:
- `doitacgiaohang.csv`
- `nhacungcap.csv`
- `sanpham.csv`

Lưu ý:
- `nhacungcap.csv` hiện chỉ còn 1 cột: `Tên nhà cung cấp`
- mã sản phẩm được sanitize trước khi ghi CSV: bỏ hậu tố ` [IMEI]`, ` [CŨ]`, ` MN`, ` ML`, sau đó unique và sort

## Cập nhật CSV trong repo

Chạy:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\fetch-doitacgiaohang.ps1
```

Script sẽ cập nhật lại cả 3 file CSV từ Google Sheet hiện tại.

## Cài đặt

1. Mở `chrome://extensions/`
2. Bật `Developer mode`
3. Chọn `Load unpacked`
4. Trỏ tới thư mục dự án

## Yêu cầu

- Google Chrome
- Quyền `clipboard`
- Quyền `storage`
