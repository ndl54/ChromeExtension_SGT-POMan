# SGT POMan Chrome Extension

Tiện ích giúp tạo nội dung cần sao chép nhanh cho bộ phận tạo Phiếu Nhập Hàng Nhà Cung Cấp (POMan) dựa trên các trường nhập liệu trong popup.

## Tính năng

- Chọn trường hợp ai lấy hàng và giao về đâu.
- Đối tác giao hàng: tùy chọn, gợi ý từ danh sách đối tác
- Nhà cung cấp: tùy chọn, gợi ý từ danh sách nhà cung cấp
- Chọn nhiều cặp `mã sản phẩm -> trạng thái kích hoạt/quà tặng`
- Tiền cước (nếu có) và ghi chú thêm
- Sao chép nhanh vào clipboard
- Tải mới dữ liệu đối tác, nhà cung cấp và mã sản phẩm từ Google Sheet rồi cache vào máy

## Trường bắt buộc

- Trường hợp
- Ít nhất 1 cặp `mã sản phẩm -> trạng thái`

## Định dạng nội dung sao chép

Ví dụ khi điền đầy đủ (Giao vận/đối tác lấy NCC):

```
Đối tác A [DT000001] | NCC A | Lấy NCC giao khách | HC-MWO381B: Còn Kích | LF-D6RCWM: Hết Kích | Cước: 100.000đ | Ghi chú tùy chọn
```

Nội dung được ghép theo thứ tự:

```
[Đối tác giao hàng] | [Nhà cung cấp] | [Trường hợp] | [Mã sản phẩm 1: Trạng thái 1] | [Mã sản phẩm 2: Trạng thái 2] | Cước: [số tiền] | [Ghi chú]
```

Áp dụng thứ tự này khi chọn Case 1 hoặc Case 2. Với Case 3 hoặc Case 4, thứ tự sẽ là:

```
[Trường hợp] | [Đối tác giao hàng] | [Nhà cung cấp] | [Mã sản phẩm 1: Trạng thái 1] | [Mã sản phẩm 2: Trạng thái 2] | Cước: [số tiền] | [Ghi chú]
```

Quy tắc ghép:

| Văn bản hiển thị (UI)                                   | Văn bản copy vào clipboard                           |
|---------------------------------------------------------|------------------------------------------------------|
| NCC giao về kho SGT                                     | `Lấy NCC giao khách`                                 |
| NCC giao tại nhà khách hàng                             | `Lấy NCC về kho`                                     |
| Giao vận/đối tác lấy NCC giao khách                     | `NCC giao về kho`                                    |
| Giao vận/đối tác lấy NCC giao về kho SGT                | `NCC giao khách`                                     |
| Đối tác giao hàng                                       | `Tên đối tác [mã đối tác]`                           |
| Mã sản phẩm + trạng thái                                | `[Mã sản phẩm]: [Giá trị đã chọn]`                   |
| Cước (ví dụ nhập `100000`)                              | `Cước: 100.000đ`                                     |
| Ghi chú                                                 | `| [nội dung ghi chú]` (nối cuối chuỗi)              |
| Không nhập các field tùy chọn                           | Những phần tương ứng sẽ bị loại bỏ                   |

Ghi chú:
- Nếu không có tiền cước hoặc ghi chú thì bỏ qua phần tương ứng.
- Đối tác sẽ tự động chuẩn hóa theo data: `Tên đối tác [mã đối tác]`.
- Có thể thêm nhiều dòng mã sản phẩm, mỗi dòng đi kèm một trạng thái riêng.

## Nhãn Trường hợp khi sao chép

- Case 1: Lấy NCC giao khách
- Case 2: Lấy NCC về kho
- Case 3: NCC giao về kho
- Case 4: NCC giao khách

## Dữ liệu đối tác giao hàng

Nguồn dữ liệu: Google Sheet (read-only).

- Nút "Mở Google Sheet" để mở nhanh link nhập liệu.
- Nút "Tải dữ liệu mới" sẽ fetch CSV từ Google Sheet và cache vào `chrome.storage`.
- File gốc trong repo: `doitacgiaohang.csv` (được tạo từ Sheet).
- File mã sản phẩm trong repo: `sanpham.csv` (được tạo từ sheet `San_Pham`).

### Cập nhật file CSV trong repo

Chạy script sau để cập nhật file CSV trong repo:

```
powershell -ExecutionPolicy Bypass -File scripts\fetch-doitacgiaohang.ps1
```

## Cài đặt

1. Mở `chrome://extensions/`.
2. Bật Developer mode.
3. Chọn Load unpacked và trỏ tới thư mục dự án.

## Yêu cầu

- Google Chrome
- Quyền clipboard
- Quyền storage (để cache dữ liệu đối tác)

## Nhà cung cấp

- Trường tùy chọn, gợi ý từ `nhacungcap.csv` hoặc sheet `Nha_Cung_Cap`.
- Khi copy, nhà cung cấp nằm sau đối tác giao hàng.
- Nút tải dữ liệu sẽ cập nhật đối tác giao hàng, nhà cung cấp và mã sản phẩm.
- Danh sách nhà cung cấp được sort theo độ dài ký tự tăng dần, sau đó theo A-Z.

## Mã sản phẩm

- Trường gợi ý từ `sanpham.csv` hoặc sheet `San_Pham`.
- Danh sách mã sản phẩm được sort theo độ dài ký tự trước, sau đó theo thứ tự A-Z để đề xuất ngắn hơn hiện lên trước.
- Trước khi ghi vào CSV, mã sản phẩm sẽ bỏ các hậu tố ` [IMEI]`, ` [CŨ]`, ` MN`, ` ML` rồi unique.
