# Quản lý Trung tâm Học thêm

Phiên bản MVP quản lý:

- Học sinh và mã tra cứu 5 chữ số.
- Lớp học, môn học, lịch học và mức học phí.
- Xếp học sinh vào lớp.
- Điểm danh theo từng buổi.
- Theo dõi học phí đã đóng/chưa đóng.
- Cổng PHHS tại đường dẫn `?phhs=1`.

## Cài đặt Google Sheet

1. Tạo một Google Sheet trống.
2. Mở **Extensions → Apps Script**.
3. Sao chép toàn bộ nội dung `Code.gs` của repository vào trình soạn thảo.
4. Trong **Project Settings → Script properties**, tạo:
   - Property: `ADMIN_KEY`
   - Value: một mã quản trị bí mật do bạn chọn.
5. Chạy hàm `setupSheets` một lần và cấp quyền. Hệ thống tự tạo:
   - `HOC_SINH`
   - `LOP_HOC`
   - `DANG_KY_LOP`
   - `DIEM_DANH`
   - `HOC_PHI`
   - `CAI_DAT_TRUNG_TAM`
6. Chọn **Deploy → New deployment → Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone**
7. Dán URL Web app và mã quản trị vào **Cài đặt kết nối** trên trang quản trị.

Trước khi gửi link PHHS, điền URL Web app vào hằng số `PUBLIC_API_URL`
trong `index.html` để trình duyệt của phụ huynh có thể đọc dữ liệu trực tiếp.

## Đường dẫn sau khi bật GitHub Pages

- Giáo viên: `https://thanhngthien84-lab.github.io/quan-ly-trung-tam-hoc-them/`
- PHHS: `https://thanhngthien84-lab.github.io/quan-ly-trung-tam-hoc-them/?phhs=1`

Không gửi mã `ADMIN_KEY` cho phụ huynh.

