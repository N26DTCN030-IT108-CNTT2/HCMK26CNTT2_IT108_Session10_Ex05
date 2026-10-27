# HCMK26CNTT2_IT108_Session10_Ex05
Bài tập số 5
# Bộ Hồ Sơ Báo Cáo Sự Cố An Ninh Mạng Trung Tâm NetGuard (Session 10 Ex05)

Bài giải chuẩn hóa hệ thống báo cáo sự cố an ninh mạng (Incident Report), tra cứu tự động mã độc từ nhật ký hệ thống và xây dựng chiến lược phòng thủ 3 tầng cho Trung tâm NetGuard.

---

# Phần 1. Báo Cáo Phân Tích Kỹ Thuật (System Analysis)

## 1. Cơ Chế "Google Search Của Bảng Tính" Thông Qua Tính Năng AutoFilter

Tính năng **AutoFilter** trong Excel đóng vai trò tương tự như một công cụ tìm kiếm "Google Search" thu nhỏ trên dữ liệu cấu trúc nhờ các nguyên lý kỹ thuật sau:

- **Khả năng quét và bóc tách chỉ mục (Indexing & Querying):** Khi bật AutoFilter (`Ctrl + Shift + L`), Excel sẽ tự động quét toàn bộ bảng dữ liệu, tạo một danh mục chứa các giá trị duy nhất (Unique Values) cho từng cột (Field). Khi người dùng nhập từ khóa tìm kiếm vào ô Search box của bộ lọc, hệ thống thực hiện thuật toán lọc tức thời và ẩn đi các dòng không thỏa mãn.
- **Hỗ trợ truy vấn nâng cao (Advanced Operators):** AutoFilter cho phép truy vấn điều kiện phức tạp giống như toán tử tìm kiếm:
  - *Text Filters:* Lọc theo các phép so sánh chuỗi `Contains` (chứa từ khóa), `Begins With` (bắt đầu bằng), `Equals` (khớp chính xác).
  - *Number Filters:* Lọc khoảng giá trị `Greater Than`, `Less Than`, `Between`.
  - *Logical AND/OR:* Lọc kết hợp nhiều điều kiện đồng thời trên nhiều cột khác nhau (ví dụ: Lọc các ca nhiễm thuộc phân vùng IP `192.168.1.x` AND có trạng thái `"Cần cách ly"`).

---

## 2. Giải Trình Thuật Toán Tra Cứu VLOOKUP Kết Hợp TRIM Chống Lỗi Log Thô

- **Hiện tượng log thô:** Dữ liệu nhật ký hệ thống (System Log) được ghi nhận tự động từ các thiết bị tường lửa thường bị dính khoảng trắng ẩn ở đầu/cuối mã cảnh báo (Ví dụ: ô `A2` chứa mã `' LOG001 '` hoặc `'LOG003 '`).
- **Rủi ro nếu dùng VLOOKUP thông thường:** Hàm `=VLOOKUP(A2, CSDL_Virus!A$2:C$5, 2, FALSE)` sẽ so sánh chuỗi `" LOG001 "` với `"LOG001"` trong CSDL. Do sự bất đồng về mã Unicode (khoảng trắng ASCII 32), VLOOKUP sẽ coi đây là 2 chuỗi khác nhau và trả về lỗi `#N/A`.
- **Giải pháp bọc TRIM + IFERROR chuẩn hóa:**
  ```excel
  =IFERROR(VLOOKUP(TRIM(A2), CSDL_Virus!A$2:C$5, 2, FALSE), "Chưa xác định")
