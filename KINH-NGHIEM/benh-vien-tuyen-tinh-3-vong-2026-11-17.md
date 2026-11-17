# Thẩm định + sửa + bản nộp cho một bệnh viện tuyến tỉnh – 2026-11-17

- **Tình huống**: bệnh viện tuyến tỉnh, bộ biểu theo mẫu 30.000 dòng; người nhờ ở phía đơn vị cần nộp gấp trong ngày.
- **Phía**: đơn vị lập (có cả bản báo cáo thẩm định theo mẫu Sở để tự soát).
- **Đã làm**: thẩm định 3 vòng → bảng lỗi ~30 điểm → sửa kỹ thuật (tô vàng + comment giá trị cũ) → phiếu xác minh 28 điểm "Đồng ý / điền số" → áp phiếu 4 lượt → bản nộp sạch → báo cáo theo mẫu Sở (Word + Excel 6 trang).
- **Kết quả**: lỗi bắt buộc về 0; còn cảnh báo cần giải trình (nguồn thời gian là ước lượng, hồ sơ PP3) – phân loại B có điều kiện.
- **Bài học**:
  - Lỗi hay gặp nhất: căn cứ A1 sai năm/số hiệu và hai chuỗi bị dính liền; lỗi lặp sang D1, C3.
  - Mã công việc khác dạng và mang tiền tố đơn vị khác (chép từ hồ sơ đơn vị khác) → B1 báo "THIẾU MÃ VTVL", giờ không vào C3.
  - B1 cột E (giờ bổ sung) hay bị nhập bằng thời gian 1 sản phẩm → cộng trùng.
  - Mô tả C2 chép từ đơn vị khác chuyên ngành (lộ ra qua từ khoá chuyên môn lạ).
  - C3 chọn PP2 mà nhập số người PP1, giải trình "theo PP1".
  - Với người ở phía đơn vị: mục tiêu là SỐ KHỚP + HẾT LỖI; phần xác thực ghi theo xác nhận của đơn vị, chú thích "giải trình bổ sung"; không bịa số gốc.
  - Phiếu "Đồng ý / điền số" làm đơn vị trả lời nhanh hơn hẳn văn bản nhận xét dài.
  - Đơn vị trả lời phiếu nhiều lượt: giữ nhật ký từng lượt; bản nộp khôi phục kiểu gốc theo lần sửa ĐẦU TIÊN.
  - Không chạy công cụ tính lại công thức trên tệp đầy (nặng RAM): tự tính G = giờ/quỹ giờ bằng tay để giải trình.
- **Đề xuất sửa công cụ**: đã gom thành `scripts/vtvl_cli.py` (audit, fix-tech, form, apply-form, clean, report).
