# Hướng dẫn cho Claude trong kho vtvl-claude

Kho này giúp Claude làm hồ sơ vị trí việc làm (VTVL) theo Hướng dẫn 07-HD/BTCTW: lập bộ biểu, thẩm định, sửa, phiếu xác minh, bản nộp, báo cáo thẩm định.

## Làm việc

1. Đọc `.claude/skills/vtvl/SKILL.md` và làm theo bước phân luồng.
2. Có công cụ sẵn → dùng (`python3 scripts/vtvl_cli.py --help`). Chạy `--dry-run` trước mọi lệnh ghi.
3. Chưa có → tự dựng script trong `scripts/` (tên hàm, biến, tệp TIẾNG ANH; chú thích, chữ hiện ra TIẾNG VIỆT), thêm kiểm thử trong `tests/` chạy trên dữ liệu GIẢ (`vtvl_cli.py sample`).
4. Mẫu bộ biểu khác (tỉnh khác, phiên bản khác) → `templates/<ten-mau>/map.json`, không sửa mẫu cũ.

## Đóng góp kinh nghiệm (PR)

Sau mỗi ca có điều đáng nhớ (lỗi hay gặp, cách sửa, mẫu mới, ngưỡng mới):
1. Tạo nhánh `kinh-nghiem/<chu-de>`.
2. Viết `KINH-NGHIEM/<chu-de>-<yyyy-mm-dd>.md` theo khuôn `KINH-NGHIEM/_KHUON.md`: tình huống (đã bỏ danh tính) / làm gì / kết quả / bài học / đề xuất sửa công cụ.
3. Sửa script/skill nếu cần, kèm kiểm thử.
4. `python3 scripts/pii_scan.py` phải OK; `python3 -m pytest -q` phải xanh.
5. `gh pr create` – mô tả nêu rõ: thay đổi gì, đã kiểm trên dữ liệu giả thế nào.

## TUYỆT ĐỐI KHÔNG đưa lên kho

- Tệp của khách/đơn vị thật (xlsx, docx, pdf, rar...), ảnh chụp, tin nhắn.
- Tên người, số điện thoại, email, CCCD, mã số thuế, tên đơn vị thật, tên tỉnh, số biên chế/số liệu thật.
- Ví dụ chỉ dùng dữ liệu giả: đơn vị "Bệnh viện Mẫu", mã `BVM-...`.
- Máy quét `scripts/pii_scan.py` chạy ở pre-commit và GitHub Actions; dính là FAIL. Không tắt, không lách. Thuật ngữ công khai bị bắt nhầm thì thêm vào `.pii/allowlist.txt` và giải thích trong PR.
- Chuỗi nhạy cảm cần chặn thêm: `python3 scripts/pii_block_add.py "<chuỗi>"` (lưu dạng băm).

## Kỹ thuật cần nhớ

- Bộ biểu đầy ~40 MB / 30.000 dòng: đọc bằng openpyxl `read_only=True`, ghi bằng sửa XML (`scripts/vtvl/xlsx_patch.py`). Không mở bằng openpyxl thường để ghi; không chạy công cụ tính lại toàn tệp (tốn nhiều GB RAM).
- Không ghi vào ô công thức; không ghi cột không phải cột NHẬP của mẫu (theo `map.json`).
- Số công thức lưu sẵn trong tệp không phải bằng chứng đã tính lại: tự tính lại từ cột nhập.
