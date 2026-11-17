---
name: vtvl
description: Hồ sơ VỊ TRÍ VIỆC LÀM (VTVL) theo Hướng dẫn 07-HD/BTCTW – bộ biểu Excel ~30 trang (A1, A2, B1, B2, C1, C2, C3, D1, D2, DM_SAN_PHAM, DU_BAO_5_NAM, K_KiemTra...). Dùng khi được nhờ LẬP bộ biểu từ quyết định chức năng nhiệm vụ, THẨM ĐỊNH hồ sơ (3 vòng, ra nhận xét + bảng lỗi), SỬA lỗi kỹ thuật, làm PHIẾU XÁC MINH cho đơn vị, ÁP phiếu đã trả lời, làm BẢN NỘP sạch, hoặc BÁO CÁO thẩm định theo mẫu Sở (Word + Excel 6 trang, phân loại A/B/C/D). Trigger: "vị trí việc làm", "VTVL", "bộ biểu", "HD07", "07-HD/BTCTW", "thẩm định hồ sơ", "phân rã nhiệm vụ", "phiếu xác minh", "bản nộp", "biên chế", "PP1 PP2 PP3".
---

# /vtvl – hồ sơ vị trí việc làm

Gốc kho = thư mục chứa `scripts/vtvl_cli.py`. Mọi lệnh: `python3 scripts/vtvl_cli.py <lệnh> ...`
(cần `pip install openpyxl python-docx lxml`). Lệnh ghi tệp đều có `--dry-run`: **chạy khô trước, đọc kế hoạch, rồi mới chạy thật.**

## Bước 0 – Phân luồng (hỏi/suy ra trước khi làm gì)

1. **Mình đứng phía nào?**
   - Phía ĐƠN VỊ lập hồ sơ (bệnh viện, trung tâm, phòng ban...) → mục tiêu: **số khớp giữa các biểu + tệp hết lỗi để nộp**.
   - Phía CƠ QUAN THẨM ĐỊNH (Sở, Ban Tổ chức...) → mục tiêu: nhận xét có căn cứ, phân loại A/B/C/D.
2. **Đang có tệp gì?** (kiểm kê, ghi lại tên tệp + phiên bản mẫu + số trang)
   - Chỉ có quyết định chức năng nhiệm vụ + mẫu trống → nhánh **L (lập)**.
   - Có bộ biểu đã điền → nhánh **T (thẩm định)** rồi tuỳ phía: **S (sửa)** / **P (phiếu)** / **N (nộp)** / **B (báo cáo Sở)**.
   - Có phiếu xác minh đơn vị đã trả lời → nhánh **P2 (áp phiếu)** rồi **N**.
3. **Mẫu nào?** So tên trang với `templates/*/map.json`. Khớp thì dùng; lệch (tỉnh khác, phiên bản khác) → chép `templates/mau-30000-dong/` thành `templates/<ten-mau>/`, sửa tên trang/chữ cột trong `map.json`, chạy `--map templates/<ten-mau>/map.json`.

## Nhánh L – Lập bộ biểu từ quyết định chức năng nhiệm vụ

1. Đọc `templates/mau-30000-dong/BAN-DO-MAU.md` (trang, cột nhập/công thức, tên vùng, chuỗi mã, quỹ giờ).
2. Đọc văn bản chức năng nhiệm vụ CỦA CHÍNH đơn vị (`pdftotext -layout`). Mỗi nhiệm vụ A1 dẫn đúng điều/khoản.
3. Viết `spec.json` (xem `examples/spec-benh-vien-mau.json`), mã tạm tiền tố `TAM-`, số liệu để trống "đơn vị cung cấp".
4. `python3 scripts/fill_template.py spec.json --no-gridcraft` (sửa XML, ~12 giây với tệp 40 MB; không ghi đè mẫu; chỉ ghi cột NHẬP).
5. Kiểm lại: `vtvl_cli.py audit` trên tệp ra.

## Nhánh T – Thẩm định 3 vòng

```
python3 scripts/vtvl_cli.py audit HOSO.xlsx -o findings.json
python3 scripts/vtvl_cli.py error-table findings.json -o Bang-loi.xlsx
python3 scripts/vtvl_cli.py review findings.json -o Nhan-xet.docx --unit "<tên đơn vị>"
```
- Vòng 1 kỹ thuật: giá trị lỗi (#REF!...), công thức bị ghi đè, mã trùng/khoảng trắng/LDQL–LĐQL, mã "VD-".
- Vòng 2 nhất quán: A2→A1, A2→C1, B1→A2, B1→DM, tiền tố mã lạc (chép từ đơn vị khác), A1 không có công việc, C1 chưa có mã / chưa có C3.
- Vòng 3 nghiệp vụ (ngưỡng là ĐỀ XUẤT): phút/SP > 480, dòng > quỹ giờ, giờ bổ sung = phút/SP (nhập nhầm cột), nguồn ghi chung chung, J/G lệch > 10%, PP2 mà nhập PP1, "Đã kiểm chứng" khi thời gian là ước lượng, PP3 > 20%, gán đồng loạt cho lãnh đạo.
- Máy chỉ khoanh vùng. **Claude phải đọc thêm** quyết định chức năng nhiệm vụ, bản mô tả VTVL, danh sách nhân sự để thêm phát hiện nghiệp vụ (nhiệm vụ ngoài chức năng, cơ cấu lấy của đơn vị khác, mô tả chép từ đơn vị khác, căn cứ sai số hiệu/năm...). Ghi vào findings.json cùng khuôn.
- Theo `docs/cau-lenh-tham-dinh.md`: phân biệt lỗi nhập liệu / lỗi công thức của mẫu / lỗi phương pháp / nội dung cần giải trình; nghi vấn thì ghi "cần xác minh", không kết luận vi phạm.

## Nhánh S – Sửa kỹ thuật (tô vàng + comment giá trị cũ)

```
python3 scripts/vtvl_cli.py fix-tech HOSO.xlsx findings.json --dry-run
python3 scripts/vtvl_cli.py fix-tech HOSO.xlsx findings.json -o HOSO-da-sua-ky-thuat.xlsx
python3 scripts/vtvl_cli.py patch HOSO.xlsx edits.json -o OUT.xlsx      # sửa tay có căn cứ
```
`edits.json` = `[{"sheet": "1. A1_NhiemVu", "cells": "D3:D17", "value": "...", "note": "căn cứ..."}]`
(`value` bỏ trống = chỉ tô; `null` = xoá trống; `color`: yellow | pink).
Chỉ sửa cái CHẮC CHẮN (số hiệu văn bản đối chiếu bản gốc, mã khác dạng, chuỗi dính liền). Không sửa số liệu nghiệp vụ, cơ cấu, định mức, phương pháp khi chưa có căn cứ: đưa vào phiếu.

## Nhánh P – Phiếu xác minh "Đồng ý / điền số"

```
python3 scripts/vtvl_cli.py form findings.json -o Phieu-xac-minh.xlsx --unit "<tên đơn vị>"
```
- Mỗi điểm một dòng, có "Đề xuất". Đơn vị chỉ ghi **Đồng ý** hoặc **số/nội dung đúng**. Nhắm ≤ 30 dòng: gộp điểm cùng gốc.
- Cột ẩn `_sheet/_cells/_proposed` để áp lại tự động – không xoá.

## Nhánh P2 – Áp phiếu đã trả lời

```
python3 scripts/vtvl_cli.py apply-form HOSO.xlsx Phieu-da-tra-loi.xlsx --source-label "phiếu xác minh ngày dd/mm" --dry-run
python3 scripts/vtvl_cli.py apply-form ... -o HOSO-theo-phieu.xlsx
```
"Đồng ý" → áp đề xuất (tô vàng); có số → ghi số; trống → tô hồng "CHỜ SỐ THẬT TỪ ĐƠN VỊ". Ô công thức / cột không phải cột NHẬP tự bỏ qua và in ra.

## Nhánh N – Bản nộp sạch

```
python3 scripts/vtvl_cli.py clean HOSO-theo-phieu.xlsx HOSO-da-sua-ky-thuat.xlsx.log.json HOSO-theo-phieu.xlsx.log.json -o HOSO-ban-nop.xlsx [--final-edits final.json]
```
Bỏ màu và comment do mình thêm (giữ comment sẵn có của mẫu/đơn vị), áp sửa cuối. Sau đó chạy lại `audit` trên bản nộp: **mục tiêu số khớp + không còn lỗi bắt buộc**. Phần chưa có minh chứng: ghi theo xác nhận của đơn vị và chú thích "giải trình bổ sung", KHÔNG bịa số liệu gốc.

## Nhánh B – Báo cáo thẩm định theo mẫu Sở

```
python3 scripts/vtvl_cli.py report findings.json --docx BaoCao.docx --xlsx ThamDinh.xlsx --meta meta.json
```
`meta.json`: `unit, date, basis, sources, owner, verify_level, task_rows` (bảng đối chiếu pháp lý ↔ A1 ↔ A2 do Claude lập từ quyết định). Khung cột ở `templates/mau-bao-cao-so/report.json`; Sở dùng khung khác thì chép thư mục và sửa.
Phân loại: A đủ điều kiện; B còn cảnh báo; C còn lỗi bắt buộc nghiệp vụ; D lỗi kỹ thuật nghiêm trọng / đứt liên kết.

## Luật cứng

1. **Không bao giờ ghi đè tệp gốc/mẫu Trung ương**; tệp ra luôn là bản mới + nhật ký `.log.json` từng ô.
2. **Không chạy công cụ tính lại công thức toàn tệp** (gridcraft convert/check, LibreOffice headless) trên bộ biểu đầy 30.000 dòng: ngốn nhiều GB RAM, từng làm sập máy. Đọc bằng openpyxl `read_only=True`; tệp ra đặt `fullCalcOnLoad` để Excel tự tính. Ghi rõ "chưa tính lại công thức trên máy".
3. **Chạy khô trước** mọi lệnh ghi; tệp sẽ gửi người khác thì mở lại kiểm (openpyxl đọc được, comment đúng ô).
4. Không tự đặt khối lượng, thời gian, số người, định mức, quỹ giờ. Thiếu thì để trống + phiếu.
5. Căn cứ phải đối chiếu **văn bản gốc** (số hiệu, ngày, điều/khoản) trước khi sửa hay khẳng định.
6. Phía đơn vị: linh động phần xác thực theo xác nhận của đơn vị; cứng ở chỗ số phải khớp và không bịa.
7. Mã nguồn: tên hàm/biến/tệp tiếng Anh, chú thích tiếng Việt. Chữ hiện cho người dùng: tiếng Việt có dấu.

## Chưa có công cụ cho việc đang làm?

Tự dựng script nhỏ (theo luật trên), chạy trên dữ liệu giả `vtvl_cli.py sample`, rồi ghi kinh nghiệm vào `KINH-NGHIEM/<chu-de>-<yyyy-mm-dd>.md` và mở PR theo `CLAUDE.md` của kho. Không đưa dữ liệu khách vào kho.
