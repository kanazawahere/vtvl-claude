# Bản đồ mẫu bộ biểu VTVL – bản 30.000 dòng (30.000 dòng)

Tệp gốc đã phân tích: `bộ biểu 30.000 dòng (bản sửa ngày 23/10/2026)` (39 MB, 31 trang).
Nguồn của bản đồ này: chính các trang hướng dẫn trong tệp (`HUONG_DAN`, `HD_DIEN_BIEU`, `QUY_TAC_MA`,
`DM_THAM_SO`, `GHI_CHU_BAN_SUA`, `BAN_DO_TRUONG`) + quét XML dòng 3 để biết cột nào là công thức (13/11/2026).
Chỗ nào là suy luận của mình thì ghi rõ "(suy ra)".

## 1. Quy ước chung (chữ của tệp)

- Ô TRẮNG = nhập/chọn; ô XÁM = công thức tự động, đã khoá, không nhập tay.
- Vùng nhập mỗi biểu: dòng 3 → 30.002 (12. CHIA_VUNG_CAP_XA: dòng 4 → 30.003). Dòng 1 = tên biểu, dòng 2 = tiêu đề cột.
- Mật khẩu khoá trang (ghi công khai trong HUONG_DAN mục "CHÚ Ý"): `VTVL2026`. Trang khoá: A2, B1, B2_CT, B2_ThayDoi, C1, C2, C3, D1, D2, DM_SAN_PHAM, CHIA_VUNG_CAP_XA, K, T, S, DM_THAM_SO, DU_BAO_5_NAM, DOI_CHIEU_NV. A1 KHÔNG khoá.
- Trang ẩn: 13. K_KiemTra, 14. T_TienDo, 15. S_NhatKy, DU_BAO_5_NAM.
- Sửa công thức: sửa dòng 3 rồi chép xuống hết 30.002, và phải báo đầu mối quản lý biểu (Tổ Công tác) theo HD 07-HD/BTCTW. Không chèn/xoá dòng, cột.
- Mọi trang có dòng "LƯU Ý: Các ví dụ chỉ để minh họa; không được dùng số liệu giả định…".

## 2. Quỹ giờ

- Tên vùng `QUY_GIO_PP2` = `DM_THAM_SO!$B$3` = **1904** (giờ/người/năm, chế độ phổ biến).
  Căn cứ ghi trong ô C3: (261 − 11 lễ Tết − 12 phép) × 8 = 1.904; "mức tham chiếu, không phải định mức biên chế".
- `C1!E` dùng `QUY_GIO_PP2` khi chế độ "Phổ biến (40 giờ/tuần)"; chế độ "Đặc thù" chỉ nhận giờ riêng ở `C1!H` khi có căn cứ `F` và `I` = "Đã kiểm chứng".
- `D1!D` lấy giờ riêng `D1!K` nếu có, không thì lấy `C1!E` của VTVL.
- BAN_DO_TRUONG: không tự giảm bằng tỷ lệ "giờ hữu ích", không mặc định 1.600 giờ (HD 07, Bước 6 điểm b).

## 3. Từng trang

| Trang | Vai trò (theo HUONG_DAN/HD_DIEN_BIEU) | Cột NHẬP | Cột CÔNG THỨC (không ghi đè) |
|---|---|---|---|
| HUONG_DAN | Mục lục, chỉ dẫn, quỹ giờ PP2, mật khẩu | – | ô quỹ giờ `=QUY_GIO_PP2` |
| HD_DIEN_BIEU | Hướng dẫn từng cột của mọi biểu (257 dòng) | – | – |
| 1. A1_NhiemVu | Danh mục nhiệm vụ: mã, tên, cơ quan, căn cứ, kỳ, trạng thái, nguồn | A–G | – |
| 2. A2_CongViec | Phân rã nhiệm vụ → công việc, sản phẩm, ĐVT, chất lượng, căn cứ; gắn mã NV và mã VTVL | A–H (B chọn DS_MA_NV, C chọn DS_MA_VTVL) | – |
| 3. B1_KhoiLuong | Số SP/năm, phút/SP, giờ bổ sung, nguồn, trạng thái căn cứ, kỳ, mã SP | A C D E H I J K L S | B F G M N O P Q R |
| 4. B2_CT_TacDong | Tác động thực tế (số hoá, thuê dịch vụ…) theo công đoạn | A C D E F G I J K L M Q R S | B H N O P |
| 5. B2_ThayDoi | Tổng hợp tự động từ B2_CT | – | A–I |
| 6. C1_VTVL | Danh mục VTVL, loại nhân lực, quyết định, quỹ giờ | A B C D F G H I J | E K |
| 7. C2_NangLuc | Trách nhiệm, năng lực theo bản mô tả | A–G | – |
| 8. C3_NhuCau | PP1/PP2/PP3, giờ ròng, người quy đổi, người cần bố trí | A B H I J K L | C D E F G M N |
| 9. D1_NhanLuc | Nguồn lực hiện có, phân bổ giờ, nơi tính biên chế | A B C E F G H I K | D J |
| 10. D2_PhuongAn | Phương án 2027–2031, biên chế giao, đề nghị | A B D E F G H I N | C J K L M |
| 11. DM_SAN_PHAM | Danh mục sản phẩm dùng chung: mã, tên, nhóm I–VI, phút chuẩn, nguồn, trạng thái | A–J (đầu mối nhập; cơ sở chỉ đọc) | – |
| 12. CHIA_VUNG_CAP_XA | Xã/phường/đặc khu theo 4 vùng | A–C (từ dòng 4); D–E để trống; F chỉ đọc | – |
| 13. K_KiemTra (ẩn) | Kiểm tra lỗi, cảnh báo nhanh | – | H … |
| 14. T_TienDo (ẩn) | Tiến độ; chỉ nhập A (đơn vị), B (kỳ) | A B | C–G |
| 15. S_NhatKy (ẩn) | Nhật ký sửa dữ liệu đã xác nhận | A–J | – |
| KT_GIA_DINH | Dữ liệu giả định để thử công thức – KHÔNG dùng làm số thật | – | – |
| Y_KIEN_DON_VI | Góp ý của đơn vị (dòng 3–201) | A–K | – |
| 00_KE_THUA | Các bước nạp/đối chiếu dữ liệu cũ (tài liệu) | – | – |
| NGUON_NHIEM_VU / NGUON_VTVL / NGUON_NHAN_LUC | Dữ liệu nguồn do đầu mối nhập | toàn bộ | – |
| DOI_CHIEU_NV | Đối chiếu nhiệm vụ từ nguồn (200 dòng) | – | A–H |
| BAN_DO_TRUONG | Nguồn nào → biểu nào, phần cần bổ sung | – | – |
| DM_06_NHOM | Khung 6 nhóm độ phức tạp sản phẩm (tham khảo) | – | – |
| QUY_TAC_MA | Quy tắc mã; mọi mã "VD-…" chỉ minh hoạ | – | – |
| THAO_TAC_NHANH | 08 bước thao tác | – | – |
| DM_THAM_SO | Tham số + danh sách chọn (khoá) | – | – |
| DU_BAO_5_NAM (ẩn) | Dự báo 2027–2031 từ C3, chỉ đầu mối tổng hợp | A–F, J | G H I |
| GHI_CHU_BAN_SUA | 37 điểm bản sửa so với bản của Ban (19/10/2026) | – | – |

## 4. Tên vùng (named range) và danh sách chọn

| Tên | Trỏ tới | Giá trị |
|---|---|---|
| QUY_GIO_PP2 | DM_THAM_SO!$B$3 | 1904 |
| DS_MA_NV | '1. A1_NhiemVu'!$A$3:$A$30002 | mã nhiệm vụ |
| DS_MA_CV | '2. A2_CongViec'!$A$3:$A$30003 | mã công việc |
| DS_MA_VTVL | '6. C1_VTVL'!$A$3:$A$30002 | mã VTVL |
| DS_MA_SP | '11. DM_SAN_PHAM'!$A$3:$A$30002 | mã sản phẩm |
| DS_TT_A1 | DM_THAM_SO!$N$9:$N$12 | Tiếp tục thực hiện / Mới tiếp nhận / Đã chuyển giao / Không còn thực hiện |
| DS_TT_B1 | $A$9:$A$10 | Có định mức / Kiểm đếm, ước lượng thực tế |
| DS_KIEM_CHUNG | $B$9:$B$10 | Đã kiểm chứng / Chưa kiểm chứng |
| DS_TT_B2CT | $C$9:$C$11 | Đã kiểm chứng / Chưa kiểm chứng / Ước lượng tham khảo |
| DS_LOAI_TD | $D$9:$D$14 | Số hóa / Cải tiến quy trình / Thuê dịch vụ/đặt hàng / Tự chủ / Xã hội hóa / Công việc mới phát sinh |
| DS_CO_SO_B1 | $E$9:$E$10 | Trước thay đổi, có điều chỉnh / Đã đo sau thay đổi (không trừ) |
| DS_DOI_CHIEU | $F$9:$F$10 | Chờ đối chiếu / Đã đối chiếu |
| DS_TT_C1 | $G$9:$G$13 | Tiếp tục sử dụng / Cần cập nhật / Đề nghị bổ sung / Đề nghị sửa đổi / Đề nghị bãi bỏ |
| DS_CHE_DO | $H$9:$H$10 | Phổ biến (40 giờ/tuần) / Đặc thù |
| DS_PP | $I$9:$I$15 | PP1 / PP2 / PP3 / PP1+PP2 / PP1+PP3 / PP2+PP3 / PP1+PP2+PP3 |
| DS_TT_C3 | $J$9:$J$12 | Tính thử / Chưa đủ dữ liệu / Đã kiểm chứng / Không áp dụng |
| DS_TT_D2 | $K$9:$K$12 | Đang xây dựng / Đã hoàn thành ở đơn vị / Đã thẩm định / Đã được quyết định |
| DS_NAM | $L$9:$L$13 | 2027…2031 |
| DS_LOAI_NL | $M$9:$M$13 | Cán bộ / Công chức / Viên chức / Hợp đồng lao động / Khác |
| DS_NHOM | $O$9:$O$15 | I…VI / Không phân nhóm |
| DS_TT_DM | $P$9:$P$11 | Đang kiểm chứng / Đã xác nhận / Không áp dụng |
| DS_TIEP_THU | $Q$9:$Q$13 | Chờ tiếp thu / Tiếp thu / Tiếp thu một phần / Không tiếp thu / Giải trình |
| DS_BIEU | $R$9:$R$20 | tên 12 biểu |

Ô chọn có bật báo lỗi (GHI_CHU_BAN_SUA mục 3–34): A1!F, A2!B, A2!C, B1!A, B1!J, B1!L, B2_CT!A C L M S,
C1!C G I J, C2!A, C3!A B L, D1!B C I, D2!A B I, DM_SAN_PHAM!D G, DU_BAO_5_NAM!A B F, Y_KIEN_DON_VI!D K.
=> Giá trị ghi vào các ô này phải đúng chữ trong danh sách (kể cả dấu), không thì Excel báo lỗi khi sửa ô.

## 5. Chuỗi liên kết mã giữa các biểu

```
A1.A (mã NV) ──> A2.B
A2.A (mã CV) ──> B1.A, B2_CT.A          (B1.B và B2_CT.B tự tra mã VTVL = A2.C theo mã CV)
A2.C (mã VTVL) <── C1.A
DM_SAN_PHAM.A (mã SP) ──> B1.L          (B1.M N O tự tra nhóm, phút chuẩn, trạng thái chuẩn)
B1.R = "ĐẠT" ──> B1.F G có giờ ──> C3.C (SUMIFS theo B1.B = mã VTVL)
B2_CT.P = "ĐẠT" ──> B2_ThayDoi ──> C3.D
C1.E (quỹ giờ) ──> C3.F, D1.D
C3.G = E/F (người quy đổi PP2, giữ số lẻ); C3.J (người cần bố trí) nhập tay sau thẩm định
C3 (Đã kiểm chứng + M = ĐẠT) ──> DU_BAO_5_NAM.G ──> D2.C ; C3.G, C3.J ──> D2.K, D2.L
```
Hệ quả: A2 thiếu mã VTVL (cột C) thì B1 báo "THIẾU MÃ VTVL", C3 không cộng được giờ. Đây là đúng thiết kế, không phải lỗi.

## 6. Lỗi đã thấy trong tệp (13/11/2026)

- `9. D1_NhanLuc!J3:J12` công thức có `#REF!`: `=IF(#REF!="","",IF(COUNTIF($A$3:$A$30002,#REF!)>1,…))`;
  `J13:J30002` là công thức dùng chung (`si=0`, gốc J13) nhưng trỏ `A3` thay vì `A13` → lệch 10 dòng.
  (suy ra) có 10 dòng bị xoá sau khi đặt công thức; đúng ra mỗi dòng kiểm `A` cùng dòng. Sửa phải qua đầu mối quản lý biểu.
- gridcraft-cli 0.8.0 trên tệp đầy 30.000 dòng: `check` chạy quá 290 giây (máy 23 GB RAM, tiến trình ~6 GB) thì bị cắt.
  Bản ít dòng thì chạy vài giây. Muốn kiểm nhanh: dùng `fill_template.py` rồi đọc lại bằng openpyxl read_only.

## 7. Script điền tái dùng

`fill_template.py <spec.json> [--no-gridcraft] [--allow-any-column]` — sửa thẳng XML, không đụng ô công thức,
chỉ cho ghi cột NHẬP ở bảng mục 3, lưu bản cũ ra trang `CU_*` trước khi xoá, thêm trang ghi chú, đặt tính lại khi mở.
Ví dụ đầy đủ: `examples/spec-benh-vien-mau.json` (dữ liệu giả, đơn vị "Bệnh viện Mẫu").

Luật nội dung khi điền (theo "Câu lệnh thẩm định hồ sơ VTVL theo HD07"):
1. Không sửa tệp mẫu; tạo bản sao ghi rõ BẢN ĐỀ XUẤT.
2. Nhiệm vụ A1 dẫn đúng điều/khoản văn bản chức năng nhiệm vụ của CHÍNH đơn vị.
3. Không tự đặt khối lượng, tần suất, thời gian, số người, định mức, quỹ giờ: để trống, ghi "đơn vị cung cấp".
4. Mã tự đặt là mã tạm (tiền tố `TAM-`), duy nhất, có bảng diễn giải.
5. Văn bản không nói phòng nào làm thì để "đơn vị xác định", không gán đồng loạt cho một chức danh.
6. Không ghi chung "thực hiện nhiệm vụ khác" ở A2 khi chưa rõ công việc thực tế (HD_DIEN_BIEU, A1 cột B).
