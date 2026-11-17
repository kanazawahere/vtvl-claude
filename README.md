# vtvl-claude – Claude Code làm hồ sơ vị trí việc làm (HD 07-HD/BTCTW)

Bộ cấu hình + công cụ để **Claude Code** giúp cán bộ làm hồ sơ **vị trí việc làm (VTVL)** theo Hướng dẫn 07-HD/BTCTW:
bộ biểu Excel ~30 trang (A1, A2, B1, B2, C1, C2, C3, D1, D2, DM_SAN_PHAM, DU_BAO_5_NAM...).

Làm được:
- **Lập** bộ biểu đề xuất từ quyết định chức năng nhiệm vụ (mã tạm, số liệu để đơn vị điền).
- **Thẩm định 3 vòng**: kỹ thuật / nhất quán giữa các biểu / nghiệp vụ → bảng lỗi Excel + nhận xét Word.
- **Sửa lỗi kỹ thuật** ngay trên tệp: ô sửa tô vàng, comment ghi giá trị cũ, nhật ký từng ô.
- **Phiếu xác minh** "Đồng ý / điền số" cho đơn vị, rồi **áp phiếu** tự động.
- **Bản nộp sạch**: bỏ màu, bỏ comment của mình, số khớp, hết lỗi bắt buộc.
- **Báo cáo thẩm định** theo khung Sở: Word + Excel 6 trang, phân loại A/B/C/D.

Chạy được với tệp đầy 30.000 dòng (~40 MB) mà không ngốn RAM: đọc chế độ chỉ-đọc, ghi bằng sửa thẳng XML.

## Cài đặt

```bash
git clone https://github.com/<tai-khoan>/vtvl-claude.git
cd vtvl-claude
python3 -m pip install -r requirements.txt
git config core.hooksPath .githooks      # bật chốt chặn dữ liệu cá nhân trước mỗi commit
```

Dùng với Claude Code: mở thư mục này rồi gõ `/vtvl` (hoặc chỉ cần nói "thẩm định bộ biểu VTVL này").
Muốn dùng skill ở mọi dự án: chép `.claude/skills/vtvl` vào `~/.claude/skills/` và sửa đường dẫn tới `scripts/`.

## Dùng thử trong 1 phút (dữ liệu giả)

```bash
python3 scripts/vtvl_cli.py sample mau.xlsx                    # bộ biểu giả "Bệnh viện Mẫu", có cài sẵn lỗi
python3 scripts/vtvl_cli.py audit mau.xlsx -o findings.json    # thẩm định 3 vòng
python3 scripts/vtvl_cli.py error-table findings.json -o Bang-loi.xlsx
python3 scripts/vtvl_cli.py fix-tech mau.xlsx findings.json --dry-run
python3 scripts/vtvl_cli.py fix-tech mau.xlsx findings.json -o mau-da-sua.xlsx
python3 scripts/vtvl_cli.py form findings.json -o Phieu-xac-minh.xlsx
# ... đơn vị ghi "Đồng ý" / số đúng vào phiếu ...
python3 scripts/vtvl_cli.py apply-form mau-da-sua.xlsx Phieu-xac-minh.xlsx -o mau-theo-phieu.xlsx
python3 scripts/vtvl_cli.py clean mau-theo-phieu.xlsx mau-da-sua.xlsx.log.json mau-theo-phieu.xlsx.log.json -o mau-ban-nop.xlsx
python3 scripts/vtvl_cli.py report findings.json --docx BaoCao.docx --xlsx ThamDinh.xlsx
```

## Mẫu của tỉnh / đơn vị khác

Mỗi mẫu là một thư mục `templates/<ten-mau>/` có `map.json` (tên trang, cột nào là gì, cột nào được nhập).
Mẫu khác thì chép `templates/mau-30000-dong/`, sửa `map.json`, chạy với `--map templates/<ten-mau>/map.json`.
Khung báo cáo của Sở: `templates/mau-bao-cao-so/report.json`.

## Cấu trúc

```
.claude/skills/vtvl/   skill /vtvl (phân luồng + các nhánh)
.claude/settings.json  quyền + hook chặn commit có dữ liệu cá nhân
CLAUDE.md              hướng dẫn cho Claude, cách đóng góp kinh nghiệm
scripts/               vtvl_cli.py, fill_template.py, pii_scan.py, vtvl/ (thư viện)
templates/             bản đồ mẫu bộ biểu, khung báo cáo
docs/                  căn cứ pháp lý (có nguồn, ngày), khung thẩm định
KINH-NGHIEM/           kinh nghiệm từng ca (đã bỏ danh tính)
tests/                 kiểm thử trên dữ liệu giả
```

## Đóng góp

Claude (hoặc bạn) mở PR thêm kinh nghiệm / công cụ theo `CLAUDE.md`. Mọi PR phải qua CI:
quét dữ liệu cá nhân (tên, số điện thoại, email, CCCD, mã số thuế, tên đơn vị thật) và kiểm thử. **Không đưa hồ sơ thật lên kho.**

## Lưu ý

Công cụ hỗ trợ rà soát, không thay quyết định của cơ quan có thẩm quyền. Ngưỡng ở vòng 3 là đề xuất nghiệp vụ để khoanh vùng.
Căn cứ pháp lý: xem `docs/can-cu-phap-ly.md`, luôn đối chiếu bản gốc.

---

Cần làm trọn gói: liên hệ **Phat Claude Architect Agent Solutions**.

Giấy phép MIT.
