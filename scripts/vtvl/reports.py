"""Sinh đầu ra từ findings: bảng lỗi Excel, nhận xét Word, phiếu xác minh, báo cáo theo mẫu Sở.

Chữ hiện cho người đọc: tiếng Việt. Màu chữ: chỉ đen / xanh đậm / đỏ (không dùng chữ xám).
"""
from __future__ import annotations

import json
import os

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

BLACK, BLUE, RED = "FF15202B", "FF0B3D6B", "FFC0392B"
THIN = Side(style="thin", color="FF15202B")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="FFDDEBF7")
GROUP_TEXT = {"bat_buoc": "Lỗi bắt buộc", "canh_bao": "Cảnh báo cần giải trình", "thong_tin": "Thông tin"}
FIX_TEXT = {"ky_thuat": "Kỹ thuật (sửa được ngay)", "don_vi": "Cần đơn vị xác minh"}
AGREE = "Đồng ý"


def _table(ws, start_row, headers, rows, widths):
    for i, h in enumerate(headers, 1):
        c = ws.cell(start_row, i, h)
        c.font = Font(bold=True, color=BLUE)
        c.fill = HEAD_FILL
        c.border = BORDER
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    for r, row in enumerate(rows, start_row + 1):
        for i, v in enumerate(row, 1):
            c = ws.cell(r, i, v)
            c.border = BORDER
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.font = Font(color=BLACK)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(start_row + 1, 1)


def _title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = Font(bold=True, size=13, color=BLUE)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = Font(italic=True, color=BLACK)


def error_table(result: dict, out: str):
    """Bảng lỗi chi tiết theo yêu cầu đầu ra mục 6.2 của câu lệnh thẩm định HD07."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bang loi"
    _title(ws, f"BẢNG LỖI – {result['file']}", f"Phân loại: {result['classification']} – {result['conclusion']}")
    rows = []
    for i, f in enumerate(result["findings"], 1):
        rows.append([i, f["code"], f"Vòng {f['round']}", GROUP_TEXT[f["group"]], f["sheet"], f["cells"],
                     f["current"], f["detail"], f["basis"], f["severity"],
                     ("Sửa thành: " + str(f["proposed"])) if f.get("proposed") not in (None, "") else "",
                     FIX_TEXT[f["fix"]], f.get("question") or ""])
    _table(ws, 4, ["STT", "Mã", "Vòng", "Nhóm", "Biểu", "Ô/dòng", "Hiện trạng", "Nội dung sai/nghi vấn",
                   "Căn cứ", "Mức độ", "Cách sửa", "Loại xử lý", "Cần đơn vị xác nhận"],
           rows, [5, 7, 8, 14, 18, 18, 22, 40, 26, 10, 24, 16, 32])
    for r in range(5, 5 + len(rows)):
        if ws.cell(r, 10).value == "Cao":
            ws.cell(r, 10).font = Font(bold=True, color=RED)
    ws2 = wb.create_sheet("Chi so")
    _title(ws2, "CHỈ SỐ THẨM ĐỊNH (ngưỡng là đề xuất nghiệp vụ, không kết luận sai chỉ vì vượt ngưỡng)")
    _indicator_sheet(ws2, result, 3)
    wb.save(out)
    return out


def _indicator_sheet(ws, result, start):
    rows = []
    for k in result["indicators"]:
        rows.append([k["stt"], k["name"], k["num"], k["den"], None, k["threshold"]])
    _table(ws, start, ["STT", "Chỉ số", "Tử số", "Mẫu số", "Giá trị", "Ngưỡng đề xuất"], rows, [6, 50, 14, 14, 12, 16])
    for r in range(start + 1, start + 1 + len(rows)):
        ws.cell(r, 5).value = f'=IF(D{r}=0,"",C{r}/D{r})'
        ws.cell(r, 5).number_format = "0.0%"


# ---------------- phiếu xác minh ----------------
FORM_HEADERS = ["STT", "Biểu", "Ô/dòng", "Vấn đề", "Giá trị hiện tại", "Đề xuất",
                "Đơn vị trả lời (Đồng ý / số đúng)", "Ghi chú",
                "_sheet", "_cells", "_proposed", "_code"]


def verification_form(result: dict, out: str, unit: str = "", max_items: int = 40):
    """Phiếu 'Đồng ý / điền số': mỗi điểm cần đơn vị xác minh là một dòng.

    Cột _sheet/_cells/_proposed (ẩn) để apply_form áp lại tự động.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Phieu xac minh"
    _title(ws, f"PHIẾU XÁC MINH HỒ SƠ VỊ TRÍ VIỆC LÀM{(' – ' + unit) if unit else ''}",
           'Đơn vị chỉ cần ghi "Đồng ý" vào cột Trả lời nếu đồng ý đề xuất; nếu không thì ghi số/nội dung đúng. Để trống = chưa xác minh.')
    items = [f for f in result["findings"] if f["fix"] == "don_vi" or f.get("question")][:max_items]
    rows = []
    for i, f in enumerate(items, 1):
        prop = f.get("proposed")
        rows.append([i, f["sheet"], f["cells"], f["detail"] + (("\n" + f["question"]) if f.get("question") else ""),
                     f["current"], "" if prop in (None, "") else str(prop), "", "",
                     f["sheet"], f["cells"], "" if prop is None else json.dumps(prop, ensure_ascii=False), f["code"]])
    _table(ws, 4, FORM_HEADERS, rows, [5, 18, 16, 48, 24, 28, 22, 22, 4, 4, 4, 4])
    for col in "IJKL":
        ws.column_dimensions[col].hidden = True
    for r in range(5, 5 + len(rows)):
        ws.cell(r, 7).fill = PatternFill("solid", fgColor="FFFFF2CC")
    if rows:
        dv = DataValidation(type="list", formula1=f'"{AGREE}"', allow_blank=True, showErrorMessage=False)
        dv.add(f"G5:G{4 + len(rows)}")
        ws.add_data_validation(dv)
    wb.save(out)
    return len(rows)


def read_form(path: str) -> list[dict]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    out = []
    head = None
    for row in ws.iter_rows(values_only=True):
        if head is None:
            if row and row[0] == "STT":
                head = list(row)
            continue
        if row[0] in (None, ""):
            continue
        rec = dict(zip(head, row))
        out.append(rec)
    return out


# ---------------- nhận xét Word ----------------
def _docx_base():
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt
    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Cm(1.8))
    st = doc.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(12)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    return doc


def _docx_table(doc, headers, rows, size=10):
    from docx.shared import Pt
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    for i, h in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.text = ""
        run = cell.paragraphs[0].add_run(str(h))
        run.bold = True
        run.font.size = Pt(size)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            run = cells[i].paragraphs[0].add_run("" if v is None else str(v))
            run.font.size = Pt(size)
    return t


def review_docx(result: dict, out: str, unit: str = "", date: str = ""):
    """Nhận xét thẩm định theo thứ tự đầu ra của câu lệnh HD07 mục 6."""
    doc = _docx_base()
    doc.add_heading(f"NHẬN XÉT THẨM ĐỊNH HỒ SƠ VỊ TRÍ VIỆC LÀM{(' – ' + unit) if unit else ''}", level=1)
    doc.add_paragraph(f"Tệp: {result['file']} ({result['sheets']} trang). {('Ngày: ' + date) if date else ''}")
    doc.add_heading("1. Kết luận tổng quát", level=2)
    doc.add_paragraph(f"Phân loại {result['classification']}: {result['conclusion']}. "
                      f"Lỗi bắt buộc: {result['counts']['bat_buoc']}; cảnh báo: {result['counts']['canh_bao']}; "
                      f"thông tin: {result['counts']['thong_tin']}.")
    doc.add_paragraph("Đây là nhận xét hỗ trợ rà soát, không thay quyết định của cơ quan có thẩm quyền. "
                      "Số liệu tự tính lại từ cột nhập; giá trị công thức lưu sẵn không được coi là đã tính lại.")
    doc.add_heading("2. Bảng lỗi chi tiết", level=2)
    _docx_table(doc, ["Mã", "Biểu / ô", "Nội dung", "Căn cứ", "Mức độ", "Cách sửa / cần xác nhận"],
                [[f["code"], f"{f['sheet']} {f['cells']}", f"{f['detail']} ({f['current']})", f["basis"], f["severity"],
                  (("Sửa thành " + str(f["proposed"]) + ". ") if f.get("proposed") not in (None, "") else "") + (f.get("question") or "")]
                 for f in result["findings"]])
    doc.add_heading("3. Chỉ số", level=2)
    _docx_table(doc, ["STT", "Chỉ số", "Tử số", "Mẫu số", "Ngưỡng"],
                [[k["stt"], k["name"], k["num"], k["den"], k["threshold"]] for k in result["indicators"]])
    doc.add_heading("4. Việc đơn vị cần làm (theo thứ tự ưu tiên)", level=2)
    order = sorted(result["findings"], key=lambda f: ({"bat_buoc": 0, "canh_bao": 1, "thong_tin": 2}[f["group"]],
                                                       {"Cao": 0, "Trung bình": 1, "Thấp": 2}[f["severity"]]))
    for f in order:
        if f["group"] == "thong_tin":
            continue
        doc.add_paragraph(f"{f['code']} – {f['sheet']} {f['cells']}: {f['detail']}", style="List Number")
    doc.save(out)
    return out


# ---------------- báo cáo theo mẫu Sở ----------------
def province_report(result: dict, template_dir: str, out_docx: str, out_xlsx: str, meta: dict):
    """Báo cáo thẩm định theo mẫu Sở: Word + Excel nhiều trang, cấu hình ở report.json của mẫu."""
    with open(os.path.join(template_dir, "report.json"), encoding="utf-8") as f:
        spec = json.load(f)
    F = result["findings"]
    hard = [x for x in F if x["group"] == "bat_buoc"]
    warn = [x for x in F if x["group"] == "canh_bao"]
    info = [x for x in F if x["group"] == "thong_tin"]
    rounds = {r: [x for x in F if x["round"] == r] for r in (1, 2, 3)}

    def round_cls(items):
        if any(x["group"] == "bat_buoc" for x in items):
            return "C" if items and items[0]["round"] == 3 else "D"
        return "B" if any(x["group"] == "canh_bao" for x in items) else "A"

    def frow(x):
        return [x["code"], f"Vòng {x['round']}", x["sheet"], x["cells"], "", str(x["current"]), x["detail"],
                x["basis"], x["severity"], meta.get("verify_level", ""), x.get("question") or (
                    ("Sửa thành " + str(x["proposed"])) if x.get("proposed") not in (None, "") else "")]

    wb = openpyxl.Workbook()
    sheets = spec["sheets"]
    ws = wb.active
    ws.title = sheets["summary"]
    _title(ws, spec["title"].format(**meta))
    summary = [["Tệp thẩm định", result["file"]], ["Căn cứ", meta.get("basis", "")],
               ["Ngày thẩm định", meta.get("date", "")], ["Nguồn đối chiếu", meta.get("sources", "")]]
    for r, items in rounds.items():
        summary.append([spec["round_names"][str(r)], f"{round_cls(items)} – {sum(1 for x in items if x['group'] == 'bat_buoc')} lỗi bắt buộc, "
                        f"{sum(1 for x in items if x['group'] == 'canh_bao')} cảnh báo"])
    summary.append(["KẾT LUẬN CHUNG", f"{result['classification']} – {result['conclusion']}"])
    summary.append(["Quy ước phân loại", spec["classes"]])
    _table(ws, 3, ["Mục", "Nội dung"], summary, [30, 120])
    cols = spec["finding_columns"]
    ws = wb.create_sheet(sheets["errors"])
    _title(ws, "BẢNG LỖI BẮT BUỘC (phải khắc phục trước khi trình)")
    _table(ws, 3, cols, [frow(x) for x in hard] or [["—", "Không có lỗi bắt buộc"] + [""] * (len(cols) - 2)], [8, 9, 20, 16, 12, 26, 40, 28, 12, 9, 34])
    ws = wb.create_sheet(sheets["warnings"])
    _title(ws, "BẢNG CẢNH BÁO CẦN GIẢI TRÌNH VÀ THÔNG TIN")
    _table(ws, 3, cols, [frow(x) for x in warn + info], [8, 9, 20, 16, 12, 26, 40, 28, 12, 9, 34])
    ws = wb.create_sheet(sheets["indicators"])
    _title(ws, "CHỈ SỐ THẨM ĐỊNH (ngưỡng là đề xuất nghiệp vụ)")
    _indicator_sheet(ws, result, 3)
    ws = wb.create_sheet(sheets["tasks"])
    _title(ws, "ĐỐI CHIẾU HAI CHIỀU: PHÁP LÝ ↔ A1 ↔ A2")
    _table(ws, 3, spec["task_columns"], meta.get("task_rows", []), [18, 50, 14, 24, 12, 18, 30, 10])
    ws = wb.create_sheet(sheets["detail"])
    _title(ws, "DANH SÁCH SAI LỆCH (chi tiết)")
    det = []
    for i, x in enumerate(F, 1):
        det.append([i, x["code"], f"Vòng {x['round']}", GROUP_TEXT[x["group"]], x["sheet"], "", x["cells"], x["cells"], "",
                    str(x["current"]), x["detail"], "", "", x["basis"], x["severity"], meta.get("verify_level", ""), "",
                    x.get("question") or "", meta.get("owner", "Đơn vị"), "Chưa xử lý"])
    _table(ws, 3, spec["detail_columns"], det, [5, 8, 8, 16, 18, 8, 12, 12, 16, 22, 34, 12, 14, 26, 10, 8, 14, 30, 14, 12])
    wb.save(out_xlsx)

    doc = _docx_base()
    doc.add_heading(spec["title"].format(**meta), level=1)
    doc.add_paragraph(f"Tệp thẩm định: {result['file']}. Căn cứ: {meta.get('basis', '')}. Ngày thẩm định: {meta.get('date', '')}.")
    doc.add_heading("1. Tóm tắt kết luận", level=2)
    _docx_table(doc, ["Vòng", "Kết quả", "Phân loại"],
                [[spec["round_names"][str(r)], f"{len(items)} phát hiện", round_cls(items)] for r, items in rounds.items()]
                + [["KẾT LUẬN CHUNG", result["conclusion"], result["classification"]]])
    doc.add_paragraph("Quy ước phân loại: " + spec["classes"])
    doc.add_heading("2. Bảng lỗi bắt buộc", level=2)
    if hard:
        _docx_table(doc, ["Mã", "Biểu / ô", "Phát hiện", "Căn cứ", "Mức độ", "Đề nghị"],
                    [[x["code"], f"{x['sheet']} {x['cells']}", x["detail"], x["basis"], x["severity"], x.get("question") or ""] for x in hard], 9)
    else:
        doc.add_paragraph("Không có lỗi bắt buộc trên tệp bộ biểu.")
    doc.add_heading("3. Bảng cảnh báo cần giải trình", level=2)
    _docx_table(doc, ["Mã", "Biểu / ô", "Phát hiện", "Căn cứ", "Mức độ", "Đề nghị"],
                [[x["code"], f"{x['sheet']} {x['cells']}", x["detail"], x["basis"], x["severity"], x.get("question") or ""] for x in warn + info], 9)
    doc.add_heading("4. Bảng chỉ số", level=2)
    doc.add_paragraph("Ngưỡng là đề xuất nghiệp vụ để khoanh vùng cần xem xét; không kết luận sai chỉ vì vượt ngưỡng.")
    _docx_table(doc, ["STT", "Chỉ số", "Tử số", "Mẫu số", "Ngưỡng"],
                [[k["stt"], k["name"], k["num"], k["den"], k["threshold"]] for k in result["indicators"]], 9)
    doc.save(out_docx)
    return out_docx, out_xlsx
