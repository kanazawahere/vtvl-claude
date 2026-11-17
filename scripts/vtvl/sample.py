"""Sinh bộ biểu VTVL GIẢ ("Bệnh viện Mẫu") để chạy thử và kiểm thử. Không có dữ liệu thật nào.

Cố ý cài sẵn lỗi để kiểm công cụ thẩm định: mã trùng, mã khác tiền tố, khoảng trắng thừa,
B1 trỏ mã công việc không tồn tại, giờ bổ sung = phút/SP, phút/SP quá lớn, C3 ghi 'Đã kiểm chứng'
khi thời gian là ước lượng, VTVL ở C1 chưa có mã.
"""
from __future__ import annotations

import json

import openpyxl
from openpyxl.comments import Comment


def build(out: str, map_path: str, seed_errors: bool = True):
    tm = json.load(open(map_path, encoding="utf-8"))
    S = {k: v["name"] for k, v in tm["sheets"].items()}
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = {}
    for k, name in S.items():
        ws[k] = wb.create_sheet(name)
        ws[k]["A1"] = f"Biểu {name} – BỆNH VIỆN MẪU (DỮ LIỆU GIẢ)"
    # tiêu đề cột dòng 2
    for k, spec in tm["sheets"].items():
        for field, col in spec["cols"].items():
            ws[k][f"{col}2"] = field
    U = "BVM"
    tasks = [(f"{U}-NV-{i:03d}", t) for i, t in enumerate(
        ["Khám bệnh, chữa bệnh", "Chăm sóc người bệnh", "Xét nghiệm, chẩn đoán hình ảnh",
         "Quản lý hành chính, tổng hợp", "Nhiệm vụ chưa phân rã"], 1)]
    for r, (c, t) in enumerate(tasks, 3):
        ws["A1"][f"A{r}"], ws["A1"][f"B{r}"] = c, t
        ws["A1"][f"D{r}"] = "Quyết định số 00/2026/QĐ-MAU (văn bản giả)"
    vtvl = [(f"{U}-BS", "Bác sĩ", "PP2"), (f"{U}-DD", "Điều dưỡng", "PP2"), (f"{U}-KTV", "Kỹ thuật y", "PP2"),
            (f"{U}-HC", "Hành chính tổng hợp", "PP3"), (f"{U}-LĐQL-01", "Giám đốc", "PP1")]
    for r, (c, n, _) in enumerate(vtvl, 3):
        ws["C1"][f"A{r}"], ws["C1"][f"B{r}"], ws["C1"][f"C{r}"] = c, n, "Viên chức"
        ws["C1"][f"E{r}"] = "=QUY_GIO"
    if seed_errors:
        ws["C1"]["B8"], ws["C1"]["C8"] = "Bảo vệ", "Hợp đồng lao động"  # chưa có mã
    jobs = []
    for i in range(1, 13):
        task = tasks[(i - 1) % 4][0]
        v = vtvl[(i - 1) % 4][0]
        jobs.append([f"{U}-CV-{i:03d}", task, v, f"Công việc giả số {i}", f"SP{i:02d}", "lượt"])
    if seed_errors:
        jobs[3][0] = jobs[2][0]                   # mã trùng
        jobs[5][0] = "XYZ-CV-006"                 # tiền tố lạc
        jobs[7][0] = f"{U}-CV-008 "               # khoảng trắng thừa
    for r, j in enumerate(jobs, 3):
        for col, v in zip("ABCDEF", j):
            ws["A2"][f"{col}{r}"] = v
    for i in range(1, 13):
        ws["DM"][f"A{i + 2}"], ws["DM"][f"B{i + 2}"] = f"SP{i:02d}", f"Sản phẩm giả {i}"
    b1 = []
    for i, j in enumerate(jobs):
        code = j[0].strip() if i != 7 else j[0]
        b1.append([code, 1000 + 100 * i, 15, None, "Báo cáo phần mềm quản lý (giả)", "Đo thực tế (giả)",
                   "Có định mức", f"SP{i + 1:02d}"])
    if seed_errors:
        b1[1][0] = f"{U}-CV-999"          # mã không có trong A2
        b1[2][3] = 15                     # giờ bổ sung = phút
        b1[4][2] = 2400                   # 40 giờ/sản phẩm
        b1[0][6] = "Kiểm đếm, ước lượng thực tế"
        b1[0][5] = None
    for r, row in enumerate(b1, 3):
        c = tm["sheets"]["B1"]["cols"]
        ws["B1"][f"{c['job']}{r}"] = row[0]
        ws["B1"][f"{c['vtvl']}{r}"] = f'=IFERROR(VLOOKUP(A{r},\'{S["A2"]}\'!A:C,3,FALSE),"THIẾU MÃ VTVL")'
        ws["B1"][f"{c['qty']}{r}"] = row[1]
        ws["B1"][f"{c['minutes']}{r}"] = row[2]
        ws["B1"][f"{c['extra_hours']}{r}"] = row[3]
        ws["B1"][f"{c['hours']}{r}"] = f"=IF(C{r}=\"\",\"\",C{r}*D{r}/60)"
        ws["B1"][f"{c['total_hours']}{r}"] = f"=IF(F{r}=\"\",\"\",F{r}+N(E{r}))"
        ws["B1"][f"{c['qty_source']}{r}"] = row[4]
        ws["B1"][f"{c['time_source']}{r}"] = row[5]
        ws["B1"][f"{c['status']}{r}"] = row[6]
        ws["B1"][f"{c['product']}{r}"] = row[7]
    c3 = tm["sheets"]["C3"]["cols"]
    for r, (c, n, m) in enumerate(vtvl, 3):
        ws["C3"][f"{c3['vtvl']}{r}"] = c
        ws["C3"][f"{c3['method']}{r}"] = m
        ws["C3"][f"{c3['b1_hours']}{r}"] = f"=SUMIFS('{S['B1']}'!G:G,'{S['B1']}'!B:B,A{r})"
        ws["C3"][f"{c3['people_need']}{r}"] = 1
        ws["C3"][f"{c3['status']}{r}"] = "Tính thử"
    if seed_errors:
        ws["C3"][f"{c3['status']}3"] = "Đã kiểm chứng"
        ws["C3"][f"{c3['people_need']}3"] = 9
        ws["C3"][f"{c3['people_pp1']}4"] = 3
    ws["A2"]["H3"].comment = Comment("Comment sẵn có của mẫu", "Mau")
    wb.save(out)
    return out
