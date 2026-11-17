"""Thẩm định bộ biểu VTVL 3 vòng (kỹ thuật / nhất quán liên biểu / nghiệp vụ) -> findings.json.

Chỉ ĐỌC tệp (openpyxl read_only), không sửa gì. Mỗi phát hiện có:
  code, round (1|2|3), group ("bat_buoc"|"canh_bao"|"thong_tin"), sheet, cells, current,
  detail, basis, severity ("Cao"|"Trung bình"|"Thấp"), fix ("ky_thuat"|"don_vi"),
  proposed (giá trị đề xuất, có thể None), question (câu hỏi cho phiếu xác minh).

Ngưỡng ở vòng 3 là ĐỀ XUẤT NGHIỆP VỤ để khoanh vùng, không kết luận sai chỉ vì vượt ngưỡng.
Giá trị công thức lưu sẵn trong tệp KHÔNG được coi là bằng chứng đã tính lại: vòng 2/3 tự tính
lại từ cột nhập (giờ = SL x phút / 60 + giờ bổ sung; người PP2 = giờ / quỹ giờ).
"""
from __future__ import annotations

import json
import re
import unicodedata
from collections import Counter, defaultdict

import openpyxl

ERROR_VALUES = ("#REF!", "#VALUE!", "#N/A", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!")
HD07 = "Hướng dẫn 07-HD/BTCTW"


def load_map(path: str) -> dict:
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def norm_code(v) -> str:
    return "" if v is None else str(v).strip()


def fold(s: str) -> str:
    """Bỏ dấu + chữ hoa, để bắt LDQL/LĐQL."""
    s = s.replace("Đ", "D").replace("đ", "d")
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").upper()


def num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip().replace(".", "").replace(",", ".") if re.fullmatch(r"[\d.]+,\d+", str(v).strip()) else str(v).strip()
    try:
        return float(t)
    except ValueError:
        return None


class Book:
    """Đọc các trang cần thiết thành list[dict] theo bản đồ cột."""

    def __init__(self, path: str, tmap: dict, max_empty: int = 200):
        self.path = path
        self.map = tmap
        self.first = tmap["first_row"]
        self.wb_v = openpyxl.load_workbook(path, read_only=True, data_only=True)
        self.wb_f = openpyxl.load_workbook(path, read_only=True, data_only=False)
        self.names = self.wb_v.sheetnames
        self.rows: dict[str, list[dict]] = {}
        self.max_empty = max_empty
        for key, spec in tmap["sheets"].items():
            if spec["name"] in self.names:
                self.rows[key] = self._read(spec)

    def _read(self, spec: dict) -> list[dict]:
        ws = self.wb_v[spec["name"]]
        cols = spec["cols"]
        idx = {k: openpyxl.utils.column_index_from_string(c) - 1 for k, c in cols.items()}
        maxc = max(idx.values()) + 1
        out, empty = [], 0
        for r, row in enumerate(ws.iter_rows(min_row=self.first, max_col=maxc, values_only=True), start=self.first):
            rec = {k: (row[i] if i < len(row) else None) for k, i in idx.items()}
            if all(v in (None, "") for v in rec.values()):
                empty += 1
                if empty >= self.max_empty:
                    break
                continue
            empty = 0
            rec["_row"] = r
            out.append(rec)
        return out

    def col(self, key: str, field: str) -> str:
        return self.map["sheets"][key]["cols"][field]

    def sheet(self, key: str) -> str:
        return self.map["sheets"][key]["name"]


class Auditor:
    def __init__(self, book: Book):
        self.b = book
        self.f: list[dict] = []
        self.n = Counter()
        self.quota = book.map.get("hours_per_year", 1904)

    def add(self, prefix, rnd, group, key, cells, current, detail, basis, severity, fix,
            proposed=None, question=None, sheet=None):
        self.n[prefix] += 1
        self.f.append({
            "code": f"{prefix}{self.n[prefix]:02d}", "round": rnd, "group": group,
            "sheet": sheet or (self.b.sheet(key) if key else ""), "cells": cells,
            "current": current, "detail": detail, "basis": basis, "severity": severity,
            "fix": fix, "proposed": proposed, "question": question,
        })

    # ================= VÒNG 1: kỹ thuật =================
    def round1(self):
        b = self.b
        # 1a. giá trị lỗi + công thức hỏng (#REF!) trên mọi trang dữ liệu
        for name in b.names:
            ws_v, ws_f = b.wb_v[name], b.wb_f[name]
            errs, refs = [], []
            for row in ws_v.iter_rows(min_row=1, max_row=min(ws_v.max_row or 0, 30005)):
                for c in row:
                    if isinstance(c.value, str) and c.value in ERROR_VALUES:
                        errs.append(c.coordinate)
            for row in ws_f.iter_rows(min_row=1, max_row=min(ws_f.max_row or 0, 30005)):
                for c in row:
                    if isinstance(c.value, str) and c.value.startswith("=") and "#REF!" in c.value:
                        refs.append(c.coordinate)
            if errs:
                self.add("KT", 1, "bat_buoc", None, _short(errs), f"{len(errs)} ô báo lỗi",
                         "Ô hiển thị giá trị lỗi công thức", "Câu lệnh thẩm định HD07 mục 5", "Cao",
                         "ky_thuat", sheet=name)
            if refs:
                self.add("KT", 1, "bat_buoc", None, _short(refs), f"{len(refs)} công thức chứa #REF!",
                         "Công thức tham chiếu vùng đã bị xoá; sửa ở dòng đầu rồi chép xuống, báo đầu mối quản lý biểu",
                         "Câu lệnh thẩm định HD07 mục 5", "Cao", "ky_thuat", sheet=name)
        # 1b. ô công thức bị ghi đè bằng hằng số (theo cột có công thức ở dòng đầu)
        inputs = b.map.get("input_columns", {})
        for key, spec in b.map["sheets"].items():
            name = spec["name"]
            if name not in b.names or name not in inputs:
                continue
            ws = b.wb_f[name]
            first = None
            over = []
            for r, row in enumerate(ws.iter_rows(min_row=b.first, max_row=b.first + 2000), start=b.first):
                if first is None:
                    first = {c.column_letter: c.value for c in row
                             if isinstance(c.value, str) and c.value.startswith("=")}
                    continue
                for c in row:
                    col = getattr(c, "column_letter", None)
                    if col in first and col not in inputs[name] and c.value not in (None, "") \
                            and not (isinstance(c.value, str) and c.value.startswith("=")):
                        over.append(c.coordinate)
            if over:
                self.add("KT", 1, "bat_buoc", key, _short(over), f"{len(over)} ô",
                         "Ô thuộc cột công thức nhưng đang chứa số/chữ nhập tay (công thức bị ghi đè)",
                         "Quy ước mẫu: ô xám là công thức, không nhập tay", "Cao", "ky_thuat")
        # 1c. mã: trùng, khoảng trắng thừa, LDQL/LĐQL, mã ví dụ VD-
        for key, field in (("A1", "code"), ("A2", "code"), ("C1", "code"), ("DM", "code")):
            rows = b.rows.get(key, [])
            col = b.col(key, field)
            seen = defaultdict(list)
            for r in rows:
                raw = r[field]
                if raw is None:
                    continue
                s = str(raw)
                cell = f"{col}{r['_row']}"
                if s != s.strip() or "  " in s:
                    self.add("KT", 1, "canh_bao", key, cell, repr(s), "Mã có khoảng trắng thừa",
                             "Câu lệnh thẩm định HD07 mục 5", "Thấp", "ky_thuat", proposed=s.strip())
                if s.strip().upper().startswith("VD-"):
                    self.add("KT", 1, "bat_buoc", key, cell, s, "Mã ví dụ 'VD-' của mẫu nằm trong dữ liệu thật",
                             "Quy tắc mã của mẫu: VD- chỉ minh hoạ", "Trung bình", "don_vi",
                             question="Thay mã ví dụ bằng mã thật hoặc xoá dòng?")
                seen[s.strip()].append(cell)
            for code, cells in seen.items():
                if len(cells) > 1:
                    self.add("KT", 1, "bat_buoc", key, ", ".join(cells), code, "Mã trùng",
                             "Mã phải duy nhất (Câu lệnh thẩm định HD07 mục 2)", "Cao", "don_vi",
                             question=f"Mã {code} trùng ở {', '.join(cells)}: giữ dòng nào, đổi mã dòng nào?")
            folded = defaultdict(set)
            for code in seen:
                folded[fold(code)].add(code)
            for variants in folded.values():
                if len(variants) > 1:
                    self.add("KT", 1, "canh_bao", key, "", " / ".join(sorted(variants)),
                             "Cùng một mã nhưng khác dấu (vd LDQL/LĐQL)", "Câu lệnh thẩm định HD07 mục 5",
                             "Trung bình", "ky_thuat")

    # ================= VÒNG 2: nhất quán liên biểu =================
    def round2(self):
        b = self.b
        codes = {k: {norm_code(r["code"]) for r in b.rows.get(k, []) if norm_code(r["code"])}
                 for k in ("A1", "A2", "C1", "DM")}
        a2_by_code = {norm_code(r["code"]): r for r in b.rows.get("A2", [])}

        def missing(key, field, target, label, basis, group="bat_buoc", sev="Cao"):
            col = b.col(key, field)
            bad = [(f"{col}{r['_row']}", norm_code(r[field])) for r in b.rows.get(key, [])
                   if norm_code(r[field]) and norm_code(r[field]) not in codes[target]]
            for cell, val in bad:
                near = _nearest(val, codes[target])
                self.add("LK", 2, group, key, cell, val, label, basis, sev,
                         "ky_thuat" if near else "don_vi", proposed=near,
                         question=None if near else f"Mã {val} không có trong biểu nguồn: mã đúng là gì?")
        if "A2" in b.rows:
            missing("A2", "task", "A1", "Mã nhiệm vụ ở A2 không có trong A1", "Chuỗi liên kết A1 -> A2")
            missing("A2", "vtvl", "C1", "Mã VTVL ở A2 không có trong C1", "Chuỗi liên kết C1 -> A2")
        if "B1" in b.rows:
            missing("B1", "job", "A2", "Mã công việc ở B1 không có trong A2 (B1 sẽ báo THIẾU MÃ VTVL, giờ không vào C3)",
                    "Chuỗi liên kết A2 -> B1")
            missing("B1", "product", "DM", "Mã sản phẩm ở B1 không có trong danh mục sản phẩm",
                    "Chuỗi liên kết DM_SAN_PHAM -> B1", group="canh_bao", sev="Trung bình")
        # tiền tố mã công việc lạc loài (vd đơn vị khác)
        prefixes = Counter(c.split("-")[0] for c in codes["A2"] if "-" in c)
        if prefixes:
            main, cnt = prefixes.most_common(1)[0]
            for c in sorted(codes["A2"]):
                p = c.split("-")[0]
                if "-" in c and p != main and prefixes[p] <= max(2, cnt // 20):
                    r = a2_by_code[c]
                    self.add("LK", 2, "canh_bao", "A2", f"{b.col('A2', 'code')}{r['_row']}", c,
                             f"Mã mang tiền tố khác ('{p}' thay vì '{main}'), nghi chép từ đơn vị khác",
                             "Mã thống nhất giữa các biểu (Câu lệnh HD07 mục 2)", "Trung bình", "ky_thuat",
                             proposed=main + c[len(p):])
        # dạng mã lạc loài (vd BVM-A01-001 giữa các mã BVM-A1-001-001)
        shapes = Counter(re.sub(r"\d", "9", c) for c in codes["A2"])
        if shapes and len(codes["A2"]) >= 10:
            main_shape, cnt = shapes.most_common(1)[0]
            for c in sorted(codes["A2"]):
                sh = re.sub(r"\d", "9", c)
                if sh != main_shape and shapes[sh] <= max(2, len(codes["A2"]) // 20):
                    r = a2_by_code[c]
                    self.add("LK", 2, "canh_bao", "A2", f"{b.col('A2', 'code')}{r['_row']}", c,
                             f"Mã khác dạng với đa số ({main_shape})", "Mã thống nhất giữa các biểu (Câu lệnh HD07 mục 2)",
                             "Thấp", "don_vi", question="Thống nhất dạng mã (sửa cả A2 và B1)?")
        # nhiệm vụ A1 không có công việc A2
        used = Counter(norm_code(r["task"]) for r in b.rows.get("A2", []))
        for r in b.rows.get("A1", []):
            c = norm_code(r["code"])
            if c and used[c] == 0:
                self.add("LK", 2, "canh_bao", "A1", f"{b.col('A1', 'code')}{r['_row']}", c,
                         "Nhiệm vụ chưa được phân rã thành công việc ở A2", f"{HD07}: nhiệm vụ phải phân tích thành công việc, sản phẩm",
                         "Trung bình", "don_vi", question="Nhiệm vụ này được thực hiện qua công việc nào, hay lồng ghép (giải trình)?")
        # C1 <-> C3
        c3 = {norm_code(r["vtvl"]) for r in b.rows.get("C3", []) if norm_code(r["vtvl"])}
        for r in b.rows.get("C1", []):
            c = norm_code(r["code"])
            if not c and r.get("name"):
                self.add("LK", 2, "bat_buoc", "C1", f"{b.col('C1', 'code')}{r['_row']}", r["name"],
                         "VTVL chưa có mã nên không vào C3, D1, D2", "Chuỗi liên kết C1 -> C3", "Cao", "don_vi",
                         question=f"Mã VTVL cho '{r['name']}'?")
            elif c and c3 and c not in c3:
                self.add("LK", 2, "canh_bao", "C1", f"{b.col('C1', 'code')}{r['_row']}", c,
                         "VTVL có ở C1 nhưng chưa có dòng nhu cầu ở C3", "Chuỗi liên kết C1 -> C3", "Trung bình", "don_vi")

    # ================= VÒNG 3: nghiệp vụ =================
    def b1_hours(self):
        """Tự tính giờ từng VTVL từ cột nhập B1 (không tin số lưu sẵn)."""
        a2 = {norm_code(r["code"]): norm_code(r["vtvl"]) for r in self.b.rows.get("A2", [])}
        hours = defaultdict(float)
        for r in self.b.rows.get("B1", []):
            v = a2.get(norm_code(r["job"]))
            q, m, e = num(r["qty"]), num(r["minutes"]), num(r["extra_hours"])
            if v and q is not None and m is not None:
                hours[v] += q * m / 60 + (e or 0)
        return hours

    def round3(self):
        b = self.b
        B1 = b.rows.get("B1", [])
        cols = b.map["sheets"]["B1"]["cols"] if "B1" in b.map["sheets"] else {}
        quota = self.quota
        for r in B1:
            q, m, e = num(r["qty"]), num(r["minutes"]), num(r["extra_hours"])
            row = r["_row"]
            if m is not None and m > 480:
                self.add("NV", 3, "canh_bao", "B1", f"{cols['minutes']}{row}", f"{q} x {m:g} phút",
                         "Thời gian 1 sản phẩm vượt 1 ngày công (480 phút): xem lại đơn vị tính", f"{HD07} Bước 3",
                         "Trung bình", "don_vi", question="Thời gian đúng cho 1 sản phẩm (phút)?")
            if q is not None and m is not None and q * m / 60 > quota:
                self.add("NV", 3, "canh_bao", "B1", f"{cols['qty']}{row}", f"{q * m / 60:,.1f} giờ",
                         f"Một dòng công việc vượt quỹ giờ 1 người ({quota} giờ)", f"{HD07} Bước 3", "Trung bình", "don_vi",
                         question="Đối chiếu số lượng và thời gian thực tế của dòng này?")
            if e is not None and m is not None and e > 0 and abs(e - m) < 1e-9:
                self.add("NV", 3, "canh_bao", "B1", f"{cols['extra_hours']}{row}", f"E={e:g}, D={m:g}",
                         "Giờ bổ sung bằng đúng thời gian 1 sản phẩm: nghi nhập nhầm cột, bị cộng trùng",
                         "HD_DIEN_BIEU B1 cột E: tổng giờ/năm của việc bổ trợ riêng", "Trung bình", "don_vi",
                         proposed=None, question="Để trống giờ bổ sung (Đồng ý) hay ghi tổng giờ/năm có căn cứ?")
        # nguồn sản lượng chung chung
        srcs = Counter(str(r["qty_source"]).strip() for r in B1 if r.get("qty_source"))
        if B1 and srcs:
            top, cnt = srcs.most_common(1)[0]
            if cnt >= 0.8 * len(B1) and len(B1) >= 5:
                self.add("NV", 3, "canh_bao", "B1", f"{cols['qty_source']}{B1[0]['_row']}:{cols['qty_source']}{B1[-1]['_row']}",
                         f'"{top}" {cnt}/{len(B1)}', "Nguồn sản lượng ghi một câu chung cho hầu hết dòng",
                         f"{HD07} Bước 3: nguồn phải phản ánh hồ sơ thực tế", "Trung bình", "don_vi",
                         question="Nguồn số liệu thật (báo cáo phần mềm, sổ theo dõi...)?")
        # cùng một số lượng dùng nhiều dòng (nghi tính trùng)
        qcount = defaultdict(list)
        for r in B1:
            q = num(r["qty"])
            if q and q >= 100:
                qcount[q].append(r["_row"])
        for q, rows in qcount.items():
            if len(rows) >= 3:
                self.add("NV", 3, "thong_tin", "B1", ", ".join(f"{cols['qty']}{x}" for x in rows[:10]), f"{q:g}",
                         "Cùng một số lượng dùng cho nhiều công việc: hợp lệ nếu công đoạn khác nhau",
                         f"{HD07}: không tính trùng", "Thấp", "don_vi")
        # C3: J/G, PP3, phân mảnh, trạng thái kiểm chứng
        hours = self.b1_hours()
        C3 = b.rows.get("C3", [])
        c3c = b.map["sheets"]["C3"]["cols"] if "C3" in b.map["sheets"] else {}
        est_words = set(b.map.get("estimate_words", []))
        est_vtvl = set()
        a2 = {norm_code(r["code"]): norm_code(r["vtvl"]) for r in b.rows.get("A2", [])}
        for r in B1:
            if str(r.get("status") or "").strip() in est_words or not r.get("time_source"):
                v = a2.get(norm_code(r["job"]))
                if v:
                    est_vtvl.add(v)
        pp3 = 0
        for r in C3:
            v = norm_code(r["vtvl"])
            meth = str(r.get("method") or "")
            if "PP3" in meth:
                pp3 += 1
            g = hours.get(v, 0) / quota if quota else 0
            j = num(r.get("people_need"))
            row = r["_row"]
            if "PP2" in meth and g and j is not None and not (0.9 <= j / g <= 1.1):
                self.add("NV", 3, "canh_bao", "C3", f"{c3c['people_need']}{row}", f"J={j:g}, G≈{g:.2f}",
                         "Số người cần bố trí lệch kết quả PP2 quá 10%", "HD_DIEN_BIEU C3 cột J: phải giải trình so với G",
                         "Trung bình", "don_vi", question="Giải trình chênh lệch (trực 24/24, kiêm nhiệm...) hay sửa J?")
            if "PP2" in meth and 0 < g < 1:
                self.add("NV", 3, "thong_tin", "C3", f"{c3c['people_pp2']}{row}", f"G≈{g:.2f}",
                         "VTVL dưới 1 người quy đổi: xem xét kiêm nhiệm", "Đề xuất nghiệp vụ", "Thấp", "don_vi")
            if "PP2" in meth and r.get("people_pp1") not in (None, ""):
                self.add("NV", 3, "canh_bao", "C3", f"{c3c['people_pp1']}{row}", str(r["people_pp1"]),
                         "Chọn PP2 nhưng vẫn nhập số người PP1", "HD_DIEN_BIEU C3 cột B, H", "Trung bình", "ky_thuat",
                         proposed=None)
            st = str(r.get("status") or "").strip()
            if st in b.map.get("verified_words", []) and v in est_vtvl:
                self.add("NV", 3, "canh_bao", "C3", f"{c3c['status']}{row}", st,
                         "Ghi 'Đã kiểm chứng' trong khi thời gian B1 là ước lượng/chưa ghi nguồn",
                         "HD_DIEN_BIEU C3 cột L", "Cao", "don_vi",
                         question="Đơn vị xác nhận đã kiểm chứng (ghi số văn bản/phiếu) hay đổi thành 'Tính thử'?")
        if C3 and pp3 / len(C3) > 0.2:
            self.add("NV", 3, "canh_bao", "C3", f"{c3c['method']}", f"{pp3}/{len(C3)} VTVL",
                     "Tỷ lệ VTVL dùng PP3 trên 20%", f"{HD07} Bước 6: PP3 phải có hồ sơ", "Trung bình", "don_vi")
        # gán đồng loạt cho một VTVL
        A2 = b.rows.get("A2", [])
        if len(A2) >= 10:
            cnt = Counter(norm_code(r["vtvl"]) for r in A2 if norm_code(r["vtvl"]))
            for v, k in cnt.items():
                if k >= 0.3 * len(A2) and re.search(r"QL|LĐ|LD", v):
                    self.add("NV", 3, "canh_bao", "A2", b.col("A2", "vtvl"), f"{v}: {k}/{len(A2)} công việc",
                             "Gán đồng loạt công việc cho một chức danh lãnh đạo", "Câu lệnh thẩm định HD07 mục 2",
                             "Trung bình", "don_vi", question="Tách công đoạn lãnh đạo / chuyên viên?")

    def indicators(self) -> list[dict]:
        b = self.b
        hours = self.b1_hours()
        total = sum(hours.values())
        C3 = b.rows.get("C3", [])
        B1 = b.rows.get("B1", [])
        est = set(b.map.get("estimate_words", []))
        est_h = 0.0
        a2 = {norm_code(r["code"]): norm_code(r["vtvl"]) for r in b.rows.get("A2", [])}
        extra = 0.0
        for r in B1:
            q, m, e = num(r["qty"]), num(r["minutes"]), num(r["extra_hours"])
            h = (q or 0) * (m or 0) / 60 + (e or 0)
            extra += e or 0
            if str(r.get("status") or "").strip() in est:
                est_h += h
        ok_rows = sum(1 for r in B1 if a2.get(norm_code(r["job"])))
        pp3 = sum(1 for r in C3 if "PP3" in str(r.get("method") or ""))
        frag = sum(1 for v, h in hours.items() if 0 < h / self.quota < 1)
        need = sum(num(r.get("people_need")) or 0 for r in C3)
        return [
            {"stt": "1", "name": "Tỷ lệ dòng B1 liên kết được A2", "num": ok_rows, "den": len(B1), "threshold": "100%"},
            {"stt": "2", "name": "Tỷ lệ giờ ước lượng / tổng giờ B1", "num": round(est_h, 2), "den": round(total, 2), "threshold": "≤ 30%"},
            {"stt": "3", "name": "Tỷ lệ giờ bổ sung / tổng giờ B1", "num": round(extra, 2), "den": round(total, 2), "threshold": "≤ 20%"},
            {"stt": "4", "name": "Tỷ lệ PP3 (theo số VTVL)", "num": pp3, "den": len(C3), "threshold": "≤ 20%"},
            {"stt": "5", "name": "VTVL phân mảnh (dưới 1 người quy đổi)", "num": frag, "den": len(hours), "threshold": "≤ 10%"},
            {"stt": "6", "name": "Tổng người cần bố trí (C3 cột J) / người quy đổi PP2", "num": need, "den": round(total / self.quota, 2), "threshold": "0,90–1,10"},
        ]


def classify(findings: list[dict]) -> str:
    """A đủ điều kiện; B còn cảnh báo; C còn lỗi bắt buộc nghiệp vụ; D lỗi kỹ thuật nghiêm trọng."""
    hard = [f for f in findings if f["group"] == "bat_buoc"]
    if any(f["round"] == 1 and f["severity"] == "Cao" for f in hard):
        return "D"
    if hard:
        return "C"
    if any(f["group"] == "canh_bao" for f in findings):
        return "B"
    return "A"


CLASS_TEXT = {
    "A": "Đạt – đủ điều kiện trình thẩm định",
    "B": "Cần bổ sung, chỉnh sửa – không còn lỗi bắt buộc, còn cảnh báo cần giải trình",
    "C": "Cần bổ sung, chỉnh sửa – kỹ thuật đạt, còn lỗi bắt buộc về nghiệp vụ",
    "D": "Chưa đủ cơ sở thẩm định – lỗi kỹ thuật nghiêm trọng hoặc đứt chuỗi liên kết",
}


def _short(cells: list[str], n: int = 12) -> str:
    return ", ".join(cells[:n]) + (f" … (+{len(cells) - n})" if len(cells) > n else "")


def _nearest(val: str, pool: set[str]):
    """Mã gần đúng duy nhất (khác dấu/khoảng trắng/hoa thường) thì đề xuất sửa kỹ thuật."""
    f = fold(val).replace(" ", "")
    hits = [p for p in pool if fold(p).replace(" ", "") == f]
    return hits[0] if len(hits) == 1 else None


def run(path: str, map_path: str) -> dict:
    book = Book(path, load_map(map_path))
    a = Auditor(book)
    a.round1()
    a.round2()
    a.round3()
    cls = classify(a.f)
    return {
        "file": path.split("/")[-1], "template": book.map["name"], "sheets": len(book.names),
        "classification": cls, "conclusion": CLASS_TEXT[cls],
        "counts": {k: sum(1 for f in a.f if f["group"] == k) for k in ("bat_buoc", "canh_bao", "thong_tin")},
        "indicators": a.indicators(), "findings": a.f,
    }
