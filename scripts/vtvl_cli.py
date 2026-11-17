#!/usr/bin/env python3
"""Lệnh chung cho hồ sơ VTVL. Mọi lệnh ghi tệp đều có --dry-run (chạy thử khô, không ghi).

  sample      OUT.xlsx                         sinh bộ biểu GIẢ có cài lỗi (Bệnh viện Mẫu)
  audit       IN.xlsx -o findings.json         thẩm định 3 vòng (chỉ đọc)
  error-table findings.json -o loi.xlsx        bảng lỗi Excel
  review      findings.json -o nhan-xet.docx   nhận xét Word
  form        findings.json -o phieu.xlsx      phiếu xác minh "Đồng ý / điền số"
  patch       IN.xlsx edits.json -o OUT.xlsx   sửa ô: tô vàng + comment giá trị cũ
  fix-tech    IN.xlsx findings.json -o OUT     tự sửa các lỗi KỸ THUẬT có đề xuất chắc chắn
  apply-form  IN.xlsx phieu.xlsx -o OUT.xlsx   áp phiếu đơn vị đã trả lời
  clean       IN.xlsx LOG.json... -o OUT.xlsx  bản nộp sạch: bỏ màu + comment của mình
  report      findings.json --docx --xlsx      báo cáo thẩm định theo mẫu Sở

Tệp lớn (30.000 dòng): mọi lệnh đọc bằng openpyxl read_only, ghi bằng sửa XML. KHÔNG chạy
công cụ tính lại công thức toàn bộ (gridcraft convert/check, LibreOffice headless) trên tệp đầy:
ngốn nhiều GB RAM. Tệp ra đặt fullCalcOnLoad để Excel tự tính khi mở.
"""
from __future__ import annotations

import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from vtvl import audit, reports, sample  # noqa: E402
from vtvl.xlsx_patch import Edit, XlsxPatcher, expand_range  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_MAP = os.path.join(ROOT, "templates", "mau-30000-dong", "map.json")
DEFAULT_REPORT = os.path.join(ROOT, "templates", "mau-bao-cao-so")
WAIT = "CHỜ SỐ THẬT TỪ ĐƠN VỊ"


def _load(p):
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def _input_guard(edits, a):
    """Chỉ cho ghi cột NHẬP của mẫu (theo map.json), trừ khi --allow-any-column."""
    if getattr(a, "allow_any_column", False):
        return edits, []
    cols = _load(a.map).get("input_columns", {})
    ok, bad = [], []
    for e in edits:
        allowed = cols.get(e.sheet)
        col = "".join(ch for ch in e.cell if ch.isalpha())
        (ok if allowed is None or col in allowed else bad).append(e)
    return ok, [(e.sheet, e.cell, "không phải cột NHẬP của mẫu") for e in bad]


def _report_skips(p, extra):
    for s, c, why in list(getattr(p, "skipped", [])) + extra:
        print(f"  BỎ QUA {s}!{c}: {why}")


def _print_plan(plan):
    for p in plan:
        print(f"  {p['sheet']}!{p['cell']}: {p['old']!r} -> {p['new']!r}  [{p['color']}] {p['note'][:60]}")
    print(f"Tổng {len(plan)} ô.")


def cmd_patch(a):
    spec = _load(a.edits)
    edits = []
    for e in spec:
        for cell in expand_range(e["cells"]):
            edits.append(Edit(e["sheet"], cell, e.get("value", "__KEEP__"), e.get("note", ""), e.get("color", "yellow")))
    edits, bad = _input_guard(edits, a)
    p = XlsxPatcher(a.input, author=a.author)
    plan = p.apply(edits, dry_run=a.dry_run, on_formula="skip")
    _print_plan(plan)
    _report_skips(p, bad)
    if not a.dry_run:
        print("Nhật ký:", p.save(a.output))


def cmd_fix_tech(a):
    res = _load(a.findings)
    edits = []
    for f in res["findings"]:
        if f["fix"] == "ky_thuat" and f.get("proposed") not in (None, "") and f["cells"] and "…" not in f["cells"]:
            for cell in expand_range(f["cells"]):
                edits.append(Edit(f["sheet"], cell, f["proposed"],
                                  f"Sửa lỗi kỹ thuật {f['code']}: {f['detail']}. Căn cứ: {f['basis']}."))
    edits, bad = _input_guard(edits, a)
    p = XlsxPatcher(a.input, author=a.author)
    plan = p.apply(edits, dry_run=a.dry_run, on_formula="skip")
    _print_plan(plan)
    _report_skips(p, bad)
    if not a.dry_run:
        print("Nhật ký:", p.save(a.output))


def _parse_answer(ans):
    if ans is None or str(ans).strip() == "":
        return "blank", None
    s = str(ans).strip()
    if s.lower() in ("đồng ý", "dong y", "ok", "đ", "x"):
        return "agree", None
    try:
        v = float(s.replace(".", "").replace(",", ".")) if "," in s else float(s)
        return "value", int(v) if v.is_integer() else v
    except ValueError:
        return "value", s


def cmd_apply_form(a):
    rows = reports.read_form(a.form)
    edits, skipped = [], []
    src = a.source_label or os.path.basename(a.form)
    for r in rows:
        kind, val = _parse_answer(r.get("Đơn vị trả lời (Đồng ý / số đúng)"))
        sheet, cells = r.get("_sheet"), str(r.get("_cells") or "")
        stt = r.get("STT")
        if not sheet or not cells or "…" in cells:
            skipped.append((stt, "không có ô cụ thể, xử lý tay"))
            continue
        try:
            targets = expand_range(cells)
        except ValueError:
            skipped.append((stt, f"ô '{cells}' không đọc được, xử lý tay"))
            continue
        if kind == "agree":
            prop = r.get("_proposed")
            value = json.loads(prop) if prop not in (None, "") else "__KEEP__"
            note = f"STT {stt}. Theo {src}: đơn vị Đồng ý."
            color = "yellow"
        elif kind == "value":
            value, note, color = val, f"STT {stt}. Theo {src}: đơn vị ghi '{val}'.", "yellow"
        else:
            value, note, color = "__KEEP__", f"STT {stt}. {WAIT}.", "pink"
        for c in targets:
            edits.append(Edit(sheet, c, value, note, color))
    edits, bad = _input_guard(edits, a)
    p = XlsxPatcher(a.input, author=a.author)
    plan = p.apply(edits, dry_run=a.dry_run, on_formula="skip")
    _print_plan(plan)
    _report_skips(p, bad)
    for s, why in skipped:
        print(f"  BỎ QUA STT {s}: {why}")
    if not a.dry_run:
        print("Nhật ký:", p.save(a.output))


def cmd_clean(a):
    recs, seen = [], set()
    for lg in a.logs:
        for r in _load(lg):
            k = (r["sheet"], r["cell"])
            if k not in seen:          # giữ kiểu GỐC của lần sửa đầu tiên
                seen.add(k)
                recs.append(r)
    authors = set(a.authors or []) | {r.get("author") for r in recs if r.get("author")}
    p = XlsxPatcher(a.input, author=a.author)
    print(f"Bỏ màu {len(recs)} ô; xoá comment của: {', '.join(sorted(authors))}")
    if a.dry_run:
        return
    p.unmark(recs, authors)
    if a.final_edits:
        spec = _load(a.final_edits)
        p.apply([Edit(e["sheet"], c, e.get("value", "__KEEP__"), "", None)
                 for e in spec for c in expand_range(e["cells"])])
    p.save(a.output, a.output + ".clean.log.json")
    print("Đã ghi", a.output)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--map", default=DEFAULT_MAP, help="bản đồ cột của mẫu (templates/<mau>/map.json)")
    ap.add_argument("--author", default="Claude VTVL", help="tên tác giả comment (để bản nộp sạch nhận ra)")
    sub = ap.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("sample"); s.add_argument("output"); s.add_argument("--clean", action="store_true")
    s = sub.add_parser("audit"); s.add_argument("input"); s.add_argument("-o", "--output", required=True)
    for name in ("error-table", "review", "form"):
        s = sub.add_parser(name); s.add_argument("findings"); s.add_argument("-o", "--output", required=True)
        s.add_argument("--unit", default=""); s.add_argument("--date", default="")
    s = sub.add_parser("patch"); s.add_argument("--allow-any-column", action="store_true"); s.add_argument("input"); s.add_argument("edits"); s.add_argument("-o", "--output")
    s.add_argument("--dry-run", action="store_true")
    s = sub.add_parser("fix-tech"); s.add_argument("--allow-any-column", action="store_true"); s.add_argument("input"); s.add_argument("findings"); s.add_argument("-o", "--output")
    s.add_argument("--dry-run", action="store_true")
    s = sub.add_parser("apply-form"); s.add_argument("--allow-any-column", action="store_true"); s.add_argument("input"); s.add_argument("form"); s.add_argument("-o", "--output")
    s.add_argument("--source-label", default=""); s.add_argument("--dry-run", action="store_true")
    s = sub.add_parser("clean"); s.add_argument("input"); s.add_argument("logs", nargs="+"); s.add_argument("-o", "--output", required=True)
    s.add_argument("--authors", nargs="*"); s.add_argument("--final-edits"); s.add_argument("--dry-run", action="store_true")
    s = sub.add_parser("report"); s.add_argument("findings"); s.add_argument("--docx", required=True); s.add_argument("--xlsx", required=True)
    s.add_argument("--template", default=DEFAULT_REPORT); s.add_argument("--meta", help="JSON: unit, date, basis, sources, owner, task_rows")
    a = ap.parse_args(argv)

    if a.cmd in ("patch", "fix-tech", "apply-form") and not a.dry_run and not a.output:
        ap.error("cần -o OUTPUT (hoặc --dry-run)")
    if a.cmd == "sample":
        print(sample.build(a.output, a.map, seed_errors=not a.clean))
    elif a.cmd == "audit":
        res = audit.run(a.input, a.map)
        with open(a.output, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1, default=str)
        print(f"Phân loại {res['classification']}: {res['conclusion']}. {res['counts']}")
    elif a.cmd == "error-table":
        print(reports.error_table(_load(a.findings), a.output))
    elif a.cmd == "review":
        print(reports.review_docx(_load(a.findings), a.output, a.unit, a.date))
    elif a.cmd == "form":
        print("Số dòng phiếu:", reports.verification_form(_load(a.findings), a.output, a.unit))
    elif a.cmd == "patch":
        cmd_patch(a)
    elif a.cmd == "fix-tech":
        cmd_fix_tech(a)
    elif a.cmd == "apply-form":
        cmd_apply_form(a)
    elif a.cmd == "clean":
        cmd_clean(a)
    elif a.cmd == "report":
        meta = {"unit": "", "date": "", "basis": "Hướng dẫn 07-HD/BTCTW", "sources": ""}
        if a.meta:
            meta.update(_load(a.meta))
        print(reports.province_report(_load(a.findings), a.template, a.docx, a.xlsx, meta))


if __name__ == "__main__":
    main()
