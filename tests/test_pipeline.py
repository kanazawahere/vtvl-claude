"""Kiểm thử toàn bộ quy trình trên dữ liệu GIẢ (Bệnh viện Mẫu)."""
import json
import os
import subprocess
import sys

import openpyxl
import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CLI = [sys.executable, os.path.join(ROOT, "scripts", "vtvl_cli.py")]
sys.path.insert(0, os.path.join(ROOT, "scripts"))
from vtvl.xlsx_patch import Edit, FormulaCellError, XlsxPatcher, expand_range  # noqa: E402


def run(*args):
    r = subprocess.run(CLI + list(args), capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    return r.stdout


@pytest.fixture(scope="module")
def work(tmp_path_factory):
    d = tmp_path_factory.mktemp("vtvl")
    run("sample", str(d / "s.xlsx"))
    run("audit", str(d / "s.xlsx"), "-o", str(d / "f.json"))
    return d


def findings(d):
    return json.load(open(d / "f.json", encoding="utf-8"))


def test_audit_finds_seeded_errors(work):
    res = findings(work)
    details = " | ".join(f["detail"] for f in res["findings"])
    for needle in ("Mã trùng", "khoảng trắng thừa", "tiền tố khác", "không có trong A2",
                   "Giờ bổ sung bằng đúng", "vượt 1 ngày công", "Đã kiểm chứng", "chưa có mã"):
        assert needle in details, needle
    assert res["classification"] in "ABCD"


def test_clean_sample_has_no_hard_errors(tmp_path):
    run("sample", str(tmp_path / "c.xlsx"), "--clean")
    run("audit", str(tmp_path / "c.xlsx"), "-o", str(tmp_path / "c.json"))
    res = json.load(open(tmp_path / "c.json", encoding="utf-8"))
    assert res["counts"]["bat_buoc"] == 0


def test_dry_run_writes_nothing(work):
    out = work / "dry.xlsx"
    run("fix-tech", str(work / "s.xlsx"), str(work / "f.json"), "--dry-run")
    assert not out.exists()


def test_fix_form_apply_clean_roundtrip(work):
    d = work
    run("fix-tech", str(d / "s.xlsx"), str(d / "f.json"), "-o", str(d / "s1.xlsx"))
    wb = openpyxl.load_workbook(d / "s1.xlsx")
    a2 = wb["2. A2_CongViec"]
    assert a2["A10"].value == "BVM-CV-008"
    assert a2["A10"].fill.fgColor.rgb == "FFFFFF00"
    assert "Giá trị cũ" in a2["A10"].comment.text
    assert a2["H3"].comment.author == "Mau"          # comment sẵn có được giữ

    run("form", str(d / "f.json"), "-o", str(d / "p.xlsx"))
    wb = openpyxl.load_workbook(d / "p.xlsx")
    ws = wb.active
    for r in range(5, ws.max_row + 1):
        ws.cell(r, 7).value = "Đồng ý" if r % 2 else None
    wb.save(d / "p2.xlsx")
    run("apply-form", str(d / "s1.xlsx"), str(d / "p2.xlsx"), "-o", str(d / "s2.xlsx"))
    run("clean", str(d / "s2.xlsx"), str(d / "s1.xlsx.log.json"), str(d / "s2.xlsx.log.json"), "-o", str(d / "nop.xlsx"))
    wb = openpyxl.load_workbook(d / "nop.xlsx")
    mine = [c for s in wb for row in s.iter_rows() for c in row if c.comment and c.comment.author == "Claude VTVL"]
    assert mine == []
    assert wb["2. A2_CongViec"]["A10"].value == "BVM-CV-008"
    assert wb["2. A2_CongViec"]["A10"].fill.fgColor.rgb in ("00000000", None)
    assert wb["2. A2_CongViec"]["H3"].comment is not None


def test_formula_cells_are_protected(work):
    p = XlsxPatcher(str(work / "s.xlsx"))
    with pytest.raises(FormulaCellError):
        p.apply([Edit("3. B1_KhoiLuong", "F3", 1)])
    p2 = XlsxPatcher(str(work / "s.xlsx"))
    p2.apply([Edit("3. B1_KhoiLuong", "F3", 1)], on_formula="skip")
    assert p2.skipped


def test_never_overwrite_source(work):
    p = XlsxPatcher(str(work / "s.xlsx"))
    p.apply([Edit("1. A1_NhiemVu", "B3", "x")])
    with pytest.raises(RuntimeError):
        p.save(str(work / "s.xlsx"))


def test_reports(work):
    d = work
    run("error-table", str(d / "f.json"), "-o", str(d / "loi.xlsx"))
    run("review", str(d / "f.json"), "-o", str(d / "nx.docx"))
    run("report", str(d / "f.json"), "--docx", str(d / "bc.docx"), "--xlsx", str(d / "bc.xlsx"))
    wb = openpyxl.load_workbook(d / "bc.xlsx")
    assert wb.sheetnames == ["01_TOM_TAT", "02_LOI_BAT_BUOC", "03_CANH_BAO", "04_CHI_SO",
                             "05_DOI_CHIEU_NHIEM_VU", "06_SAI_LECH_20_COT"]
    assert (d / "bc.docx").stat().st_size > 0


def test_expand_range():
    assert expand_range("A1:B2, C5") == ["A1", "A2", "B1", "B2", "C5"]


def test_fill_template_on_sample(tmp_path):
    run("sample", str(tmp_path / "m.xlsx"), "--clean")
    spec = json.load(open(os.path.join(ROOT, "examples", "spec-benh-vien-mau.json"), encoding="utf-8"))
    spec["template"], spec["output"] = str(tmp_path / "m.xlsx"), str(tmp_path / "o.xlsx")
    (tmp_path / "spec.json").write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
    r = subprocess.run([sys.executable, os.path.join(ROOT, "scripts", "fill_template.py"),
                        str(tmp_path / "spec.json"), "--no-gridcraft"], capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    wb = openpyxl.load_workbook(tmp_path / "o.xlsx")
    assert wb["1. A1_NhiemVu"]["A3"].value == "TAM-BVM-A1-001"
    assert "GHI_CHU_DE_XUAT" in wb.sheetnames
