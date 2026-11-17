"""Máy quét dữ liệu cá nhân phải bắt được dữ liệu thật và để qua dữ liệu giả."""
import base64
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "scripts"))
import pii_scan  # noqa: E402


def scan_text(tmp_path, text, name="x.md"):
    p = tmp_path / name
    p.write_text(text, encoding="utf-8")
    return pii_scan.scan_file(str(p), name)


def test_catches_personal_data(tmp_path):
    # Mã hoá base64 để máy quét không bắt nhầm chính tệp kiểm thử (dữ liệu bịa).
    bad = [base64.b64decode(x).decode() for x in [
           'TmfGsOG7nWkgbOG6rXA6IFRy4bqnbiBUaOG7iyBCw61jaCBOZ+G7jWM=',
           'R+G7jWkgY2jhu4sgSOG6oW5oIMSR4buDIHjDoWMgbmjhuq1u',
           'U8SQVCAwOTEyIDM0NSA2Nzg=',
           'RW1haWwgYS5iQGV4YW1wbGUuY29t',
           'Q0NDRCAwMTIzNDU2Nzg5MDE=',
           'TVNUIDAxMDEyMzQ1Njc=',
           'QuG7h25oIHZp4buHbiDEkGEgS2hvYSBIw7JhIELDrG5o',
           'U+G7nyBO4buZaSBW4bul',
           'dOG7iW5oIELDrG5oIETGsMahbmc=']]
    for t in bad:
        assert scan_text(tmp_path, t), t


def test_allows_fake_data(tmp_path):
    ok = ["Đơn vị: Bệnh viện Mẫu", "Mã BVM-CV-001, quỹ giờ 1.904", "Hướng dẫn 07-HD/BTCTW ngày 19/10/2026",
          "Ban Tổ chức Trung ương"]
    for t in ok:
        assert not scan_text(tmp_path, t), (t, scan_text(tmp_path, t))


def test_blocks_binary(tmp_path):
    assert scan_text(tmp_path, "x", "hoso.rar")
