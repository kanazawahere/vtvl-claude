#!/usr/bin/env python3
"""Quét dữ liệu cá nhân / dữ liệu khách trước khi commit hoặc gộp PR. Dính là FAIL (mã thoát 1).

Bắt:
  - họ tên người Việt (họ phổ biến + 1-3 chữ viết hoa), cách gọi "anh/chị/ông/bà + Tên"
  - số điện thoại VN, email, số CCCD (12 số), mã số thuế (10 số hoặc 10-3)
  - tên đơn vị thật: "Bệnh viện/Trung tâm/Sở/Chi cục/UBND/Trường ... + tên riêng"
    (chỉ cho phép đơn vị giả có chữ "Mẫu"/"Giả"/"Ví dụ" hoặc tên cơ quan Trung ương trong allowlist)
  - danh sách chặn dạng BĂM (.pii/blocklist.sha256): tên khách cũ, mã đơn vị... không lộ chữ gốc
Đọc cả nội dung trong .xlsx/.docx (giải nén XML). Tệp nhị phân khác (.xls, .rar, .pdf) bị chặn hẳn.

Dùng:  python3 scripts/pii_scan.py            # quét mọi tệp git đang theo dõi + tệp đã stage
       python3 scripts/pii_scan.py FILE...    # quét tệp chỉ định
Ngoại lệ hợp lệ: thêm biểu thức chính quy vào .pii/allowlist.txt (PR phải nêu lý do).
"""
from __future__ import annotations

import hashlib
import html
import os
import re
import subprocess
import sys
import unicodedata
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SURNAMES = ("Nguyễn Trần Lê Phạm Hoàng Huỳnh Phan Vũ Võ Đặng Bùi Đỗ Hồ Ngô Dương Lý Đinh Trương "
            "Lâm Mai Tô Tạ Châu Quách Lương Lưu Thái Cao Hà Đoàn Kiều Tăng Triệu Văn Từ Diệp La").split()
UP = "A-ZÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬĐÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴ"
LO = "a-zàáảãạăằắẳẵặâầấẩẫậđèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵ"
WORD = f"[{UP}][{LO}]+"
PATTERNS = [
    ("họ tên", re.compile(rf"\b(?:{'|'.join(SURNAMES)})(?:\s+{WORD}){{1,3}}\b")),
    ("xưng hô + tên", re.compile(rf"\b(?:anh|chị|ông|bà|cô|chú|bác|em|bạn)\s+{WORD}(?:\s+{WORD})?\b")),
    ("số điện thoại", re.compile(r"(?<![\d.])(?:\+?84|0)(?:[\s.-]?\d){9}(?![\d])")),
    ("email", re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")),
    ("CCCD/MST", re.compile(r"(?<![\d.,/-])(?:\d{12}|\d{10}(?:-\d{3})?)(?![\d.,/-])")),
    ("tên đơn vị", re.compile(rf"\b(?:Bệnh viện|BV|Trung tâm|Chi cục|Sở|UBND|Ủy ban nhân dân|Trường|Phòng khám|Công ty|Ban)\s+(?:{WORD}\s*){{1,6}}")),
    ("tên tỉnh", re.compile(rf"\b(?:tỉnh|thành phố|TP\.?)\s+{WORD}(?:\s+{WORD})?")),
]
FORBIDDEN_EXT = {".xls", ".rar", ".7z", ".zip", ".pdf", ".doc", ".png", ".jpg", ".jpeg", ".msg", ".eml"}
OFFICE_EXT = {".xlsx", ".docx", ".xlsm"}
SKIP_FILES = {"scripts/pii_scan.py", ".pii/allowlist.txt", ".pii/blocklist.sha256", "LICENSE"}


def fold(s: str) -> str:
    s = s.replace("Đ", "D").replace("đ", "d")
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9 ]+", " ", s.lower()).strip()


def load_lines(p):
    path = os.path.join(ROOT, p)
    if not os.path.exists(path):
        return []
    return [ln.strip() for ln in open(path, encoding="utf-8") if ln.strip() and not ln.startswith("#")]


ALLOW = [re.compile(x) for x in load_lines(".pii/allowlist.txt")]
BLOCK = set(load_lines(".pii/blocklist.sha256"))


def allowed(hit: str) -> bool:
    return any(a.search(hit) for a in ALLOW)


def text_of(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext in OFFICE_EXT:
        out = []
        with zipfile.ZipFile(path) as z:
            for n in z.namelist():
                if n.endswith(".xml") or n.endswith(".vml"):
                    raw = z.read(n).decode("utf-8", "ignore")
                    out.append(html.unescape(re.sub(r"<[^>]+>", " ", raw)))
        return "\n".join(out)
    with open(path, encoding="utf-8", errors="ignore") as f:
        return f.read()


def ngrams(words, n_max=4):
    for n in range(1, n_max + 1):
        for i in range(len(words) - n + 1):
            yield " ".join(words[i:i + n])


def scan_file(path: str, rel: str) -> list[str]:
    ext = os.path.splitext(path)[1].lower()
    if ext in FORBIDDEN_EXT:
        return [f"{rel}: loại tệp {ext} không được đưa lên kho (chỉ dữ liệu giả dạng .xlsx/.docx/.json/.md)"]
    try:
        txt = text_of(path)
    except Exception as e:  # tệp hỏng cũng coi là lỗi
        return [f"{rel}: không đọc được ({e})"]
    issues = []
    for ln_no, line in enumerate(txt.splitlines(), 1):
        for label, pat in PATTERNS:
            for m in pat.finditer(line):
                hit = m.group(0).strip()
                if not allowed(hit):
                    issues.append(f"{rel}:{ln_no}: [{label}] {hit}")
        if BLOCK:
            for g in ngrams(fold(line).split()):
                if hashlib.sha256(g.encode()).hexdigest() in BLOCK:
                    issues.append(f"{rel}:{ln_no}: [danh sách chặn] (chuỗi bị chặn, xem .pii/blocklist.sha256)")
                    break
    return issues


def tracked_files() -> list[str]:
    try:
        out = subprocess.run(["git", "ls-files", "-co", "--exclude-standard"], cwd=ROOT,
                             capture_output=True, text=True, check=True).stdout.split("\n")
    except Exception:
        out = []
        for d, _, fs in os.walk(ROOT):
            if "/.git" in d:
                continue
            out += [os.path.relpath(os.path.join(d, f), ROOT) for f in fs]
    return [f for f in out if f]


def main(argv):
    files = argv or tracked_files()
    issues = []
    for rel in files:
        rel = os.path.relpath(os.path.abspath(rel), ROOT) if os.path.isabs(rel) else rel
        if rel in SKIP_FILES or rel.startswith(".git/"):
            continue
        path = os.path.join(ROOT, rel)
        if os.path.isfile(path):
            issues += scan_file(path, rel)
    for i in issues:
        print(i)
    if issues:
        print(f"\nFAIL: {len(issues)} điểm nghi dữ liệu cá nhân/dữ liệu khách. "
              "Thay bằng dữ liệu giả ('Bệnh viện Mẫu'), hoặc nếu là thuật ngữ công khai thì thêm vào .pii/allowlist.txt kèm lý do.")
        return 1
    print(f"OK: {len(files)} tệp, không thấy dữ liệu cá nhân.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
