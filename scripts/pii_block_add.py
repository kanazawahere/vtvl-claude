#!/usr/bin/env python3
"""Thêm chuỗi cần chặn vào .pii/blocklist.sha256 dưới dạng BĂM (không lộ chữ gốc lên kho)."""
import hashlib
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pii_scan import ROOT, fold  # noqa: E402

path = os.path.join(ROOT, ".pii", "blocklist.sha256")
with open(path, "a", encoding="utf-8") as f:
    for t in sys.argv[1:]:
        g = fold(t)
        if len(g.split()) > 4:
            sys.exit("Chuỗi tối đa 4 từ (máy quét so theo cụm 1-4 từ).")
        f.write(hashlib.sha256(g.encode()).hexdigest() + "\n")
print("Đã thêm", len(sys.argv) - 1, "chuỗi.")
