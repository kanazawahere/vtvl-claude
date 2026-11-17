#!/usr/bin/env python3
"""Điền bộ biểu VTVL (mẫu 30.000 dòng, 30.000 dòng) bằng cách sửa thẳng XML trong tệp xlsx.

Vì sao sửa XML mà không dùng openpyxl:
- Tệp mẫu ~39 MB, ~4 triệu ô, có calcChain, comment, khoá trang, danh sách chọn.
  openpyxl nạp hết vào RAM (vài GB) và khi lưu dễ làm rơi chi tiết định dạng.
- Sửa XML chỉ đụng đúng các ô cần ghi; mọi thứ khác (công thức, khoá, danh sách chọn,
  tên vùng, comment, định dạng) giữ nguyên từng byte.

Cách dùng:
    python3 fill_template.py spec.json            # điền + gridcraft convert + check
    python3 fill_template.py spec.json --no-gridcraft

Tệp spec.json (xem templates/mau-30000-dong/BAN-DO-MAU.md, mục "Script điền"):
{
  "template": "/đường/dẫn/goc.xlsx",            # không bao giờ bị ghi đè
  "output":   "/đường/dẫn/ban-de-xuat.xlsx",
  "archive":  [ {"source": "2. A2_CongViec", "target": "CU_A2", "first_row": 1,
                 "last_row": 60, "columns": "A:H", "title": "..."} ],
  "fill":     [ {"sheet": "2. A2_CongViec", "start_row": 3, "clear_to_row": 60,
                 "columns": ["A","B",...], "rows": [ {"A": "...", "B": "..."}, ... ]} ],
  "cells":    [ {"sheet": "1. A1_NhiemVu", "cell": "F1", "value": "BẢN ĐỀ XUẤT"} ],
  "new_sheets": [ {"name": "GHI_CHU_DE_XUAT", "rows": [["..",".."], ...],
                   "col_widths": [30, 80], "header_rows": [1]} ],
  "styles":   {"header": "1. A1_NhiemVu!A2", "body": "1. A1_NhiemVu!B3"},
  "active_sheet": "GHI_CHU_DE_XUAT",
  "gridcraft": {"convert_to": "/đường/dẫn/ban-de-xuat-tinh-lai.xlsx", "timeout": 1500}
}

Luật an toàn cài sẵn:
- Không ghi đè ô có công thức (báo lỗi và dừng).
- Với mẫu 30.000 dòng: chỉ cho ghi vào cột NHẬP đã khai trong INPUT_COLUMNS (trừ khi --allow-any-column).
- Không bao giờ ghi vào tệp mẫu; tệp ra phải khác tệp mẫu.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from xml.sax.saxutils import escape

from lxml import etree

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
WORKSHEET_TYPE = NS_REL + "/worksheet"
WORKSHEET_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"

# Cột NHẬP của mẫu 30.000 dòng (theo HD_DIEN_BIEU + GHI_CHU_BAN_SUA). Cột còn lại là công thức hoặc chỉ đọc.
INPUT_COLUMNS = {
    "1. A1_NhiemVu": "ABCDEFG",
    "2. A2_CongViec": "ABCDEFGH",
    "3. B1_KhoiLuong": "ACDEHIJKLS",
    "4. B2_CT_TacDong": "ACDEFGIJKLMQRS",
    "6. C1_VTVL": "ABCDFGHIJ",
    "7. C2_NangLuc": "ABCDEFG",
    "8. C3_NhuCau": "ABHIJKL",
    "9. D1_NhanLuc": "ABCEFGHIK",
    "10. D2_PhuongAn": "ABDEFGHIN",
    "11. DM_SAN_PHAM": "ABCDEFGHIJ",
    "12. CHIA_VUNG_CAP_XA": "ABC",
    "15. S_NhatKy": "ABCDEFGHIJ",
    "DU_BAO_5_NAM": "ABCDEFJ",
    "Y_KIEN_DON_VI": "ABCDEFGHIJK",
    "NGUON_NHIEM_VU": "ABCDEFGHIJKL",
    "NGUON_VTVL": "ABCDEFGHIJ",
    "NGUON_NHAN_LUC": "ABCDEFGHIJK",
}

ROW_RE = re.compile(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)', re.S)
CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)


# ---------------------------------------------------------------- tiện ích cột

def column_to_index(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def index_to_column(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def expand_columns(spec: str) -> list[str]:
    """'A:H' -> ['A',...,'H']; 'A,C,E' -> ['A','C','E']."""
    if ":" in spec:
        a, b = spec.split(":")
        return [index_to_column(i) for i in range(column_to_index(a), column_to_index(b) + 1)]
    return [c.strip() for c in spec.split(",") if c.strip()]


def clean_xml_text(text: str) -> str:
    # bỏ ký tự điều khiển XML không cho phép
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)


# ---------------------------------------------------------------- đọc gói xlsx

class Workbook:
    """Gói xlsx đọc lười: chỉ nạp phần XML nào cần sửa."""

    def __init__(self, path: str):
        self.path = path
        self.zip = zipfile.ZipFile(path)
        self.modified: dict[str, bytes] = {}
        self.added: dict[str, bytes] = {}
        self._load_workbook_map()
        self.shared_strings = self._load_shared_strings()

    def read(self, name: str) -> bytes:
        if name in self.modified:
            return self.modified[name]
        if name in self.added:
            return self.added[name]
        return self.zip.read(name)

    def _load_workbook_map(self):
        self.workbook_xml = self.read("xl/workbook.xml").decode("utf-8")
        rels = etree.fromstring(self.read("xl/_rels/workbook.xml.rels"))
        rid_to_target = {r.get("Id"): r.get("Target") for r in rels}
        self.sheet_order: list[str] = []
        self.sheet_path: dict[str, str] = {}
        for m in re.finditer(r'<sheet [^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"[^>]*/>', self.workbook_xml):
            name = unescape_attr(m.group(1))
            target = rid_to_target[m.group(2)]
            path = target.lstrip("/") if target.startswith("/") else "xl/" + target
            self.sheet_order.append(name)
            self.sheet_path[name] = path

    def _load_shared_strings(self) -> list[str]:
        try:
            root = etree.fromstring(self.read("xl/sharedStrings.xml"))
        except KeyError:
            return []
        out = []
        for si in root.findall(f"{{{NS_MAIN}}}si"):
            out.append("".join(t.text or "" for t in si.iter(f"{{{NS_MAIN}}}t")))
        return out

    def sheet_xml(self, name: str) -> str:
        return self.read(self.sheet_path[name]).decode("utf-8")

    def set_sheet_xml(self, name: str, xml: str):
        self.modified[self.sheet_path[name]] = xml.encode("utf-8")


def unescape_attr(s: str) -> str:
    return (s.replace("&quot;", '"').replace("&apos;", "'").replace("&lt;", "<")
            .replace("&gt;", ">").replace("&amp;", "&"))


# ---------------------------------------------------------------- đọc/ghi ô trong một trang

class SheetEditor:
    """Sửa sheetData của một trang ở mức chuỗi, chỉ tách các dòng cần đụng tới."""

    def __init__(self, wb: Workbook, sheet: str):
        self.wb = wb
        self.sheet = sheet
        self.xml = wb.sheet_xml(sheet)
        sd_start = self.xml.find("<sheetData")
        sd_open_end = self.xml.find(">", sd_start) + 1
        if self.xml[sd_open_end - 2] == "/":  # <sheetData/>
            self.head = self.xml[:sd_start] + "<sheetData>"
            self.body = ""
            self.tail = "</sheetData>" + self.xml[sd_open_end:]
        else:
            sd_end = self.xml.find("</sheetData>")
            self.head = self.xml[:sd_open_end]
            self.body = self.xml[sd_open_end:sd_end]
            self.tail = self.xml[sd_end:]
        self._index_rows()
        self.row_edits: dict[int, dict[str, tuple]] = {}  # row -> col -> ("set", value) | ("clear",)

    def _index_rows(self):
        self.rows: dict[int, tuple[int, int]] = {}
        for m in ROW_RE.finditer(self.body):
            self.rows[int(m.group(1))] = (m.start(), m.end())

    def row_cells(self, row: int) -> dict[str, dict]:
        """Trả {cột: {attrs, inner, formula, value}} của một dòng (giá trị đã giải mã)."""
        if row not in self.rows:
            return {}
        a, b = self.rows[row]
        row_xml = self.body[a:b]
        cells = {}
        for m in CELL_RE.finditer(row_xml):
            col, attrs, inner = m.group(1), m.group(3), m.group(4) or ""
            cells[col] = {
                "attrs": attrs,
                "inner": inner,
                "formula": "<f" in inner,
                "value": self._decode_value(attrs, inner),
            }
        return cells

    def _decode_value(self, attrs: str, inner: str):
        t = re.search(r'\bt="([^"]+)"', attrs)
        t = t.group(1) if t else "n"
        if t == "inlineStr":
            return unescape_attr("".join(re.findall(r"<t[^>]*>(.*?)</t>", inner, re.S)))
        v = re.search(r"<v>(.*?)</v>", inner, re.S)
        if not v:
            return None
        raw = v.group(1)
        if t == "s":
            return self.wb.shared_strings[int(raw)]
        if t in ("str", "e"):
            return unescape_attr(raw)
        if t == "b":
            return raw == "1"
        try:
            f = float(raw)
            return int(f) if f.is_integer() else f
        except ValueError:
            return raw

    def get_value(self, ref: str):
        col, row = re.match(r"([A-Z]+)(\d+)", ref).groups()
        cell = self.row_cells(int(row)).get(col)
        return None if cell is None else cell["value"]

    def get_style(self, ref: str) -> str | None:
        col, row = re.match(r"([A-Z]+)(\d+)", ref).groups()
        cell = self.row_cells(int(row)).get(col)
        if not cell:
            return None
        s = re.search(r'\bs="(\d+)"', cell["attrs"])
        return s.group(1) if s else None

    def set_value(self, row: int, col: str, value):
        self.row_edits.setdefault(row, {})[col] = ("set", value)

    def clear(self, row: int, col: str):
        self.row_edits.setdefault(row, {}).setdefault(col, ("clear",))

    # -- dựng lại XML
    def _cell_xml(self, col: str, row: int, style: str | None, value) -> str:
        s_attr = f' s="{style}"' if style else ""
        ref = f"{col}{row}"
        if value is None or value == "":
            return f'<c r="{ref}"{s_attr}/>'
        if isinstance(value, bool):
            return f'<c r="{ref}"{s_attr} t="b"><v>{int(value)}</v></c>'
        if isinstance(value, (int, float)):
            return f'<c r="{ref}"{s_attr}><v>{value}</v></c>'
        text = escape(clean_xml_text(str(value)))
        return f'<c r="{ref}"{s_attr} t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'

    def _render_row(self, row: int, edits: dict) -> str:
        if row in self.rows:
            a, b = self.rows[row]
            row_xml = self.body[a:b]
            open_tag = re.match(r"<row [^>]*?/?>", row_xml).group(0)
            if open_tag.endswith("/>"):
                open_tag = open_tag[:-2] + ">"
        else:
            open_tag = f'<row r="{row}">'
        existing = self.row_cells(row)
        out_cells: dict[int, str] = {}
        for col, cell in existing.items():
            out_cells[column_to_index(col)] = None  # đánh dấu giữ nguyên, điền sau
        # giữ nguyên ô không sửa (lấy lại đúng chuỗi gốc)
        if row in self.rows:
            a, b = self.rows[row]
            for m in CELL_RE.finditer(self.body[a:b]):
                out_cells[column_to_index(m.group(1))] = m.group(0)
        for col, edit in edits.items():
            cell = existing.get(col)
            style = None
            if cell:
                s = re.search(r'\bs="(\d+)"', cell["attrs"])
                style = s.group(1) if s else None
                if cell["formula"]:
                    if edit[0] == "set":
                        raise RuntimeError(
                            f"[{self.sheet}] Ô {col}{row} là CÔNG THỨC, không được ghi đè.")
                    continue  # xoá: bỏ qua ô công thức
            elif edit[0] == "clear":
                continue
            value = edit[1] if edit[0] == "set" else None
            out_cells[column_to_index(col)] = self._cell_xml(col, row, style, value)
        cells_xml = "".join(out_cells[k] for k in sorted(out_cells))
        return f"{open_tag}{cells_xml}</row>"

    def save(self):
        if not self.row_edits:
            return
        pieces = []
        cursor = 0
        existing_rows = sorted(self.rows)
        new_rows = sorted(r for r in self.row_edits if r not in self.rows)
        # chèn dòng mới theo thứ tự: tìm vị trí trước dòng lớn hơn đầu tiên
        insert_map: dict[int, list[int]] = {}
        for r in new_rows:
            nxt = next((x for x in existing_rows if x > r), None)
            insert_map.setdefault(nxt, []).append(r)
        for r in existing_rows:
            a, b = self.rows[r]
            for nr in insert_map.get(r, []):
                pieces.append(self.body[cursor:a])
                cursor = a
                pieces.append(self._render_row(nr, self.row_edits[nr]))
            if r in self.row_edits:
                pieces.append(self.body[cursor:a])
                pieces.append(self._render_row(r, self.row_edits[r]))
                cursor = b
        pieces.append(self.body[cursor:])
        for nr in insert_map.get(None, []):
            pieces.append(self._render_row(nr, self.row_edits[nr]))
        self.body = "".join(pieces)
        self.wb.set_sheet_xml(self.sheet, self.head + self.body + self.tail)


# ---------------------------------------------------------------- các bước chính

def archive_values(wb: Workbook, item: dict) -> list[list]:
    """Chép GIÁ TRỊ (công thức lấy giá trị lưu sẵn) của một vùng sang danh sách dòng."""
    ed = SheetEditor(wb, item["source"])
    cols = expand_columns(item.get("columns", "A:Z"))
    out = []
    for r in range(item.get("first_row", 1), item["last_row"] + 1):
        cells = ed.row_cells(r)
        vals = [cells.get(c, {}).get("value") for c in cols]
        if item.get("skip_empty", True) and all(v in (None, "") for v in vals):
            continue
        out.append([r] + vals)
    return out


def fill_sheet(wb: Workbook, item: dict, strict: bool, log: list):
    sheet = item["sheet"]
    allowed = INPUT_COLUMNS.get(sheet)
    ed = SheetEditor(wb, sheet)
    start = item.get("start_row", 3)
    cols = item.get("columns") or sorted({c for row in item["rows"] for c in row})
    if strict and allowed:
        bad = [c for c in cols if c not in allowed]
        if bad:
            raise RuntimeError(f"[{sheet}] Cột {bad} không phải cột NHẬP của mẫu 30.000 dòng.")
    clear_to = item.get("clear_to_row")
    if clear_to:
        for r in range(start, clear_to + 1):
            for c in cols:
                ed.clear(r, c)
    for i, row in enumerate(item["rows"]):
        r = start + i
        for c, v in row.items():
            if strict and allowed and c not in allowed:
                raise RuntimeError(f"[{sheet}] Cột {c} không phải cột NHẬP.")
            ed.set_value(r, c, v)
    ed.save()
    log.append(f"{sheet}: ghi {len(item['rows'])} dòng từ dòng {start}"
               + (f", xoá trước vùng {start}-{clear_to} cột {''.join(cols)}" if clear_to else ""))


def set_cells(wb: Workbook, items: list, log: list):
    by_sheet: dict[str, list] = {}
    for it in items:
        by_sheet.setdefault(it["sheet"], []).append(it)
    for sheet, its in by_sheet.items():
        ed = SheetEditor(wb, sheet)
        for it in its:
            col, row = re.match(r"([A-Z]+)(\d+)", it["cell"]).groups()
            ed.set_value(int(row), col, it["value"])
        ed.save()
        log.append(f"{sheet}: ghi {len(its)} ô lẻ")


def build_sheet_xml(rows: list[list], col_widths: list | None, header_rows: set,
                    header_style: str | None, body_style: str | None, tab_color: str | None) -> str:
    parts = [f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
             f'<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_REL}">']
    if tab_color:
        parts.append(f'<sheetPr><tabColor rgb="{tab_color}"/></sheetPr>')
    ncol = max((len(r) for r in rows), default=1)
    parts.append(f'<dimension ref="A1:{index_to_column(max(ncol,1))}{max(len(rows),1)}"/>')
    parts.append('<sheetViews><sheetView workbookViewId="0"/></sheetViews>')
    parts.append('<sheetFormatPr defaultRowHeight="15"/>')
    if col_widths:
        parts.append("<cols>" + "".join(
            f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>'
            for i, w in enumerate(col_widths, 1)) + "</cols>")
    parts.append("<sheetData>")
    for ri, row in enumerate(rows, 1):
        style = header_style if ri in header_rows else body_style
        s_attr = f' s="{style}"' if style else ""
        cells = []
        for ci, v in enumerate(row, 1):
            ref = f"{index_to_column(ci)}{ri}"
            if v is None or v == "":
                cells.append(f'<c r="{ref}"{s_attr}/>')
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                cells.append(f'<c r="{ref}"{s_attr}><v>{v}</v></c>')
            else:
                t = escape(clean_xml_text(str(v)))
                cells.append(f'<c r="{ref}"{s_attr} t="inlineStr"><is><t xml:space="preserve">{t}</t></is></c>')
        parts.append(f'<row r="{ri}">' + "".join(cells) + "</row>")
    parts.append("</sheetData>")
    parts.append('<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>')
    parts.append("</worksheet>")
    return "".join(parts)


def add_sheets(wb: Workbook, sheets: list, styles: dict, log: list):
    header_style = body_style = None
    if styles.get("header"):
        sh, ref = styles["header"].split("!")
        header_style = SheetEditor(wb, sh).get_style(ref)
    if styles.get("body"):
        sh, ref = styles["body"].split("!")
        body_style = SheetEditor(wb, sh).get_style(ref)

    rels_xml = wb.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    ct_xml = wb.read("[Content_Types].xml").decode("utf-8")
    wb_xml = wb.workbook_xml
    existing_files = set(wb.zip.namelist()) | set(wb.added)
    max_sheet_id = max(int(x) for x in re.findall(r'<sheet [^>]*?sheetId="(\d+)"', wb_xml))
    rid_nums = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels_xml)]
    next_rid = max(rid_nums) + 1
    n = 1
    for spec in sheets:
        name = spec["name"]
        if name in wb.sheet_order:
            raise RuntimeError(f"Trang '{name}' đã có trong mẫu.")
        if len(name) > 31 or re.search(r"[\[\]:*?/\\]", name):
            raise RuntimeError(f"Tên trang '{name}' không hợp lệ cho Excel.")
        while f"xl/worksheets/sheet{n}.xml" in existing_files:
            n += 1
        part = f"xl/worksheets/sheet{n}.xml"
        existing_files.add(part)
        xml = build_sheet_xml(spec["rows"], spec.get("col_widths"), set(spec.get("header_rows", [1])),
                              header_style, body_style, spec.get("tab_color", "FFC00000"))
        wb.added[part] = xml.encode("utf-8")
        rid = f"rId{next_rid}"
        next_rid += 1
        max_sheet_id += 1
        rels_xml = rels_xml.replace(
            "</Relationships>",
            f'<Relationship Id="{rid}" Type="{WORKSHEET_TYPE}" Target="worksheets/sheet{n}.xml"/></Relationships>')
        ct_xml = ct_xml.replace(
            "</Types>", f'<Override PartName="/{part}" ContentType="{WORKSHEET_CT}"/></Types>')
        wb_xml = wb_xml.replace(
            "</sheets>",
            f'<sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="{escape(name, {chr(34): "&quot;"})}" sheetId="{max_sheet_id}" r:id="{rid}"/></sheets>')
        wb.sheet_order.append(name)
        wb.sheet_path[name] = part
        log.append(f"Thêm trang {name} ({len(spec['rows'])} dòng)")
    wb.modified["xl/_rels/workbook.xml.rels"] = rels_xml.encode("utf-8")
    wb.modified["[Content_Types].xml"] = ct_xml.encode("utf-8")
    wb.workbook_xml = wb_xml


def set_active_sheet(wb: Workbook, name: str):
    idx = wb.sheet_order.index(name)
    wb_xml = wb.workbook_xml
    wb_xml = re.sub(r'activeTab="\d+"', f'activeTab="{idx}"', wb_xml)
    if 'activeTab="' not in wb_xml:
        wb_xml = wb_xml.replace("<workbookView ", f'<workbookView activeTab="{idx}" ', 1)
    wb_xml = re.sub(r'firstSheet="\d+"', 'firstSheet="0"', wb_xml)
    wb.workbook_xml = wb_xml
    # bỏ tabSelected ở các trang khác để Excel không gộp nhóm trang
    for sheet, path in wb.sheet_path.items():
        data = wb.read(path)
        head = data[:6000]
        if b'tabSelected="1"' in head:
            text = data.decode("utf-8")
            text = text.replace(' tabSelected="1"', "", 1)
            wb.modified[path] = text.encode("utf-8")
    path = wb.sheet_path[name]
    text = wb.read(path).decode("utf-8")
    text = text.replace("<sheetView ", '<sheetView tabSelected="1" ', 1)
    target = wb.added if path in wb.added else wb.modified
    target[path] = text.encode("utf-8")


def force_recalc_on_open(wb: Workbook):
    xml = wb.workbook_xml
    if "<calcPr" in xml and "fullCalcOnLoad" not in xml:
        xml = xml.replace("<calcPr ", '<calcPr fullCalcOnLoad="1" ', 1)
    wb.workbook_xml = xml


def write_package(wb: Workbook, out_path: str):
    wb.modified["xl/workbook.xml"] = wb.workbook_xml.encode("utf-8")
    tmp = out_path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zout:
        for info in wb.zip.infolist():
            data = wb.modified.get(info.filename)
            if data is None:
                data = wb.zip.read(info.filename)
            zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zout.writestr(zi, data)
        for name, data in wb.added.items():
            zi = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            zout.writestr(zi, data)
    os.replace(tmp, out_path)


def run_gridcraft(out_path: str, cfg: dict, log: list) -> dict:
    result = {}
    timeout = cfg.get("timeout", 1500)
    convert_to = cfg.get("convert_to")
    if convert_to:
        try:
            p = subprocess.run(["gridcraft-cli", "convert", out_path, convert_to],
                               capture_output=True, text=True, timeout=timeout)
            result["convert"] = {"code": p.returncode, "out": (p.stdout + p.stderr)[-4000:]}
        except subprocess.TimeoutExpired:
            result["convert"] = {"code": None, "out": f"QUÁ {timeout}s – dừng"}
        log.append(f"gridcraft convert -> {convert_to}: {result['convert']['code']}")
    check_target = convert_to if convert_to and os.path.exists(convert_to) else out_path
    try:
        p = subprocess.run(["gridcraft-cli", "check", check_target, "--max-report", "60"],
                           capture_output=True, text=True, timeout=timeout)
        result["check"] = {"file": check_target, "code": p.returncode, "out": (p.stdout + p.stderr)[-12000:]}
    except subprocess.TimeoutExpired:
        result["check"] = {"file": check_target, "code": None, "out": f"QUÁ {timeout}s – dừng"}
    log.append(f"gridcraft check {check_target}: {result['check']['code']}")
    return result


def main():
    ap = argparse.ArgumentParser(description="Điền bộ biểu VTVL (mẫu 30.000 dòng) từ tệp spec JSON.")
    ap.add_argument("spec")
    ap.add_argument("--no-gridcraft", action="store_true")
    ap.add_argument("--allow-any-column", action="store_true",
                    help="cho ghi cả cột ngoài danh sách cột NHẬP (vẫn cấm ghi đè công thức)")
    args = ap.parse_args()
    spec = json.load(open(args.spec, encoding="utf-8"))
    template, output = spec["template"], spec["output"]
    if os.path.abspath(template) == os.path.abspath(output):
        sys.exit("Tệp ra trùng tệp mẫu: không được ghi đè mẫu.")
    log: list[str] = []
    wb = Workbook(template)

    archived = []
    for item in spec.get("archive", []):
        rows = archive_values(wb, item)
        header = ["Dòng gốc"] + expand_columns(item.get("columns", "A:Z"))
        title = item.get("title", f"Bản lưu giá trị cũ của {item['source']}")
        archived.append({"name": item["target"],
                         "rows": [[title], header] + rows,
                         "header_rows": [1, 2],
                         "col_widths": item.get("col_widths"),
                         "tab_color": "FF7F7F7F"})
        log.append(f"Lưu {len(rows)} dòng của {item['source']} sang {item['target']}")

    for item in spec.get("fill", []):
        fill_sheet(wb, item, strict=not args.allow_any_column, log=log)
    if spec.get("cells"):
        set_cells(wb, spec["cells"], log)
    add_sheets(wb, spec.get("new_sheets", []) + archived, spec.get("styles", {}), log)
    if spec.get("active_sheet"):
        set_active_sheet(wb, spec["active_sheet"])
    force_recalc_on_open(wb)
    write_package(wb, output)
    log.append(f"Đã ghi {output}")

    report = {"log": log}
    if not args.no_gridcraft and spec.get("gridcraft") is not None:
        report["gridcraft"] = run_gridcraft(output, spec["gridcraft"], log)
    print(json.dumps(report, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
