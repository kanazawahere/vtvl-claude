"""Sửa ô trong tệp xlsx ở mức XML: tô màu + comment ghi giá trị cũ, không đụng phần còn lại.

Vì sao không dùng openpyxl để GHI: bộ biểu VTVL 30.000 dòng nặng ~40 MB, openpyxl nạp
vài GB RAM và khi lưu làm rơi khoá trang, danh sách chọn, comment cũ. Ở đây chỉ thay đúng
thẻ <c> của ô cần sửa, thêm kiểu tô màu mới vào styles.xml, thêm comment (tự tạo phần
comments/vml nếu trang chưa có).

Luật an toàn:
- Không bao giờ ghi đè ô có công thức (raise FormulaCellError).
- Tệp ra phải khác tệp vào.
- Mỗi ô sửa được ghi vào nhật ký (old, new, old_style) để bản nộp sạch khôi phục màu gốc.
"""
from __future__ import annotations

import html
import json
import os
import re
import zipfile
from dataclasses import dataclass, field

NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_COMMENTS = NS_REL + "/comments"
REL_VML = NS_REL + "/vmlDrawing"
CT_COMMENTS = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
CT_VML = "application/vnd.openxmlformats-officedocument.vmlDrawing"

COLORS = {"yellow": "FFFFFF00", "pink": "FFF8C8C8", "green": "FFC6EFCE"}

# Thứ tự phần tử con của <worksheet> sau legacyDrawing (chèn legacyDrawing TRƯỚC các thẻ này)
_AFTER_LEGACY = ("legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls",
                 "webPublishItems", "tableParts", "extLst")


class FormulaCellError(RuntimeError):
    pass


def col_to_index(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def index_to_col(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def split_ref(ref: str) -> tuple[str, int]:
    m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
    if not m:
        raise ValueError(f"Ô không hợp lệ: {ref}")
    return m.group(1), int(m.group(2))


def expand_range(spec: str) -> list[str]:
    """'E3:E5, H7' -> ['E3','E4','E5','H7']"""
    out = []
    for part in re.split(r"[,;]\s*", spec.strip()):
        if not part:
            continue
        if ":" in part:
            a, b = part.split(":")
            ca, ra = split_ref(a)
            cb, rb = split_ref(b)
            for c in range(col_to_index(ca), col_to_index(cb) + 1):
                for r in range(ra, rb + 1):
                    out.append(f"{index_to_col(c)}{r}")
        else:
            split_ref(part)
            out.append(part)
    return out


def _cell_re(ref: str):
    return re.compile(r'<c r="' + ref + r'"([^>]*?)(?:/>|>(.*?)</c>)', re.S)


@dataclass
class Edit:
    sheet: str
    cell: str
    value: object = "__KEEP__"   # "__KEEP__" = giữ giá trị, chỉ tô; None = xoá trống
    note: str = ""
    color: str = "yellow"


@dataclass
class XlsxPatcher:
    path: str
    author: str = "Claude VTVL"
    files: dict = field(default_factory=dict)
    log: list = field(default_factory=list)

    def __post_init__(self):
        self._zin = zipfile.ZipFile(self.path)
        self.files = {n: self._zin.read(n) for n in self._zin.namelist()}
        self._sheet_paths = self._map_sheets()
        self._shared = self._load_shared()
        self._styles = self.files["xl/styles.xml"].decode()
        self._fill_ids: dict[str, int] = {}
        self._new_xfs: list[str] = []
        self._xf_cache: dict[tuple, int] = {}
        m = re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", self._styles, re.S)
        self._xfs = re.findall(r"<xf\b[^>]*?(?:/>|>.*?</xf>)", m.group(1), re.S)

    # ---------- đọc cấu trúc ----------
    def _map_sheets(self) -> dict[str, str]:
        wb = self.files["xl/workbook.xml"].decode()
        rels = self.files["xl/_rels/workbook.xml.rels"].decode()
        rid = {}
        for m in re.finditer(r"<Relationship\b[^>]*>", rels):
            tag = m.group(0)
            i = re.search(r'Id="([^"]+)"', tag).group(1)
            t = re.search(r'Target="([^"]+)"', tag).group(1)
            t = t.lstrip("/")
            rid[i] = t if t.startswith("xl/") else "xl/" + t
        out = {}
        for m in re.finditer(r"<sheet\b[^>]*>", wb):
            tag = m.group(0)
            name = html.unescape(re.search(r'name="([^"]*)"', tag).group(1))
            r = re.search(r'r:id="([^"]+)"', tag).group(1)
            out[name] = rid[r]
        return out

    def sheet_names(self) -> list[str]:
        return list(self._sheet_paths)

    def _load_shared(self) -> list[str]:
        raw = self.files.get("xl/sharedStrings.xml")
        if not raw:
            return []
        items = re.findall(r"<si>(.*?)</si>", raw.decode(), re.S)
        return [html.unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", i, re.S))) for i in items]

    def _sheet_xml(self, sheet: str) -> str:
        if sheet not in self._sheet_paths:
            raise KeyError(f"Không có trang '{sheet}'. Có: {', '.join(self._sheet_paths)}")
        return self.files[self._sheet_paths[sheet]].decode()

    def read_cell(self, sheet: str, ref: str):
        x = self._sheet_xml(sheet)
        m = _cell_re(ref).search(x)
        if not m:
            return None
        return self._cell_value(m.group(1), m.group(2) or "")

    def _cell_value(self, attrs: str, inner: str):
        if ' t="s"' in attrs:
            v = re.search(r"<v>(\d+)</v>", inner)
            return self._shared[int(v.group(1))] if v else None
        if ' t="inlineStr"' in attrs:
            return html.unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", inner, re.S)))
        v = re.search(r"<v>(.*?)</v>", inner, re.S)
        if not v:
            return None
        t = v.group(1)
        try:
            f = float(t)
            return int(f) if f.is_integer() else f
        except ValueError:
            return html.unescape(t)

    # ---------- kiểu tô màu ----------
    def _fill_id(self, color: str) -> int:
        if color not in self._fill_ids:
            m = re.search(r'<fills count="(\d+)">(.*?)</fills>', self._styles, re.S)
            n = int(m.group(1))
            rgb = COLORS.get(color, color)
            new = (f'<fills count="{n + 1}">{m.group(2)}<fill><patternFill patternType="solid">'
                   f'<fgColor rgb="{rgb}"/><bgColor indexed="64"/></patternFill></fill></fills>')
            self._styles = self._styles.replace(m.group(0), new, 1)
            self._fill_ids[color] = n
        return self._fill_ids[color]

    def _colored_style(self, s: int, color: str) -> int:
        key = (s, color)
        if key not in self._xf_cache:
            xf = (self._xfs + self._new_xfs)[s]
            xf = re.sub(r' (fillId|applyFill)="\d+"', "", xf)
            xf = xf.replace("<xf ", f'<xf fillId="{self._fill_id(color)}" applyFill="1" ', 1)
            self._new_xfs.append(xf)
            self._xf_cache[key] = len(self._xfs) + len(self._new_xfs) - 1
        return self._xf_cache[key]

    # ---------- ghi ô ----------
    @staticmethod
    def _ensure_cell(x: str, ref: str) -> str:
        if _cell_re(ref).search(x):
            return x
        col, row = split_ref(ref)
        rm = re.search(r'<row r="' + str(row) + r'"([^>]*?)(/>|>(.*?)</row>)', x, re.S)
        if not rm:
            # chèn dòng mới đúng thứ tự
            pos = None
            for m in re.finditer(r'<row r="(\d+)"', x):
                if int(m.group(1)) > row:
                    pos = m.start()
                    break
            if pos is None:
                pos = x.index("</sheetData>") if "</sheetData>" in x else None
                if pos is None:  # <sheetData/>
                    x = x.replace("<sheetData/>", "<sheetData></sheetData>", 1)
                    pos = x.index("</sheetData>")
            return x[:pos] + f'<row r="{row}"><c r="{ref}"/></row>' + x[pos:]
        if rm.group(2) == "/>":
            new = f'<row r="{row}"{rm.group(1)}><c r="{ref}"/></row>'
            return x[:rm.start()] + new + x[rm.end():]
        inner = rm.group(3)
        pos = len(inner)
        for cm in re.finditer(r'<c r="([A-Z]+)\d+"', inner):
            if col_to_index(cm.group(1)) > col_to_index(col):
                pos = cm.start()
                break
        start = rm.start(3)
        return x[:start] + inner[:pos] + f'<c r="{ref}"/>' + inner[pos:] + x[rm.end(3):]

    def apply(self, edits: list[Edit], dry_run: bool = False, on_formula: str = "raise") -> list[dict]:
        """on_formula='raise' dừng khi gặp ô công thức; 'skip' bỏ qua và ghi vào self.skipped."""
        self.skipped = getattr(self, "skipped", [])
        by_sheet: dict[str, list[Edit]] = {}
        for e in edits:
            by_sheet.setdefault(e.sheet, []).append(e)
        plan = []
        for sheet, items in by_sheet.items():
            x = self._sheet_xml(sheet)
            comments = []
            for e in items:
                x = self._ensure_cell(x, e.cell)
                m = _cell_re(e.cell).search(x)
                attrs, inner = m.group(1), m.group(2) or ""
                if "<f" in inner:
                    if on_formula == "skip":
                        self.skipped.append((sheet, e.cell, "ô công thức"))
                        continue
                    raise FormulaCellError(f"Ô công thức, không sửa: {sheet}!{e.cell}")
                old = self._cell_value(attrs, inner)
                sm = re.search(r' s="(\d+)"', attrs)
                old_style = int(sm.group(1)) if sm else 0
                rec = {"sheet": sheet, "cell": e.cell, "old": old,
                       "new": old if e.value == "__KEEP__" else e.value,
                       "old_style": old_style, "note": e.note, "color": e.color,
                       "author": self.author}
                plan.append(rec)
                if dry_run:
                    continue
                ns = self._colored_style(old_style, e.color) if e.color else old_style
                if e.value == "__KEEP__":
                    rest = re.sub(r' s="\d+"', "", attrs)
                    cell = f'<c r="{e.cell}" s="{ns}"{rest}' + (f">{inner}</c>" if inner else "/>")
                elif e.value is None or e.value == "":
                    cell = f'<c r="{e.cell}" s="{ns}"/>'
                elif isinstance(e.value, (int, float)) and not isinstance(e.value, bool):
                    cell = f'<c r="{e.cell}" s="{ns}"><v>{e.value}</v></c>'
                else:
                    txt = html.escape(str(e.value), quote=False)
                    cell = (f'<c r="{e.cell}" s="{ns}" t="inlineStr"><is><t xml:space="preserve">'
                            f"{txt}</t></is></c>")
                x = x[:m.start()] + cell + x[m.end():]
                old_txt = "(trống)" if old in (None, "") else str(old)
                comments.append((e.cell, f"Giá trị cũ: {old_txt}\n{e.note}".strip()))
            if not dry_run:
                self.files[self._sheet_paths[sheet]] = x.encode()
                if comments:
                    self._add_comments(sheet, comments)
        self.log.extend(plan)
        return plan

    # ---------- comment ----------
    def _rels_path(self, sheet: str) -> str:
        p = self._sheet_paths[sheet]
        d, f = os.path.split(p)
        return f"{d}/_rels/{f}.rels"

    def _rel_target(self, sheet: str, rel_type: str):
        rp = self._rels_path(sheet)
        if rp not in self.files:
            return None
        rels = self.files[rp].decode()
        for m in re.finditer(r"<Relationship\b[^>]*>", rels):
            tag = m.group(0)
            if f'Type="{rel_type}"' in tag:
                t = re.search(r'Target="([^"]+)"', tag).group(1)
                if t.startswith("/"):
                    return t.lstrip("/")
                base = os.path.dirname(self._sheet_paths[sheet])
                return os.path.normpath(os.path.join(base, t)).replace("\\", "/")
        return None

    def _add_rel(self, sheet: str, rel_type: str, target_abs: str) -> str:
        rp = self._rels_path(sheet)
        if rp in self.files:
            rels = self.files[rp].decode()
        else:
            rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships '
                    'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
        ids = [int(i) for i in re.findall(r'Id="rId(\d+)"', rels)]
        rid = f"rId{max(ids, default=0) + 1}"
        base = os.path.dirname(self._sheet_paths[sheet])
        rel_t = os.path.relpath(target_abs, base).replace("\\", "/")
        rels = rels.replace("</Relationships>",
                            f'<Relationship Id="{rid}" Type="{rel_type}" Target="{rel_t}"/></Relationships>')
        self.files[rp] = rels.encode()
        return rid

    def _content_type(self, part: str, ct: str, ext: str | None = None):
        c = self.files["[Content_Types].xml"].decode()
        if ext:
            if f'Extension="{ext}"' not in c:
                c = c.replace("</Types>", f'<Default Extension="{ext}" ContentType="{ct}"/></Types>')
        elif f'PartName="/{part}"' not in c:
            c = c.replace("</Types>", f'<Override PartName="/{part}" ContentType="{ct}"/></Types>')
        self.files["[Content_Types].xml"] = c.encode()

    def _new_part_name(self, pattern: str) -> str:
        i = 1
        while pattern.format(i) in self.files:
            i += 1
        return pattern.format(i)

    def _add_comments(self, sheet: str, comments: list[tuple[str, str]]):
        cpath = self._rel_target(sheet, REL_COMMENTS)
        vpath = self._rel_target(sheet, REL_VML)
        if cpath is None:
            cpath = self._new_part_name("xl/comments{}.xml")
            self.files[cpath] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                                 '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                                 "<authors></authors><commentList></commentList></comments>").encode()
            self._add_rel(sheet, REL_COMMENTS, cpath)
            self._content_type(cpath, CT_COMMENTS)
        if vpath is None:
            vpath = self._new_part_name("xl/drawings/vmlDrawing{}.vml")
            self.files[vpath] = ('<xml xmlns:v="urn:schemas-microsoft-com:vml" '
                                 'xmlns:o="urn:schemas-microsoft-com:office:office" '
                                 'xmlns:x="urn:schemas-microsoft-com:office:excel">'
                                 '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>'
                                 '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" '
                                 'path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/>'
                                 '<v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype></xml>').encode()
            rid = self._add_rel(sheet, REL_VML, vpath)
            self._content_type(vpath, CT_VML, ext="vml")
            sp = self._sheet_paths[sheet]
            x = self.files[sp].decode()
            if "xmlns:r=" not in x[:2000]:
                x = re.sub(r"<worksheet\b", f'<worksheet xmlns:r="{NS_REL}"', x, count=1)
            tag = f'<legacyDrawing r:id="{rid}"/>'
            pos = None
            for name in _AFTER_LEGACY:
                m = re.search(r"<" + name + r"\b", x)
                if m and (pos is None or m.start() < pos):
                    pos = m.start()
            if pos is None:
                pos = x.rindex("</worksheet>")
            x = x[:pos] + tag + x[pos:]
            self.files[sp] = x.encode()
        cx = self.files[cpath].decode()
        vx = self.files[vpath].decode()
        author_x = html.escape(self.author, quote=False)
        if f"<author>{author_x}</author>" not in cx:
            if "<authors/>" in cx:
                cx = cx.replace("<authors/>", "<authors></authors>")
            cx = cx.replace("</authors>", f"<author>{author_x}</author></authors>")
        authors = re.findall(r"<author>(.*?)</author>", cx)
        aid = authors.index(author_x)
        if "<commentList/>" in cx:
            cx = cx.replace("<commentList/>", "<commentList></commentList>")
        if '_x0000_t202"' not in vx:
            vx = vx.replace("</xml>", '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" '
                            'path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/>'
                            '<v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype></xml>')
        next_id = max([int(n) for n in re.findall(r"_x0000_s(\d+)", vx)] + [1024]) + 1
        for ref, text in comments:
            # comment cũ cùng ô thì thay (Excel không cho 2 comment một ô)
            cx = re.sub(r'<comment ref="' + ref + r'"[^>]*>.*?</comment>', "", cx, flags=re.S)
            cx = cx.replace("</commentList>",
                            f'<comment ref="{ref}" authorId="{aid}"><text><r><rPr><sz val="9"/>'
                            f'<rFont val="Tahoma"/></rPr><t xml:space="preserve">'
                            f"{html.escape(text, quote=False)}</t></r></text></comment></commentList>")
            col, row = split_ref(ref)
            c = col_to_index(col) - 1
            vx = self._drop_shape(vx, row - 1, c)
            vx = vx.replace("</xml>", (
                f'<v:shape id="_x0000_s{next_id}" type="#_x0000_t202" '
                'style="position:absolute;margin-left:80pt;margin-top:2pt;width:220pt;height:90pt;'
                'z-index:1;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">'
                '<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/>'
                '<v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"/>'
                '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>'
                f"<x:Anchor>{c + 1}, 15, {max(row - 2, 0)}, 2, {c + 4}, 15, {row + 4}, 2</x:Anchor>"
                f"<x:AutoFill>False</x:AutoFill><x:Row>{row - 1}</x:Row><x:Column>{c}</x:Column>"
                "</x:ClientData></v:shape></xml>"))
            next_id += 1
        self.files[cpath] = cx.encode()
        self.files[vpath] = vx.encode()

    @staticmethod
    def _drop_shape(vx: str, row0: int, col0: int) -> str:
        def keep(m):
            s = m.group(0)
            r = re.search(r"<x:Row>(\d+)</x:Row>", s)
            c = re.search(r"<x:Column>(\d+)</x:Column>", s)
            if r and c and int(r.group(1)) == row0 and int(c.group(1)) == col0:
                return ""
            return s
        return re.sub(r"<v:shape\b.*?</v:shape>", keep, vx, flags=re.S)

    # ---------- bỏ dấu (bản nộp sạch) ----------
    def unmark(self, records: list[dict], authors: set[str]):
        """Trả kiểu gốc cho các ô trong nhật ký và xoá comment của các tác giả `authors`."""
        by_sheet: dict[str, list[dict]] = {}
        for r in records:
            by_sheet.setdefault(r["sheet"], []).append(r)
        for sheet, recs in by_sheet.items():
            if sheet not in self._sheet_paths:
                continue
            x = self._sheet_xml(sheet)
            for r in recs:
                m = _cell_re(r["cell"]).search(x)
                if not m:
                    continue
                rest = re.sub(r' s="\d+"', "", m.group(1))
                s = f' s="{r["old_style"]}"' if r.get("old_style") else ""
                cell = f'<c r="{r["cell"]}"{s}{rest}' + (f">{m.group(2)}</c>" if m.group(2) is not None else "/>")
                x = x[:m.start()] + cell + x[m.end():]
            self.files[self._sheet_paths[sheet]] = x.encode()
            cpath = self._rel_target(sheet, REL_COMMENTS)
            vpath = self._rel_target(sheet, REL_VML)
            if not cpath:
                continue
            cx = self.files[cpath].decode()
            names = [html.unescape(a) for a in re.findall(r"<author>(.*?)</author>", cx)]
            mine = {str(i) for i, a in enumerate(names) if a in authors}
            dropped = []

            def drop(m):
                if m.group(2) in mine:
                    dropped.append(m.group(1))
                    return ""
                return m.group(0)
            cx = re.sub(r'<comment ref="([A-Z]+\d+)" authorId="(\d+)".*?</comment>', drop, cx, flags=re.S)
            self.files[cpath] = cx.encode()
            if vpath and dropped:
                vx = self.files[vpath].decode()
                for ref in dropped:
                    col, row = split_ref(ref)
                    vx = self._drop_shape(vx, row - 1, col_to_index(col) - 1)
                self.files[vpath] = vx.encode()

    # ---------- lưu ----------
    def save(self, dst: str, log_path: str | None = None):
        if os.path.abspath(dst) == os.path.abspath(self.path):
            raise RuntimeError("Tệp ra phải khác tệp vào (không ghi đè bản gốc).")
        if self._new_xfs:
            m = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', self._styles, re.S)
            self._styles = self._styles.replace(
                m.group(0),
                f'<cellXfs count="{len(self._xfs) + len(self._new_xfs)}">{m.group(2)}{"".join(self._new_xfs)}</cellXfs>', 1)
        self.files["xl/styles.xml"] = self._styles.encode()
        wb = self.files["xl/workbook.xml"].decode()
        if "fullCalcOnLoad" not in wb:
            if "<calcPr" in wb:
                wb = re.sub(r"<calcPr\b", '<calcPr fullCalcOnLoad="1"', wb, count=1)
            else:
                wb = wb.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
        self.files["xl/workbook.xml"] = wb.encode()
        names = [i.filename for i in self._zin.infolist()]
        extra = [n for n in self.files if n not in names]
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
            for n in names + extra:
                z.writestr(n, self.files[n])
        if log_path is None:
            log_path = dst + ".log.json"
        with open(log_path, "w", encoding="utf-8") as f:
            json.dump(self.log, f, ensure_ascii=False, indent=1, default=str)
        return log_path
