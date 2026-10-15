"""เขียนข้อความภาษาไทยลงใน PDF ให้ถูกต้องทั้ง "ตาเห็น" และ "ข้อความที่ค้นหา/คัดลอกได้"

วิธีทำ
  1. จัดรูปตัวอักษรด้วย HarfBuzz (เครื่องมือเดียวกับที่เบราว์เซอร์และ Word ใช้) -> สระ/วรรณยุกต์ซ้อนถูกตำแหน่ง
  2. วาดเส้นขอบตัวอักษรเป็นเวกเตอร์ลงในหน้า PDF -> ตาเห็นถูก 100% ทุกโปรแกรม
  3. ใส่ข้อความ Unicode ที่ถูกต้องแบบมองไม่เห็นทับตำแหน่งเดียวกัน -> ค้นหา / copy ได้ถูก

ฟอนต์
  sarabun : "Sarabun TH Compat" ขนาดและระยะเท่า TH Sarabun New / TH SarabunIT๙ (ใช้กับเอกสารราชการไทยส่วนใหญ่)
  noto    : Noto Sans Thai (ใช้กับเอกสารที่เป็นฟอนต์ไม่มีหัว เช่น Tahoma / Leelawadee)
"""
import os
import re
import uuid
from functools import lru_cache

import pymupdf as fitz
import uharfbuzz as hb
from fontTools.pens.basePen import BasePen
from fontTools.ttLib import TTFont

FONT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static", "fonts")
FACE_FILES = {
    ("sarabun", False, False): "SarabunTHCompat-Regular.ttf",
    ("sarabun", True, False): "SarabunTHCompat-Bold.ttf",
    ("sarabun", False, True): "SarabunTHCompat-Italic.ttf",
    ("sarabun", True, True): "SarabunTHCompat-BoldItalic.ttf",
    ("noto", False, False): "NotoSansThai-Regular.ttf",
    ("noto", True, False): "NotoSansThai-Bold.ttf",
}
SARABUN_LIKE = re.compile(r"sarabun|angsana|cordia|browallia|eucrosia|freesia|iris|jasmine|kodchiang|lily|"
                          r"niramit|kodchasan|krub|k2d|charm|baijam|fahkwang|srisakdi|chakra|mali|"
                          r"^th\s|^th[A-Z]|\bTH ", re.I)
SANS_LIKE = re.compile(r"tahoma|leelawadee|thonburi|sukhumvit|noto ?sans|arial|helvet|segoe|loma|garuda|"
                       r"kanit|prompt|mitr|sarabun-?ui", re.I)
INVISIBLE_FACE = ("sarabun", False, False)          # ฟอนต์สำหรับชั้นข้อความที่มองไม่เห็น (มีทั้งไทย/อังกฤษ)


class _Pen(BasePen):
    def __init__(self, glyphset):
        super().__init__(glyphset)
        self.ops = []

    def _moveTo(self, p):
        self.ops.append(("M", p))

    def _lineTo(self, p):
        self.ops.append(("L", p))

    def _curveToOne(self, p1, p2, p3):
        self.ops.append(("C", p1, p2, p3))

    def _closePath(self):
        self.ops.append(("Z",))

    def _endPath(self):
        self.ops.append(("Z",))


class Face:
    """ฟอนต์หนึ่งแบบ: จัดรูปด้วย HarfBuzz + เส้นขอบตัวอักษรจาก fontTools"""

    def __init__(self, path):
        self.path = path
        data = open(path, "rb").read()
        self.hb_face = hb.Face(data)
        self.hb_font = hb.Font(self.hb_face)
        self.upm = self.hb_face.upem
        self.hb_font.scale = (self.upm, self.upm)
        self.tt = TTFont(path)
        self.order = self.tt.getGlyphOrder()
        self.glyphset = self.tt.getGlyphSet()
        self.cmap = self.tt.getBestCmap()
        hh = self.tt["hhea"]
        self.asc = hh.ascent / self.upm
        self.desc = -hh.descent / self.upm
        self.gap = hh.lineGap / self.upm
        self._paths = {}

    def covers(self, text):
        return all(ord(c) in self.cmap or c in "\n\r\t​‌‍" for c in text)

    def shape(self, text):
        buf = hb.Buffer()
        buf.add_str(text)
        buf.guess_segment_properties()
        hb.shape(self.hb_font, buf, {"kern": True, "liga": True})
        return [(i.codepoint, p.x_advance, p.y_advance, p.x_offset, p.y_offset)
                for i, p in zip(buf.glyph_infos, buf.glyph_positions)]

    def advance(self, text):
        """ความกว้างหน่วย em"""
        return sum(g[1] for g in self.shape(text)) / self.upm

    def outline(self, gid):
        if gid not in self._paths:
            pen = _Pen(self.glyphset)
            try:
                self.glyphset[self.order[gid]].draw(pen)
            except Exception:  # noqa: BLE001
                pen.ops = []
            self._paths[gid] = pen.ops
        return self._paths[gid]


@lru_cache(maxsize=16)
def get_face(family="sarabun", bold=False, italic=False):
    """คืน (Face, ต้องทำตัวเอียงเอง, ต้องทำตัวหนาเอง)"""
    for key, synth_i, synth_b in (((family, bold, italic), False, False),
                                  ((family, bold, False), italic, False),
                                  ((family, False, False), italic, bold),
                                  (("sarabun", bold, italic), False, False)):
        if key in FACE_FILES:
            return Face(os.path.join(FONT_DIR, FACE_FILES[key])), synth_i, synth_b
    raise KeyError(family)


def pick_family(font_name="", text="", width=None, size=None, bold=False):
    """เลือกฟอนต์ที่ใกล้ต้นฉบับที่สุด: ดูจากชื่อฟอนต์ก่อน ถ้าไม่รู้จักให้เทียบความกว้างจริง"""
    name = re.sub(r"^[A-Z]{6}\+", "", font_name or "")
    if SARABUN_LIKE.search(name):
        return "sarabun"
    if SANS_LIKE.search(name):
        return "noto"
    if text and width and size:
        best, best_err = "sarabun", 9e9
        for fam in ("sarabun", "noto"):
            w = get_face(fam, bold)[0].advance(text) * size
            if w > 0:
                err = abs(1 - width / w)
                if err < best_err:
                    best, best_err = fam, err
        return best
    return "sarabun"


def line_metrics(family="sarabun"):
    f = get_face(family)[0]
    return {"asc": f.asc, "desc": f.desc, "gap": f.gap}


# ---------------------------------------------------------------------------
# จัดบรรทัด
# ---------------------------------------------------------------------------

def _tokens(text):
    """แบ่งคำสำหรับตัดบรรทัด (ภาษาไทยไม่มีช่องว่างระหว่างคำ)"""
    try:
        from pythainlp.tokenize import word_tokenize
        return [t for t in word_tokenize(text, engine="newmm", keep_whitespace=True) if t]
    except Exception:  # noqa: BLE001
        return re.findall(r"\s+|\S+", text)


def _norm_runs(runs, default):
    out = []
    for r in runs:
        if not r.get("text"):
            continue
        rr = dict(default)
        rr.update({k: v for k, v in r.items() if v is not None})
        out.append(rr)
    return out


def layout(runs, width=None, family="sarabun", size=14, color="#000000", line_height=1.0):
    """runs = [{text, bold, italic, underline, color, size}] -> รายการบรรทัด
    แต่ละบรรทัด = {"items": [...], "width": pt, "text": str}"""
    default = {"bold": False, "italic": False, "underline": False, "color": color, "size": size}
    runs = _norm_runs(runs, default)
    pieces = []                                  # (ข้อความ, style) ทีละคำ / ขึ้นบรรทัดใหม่
    for r in runs:
        parts = r["text"].replace("\r", "").split("\n")
        for i, part in enumerate(parts):
            if i:
                pieces.append(("\n", r))
            for tok in _tokens(part):
                pieces.append((tok, r))

    lines, cur, cur_w = [], [], 0.0

    def item(tok, r):
        face, si, sb = get_face(family, bool(r.get("bold")), bool(r.get("italic")))
        if not face.covers(tok.strip()):
            face, si, sb = get_face("sarabun", bool(r.get("bold")), bool(r.get("italic")))
        sz = float(r.get("size") or size)
        glyphs = face.shape(tok)
        w = sum(g[1] for g in glyphs) / face.upm * sz
        return {"text": tok, "face": face, "size": sz, "color": r.get("color") or color,
                "underline": bool(r.get("underline")), "synth_italic": si, "synth_bold": sb,
                "glyphs": glyphs, "width": w}

    def flush():
        nonlocal cur, cur_w
        while cur and not cur[-1]["text"].strip():      # ตัดช่องว่างท้ายบรรทัด
            cur_w -= cur[-1]["width"]
            cur.pop()
        lines.append({"items": cur, "width": cur_w, "text": "".join(i["text"] for i in cur)})
        cur, cur_w = [], 0.0

    for tok, r in pieces:
        if tok == "\n":
            flush()
            continue
        it = item(tok, r)
        if width and cur and cur_w + it["width"] > width + 0.01 and tok.strip():
            flush()
        if not cur and not tok.strip() and lines:       # ไม่ขึ้นบรรทัดใหม่ด้วยช่องว่าง
            continue
        if width and not cur and it["width"] > width and len(tok) > 1:
            # คำยาวกว่าบรรทัด: ตัดทีละตัวอักษร
            for ch in tok:
                ci = item(ch, r)
                if cur and cur_w + ci["width"] > width:
                    flush()
                cur.append(ci)
                cur_w += ci["width"]
            continue
        cur.append(it)
        cur_w += it["width"]
    if cur or not lines:
        flush()
    return lines


# ---------------------------------------------------------------------------
# วาดลงหน้า PDF
# ---------------------------------------------------------------------------

def _rgb(color):
    if isinstance(color, (tuple, list)):
        return tuple(color)
    h = (color or "#000000").lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


def _n(v):
    return f"{v:.4f}".rstrip("0").rstrip(".") if abs(v) < 1e6 else "0"


# ---------- ทรัพยากรของหน้า (XObject / ExtGState) ----------

def _page_res(doc, page):
    t, v = doc.xref_get_key(page.xref, "Resources")
    if t == "xref":
        return int(v.split()[0]), ""
    if t != "dict":
        doc.xref_set_key(page.xref, "Resources", "<<>>")
    return page.xref, "Resources/"


def _add_res(doc, page, cat, name, ref):
    rx, pre = _page_res(doc, page)
    t, v = doc.xref_get_key(rx, pre + cat)
    if t == "xref":
        doc.xref_set_key(int(v.split()[0]), name, ref)
    elif t == "dict":
        doc.xref_set_key(rx, f"{pre}{cat}/{name}", ref)
    else:
        doc.xref_set_key(rx, pre + cat, f"<</{name} {ref}>>")


def _glyph_xobject(doc, face, gid, stroke):
    """ตัวอักษรหนึ่งตัวเป็น Form XObject (สร้างครั้งเดียวต่อเอกสาร ใช้ซ้ำได้ทุกหน้า) -> ลดขนาดไฟล์มาก"""
    cache = getattr(doc, "_th_glyph_cache", None)
    if cache is None:
        cache = {}
        try:
            doc._th_glyph_cache = cache
        except Exception:  # noqa: BLE001
            pass
    key = (face.path, gid, stroke)
    if key in cache:
        return cache[key]
    ops, xs, ys = [], [], []
    for op in face.outline(gid):
        if op[0] == "Z":
            ops.append("h")
            continue
        for x, y in op[1:]:
            xs.append(x)
            ys.append(y)
        pts = " ".join(f"{_n(x)} {_n(y)}" for x, y in op[1:])
        ops.append(f"{pts} {'m' if op[0] == 'M' else 'l' if op[0] == 'L' else 'c'}")
    if not ops:
        cache[key] = None
        return None
    pad = face.upm * 0.05
    bbox = f"[{_n(min(xs) - pad)} {_n(min(ys) - pad)} {_n(max(xs) + pad)} {_n(max(ys) + pad)}]"
    body = "\n".join(ops) + ("\nB\n" if stroke else "\nf\n")
    xref = doc.get_new_xref()
    doc.update_object(xref, f"<</Type/XObject/Subtype/Form/BBox{bbox}/Resources<<>>>>")
    doc.update_stream(xref, body.encode())
    cache[key] = xref
    return xref


def draw_lines(page, lines, x, first_baseline, line_step, align="left", box_width=None,
               opacity=1.0, rotate=0.0, pivot=None, invisible=True):
    """วาดบรรทัดที่จัดแล้วลงหน้า คืนกรอบที่วาด (Rect)"""
    doc = page.parent
    to_pdf = ~page.transformation_matrix                    # พิกัดหน้า (บนลงล่าง) -> พิกัด PDF
    post = fitz.Matrix(1, 1)
    if rotate:
        pivot = fitz.Point(pivot)
        post = fitz.Matrix(1, 0, 0, 1, -pivot.x, -pivot.y) * fitz.Matrix(rotate) * \
            fitz.Matrix(1, 0, 0, 1, pivot.x, pivot.y)

    def tp(px, py):
        return fitz.Point(px, py) * post

    bounds = fitz.Rect()
    out = ["q"]
    if opacity < 1:
        gs = page._set_opacity(CA=opacity, ca=opacity)
        if gs:
            out.append(f"/{gs} gs")
    names = {}
    inv = []
    cur_color = None
    baseline = first_baseline
    for ln in lines:
        lx = x
        if box_width and align in ("center", "right"):
            free = box_width - ln["width"]
            lx = x + (free / 2 if align == "center" else free)
        pen = lx
        for it in ln["items"]:
            face, size = it["face"], it["size"]
            s = size / face.upm
            sh = 0.21 if it["synth_italic"] else 0.0
            rgb = _rgb(it["color"])
            if rgb != cur_color:
                out.append(f"{_n(rgb[0])} {_n(rgb[1])} {_n(rgb[2])} rg {_n(rgb[0])} {_n(rgb[1])} {_n(rgb[2])} RG")
                cur_color = rgb
            gx = pen
            for gid, xa, ya, xo, yo in it["glyphs"]:
                xref = _glyph_xobject(doc, face, gid, it["synth_bold"])
                if xref:
                    name = names.get(xref)
                    if name is None:
                        name = names[xref] = f"TG{xref}"
                        _add_res(doc, page, "XObject", name, f"{xref} 0 R")
                    ox, oy = gx + xo * s, baseline - yo * s
                    m = fitz.Matrix(s, 0, sh * s, -s, ox, oy) * post * to_pdf
                    w = f"{_n(max(0.3, size * 0.03) / s)} w " if it["synth_bold"] else ""
                    out.append(f"q {_n(m.a)} {_n(m.b)} {_n(m.c)} {_n(m.d)} {_n(m.e)} {_n(m.f)} cm {w}/{name} Do Q")
                    bounds |= tp(ox, oy - size * face.asc)
                    bounds |= tp(ox + xa * s, oy + size * face.desc)
                gx += xa * s
            if it["underline"] and it["text"].strip():
                uy, th = baseline + size * 0.12, max(0.4, size * 0.05)
                pts = [tp(pen, uy) * to_pdf, tp(gx, uy) * to_pdf, tp(gx, uy + th) * to_pdf, tp(pen, uy + th) * to_pdf]
                out.append(" ".join(f"{_n(p.x)} {_n(p.y)} {'m' if i == 0 else 'l'}" for i, p in enumerate(pts)) + " h f")
            pen = gx
        if ln["text"].strip():
            inv.append((ln["text"], lx, baseline, pen - lx, ln["items"][0]["size"] if ln["items"] else 12))
        baseline += line_step
    out.append("Q")

    if len(out) > 2:
        if not page.is_wrapped:
            page.wrap_contents()
        nx = doc.get_new_xref()
        doc.update_object(nx, "<<>>")
        doc.update_stream(nx, ("\n".join(out) + "\n").encode())
        conts = page.get_contents()
        doc.xref_set_key(page.xref, "Contents", "[" + " ".join(f"{c} 0 R" for c in conts + [nx]) + "]")

    if invisible and inv:
        iface = get_face(*INVISIBLE_FACE)[0]
        fname = "th" + uuid.uuid4().hex[:8]       # ชื่อไม่ซ้ำ กันชนกับฟอนต์ใน XObject
        page.insert_font(fontname=fname, fontfile=iface.path)
        mfont = _measure(iface.path)
        for text, lx, base, w, size in inv:
            tw = mfont.text_length(text, fontsize=size) or 1
            m = fitz.Matrix(max(0.05, w / tw), 0, 0, 1, 0, 0)
            start = fitz.Point(lx, base)
            if rotate:
                start = tp(lx, base)
                m = m * fitz.Matrix(rotate)
            try:
                page.insert_text(start, text, fontname=fname, fontsize=size, render_mode=3, morph=(start, m))
            except Exception:  # noqa: BLE001
                pass
    return bounds


@lru_cache(maxsize=8)
def _measure(path):
    return fitz.Font(fontfile=path)


# ---------------------------------------------------------------------------
# ฟังก์ชันที่ระบบเรียกใช้
# ---------------------------------------------------------------------------

def add_runs_at_baseline(page, runs, x, baseline, size, family="sarabun", color="#000000", max_x=None,
                         line_height=1.0):
    """เขียนข้อความ (แก้ข้อความเดิม) เริ่มที่ x บนเส้นฐาน baseline เดิมพอดี"""
    lines = layout(runs, None, family, size, color)
    m = line_metrics(family)
    widest = max((ln["width"] for ln in lines), default=0)
    if max_x and widest > 0 and x + widest > max_x:          # ยาวเกินขอบหน้า -> ย่อเล็กน้อย
        k = max(0.7, (max_x - x) / widest)
        runs = [dict(r, size=float(r.get("size") or size) * k) for r in runs]
        size *= k
        lines = layout(runs, None, family, size, color)
    step = size * (m["asc"] + m["desc"] + m["gap"]) * line_height
    return draw_lines(page, lines, x, baseline, step)


def add_text_box(page, rect, runs, size=14, family="sarabun", color="#000000", align="left",
                 line_height=1.0, opacity=1.0, rotate=0.0, pivot=None):
    """เขียนข้อความในกรอบ (ตัดบรรทัดอัตโนมัติ) ตำแหน่งบรรทัดตรงกับที่หน้าเว็บแสดง"""
    rect = fitz.Rect(rect)
    if isinstance(runs, str):
        runs = [{"text": runs}]
    lines = layout(runs, rect.width, family, size, color)
    m = line_metrics(family)
    lh = (m["asc"] + m["desc"] + m["gap"]) * line_height
    first = rect.y0 + size * ((lh - m["asc"] - m["desc"]) / 2 + m["asc"])
    return draw_lines(page, lines, rect.x0, first, size * lh, align, rect.width, opacity, rotate, pivot)


def add_thai_text(page, rect, text, size=12, color="#000000", bold=False, italic=False, underline=False,
                  align="left", line_height=1.0, family="sarabun", opacity=1.0, **_):
    """ใช้แทนของเดิม (ข้อความธรรมดาในกรอบ)"""
    return add_text_box(page, rect, [{"text": text, "bold": bold, "italic": italic, "underline": underline}],
                        size, family, color, align, line_height, opacity)


def text_width(text, size, bold=False, family="sarabun"):
    return get_face(family, bold)[0].advance(text) * size
