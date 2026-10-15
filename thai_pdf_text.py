"""เขียนข้อความภาษาไทยลงใน PDF ให้ถูกต้องทั้ง "ตาเห็น" และ "ข้อความที่ค้นหา/คัดลอกได้"

ปัญหา: เวลาเขียนภาษาไทยลง PDF ส่วนใหญ่ (รวมถึง Word ที่ใช้ Cordia New) ฟอนต์จะเปลี่ยนรูปสระ/วรรณยุกต์
เพื่อจัดตำแหน่ง ทำให้ข้อความที่ซ่อนอยู่ใน PDF เพี้ยน (ค้นหาไม่เจอ, copy แล้วได้ "กำร" แทน "การ")

วิธีแก้ (แบบเดียวกับ PDF ที่ผ่าน OCR):
  1. วาดข้อความด้วย HTML/CSS (HarfBuzz จัดรูปสระ/วรรณยุกต์ไทยถูกต้อง) แล้วแปลงเป็นเส้นเวกเตอร์ -> ตาเห็นถูก 100%
  2. วางข้อความที่มองไม่เห็นทับในตำแหน่งเดียวกันด้วยตัวอักษร Unicode ที่ถูกต้อง -> ค้นหา/คัดลอกได้ถูก 100%
"""
import html
import os
import uuid
from functools import lru_cache

import pymupdf as fitz
from fontTools.ttLib import TTFont

from converter import fix_thai

FONT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static", "fonts")
FONT_REGULAR = os.path.join(FONT_DIR, "NotoSansThai-Regular.ttf")
FONT_BOLD = os.path.join(FONT_DIR, "NotoSansThai-Bold.ttf")
_ARCHIVE = None


def _archive():
    global _ARCHIVE
    if _ARCHIVE is None:
        _ARCHIVE = fitz.Archive(FONT_DIR)
    return _ARCHIVE


@lru_cache(maxsize=4)
def _measure_font(path=FONT_REGULAR):
    return fitz.Font(fontfile=path)


def text_width(text, size, bold=False):
    return _measure_font(FONT_BOLD if bold else FONT_REGULAR).text_length(text, fontsize=size)


@lru_cache(maxsize=4)
def glyph_unicode_map(font_path):
    """GID -> ตัวอักษร Unicode รวมถึงรูปสระ/วรรณยุกต์ที่ฟอนต์เปลี่ยนตำแหน่ง (GSUB)"""
    f = TTFont(font_path)
    order = f.getGlyphOrder()
    gid = {g: i for i, g in enumerate(order)}
    m = {}
    for cp, g in f.getBestCmap().items():
        m.setdefault(g, chr(cp))
    if "GSUB" in f:
        lookups = f["GSUB"].table.LookupList.Lookup
        changed = True
        while changed:
            changed = False
            for lk in lookups:
                for st in lk.SubTable:
                    if lk.LookupType == 7:
                        st = st.ExtSubTable
                    t = getattr(st, "LookupType", lk.LookupType)
                    if t == 1:
                        for a, b in st.mapping.items():
                            if a in m and b not in m:
                                m[b] = m[a]
                                changed = True
                    elif t == 2:
                        for a, seq in st.mapping.items():
                            if a in m:
                                for k, b in enumerate(seq):
                                    if b not in m:
                                        m[b] = m[a] if k == 0 else ""
                                        changed = True
                    elif t == 4:
                        for first, ligs in st.ligatures.items():
                            for lig in ligs:
                                comps = [first] + lig.Component
                                if all(c in m for c in comps) and lig.LigGlyph not in m:
                                    m[lig.LigGlyph] = "".join(m[c] for c in comps)
                                    changed = True
    return {gid[g]: u for g, u in m.items() if g in gid}


def _fix_tounicode(doc):
    """เขียน ToUnicode ใหม่ให้ฟอนต์ Noto Sans Thai ในเอกสาร (ใช้กับหน้าชั่วคราวเพื่ออ่านตำแหน่งบรรทัด)"""
    fonts = {"NotoSansThai-Bold": FONT_BOLD, "NotoSansThaiBold": FONT_BOLD,
             "NotoSansThai": FONT_REGULAR}
    for xref in range(1, doc.xref_length()):
        try:
            if doc.xref_get_key(xref, "Subtype")[1] != "/Type0":
                continue
        except Exception:  # noqa: BLE001
            continue
        base = doc.xref_get_key(xref, "BaseFont")[1].replace(" ", "")
        path = next((p for k, p in fonts.items() if k in base), None)
        if not path:
            continue
        entries = []
        for g, u in sorted(glyph_unicode_map(path).items()):
            hexu = "".join(f"{ord(c):04X}" for c in (u or "​") if ord(c) < 0x10000)
            entries.append(f"<{g:04X}> <{hexu}>")
        body = "\n".join(f"{len(entries[i:i + 100])} beginbfchar\n" + "\n".join(entries[i:i + 100]) + "\nendbfchar"
                         for i in range(0, len(entries), 100))
        cmap = ("/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n"
                "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n"
                "/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n1 begincodespacerange\n<0000> <FFFF>\n"
                "endcodespacerange\n" + body + "\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n")
        tu = doc.xref_get_key(xref, "ToUnicode")
        if tu[0] == "xref":
            doc.update_stream(int(tu[1].split()[0]), cmap.encode())
        else:
            nx = doc.get_new_xref()
            doc.update_object(nx, "<<>>")
            doc.update_stream(nx, cmap.encode())
            doc.xref_set_key(xref, "ToUnicode", f"{nx} 0 R")


def _css(size, color, align, line_height):
    return (f"@font-face{{font-family:th;src:url({os.path.basename(FONT_REGULAR)});}}"
            f"@font-face{{font-family:th;src:url({os.path.basename(FONT_BOLD)});font-weight:bold;}}"
            f"*{{font-family:th;font-size:{size:.2f}pt;color:{color};margin:0;padding:0;"
            f"line-height:{line_height};text-align:{align};}}")


def to_html(text, bold=False, italic=False, underline=False):
    body = "<br>".join(html.escape(line) for line in text.split("\n"))
    if bold:
        body = f"<b>{body}</b>"
    if italic:
        body = f"<i>{body}</i>"
    if underline:
        body = f"<u>{body}</u>"
    return body


def add_thai_text(page, rect, text, size=12, color="#000000", bold=False, italic=False,
                  underline=False, align="left", line_height=1.25, shrink=True, opacity=1.0):
    """เขียนข้อความไทยในกรอบ rect ของ page (หน่วย point, หน้าต้องไม่หมุน)"""
    rect = fitz.Rect(rect)
    if not text.strip():
        return
    tmp = fitz.open()
    tp = tmp.new_page(width=page.rect.width, height=page.rect.height)
    tp.insert_htmlbox(rect, to_html(text, bold, italic, underline), css=_css(size, color, align, line_height),
                      archive=_archive(), scale_low=0.2 if shrink else 1, opacity=opacity)
    _fix_tounicode(tmp)
    tmp = fitz.open("pdf", tmp.tobytes())
    tp = tmp[0]

    # อ่านตำแหน่งแต่ละบรรทัด (สำหรับชั้นข้อความที่มองไม่เห็น)
    lines = []
    for b in tp.get_text("dict")["blocks"]:
        for ln in b.get("lines", []):
            t = fix_thai("".join(s["text"] for s in ln["spans"]).replace("​", ""))
            if t.strip():
                sp = ln["spans"][0]
                lines.append((t, fitz.Rect(ln["bbox"]), sp["size"], sp["origin"][1]))

    # 1) ตัวที่ตาเห็น: แปลงเป็นเส้นเวกเตอร์
    svg = tp.get_svg_image(text_as_path=True)
    src = fitz.open("pdf", fitz.open(stream=svg.encode(), filetype="svg").convert_to_pdf())
    page.show_pdf_page(page.rect, src, 0, overlay=True)

    # 2) ชั้นข้อความที่มองไม่เห็น: Unicode ถูกต้อง ค้นหา/คัดลอกได้
    # ชื่อฟอนต์ไม่ซ้ำ: ถ้าซ้ำกับฟอนต์ใน XObject (เช่นลายน้ำ) PyMuPDF จะไม่ผูกฟอนต์กับหน้า ทำให้ข้อความเพี้ยน
    fname = "th" + uuid.uuid4().hex[:8]
    page.insert_font(fontname=fname, fontfile=FONT_REGULAR)
    for t, bb, sz, base in lines:
        w = text_width(t, sz) or 1
        page.insert_text((bb.x0, base), t, fontname=fname, fontsize=sz, render_mode=3,
                         morph=(fitz.Point(bb.x0, base), fitz.Matrix(bb.width / w, 1)))
