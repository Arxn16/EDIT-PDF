"""แปลง PDF ภาษาไทยเป็น Word โดยไม่ให้สระ/วรรณยุกต์เพี้ยน

โหมดที่รองรับ
  auto  - ตรวจไฟล์ก่อน แล้วเลือก text หรือ ocr ให้เอง
  text  - ดึงข้อความจาก PDF (pdf2docx) เก็บรูปแบบหน้า + ซ่อมรหัสภาษาไทย
  ocr   - อ่านจากภาพด้วย Tesseract (ใช้กับไฟล์สแกน / ข้อความในไฟล์เสีย)
  image - ใส่แต่ละหน้าเป็นรูปภาพ เหมือนต้นฉบับ 100% แต่แก้ข้อความไม่ได้
"""
import copy
import io
import re
import statistics
import unicodedata

import pymupdf as fitz
import pytesseract
from PIL import Image
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu, Pt
from pdf2docx import Converter

from docx_fix import embed_fonts, fix_table_outdent, normalize_ooxml, sanitize_ooxml

DEFAULT_FONT = "Sarabun TH Compat"     # ขนาดเท่า TH Sarabun New และฝังในไฟล์ได้ (แสดงถูกทุกเครื่อง)

# ---------------------------------------------------------------------------
# ซ่อมข้อความภาษาไทย
# ---------------------------------------------------------------------------

# รหัส Private Use Area ที่ฟอนต์ไทยหลายตัวใช้แทนสระ/วรรณยุกต์ที่ "เลื่อนตำแหน่ง"
PUA_MAP = {
    0xF700: "ฐ", 0xF70F: "ญ",
    0xF701: "ิ", 0xF702: "ี", 0xF703: "ึ", 0xF704: "ื",
    0xF705: "่", 0xF706: "้", 0xF707: "๊", 0xF708: "๋", 0xF709: "์",
    0xF70A: "่", 0xF70B: "้", 0xF70C: "๊", 0xF70D: "๋", 0xF70E: "์",
    0xF710: "ั", 0xF711: "ํ", 0xF712: "็",
    0xF713: "่", 0xF714: "้", 0xF715: "๊", 0xF716: "๋", 0xF717: "์",
    0xF718: "ุ", 0xF719: "ู", 0xF71A: "ฺ",
}

TONES = "่้๊๋์"
VOWELS_AB = "ัิีึืฺุู"            # สระบน/ล่าง
COMBINING = set(VOWELS_AB + TONES + "็ํ๎")
THAI_RE = re.compile(r"[฀-๿]")
DUP_MARKS = re.compile("([\u0E31\u0E34-\u0E3A\u0E47-\u0E4E]+)\\1+")
# สระอำที่นิคหิตหายกลายเป็นช่องว่าง เช่น "ส าหรับ" -> "สำหรับ", "ต่ ากว่า" -> "ต่ำกว่า"
# (คำไทยขึ้นต้นด้วย "า" ไม่ได้ ช่องว่าง+า หลังพยัญชนะจึงเป็นสระอำที่เสียแน่นอน)
LOST_AM = re.compile("([\u0E01-\u0E2E][\u0E48-\u0E4B]?)[ \t\u00A0\u200B\uFFFD]+\u0E32")
_SP = "[ \\t\\u00A0\\u200B]+"
# สระ/วรรณยุกต์ที่ขึ้นต้นคำไม่ได้ (ยกเว้น า ซึ่งแยกจัดการ เพราะอาจเป็นสระอำที่นิคหิตหาย)
SPACED_MARK = re.compile(f"([\\u0E01-\\u0E4E]){_SP}([\\u0E30\\u0E31\\u0E33-\\u0E3A\\u0E45\\u0E47-\\u0E4E])")
SPACED_LEAD = re.compile(f"([เแโใไ]){_SP}(?=[\\u0E01-\\u0E2E])")   # สระหน้าลงท้ายคำไม่ได้
SPACED_AFTER_HAN = re.compile(f"(\\u0E31[\\u0E48-\\u0E4B]?){_SP}(?=[\\u0E01-\\u0E2E])")  # ั ต้องมีตัวสะกดเสมอ
SPACED_NIKHAHIT = re.compile(f"ํ{_SP}?า")
SPACED_AA = re.compile(f"([\\u0E01-\\u0E2E][\\u0E31-\\u0E3A\\u0E47-\\u0E4E]*){_SP}า")
TONE_BEFORE_VOWEL = re.compile(f"([{TONES}])([{VOWELS_AB}])")


def fix_mojibake(text: str) -> str:
    """แก้ข้อความไทยรหัสเก่า (TIS-620) ที่ถูกอ่านเป็นภาษาต่างดาว เช่น '¡ÒÃ»ÃÐªØÁ' -> 'การประชุม'"""
    if not text or THAI_RE.search(text):
        return text
    hi = sum(1 for c in text if 0xA1 <= ord(c) <= 0xFB)
    letters = len(text.replace(" ", ""))
    if hi < 2 or hi < letters * 0.3:
        return text
    try:
        return text.encode("latin-1").decode("cp874")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return text


def fix_spacing(text: str, lost_am: bool = True) -> str:
    """ลบช่องว่างที่แทรกผิดตำแหน่งระหว่างพยัญชนะกับสระ/วรรณยุกต์ (ครอบคลุมทุกตัว)"""
    prev = None
    while prev != text:                                   # วนจนไม่มีอะไรให้แก้ (กรณีแตกหลายชิ้น)
        prev = text
        text = SPACED_MARK.sub(r"\1\2", text)            # ท ี่ / ที ่ / ก ิน / ป ็ / ส ำ -> ติดกัน
        text = SPACED_LEAD.sub(r"\1", text)
        text = SPACED_AFTER_HAN.sub(r"\1", text)         # ค ั บ / คั บ -> คับ              # เ ป็น / ไ ป -> เป็น / ไป
        text = SPACED_NIKHAHIT.sub("ำ", text)             # น ํ า -> นำ
    if lost_am:
        text = LOST_AM.sub(r"\1ำ", text)                 # ส าหรับ -> สำหรับ (นิคหิตหาย)
    else:
        text = SPACED_AA.sub(r"\1า", text)               # ส าย -> สาย (แค่ช่องว่างเกิน)
    return text


def fix_thai(text: str, lost_am: bool = True) -> str:
    if not text:
        return text
    text = fix_mojibake(text)
    text = text.translate(PUA_MAP)
    text = fix_spacing(text, lost_am)
    text = TONE_BEFORE_VOWEL.sub(r"\2\1", text)        # วรรณยุกต์ต้องตามหลังสระบน/ล่าง
    text = re.sub(f"ํ([{TONES}]?)า", r"\1ำ", text)     # นิคหิต + (วรรณยุกต์) + า -> ำ
    text = DUP_MARKS.sub(r"\1", text)                    # สระ/วรรณยุกต์ซ้ำ เช่น ทั้ั้ง -> ทั้ง
    return text


def detect_lost_am(raw: str) -> bool:
    """ตัดสินว่า 'ส าหรับ' ในไฟล์นี้คือสระอำที่นิคหิตหาย (True) หรือแค่ช่องว่างเกินหน้า า (False)
    ถ้าในไฟล์มีสระอำปกติมากกว่าจุดที่เสีย แปลว่าสระอำไม่ได้พัง ช่องว่างหน้า า จึงเป็นแค่ช่องว่างเกิน"""
    return raw.count("ำ") < len(LOST_AM.findall(raw)) or "ำ" not in raw


def thai_problems(text: str):
    """นับตัวอักษรไทยทั้งหมด และตัวที่ผิดรูป (สระลอย, รหัสเสีย)"""
    total = bad = 0
    prev = ""
    for c in text:
        o = ord(c)
        if 0x0E00 <= o <= 0x0E7F:
            total += 1
        if c in COMBINING:
            ok_prev = prev and (0x0E01 <= ord(prev) <= 0x0E2E or prev in COMBINING)
            if not ok_prev:
                bad += 1
        if c == "�" or 0xF700 <= o <= 0xF71A:
            bad += 1
            total += 1
        prev = c
    return total, bad


# ---------------------------------------------------------------------------
# ตรวจไฟล์
# ---------------------------------------------------------------------------

def analyze(pdf_path: str) -> dict:
    with fitz.open(pdf_path) as pdf:
        pages = len(pdf)
        raw = "".join(p.get_text() for p in pdf)
    fixed = fix_thai(raw, detect_lost_am(raw))
    chars = len(re.sub(r"\s", "", fixed))
    total, bad = thai_problems(fixed)
    pua = sum(1 for c in raw if 0xF700 <= ord(c) <= 0xF71A)
    mojibake = raw != fix_mojibake(raw)
    am_lost_mode = detect_lost_am(raw)
    lost_am = len(LOST_AM.findall(raw)) if am_lost_mode else 0
    spaced = len(SPACED_MARK.findall(raw)) + len(SPACED_LEAD.findall(raw))
    leftover_latin = sum(1 for c in fixed if 0xA1 <= ord(c) <= 0xFB)
    bad_ratio = bad / total if total else 0.0

    if chars < 30 * max(pages, 1):
        method, reason = "ocr", "แทบไม่มีข้อความในไฟล์ (น่าจะเป็นไฟล์สแกน) จึงใช้ OCR"
    elif bad_ratio > 0.01:
        method, reason = "ocr", f"ข้อความในไฟล์เสีย {bad_ratio:.1%} แม้ซ่อมแล้ว จึงใช้ OCR"
    elif leftover_latin > chars * 0.05:
        method, reason = "ocr", "ไฟล์ใช้ฟอนต์เข้ารหัสแบบพิเศษที่ถอดไม่ได้ จึงใช้ OCR"
    else:
        method = "text"
        notes = []
        if pua:
            notes.append(f"ซ่อมสระ/วรรณยุกต์รหัส PUA {pua} ตัว")
        if spaced:
            notes.append(f"ลบช่องว่างที่แทรกในสระ/วรรณยุกต์ {spaced} จุด")
        if lost_am:
            notes.append(f"ซ่อมสระอำที่หาย {lost_am} จุด")
        if mojibake:
            notes.append("แปลงรหัสไทยแบบเก่า (TIS-620)")
        reason = "ข้อความในไฟล์ใช้ได้ ดึงตรงจาก PDF" + (f" ({', '.join(notes)})" if notes else "")

    return {
        "pages": pages, "chars": chars, "thai_chars": total,
        "bad_ratio": round(bad_ratio, 4), "pua": pua, "mojibake": mojibake,
        "method": method, "reason": reason, "lost_am": am_lost_mode,
    }


# ---------------------------------------------------------------------------
# จัดฟอนต์ใน Word (ภาษาไทยใช้ฟอนต์/ขนาดฝั่ง "complex script")
# ---------------------------------------------------------------------------

def _child(parent, tag):
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        parent.append(el)
    return el


def _hf_roots(doc):
    """หัว/ท้ายกระดาษที่มีอยู่จริงในไฟล์ (ห้ามใช้ section.header._element ตรงๆ: python-docx จะสร้าง
    หัวกระดาษว่างใหม่ให้ ทำให้เนื้อหาถูกดันลงและล้นไปหน้าใหม่)"""
    out, seen = [], set()
    for s in doc.sections:
        for hf in (s.header, s.footer, s.first_page_header, s.first_page_footer,
                   s.even_page_header, s.even_page_footer):
            try:
                if hf.is_linked_to_previous:
                    continue
                part = hf._get_or_add_definition()
            except Exception:  # noqa: BLE001
                continue
            if id(part) not in seen:
                seen.add(id(part))
                out.append(part.element)
    return out


def finalize_docx(doc):
    """ขั้นสุดท้ายก่อนบันทึก: ตารางอยู่ตำแหน่งเดียวกับ PDF + ลำดับ element ตามมาตรฐานที่ Word ต้องการ"""
    styles = doc.styles.element
    for root in [doc.element.body] + _hf_roots(doc):
        fix_table_outdent(root, styles)
        sanitize_ooxml(root)
        normalize_ooxml(root)
    sanitize_ooxml(styles)
    normalize_ooxml(styles)
    try:                                      # แม่แบบของ python-docx ไม่มี w:percent (บังคับตามมาตรฐาน)
        zoom = doc.settings.element.find(qn("w:zoom"))
        if zoom is not None and zoom.get(qn("w:percent")) is None:
            zoom.set(qn("w:percent"), "100")
    except Exception:  # noqa: BLE001
        pass


def apply_thai_font(doc: Document, font: str):
    """font = ชื่อฟอนต์ หรือ 'original' เพื่อใช้ฟอนต์เดิมจาก PDF"""
    roots = [doc.element.body] + _hf_roots(doc)

    def fix_rpr(rpr):
        fonts = _child(rpr, "w:rFonts")
        if font == "original":
            base = fonts.get(qn("w:ascii")) or fonts.get(qn("w:hAnsi"))
            if base:
                fonts.set(qn("w:cs"), base)
        else:
            for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
                fonts.set(qn(a), font)
        sz = rpr.find(qn("w:sz"))
        if sz is not None and rpr.find(qn("w:szCs")) is None:
            szcs = OxmlElement("w:szCs")
            szcs.set(qn("w:val"), sz.get(qn("w:val")))
            sz.addnext(szcs)
        # ตัวหนา/เอียงของภาษาไทยใช้ bCs/iCs แยกจากภาษาอังกฤษ (pdf2docx ใส่แค่ b/i ทำให้ไทยไม่หนา)
        for tag, cs in (("w:b", "w:bCs"), ("w:i", "w:iCs")):
            el = rpr.find(qn(tag))
            if el is not None and rpr.find(qn(cs)) is None:
                new = OxmlElement(cs)
                if el.get(qn("w:val")) is not None:
                    new.set(qn("w:val"), el.get(qn("w:val")))
                el.addnext(new)
        _child(rpr, "w:lang").set(qn("w:bidi"), "th-TH")

    for root in roots:
        for r in root.iter(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                r.insert(0, rpr)
            fix_rpr(rpr)

    # ค่าเริ่มต้นของสไตล์ Normal
    normal = doc.styles["Normal"].element
    fix_rpr(_child(normal, "w:rPr"))


# ---------------------------------------------------------------------------
# ตัวแปลงแต่ละโหมด
# ---------------------------------------------------------------------------

_NON_INITIAL = "\u0E30\u0E31\u0E33\u0E34\u0E35\u0E36\u0E37\u0E38\u0E39\u0E3A\u0E45\u0E47\u0E48\u0E49\u0E4A\u0E4B\u0E4C\u0E4D\u0E4E"
_SPACES = " \u00A0\u200B\t"


def _fix_boundary(prev, cur, lost_am):
    """ซ่อมช่องว่างผิดตำแหน่งที่คร่อมระหว่าง 2 run"""
    p_core = prev.text.rstrip(_SPACES)
    c_core = cur.text.lstrip(_SPACES)
    gap = len(prev.text) - len(p_core) + len(cur.text) - len(c_core)
    if not gap or not p_core or not c_core:
        return
    last, first = p_core[-1], c_core[0]
    thai_last = "\u0E01" <= last <= "\u0E4E"
    if thai_last and first in _NON_INITIAL:
        new_prev, new_cur = p_core, c_core
    elif thai_last and first == "า" and "\u0E01" <= last <= "\u0E2E" or (last in TONES and first == "า"):
        new_prev, new_cur = p_core, ("ำ" + c_core[1:]) if lost_am else c_core
    elif (last in "เแโใไ" or p_core[-1:] == "\u0E31" or p_core[-2:-1] == "\u0E31" and last in TONES) and "\u0E01" <= first <= "\u0E2E":
        new_prev, new_cur = p_core, c_core
    else:
        return
    prev.text, cur.text = new_prev, new_cur
    for t in (prev, cur):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _restore_spaces(root, pairs, known):
    """pdf2docx บางครั้งแยกบรรทัดเป็นหลาย run แล้วทิ้งช่องว่างระหว่าง run ทำให้คำติดกัน
    คืนช่องว่างเมื่อใน PDF ต้นฉบับสองคำนั้นมีช่องว่างคั่นจริง"""
    for para in root.iter(qn("w:p")):
        # คู่ข้อความที่อยู่บรรทัดเดียวกันเท่านั้น (ข้าม <w:br/> หรือแท็บ = ขึ้นบรรทัดใหม่/เว้นระยะอยู่แล้ว
        # ถ้าใส่ช่องว่างตรงนั้น บรรทัดถัดไปจะเยื้องขวา 1 ช่องว่าง)
        ts, prev_t = [], None
        for el in para.iter(qn("w:t"), qn("w:br"), qn("w:cr"), qn("w:tab")):
            if el.tag == qn("w:t"):
                if el.text:
                    if prev_t is not None:
                        ts.append((prev_t, el))
                    prev_t = el
            else:
                prev_t = None
        for prev, cur in ts:
            if prev.text[-1].isspace() or cur.text[0].isspace():
                continue
            tail, head = prev.text.split()[-1], cur.text.split()[0]
            if (tail, head) in pairs and (tail + head) not in known:
                cur.text = " " + cur.text
                cur.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _fix_tables(root):
    """ซ่อมตารางจาก pdf2docx ให้เปิดได้ถูกในทุกโปรแกรม (Word, LibreOffice, Pages)
    - ความกว้างคอลัมน์ใน tblGrid ไม่ตรงกับความกว้างช่องจริง -> ตั้งให้ตรง
    - ความสูงแถวแบบ exact ทำให้ข้อความที่ล้นถูกตัดหาย -> เปลี่ยนเป็น atLeast (ขยายได้)"""
    for h in root.iter(qn("w:trHeight")):
        if h.get(qn("w:hRule")) == "exact":
            h.set(qn("w:hRule"), "atLeast")
    for tbl in root.iter(qn("w:tbl")):
        grid = tbl.find(qn("w:tblGrid"))
        if grid is None:
            continue
        cols = grid.findall(qn("w:gridCol"))
        # หาแถวที่ทุกช่องไม่ผสาน และจำนวนช่องเท่ากับจำนวนคอลัมน์
        for tr in tbl.findall(qn("w:tr")):
            tcs = tr.findall(qn("w:tc"))
            if len(tcs) != len(cols):
                continue
            widths = []
            for tc in tcs:
                pr = tc.find(qn("w:tcPr"))
                if pr is None or pr.find(qn("w:gridSpan")) is not None or pr.find(qn("w:tcW")) is None:
                    break
                widths.append(pr.find(qn("w:tcW")).get(qn("w:w")))
            if len(widths) == len(cols):
                for col, w in zip(cols, widths):
                    col.set(qn("w:w"), w)
                break


THAI_DIGITS = str.maketrans("0123456789", "๐๑๒๓๔๕๖๗๘๙")


def _keep_thai_digits(root, target_font):
    """ฟอนต์ TH SarabunIT๙ / TH Niramit AS ฯลฯ (ที่ลงท้าย IT๙) แสดงเลข 0-9 เป็นเลขไทย
    ถ้าเปลี่ยนไปใช้ฟอนต์อื่น ต้องแปลงตัวเลขเป็นเลขไทยจริง ไม่งั้นจะกลายเป็นเลขอารบิก"""
    if "IT" in target_font and "๙" in target_font:
        return
    for r in root.iter(qn("w:r")):
        fonts = r.find(qn("w:rPr") + "/" + qn("w:rFonts")) if r.find(qn("w:rPr")) is not None else None
        name = ""
        if fonts is not None:
            name = fonts.get(qn("w:ascii")) or fonts.get(qn("w:hAnsi")) or ""
        if re.search(r"IT(\?|๙|9)?$|IT[\?๙9]?,", name) or name.rstrip(",Bold").endswith("IT?"):
            for t in r.iter(qn("w:t")):
                if t.text:
                    t.text = t.text.translate(THAI_DIGITS)


# ---------------------------------------------------------------------------
# เส้นที่ pdf2docx ทำหาย (เส้นประเติมข้อความในแบบฟอร์ม, เส้นคั่น, กรอบที่ไม่ใช่ตาราง)
# ---------------------------------------------------------------------------

def _pdf_segments(page):
    """เส้นตรงแนวนอน/แนวตั้งทั้งหมดในหน้า (รวมเส้นที่วาดเป็นสี่เหลี่ยมบาง ๆ แบบที่ Word ส่งออก)"""
    segs = []
    for d in page.get_drawings():
        color, fill = d.get("color"), d.get("fill")
        width = d.get("width") or 0.5
        dashes = d.get("dashes") or ""
        dash = dashes not in ("", "[] 0", "[] 0.0")
        for it in d["items"]:
            if it[0] == "l" and color is not None:
                p1, p2 = it[1], it[2]
                if abs(p1.y - p2.y) < 0.8 and abs(p1.x - p2.x) > 2:
                    segs.append(["h", min(p1.x, p2.x), max(p1.x, p2.x), (p1.y + p2.y) / 2, width, color, dash, dashes])
                elif abs(p1.x - p2.x) < 0.8 and abs(p1.y - p2.y) > 2:
                    segs.append(["v", min(p1.y, p2.y), max(p1.y, p2.y), (p1.x + p2.x) / 2, width, color, dash, dashes])
            elif it[0] == "re":
                r = fitz.Rect(it[1])
                if fill is not None and "f" in (d.get("type") or ""):
                    if r.height <= 2.5 and r.width > 2.5:
                        segs.append(["h", r.x0, r.x1, (r.y0 + r.y1) / 2, max(r.height, 0.25), fill, False, ""])
                    elif r.width <= 2.5 and r.height > 2.5:
                        segs.append(["v", r.y0, r.y1, (r.x0 + r.x1) / 2, max(r.width, 0.25), fill, False, ""])
                if color is not None and "s" in (d.get("type") or "") and r.width > 2.5 and r.height > 2.5:
                    for sg in (["h", r.x0, r.x1, r.y0], ["h", r.x0, r.x1, r.y1], ["v", r.y0, r.y1, r.x0], ["v", r.y0, r.y1, r.x1]):
                        segs.append(sg + [width, color, dash, dashes])
    # รวมเส้นที่ต่อกัน (Word วาดเส้นตารางเป็นชิ้นเล็ก ๆ หลายชิ้น)
    segs.sort(key=lambda s: (s[0], round(s[3] * 2), s[1]))
    merged = []
    for sg in segs:
        m = merged[-1] if merged else None
        if m and m[0] == sg[0] and abs(m[3] - sg[3]) < 0.6 and sg[1] <= m[2] + 1.5 and m[6] == sg[6] \
                and tuple(m[5]) == tuple(sg[5]):
            m[2] = max(m[2], sg[2])
        else:
            merged.append(list(sg))
    return merged


def _table_segments(blocks, out):
    for b in blocks:
        for r in b.get("rows", []):
            for c in r.get("cells", []):
                if not c or not c.get("bbox"):                # ช่องที่ถูกผสาน (pdf2docx ไม่มีกรอบ)
                    continue
                x0, y0, x1, y1 = c["bbox"]
                bw = c.get("border_width") or (0, 0, 0, 0)
                for side, wd in zip(("top", "right", "bottom", "left"), bw):
                    if wd and wd > 0:
                        out.append({"top": ("h", x0, x1, y0), "bottom": ("h", x0, x1, y1),
                                    "left": ("v", y0, y1, x0), "right": ("v", y0, y1, x1)}[side])
                _table_segments(c.get("blocks", []), out)
        _table_segments(b.get("blocks", []), out)


def _covered(seg, tsegs):
    o, a, b, pos = seg[:4]
    need = (b - a) * 0.7
    got = 0.0
    for t in tsegs:
        if t[0] == o and abs(t[3] - pos) <= 2.5:
            got += max(0.0, min(b, t[2]) - max(a, t[1]))
    return got >= need


def _is_text_decoration(seg, lines):
    """เส้นใต้/ขีดฆ่าตัวอักษร (pdf2docx แปลงเป็นรูปแบบตัวอักษรให้แล้ว)"""
    if seg[0] != "h":
        return False
    a, b, y = seg[1], seg[2], seg[3]
    for (x0, y0, x1, y1) in lines:
        if y0 - 1 <= y <= y1 + 2.5:
            ov = max(0.0, min(b, x1) - max(a, x0))
            if ov >= (b - a) * 0.6:
                return True
    return False


def _orphan_lines(pdf_path, parsed):
    out = {}
    with fitz.open(pdf_path) as pdf:
        for i, page in enumerate(pdf):
            tsegs = []
            try:
                for sec in parsed["pages"][i]["sections"]:
                    for col in sec["columns"]:
                        _table_segments(col["blocks"], tsegs)
            except (KeyError, IndexError, TypeError):
                tsegs = None
            if tsegs is None:                                  # อ่านโครงสร้างตารางไม่ได้ -> ไม่เติมเส้น (กันเส้นซ้ำ)
                continue
            tl = [(tuple(l["bbox"]), "".join(sp["text"] for sp in l["spans"]))
                  for b in page.get_text("dict")["blocks"] for l in b.get("lines", [])
                  if any(sp["text"].strip() for sp in l["spans"])]
            text_lines = [t[0] for t in tl]
            orphans = [s for s in _pdf_segments(page)
                       if (s[2] - s[1]) > 4 and not _covered(s, tsegs) and not _is_text_decoration(s, text_lines)]
            if orphans:
                out[i] = {"segs": orphans, "lines": tl}
    return out


# ---------------------------------------------------------------------------
# เส้นคู่ (double border): pdf2docx รวมเส้นขนาน 2 เส้นเป็นเส้นเดียว -> ตั้งขอบช่องเป็นแบบเส้นคู่ของ Word
# ---------------------------------------------------------------------------

def _double_pairs(segs):
    """เส้นขนาน 2 เส้นห่างกัน 0.9–4pt ยาวทับกัน >= 70% (เส้นคู่แบบ Word/LibreOffice)"""
    pairs = []
    solid = [sg for sg in segs if not sg[6] and sg[4] <= 1.6]
    for i, a in enumerate(solid):
        for b in solid[i + 1:]:
            if a[0] != b[0]:
                continue
            d = abs(a[3] - b[3])
            if not 0.9 <= d <= 4.0:
                continue
            ov = min(a[2], b[2]) - max(a[1], b[1])
            if ov < 0.7 * min(a[2] - a[1], b[2] - b[1]):
                continue
            pairs.append((a[0], max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3]), max(a[3], b[3]),
                          (a[4] + b[4]) / 2))
    return pairs


def _cell_sides(c):
    x0, y0, x1, y1 = c["bbox"]
    bw = c.get("border_width") or (0, 0, 0, 0)
    geo = {"top": ("h", x0, x1, y0), "right": ("v", y0, y1, x1), "bottom": ("h", x0, x1, y1), "left": ("v", y0, y1, x0)}
    return [(side, geo[side]) for side, wd in zip(("top", "right", "bottom", "left"), bw) if wd and wd > 0]


def _tables_in_order(blocks, out):
    """ตารางตามลำดับเดียวกับ w:tbl ใน .docx: [(ช่องที่มีกรอบ [(แถว, คอลัมน์, ช่อง)])]"""
    for b in blocks:
        if "rows" in b:
            cells = []
            out.append(cells)
            for i, r in enumerate(b.get("rows", [])):
                for j, c in enumerate(r.get("cells", [])):
                    if c and c.get("bbox"):
                        cells.append((i, j, c))
                    if c:
                        _tables_in_order(c.get("blocks", []), out)
        else:
            _tables_in_order(b.get("blocks", []), out)


def _double_borders(pdf_path, parsed):
    """{หน้า: {"n": จำนวนตาราง, "items": [(ตาราง, แถว, คอลัมน์, ด้าน, ขนาด, จำนวนแถวที่ผสาน)]}}"""
    out = {}
    with fitz.open(pdf_path) as pdf:
        for pno, page in enumerate(pdf):
            tables = []
            try:
                for sec in parsed["pages"][pno]["sections"]:
                    for col in sec["columns"]:
                        _tables_in_order(col["blocks"], tables)
            except (KeyError, IndexError, TypeError):
                continue
            if not tables:
                continue
            pairs = _double_pairs(_pdf_segments(page))
            if not pairs:
                continue
            tsegs = [sg for cells in tables for _, _, c in cells for _, sg in _cell_sides(c)]
            items = []
            for k, cells in enumerate(tables):
                for i, j, c in cells:
                    rowspan = (c.get("merged_cells") or (1, 1))[0] or 1
                    for side, (o, a, b, pos) in _cell_sides(c):
                        for po, pa, pb, lo, hi, wd in pairs:
                            if po != o or not (lo - 1.0 <= pos <= hi + 1.0):
                                continue
                            if min(b, pb) - max(a, pa) < 0.7 * (b - a):
                                continue
                            tol = min(0.6, (hi - lo) * 0.3)

                            def near(y):
                                return any(t[0] == o and abs(t[3] - y) <= tol and
                                           min(t[2], pb) - max(t[1], pa) > 0.5 * (pb - pa) for t in tsegs)
                            if near(lo) and near(hi):                # 2 เส้นเป็นขอบจริงคนละช่องอยู่แล้ว
                                continue
                            items.append((k, i, j, side, max(2, min(96, round(wd * 8))), rowspan))
                            break
            if items:
                out[pno] = {"n": len(tables), "items": items}
    return out


def _tc_at(tr, col):
    start = 0
    for tc in tr.findall(qn("w:tc")):
        pr = tc.find(qn("w:tcPr"))
        span = pr.find(qn("w:gridSpan")) if pr is not None else None
        n = int(span.get(qn("w:val"), "1")) if span is not None else 1
        if start <= col < start + n:
            return tc
        start += n
    return None


def _apply_double_borders(doc, doubles):
    if not doubles:
        return 0
    pages, page = {}, 0
    for child in doc.element.body:
        pages.setdefault(page, []).append(child)
        if child.tag == qn("w:p") and child.find(qn("w:pPr") + "/" + qn("w:sectPr")) is not None:
            page += 1
    done = 0
    for pno, data in doubles.items():
        tbls = [t for el in pages.get(pno, []) for t in el.iter(qn("w:tbl"))]
        if len(tbls) != data["n"]:                             # โครงสร้างไม่ตรงกัน -> ไม่แตะ
            continue
        for k, i, j, side, sz, rowspan in data["items"]:
            rows = tbls[k].findall(qn("w:tr"))
            if side == "top":
                targets = [i]
            elif side == "bottom":
                targets = [i + rowspan - 1]
            else:
                targets = list(range(i, i + rowspan))
            for ri in targets:
                tc = _tc_at(rows[ri], j) if ri < len(rows) else None
                if tc is None:
                    continue
                pr = tc.find(qn("w:tcPr"))
                if pr is None:
                    pr = OxmlElement("w:tcPr")
                    tc.insert(0, pr)
                bd = pr.find(qn("w:tcBorders"))
                if bd is None:
                    bd = OxmlElement("w:tcBorders")
                    pr.append(bd)
                alt = {"left": "start", "right": "end"}.get(side)
                el = bd.find(qn("w:" + side))
                if el is None and alt:
                    el = bd.find(qn("w:" + alt))
                if el is None:
                    el = OxmlElement("w:" + side)
                    bd.append(el)
                el.set(qn("w:val"), "double")
                el.set(qn("w:sz"), str(sz))
                el.set(qn("w:space"), "0")
                if el.get(qn("w:color")) is None:
                    el.set(qn("w:color"), "000000")
                done += 1
    return done


_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_NS_WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
_NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
EMU = 12700


def _line_drawing_xml(seg, shape_id, rel_line_top=None):
    o, a, b, pos, width, color, dash, dashes = seg
    if o == "h":
        x, y, cx, cy = a, pos, b - a, 0
    else:
        x, y, cx, cy = pos, a, 0, b - a
    vrel = "page"
    if rel_line_top is not None:                       # ยึดกับบรรทัดข้อความ: ขยับตามข้อความถ้าเลย์เอาต์เลื่อน
        vrel, y = "line", y - rel_line_top
    rgb = "".join(f"{int(round(c * 255)):02X}" for c in (list(color) + [0, 0, 0])[:3])
    dash_xml = ""
    if dash:
        on = re.findall(r"[\d.]+", dashes.split("]")[0]) if dashes else []
        dotted = bool(on) and float(on[0]) <= max(1.6, width * 1.6)
        dash_xml = f'<a:prstDash val="{"sysDot" if dotted else "dash"}"/>'
    lw = max(int(width * EMU), 3175)
    return (f'<mc:AlternateContent xmlns:mc="{_NS_MC}"><mc:Choice Requires="wps"><w:drawing xmlns:w="{W_NS}">'
            f'<wp:anchor xmlns:wp="{_NS_WP}" distT="0" distB="0" distL="0" distR="0" simplePos="0" '
            f'relativeHeight="{251658240 + shape_id}" behindDoc="1" locked="0" layoutInCell="0" allowOverlap="1">'
            f'<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>{int(x * EMU)}</wp:posOffset>'
            f'</wp:positionH><wp:positionV relativeFrom="{vrel}"><wp:posOffset>{int(y * EMU)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{int(cx * EMU)}" cy="{int(cy * EMU)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:wrapNone/><wp:docPr id="{shape_id}" name="Line {shape_id}"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic xmlns:a="{_NS_A}"><a:graphicData uri="{_NS_WPS}"><wps:wsp xmlns:wps="{_NS_WPS}">'
            f'<wps:cNvCnPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(cx * EMU)}" cy="{int(cy * EMU)}"/>'
            f'</a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="{lw}"><a:solidFill>'
            f'<a:srgbClr val="{rgb}"/></a:solidFill>{dash_xml}</a:ln></wps:spPr><wps:bodyPr/></wps:wsp>'
            f'</a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback/></mc:AlternateContent>')


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


_NORM_DIGITS = str.maketrans("๐๑๒๓๔๕๖๗๘๙", "0123456789")


def _norm(t):
    return re.sub(r"\s+", "", fix_thai(t)).translate(_NORM_DIGITS)


# ---------------------------------------------------------------------------
# จัดบรรทัดให้ตรง PDF: ตำแหน่งเริ่มบรรทัด และระยะห่างตัวอักษร (บรรทัดที่ต้นฉบับบีบตัวอักษรให้พอดีบรรทัด)
# ---------------------------------------------------------------------------

def _line_key(t):
    t = re.sub(r"\s+", "", fix_thai(t)).translate(_NORM_DIGITS)
    return t.replace("\u0e33", "\u0e32").replace("\u0e4d", "")


def _pdf_line_info(pdf_path):
    """{หน้า: [(key, x ตัวอักษรแรก, ระยะบีบ/ขยายตัวอักษร, x ท้ายบรรทัด)]} (หน่วย pt) ตามลำดับในหน้า"""
    out = {}
    with fitz.open(pdf_path) as pdf:
        for pno, page in enumerate(pdf):
            items = []
            for b in page.get_text("rawdict")["blocks"]:
                for ln in b.get("lines", []):
                    d = ln.get("dir", (1, 0))
                    if abs(d[0] - 1) > 0.01 or abs(d[1]) > 0.01:
                        continue
                    chars = [c for sp in ln["spans"] for c in sp["chars"]]
                    ink = [c for c in chars if not c["c"].isspace()]
                    if not ink:
                        continue
                    key = _line_key("".join(c["c"] for c in chars))
                    if len(key) < 2:
                        continue
                    gaps = []
                    for a, c in zip(chars, chars[1:]):
                        if a["c"].isspace() or c["c"].isspace() or abs(c["origin"][1] - a["origin"][1]) > 0.5 \
                                or unicodedata.category(a["c"])[0] == "M" or unicodedata.category(c["c"])[0] == "M":
                            continue
                        gaps.append(c["origin"][0] - a["origin"][0] - (a["bbox"][2] - a["bbox"][0]))
                    tc = 0.0
                    if len(gaps) >= 8:
                        med = statistics.median(gaps)
                        if abs(med) >= 0.1 and sum(abs(g - med) <= 0.08 for g in gaps) >= 0.7 * len(gaps):
                            tc = med
                    items.append((key, ink[0]["bbox"][0], tc, ink[-1]["bbox"][2]))
            out[pno] = items
    return out


def _page_groups(doc):
    """[(องค์ประกอบในหน้า, sectPr ของหน้า)] (pdf2docx: 1 หน้า PDF = 1 section)"""
    groups, cur = [], []
    body = doc.element.body
    for child in body:
        if child.tag == qn("w:sectPr"):
            continue
        cur.append(child)
        sp = child.find(qn("w:pPr") + "/" + qn("w:sectPr")) if child.tag == qn("w:p") else None
        if sp is not None:
            groups.append((cur, sp))
            cur = []
    if cur:
        groups.append((cur, body.find(qn("w:sectPr"))))
    return groups


def _split_lines(p):
    """แบ่งย่อหน้าเป็นบรรทัดตาม <w:br/> คืน None ถ้าโครงสร้างซับซ้อน (ไม่แตะ)"""
    lines = [{"lead": [], "runs": [], "text": "", "mid": False}]
    for r in p.iter(qn("w:r")):
        if next(r.iterancestors(qn("w:p")), None) is not p:
            continue
        kids = [c for c in r if c.tag != qn("w:rPr")]
        tags = {c.tag for c in kids}
        cur = lines[-1]
        if tags & {qn("w:br"), qn("w:cr")}:
            if tags - {qn("w:br"), qn("w:cr")} or any(c.get(qn("w:type")) not in (None, "textWrapping") for c in kids):
                return None
            lines.append({"lead": [], "runs": [], "text": "", "mid": False})
            continue
        if tags == {qn("w:tab")}:
            if cur["text"].strip():
                cur["mid"] = True
            else:
                cur["lead"].append(r)
            continue
        if qn("w:tab") in tags:
            cur["mid"] = True
        if qn("w:t") in tags:
            cur["runs"].append(r)
            cur["text"] += "".join(t.text or "" for t in r.findall(qn("w:t")))
    return lines


def _tab_run(like):
    r = OxmlElement("w:r")
    rpr = like.find(qn("w:rPr")) if like is not None else None
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    r.append(OxmlElement("w:tab"))
    return r


def _retab(p, lines, xs, margin_tw):
    """ตั้งแท็บใหม่ให้ทุกบรรทัดเริ่มตรงตำแหน่งใน PDF (แท็บวัดจากขอบกระดาษซ้าย)"""
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        return False
    jc = ppr.find(qn("w:jc"))
    if jc is not None and jc.get(qn("w:val")) not in ("left", "start"):   # จัดกลาง/ขวา/เต็มบรรทัด: แท็บใช้ไม่ได้ตรง
        return False
    ind = ppr.find(qn("w:ind"))
    base = 0
    if ind is not None:
        if any(int(float(ind.get(qn(a)) or 0)) for a in ("w:firstLine", "w:hanging")):
            return False
        base = int(float(ind.get(qn("w:left")) or ind.get(qn("w:start")) or 0))
    targets = [round(x * 20) - margin_tw for x in xs]
    if any(t < base - 30 for t in targets):
        return False
    stops = []
    for t in sorted(t for t in targets if t > base + 20):
        if not stops or t - stops[-1][-1] > 12:
            stops.append([t])
        else:
            stops[-1].append(t)
    pos = [round(sum(c) / len(c)) for c in stops]
    for ln, t in zip(lines, targets):
        k = 0 if t <= base + 20 else next(i + 1 for i, c in enumerate(stops) if t in c)
        first = ln["runs"][0]
        like = ln["lead"][0] if ln["lead"] else first
        for r in ln["lead"]:
            r.getparent().remove(r)
        for _ in range(k):
            first.addprevious(_tab_run(like))
        t0 = first.find(qn("w:t"))                       # ช่องว่างนำหน้า -> ใช้แท็บวางตำแหน่งแทน
        if t0 is not None and t0.text:
            t0.text = t0.text.lstrip(" ")
    old = ppr.find(qn("w:tabs"))
    if old is not None:
        ppr.remove(old)
    if pos:
        tabs = OxmlElement("w:tabs")
        for v in pos:
            tab = OxmlElement("w:tab")
            tab.set(qn("w:val"), "left")
            tab.set(qn("w:pos"), str(v))
            tabs.append(tab)
        ppr.append(tabs)                                   # ลำดับที่ถูกต้องจัดให้ตอน finalize_docx
    return True


def _uncenter(p):
    ppr = p.find(qn("w:pPr"))
    jc = ppr.find(qn("w:jc")) if ppr is not None else None
    if jc is None or jc.get(qn("w:val")) != "center":
        return
    jc.set(qn("w:val"), "left")
    ind = ppr.find(qn("w:ind"))
    if ind is not None:
        for a in ("w:right", "w:end"):
            if ind.get(qn(a)) is not None:
                ind.set(qn(a), "0")


def _align_with_pdf(doc, pdf_path):
    """ให้แต่ละบรรทัดใน Word เริ่มตรงตำแหน่งเดียวกับ PDF และบีบตัวอักษรเท่าต้นฉบับ
    (pdf2docx ใส่จำนวนแท็บผิดในบางบรรทัด และไม่เก็บค่าการบีบตัวอักษร ทำให้บรรทัดยาวเกินแล้วตกบรรทัด)"""
    info = _pdf_line_info(pdf_path)
    stats = {"retab": 0, "spacing": 0}
    for pno, (elems, sect) in enumerate(_page_groups(doc)):
        items = info.get(pno)
        if not items:
            continue
        used = [False] * len(items)
        mar = sect.find(qn("w:pgMar")) if sect is not None else None
        margin_tw = int(float(mar.get(qn("w:left")) or 0)) if mar is not None else None
        for el in elems:
            for p in el.iter(qn("w:p")):
                lines = _split_lines(p)
                if lines is None:
                    continue
                lines = [ln for ln in lines if ln["text"].strip() and ln["runs"]]
                xs, ends, ok = [], [], bool(lines)
                for ln in lines:
                    key = _line_key(ln["text"])
                    k = next((i for i, it in enumerate(items) if not used[i] and it[0] == key), None)
                    if k is None:
                        ok = False
                        continue
                    used[k] = True
                    xs.append(items[k][1])
                    ends.append(items[k][3])
                    tc = items[k][2]
                    if tc:
                        val = str(max(-40, min(40, round(tc * 20))))
                        for r in ln["runs"]:
                            rpr = r.find(qn("w:rPr"))
                            if rpr is None:
                                rpr = OxmlElement("w:rPr")
                                r.insert(0, rpr)
                            sp = rpr.find(qn("w:spacing"))
                            if sp is None:
                                sp = OxmlElement("w:spacing")
                                rpr.append(sp)
                            sp.set(qn("w:val"), val)
                        stats["spacing"] += 1
                in_table = next(p.iterancestors(qn("w:tbl")), None) is not None
                if ok and not in_table and len(xs) >= 2 and max(xs) - min(xs) < 1.0 and max(ends) - min(ends) > 4.0:
                    _uncenter(p)          # pdf2docx เดาว่าจัดกลาง แต่ใน PDF ทุกบรรทัดเริ่มตรงกัน = ชิดซ้าย
                if ok and margin_tw is not None and not in_table and not any(ln["mid"] for ln in lines) \
                        and len(xs) == len(lines) and _retab(p, lines, xs, margin_tw):
                    stats["retab"] += 1
    return stats


def _restore_lines(doc, orphans):
    """ใส่เส้นที่หายกลับเป็นรูปเส้นตรง: ยึดกับบรรทัดข้อความที่อยู่เหนือเส้น (ขยับตามข้อความ)
    ถ้าหาบรรทัดไม่เจอ ยึดกับตำแหน่งบนหน้าแทน"""
    if not orphans:
        return 0
    from lxml import etree
    body = doc.element.body
    # แบ่งย่อหน้าตามหน้า (pdf2docx: 1 หน้า PDF = 1 section)
    pages, page = {}, 0
    for child in body:
        pages.setdefault(page, []).append(child)
        if child.tag == qn("w:p") and child.find(qn("w:pPr") + "/" + qn("w:sectPr")) is not None:
            page += 1
    ids = [int(x) for x in re.findall(r'<wp:docPr[^>]*\sid="(\d+)"', etree.tostring(body).decode())] or [0]
    next_id = max(ids) + 1
    added = 0
    for pno, data in orphans.items():
        elems = pages.get(pno, [])
        # ข้อความทั้งหน้าใน .docx พร้อมตำแหน่ง run (รวมในตาราง)
        stream, owners = [], []
        for el in elems:
            for r in el.iter(qn("w:r")):
                for t in r.iter(qn("w:t")):
                    for ch in _norm(t.text or ""):
                        stream.append(ch)
                        owners.append(r)
        text = "".join(stream)
        first_para = next((e for e in elems if e.tag == qn("w:p")), None)
        for sg in data["segs"]:
            top_y = sg[3] if sg[0] == "h" else sg[1]
            # บรรทัดข้อความที่ใกล้ที่สุดซึ่งอยู่เหนือ (หรือระดับเดียวกับ) เส้น
            cands = [(top_y - bb[1], bb, t) for bb, t in data["lines"] if bb[1] <= top_y + 1 and len(_norm(t)) >= 2]
            anchor_run, rel = None, None
            for _, bb, t in sorted(cands)[:3]:
                key = _norm(t)[:14]
                k = text.find(key)
                if k >= 0:
                    anchor_run, rel = owners[k], bb[1]
                    break
            xml = _line_drawing_xml(sg, next_id, rel)
            next_id += 1
            run = OxmlElement("w:r")
            run.append(etree.fromstring(xml))
            if anchor_run is not None and anchor_run.getparent() is not None:
                anchor_run.addprevious(run)
            elif first_para is not None:                      # ไม่เจอบรรทัด -> ยึดกับหน้า
                run = OxmlElement("w:r")
                run.append(etree.fromstring(_line_drawing_xml(sg, next_id - 1, None)))
                ppr = first_para.find(qn("w:pPr"))
                (ppr.addnext(run) if ppr is not None else first_para.insert(0, run))
            else:
                continue
            added += 1
    return added


def _relax_right_indent(root):
    """บรรทัดถูกตัดตรงตำแหน่งเดียวกับ PDF อยู่แล้ว (ขึ้นบรรทัดใหม่ทุกบรรทัด) -> ระยะเยื้องขวาของย่อหน้าชิดซ้าย
    ไม่จำเป็น และทำให้คำสุดท้ายตกบรรทัดถ้าฟอนต์ในเครื่องที่เปิดกว้างกว่าเล็กน้อย"""
    for p in root.iter(qn("w:p")):
        ppr = p.find(qn("w:pPr"))
        if ppr is None:
            continue
        jc = ppr.find(qn("w:jc"))
        if jc is not None and jc.get(qn("w:val")) not in ("left", "start", None):
            continue
        ind = ppr.find(qn("w:ind"))
        for attr in ("w:right", "w:end"):
            if ind is not None and ind.get(qn(attr)) not in (None, "0"):
                ind.set(qn(attr), "0")


def _shrink_empty_headers(doc):
    """pdf2docx ใส่หัว/ท้ายกระดาษว่างที่สูงเกินขอบกระดาษ -> เนื้อหาถูกดันลง ~12pt และล้นไปหน้าใหม่
    ทำให้หัว/ท้ายกระดาษที่ว่างเตี้ยที่สุด และชิดขอบ"""
    for sec in doc.sections:
        for part, dist in ((sec.header, "header_distance"), (sec.footer, "footer_distance")):
            try:
                if part.is_linked_to_previous:            # ไม่มีหัว/ท้ายกระดาษ (อย่าสร้างใหม่)
                    continue
                el = part._element
            except Exception:  # noqa: BLE001
                continue
            if "".join(t.text or "" for t in el.iter(qn("w:t"))).strip():
                continue                                      # มีข้อความจริง ไม่แตะ
            for p in el.iter(qn("w:p")):
                ppr = p.find(qn("w:pPr"))
                if ppr is None:
                    ppr = OxmlElement("w:pPr")
                    p.insert(0, ppr)
                sp = _child(ppr, "w:spacing")
                sp.set(qn("w:before"), "0")
                sp.set(qn("w:after"), "0")
                sp.set(qn("w:line"), "20")
                sp.set(qn("w:lineRule"), "exact")
                rpr = _child(ppr, "w:rPr")
                for tag in ("w:sz", "w:szCs"):
                    _child(rpr, tag).set(qn("w:val"), "2")
            try:
                cur = getattr(sec, dist)
                if cur is None or cur > Pt(12):
                    setattr(sec, dist, Pt(12))
            except Exception:  # noqa: BLE001
                pass


def convert_text(pdf_path, out_path, font, lost_am=True):
    cv = Converter(pdf_path)
    try:
        # ขึ้นบรรทัดใหม่ตรงกับ PDF ทุกบรรทัด (ค่าเริ่มต้นของ pdf2docx รวมบรรทัดที่เกือบเต็มเข้าด้วยกัน ทำให้ข้อ (2) (4) หาย/ติดกัน)
        cv.convert(out_path, line_break_free_space_ratio=0.0, line_break_width_ratio=1.0)
        try:
            parsed = cv.store()
        except Exception:  # noqa: BLE001
            parsed = {"pages": []}
    finally:
        cv.close()
    orphans = _orphan_lines(pdf_path, parsed)
    try:
        doubles = _double_borders(pdf_path, parsed)
    except Exception:  # noqa: BLE001
        doubles = {}
    # คู่คำที่อยู่ติดกันใน PDF ต้นฉบับ (ใช้คืนช่องว่างที่ pdf2docx ทำหาย)
    with fitz.open(pdf_path) as pdf:
        raw_tokens = fix_thai(" ".join(p.get_text() for p in pdf), lost_am).split()
    pairs = set(zip(raw_tokens, raw_tokens[1:]))
    known = set(raw_tokens)

    doc = Document(out_path)
    roots = [doc.element.body] + _hf_roots(doc)
    for root in roots:
        for t in root.iter(qn("w:t")):
            if t.text:
                t.text = fix_thai(t.text, lost_am)
        _restore_spaces(root, pairs, known)
        _keep_thai_digits(root, font)
        _fix_tables(root)
        _relax_right_indent(root)
        sanitize_ooxml(root)
        # คำที่ถูกแยกคนละ run เช่น "ส" | " าหรับ", "ท" | " ี่", "เ " | "ป็น"
        for para in root.iter(qn("w:p")):
            ts = [t for t in para.iter(qn("w:t")) if t.text]
            for prev, cur in zip(ts, ts[1:]):
                _fix_boundary(prev, cur, lost_am)
    _restore_lines(doc, orphans)
    _apply_double_borders(doc, doubles)
    try:
        _align_with_pdf(doc, pdf_path)
    except Exception:  # noqa: BLE001
        pass
    _shrink_empty_headers(doc)
    apply_thai_font(doc, font)
    finalize_docx(doc)
    doc.save(out_path)
    embed_fonts(out_path)


def _page_image(page, dpi, gray=False):
    pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY if gray else fitz.csRGB)
    mode = "L" if gray else "RGB"
    return Image.frombytes(mode, (pix.width, pix.height), pix.samples)


def _setup_section(section, page, margin):
    w, h = page.rect.width, page.rect.height     # หน่วย point
    section.page_width, section.page_height = Pt(w), Pt(h)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(section, side, margin)
    section.header_distance = section.footer_distance = Emu(0)


def convert_ocr(pdf_path, out_path, font, dpi=300, dict_fix=True):
    """OCR ภาษาไทย (+อังกฤษ) แล้วแก้คำผิดด้วยพจนานุกรม คำที่ถูกแก้จะไฮไลต์สีเหลืองให้ตรวจ"""
    if dict_fix:
        from thai_ocr_fix import fix_ocr_segments
    doc = Document()
    fixed_count = 0
    with fitz.open(pdf_path) as pdf:
        _setup_section(doc.sections[0], pdf[0], Pt(54))
        for i, page in enumerate(pdf):
            img = _page_image(page, dpi, gray=True)
            text = pytesseract.image_to_string(img, lang="tha+eng", config="--oem 1 --psm 3")
            for block in re.split(r"\n\s*\n", text):
                lines = [ln.rstrip() for ln in block.splitlines() if ln.strip()]
                segs_per_line = []
                for ln in lines:
                    ln = fix_thai(ln)
                    segs = fix_ocr_segments(ln) if dict_fix else [(ln, False)]
                    if segs:
                        segs_per_line.append(segs)
                if not segs_per_line:
                    continue
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(6)
                for j, segs in enumerate(segs_per_line):
                    for t, changed in segs:
                        run = p.add_run(t)
                        if changed:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                            fixed_count += 1
                    if j < len(segs_per_line) - 1:
                        run.add_break()
            if i < len(pdf) - 1:
                doc.add_page_break()
    apply_thai_font(doc, DEFAULT_FONT if font == "original" else font)
    finalize_docx(doc)
    doc.save(out_path)
    embed_fonts(out_path)
    return fixed_count


def convert_image(pdf_path, out_path, dpi=200):
    doc = Document()
    with fitz.open(pdf_path) as pdf:
        for i, page in enumerate(pdf):
            section = doc.sections[0] if i == 0 else doc.add_section(WD_SECTION.NEW_PAGE)
            _setup_section(section, page, Emu(0))
            buf = io.BytesIO()
            _page_image(page, dpi).save(buf, format="PNG", optimize=True)
            buf.seek(0)
            p = doc.paragraphs[0] if i == 0 and doc.paragraphs else doc.add_paragraph()
            fmt = p.paragraph_format
            fmt.space_before = fmt.space_after = Pt(0)
            fmt.line_spacing = 1.0
            # ย่อเล็กน้อยกันล้นไปหน้าใหม่
            p.add_run().add_picture(buf, height=Pt(page.rect.height * 0.985))
    finalize_docx(doc)
    doc.save(out_path)


def convert(pdf_path: str, out_path: str, mode: str = "auto", font: str = DEFAULT_FONT,
            dict_fix: bool = True) -> dict:
    info = analyze(pdf_path)
    if mode == "auto":
        mode = info["method"]
        detail = info["reason"]
    else:
        detail = {
            "text": "ดึงข้อความจาก PDF (เลือกเอง)",
            "ocr": "อ่านด้วย OCR (เลือกเอง)",
            "image": "ใส่เป็นรูปภาพ เหมือนต้นฉบับ แก้ไขข้อความไม่ได้",
        }.get(mode)
        if detail is None:
            raise ValueError(f"ไม่รู้จักโหมด: {mode}")

    if mode == "text":
        convert_text(pdf_path, out_path, font, info["lost_am"])
    elif mode == "ocr":
        n = convert_ocr(pdf_path, out_path, font, dict_fix=dict_fix)
        if dict_fix:
            detail += f" · แก้คำด้วยพจนานุกรม {n} จุด (ไฮไลต์สีเหลือง)" if n else " · ไม่พบคำที่ต้องแก้"
    else:
        convert_image(pdf_path, out_path)

    info.update(method=mode, detail=detail)
    return info


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="PDF -> Word (ภาษาไทย)")
    ap.add_argument("pdf")
    ap.add_argument("docx")
    ap.add_argument("--mode", default="auto", choices=["auto", "text", "ocr", "image"])
    ap.add_argument("--font", default=DEFAULT_FONT)
    ap.add_argument("--no-dict", action="store_true", help="ปิดการแก้คำผิดจาก OCR ด้วยพจนานุกรม")
    a = ap.parse_args()
    print(convert(a.pdf, a.docx, a.mode, a.font, not a.no_dict))
