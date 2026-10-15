"""แปลง PDF ภาษาไทยเป็น Word โดยไม่ให้สระ/วรรณยุกต์เพี้ยน

โหมดที่รองรับ
  auto  - ตรวจไฟล์ก่อน แล้วเลือก text หรือ ocr ให้เอง
  text  - ดึงข้อความจาก PDF (pdf2docx) เก็บรูปแบบหน้า + ซ่อมรหัสภาษาไทย
  ocr   - อ่านจากภาพด้วย Tesseract (ใช้กับไฟล์สแกน / ข้อความในไฟล์เสีย)
  image - ใส่แต่ละหน้าเป็นรูปภาพ เหมือนต้นฉบับ 100% แต่แก้ข้อความไม่ได้
"""
import io
import re

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


def apply_thai_font(doc: Document, font: str):
    """font = ชื่อฟอนต์ หรือ 'original' เพื่อใช้ฟอนต์เดิมจาก PDF"""
    roots = [doc.element.body]
    for s in doc.sections:
        roots += [s.header._element, s.footer._element]

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
        ts = [t for t in para.iter(qn("w:t")) if t.text]
        for prev, cur in zip(ts, ts[1:]):
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


def convert_text(pdf_path, out_path, font, lost_am=True):
    cv = Converter(pdf_path)
    try:
        # ขึ้นบรรทัดใหม่ตรงกับ PDF ทุกบรรทัด (ค่าเริ่มต้นของ pdf2docx รวมบรรทัดที่เกือบเต็มเข้าด้วยกัน ทำให้ข้อ (2) (4) หาย/ติดกัน)
        cv.convert(out_path, line_break_free_space_ratio=0.0, line_break_width_ratio=1.0)
    finally:
        cv.close()
    # คู่คำที่อยู่ติดกันใน PDF ต้นฉบับ (ใช้คืนช่องว่างที่ pdf2docx ทำหาย)
    with fitz.open(pdf_path) as pdf:
        raw_tokens = fix_thai(" ".join(p.get_text() for p in pdf), lost_am).split()
    pairs = set(zip(raw_tokens, raw_tokens[1:]))
    known = set(raw_tokens)

    doc = Document(out_path)
    roots = [doc.element.body] + [x for s in doc.sections for x in (s.header._element, s.footer._element)]
    for root in roots:
        for t in root.iter(qn("w:t")):
            if t.text:
                t.text = fix_thai(t.text, lost_am)
        _restore_spaces(root, pairs, known)
        _keep_thai_digits(root, font)
        _fix_tables(root)
        # คำที่ถูกแยกคนละ run เช่น "ส" | " าหรับ", "ท" | " ี่", "เ " | "ป็น"
        for para in root.iter(qn("w:p")):
            ts = [t for t in para.iter(qn("w:t")) if t.text]
            for prev, cur in zip(ts, ts[1:]):
                _fix_boundary(prev, cur, lost_am)
    apply_thai_font(doc, font)
    doc.save(out_path)


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
    apply_thai_font(doc, "TH Sarabun New" if font == "original" else font)
    doc.save(out_path)
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
    doc.save(out_path)


def convert(pdf_path: str, out_path: str, mode: str = "auto", font: str = "TH Sarabun New",
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
    ap.add_argument("--font", default="TH Sarabun New")
    ap.add_argument("--no-dict", action="store_true", help="ปิดการแก้คำผิดจาก OCR ด้วยพจนานุกรม")
    a = ap.parse_args()
    print(convert(a.pdf, a.docx, a.mode, a.font, not a.no_dict))
