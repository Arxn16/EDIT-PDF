"""วัดคุณภาพการแปลง PDF -> Word เทียบกับต้นฉบับ

แปลง PDF เป็น .docx แล้วให้ LibreOffice แปลง .docx กลับเป็น PDF จากนั้นเทียบ
  - ข้อความ: คำที่หาย / เกิน (ไม่สนลำดับ)
  - เส้นตาราง: ความยาวรวมของเส้นแนวนอน/แนวตั้ง
  - พื้นหลังสี (shading): พื้นที่รวมแยกตามสี
  - ตำแหน่งเส้นตาราง: เส้นแนวตั้งเลื่อนจากต้นฉบับกี่ pt (ตารางเยื้องซ้าย/ขวา)
  - ตำแหน่งบรรทัด: ทุกบรรทัดอยู่ที่เดิมไหม (ตกบรรทัด/เยื้องผิด/เลื่อนลง)
  - ไฟล์ Word ถูกต้องตามมาตรฐาน: ไม่มีค่าที่ Word อ่านไม่ได้ (ทศนิยม, สี #xxxxxx),
    ลำดับ element ถูกต้อง, ไม่มี element ซ้ำ, ใช้ w:left/w:right แบบที่ Word เขียนเอง

รัน:  python tests/eval_convert.py file1.pdf [file2.pdf ...]
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from collections import Counter

import pymupdf as fitz

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, os.path.dirname(__file__))
from converter import convert, fix_thai  # noqa: E402
from docx_fix import ooxml_issues  # noqa: E402
from lxml import etree  # noqa: E402


def drawings_summary(doc):
    """ความยาวรวมเส้นแนวนอน/แนวตั้ง (pt) และพื้นที่ shading แยกสี"""
    h = v = 0.0
    shade = Counter()
    for page in doc:
        for d in page.get_drawings():
            fill = d.get("fill")
            for it in d["items"]:
                if it[0] == "re":
                    r = fitz.Rect(it[1])
                    if r.width <= 2.5 and r.height > 2.5:
                        v += r.height
                    elif r.height <= 2.5 and r.width > 2.5:
                        h += r.width
                    elif fill and r.width > 2.5 and r.height > 2.5 and min(fill) < 0.995:
                        shade[tuple(round(c, 1) for c in fill)] += r.width * r.height
                elif it[0] == "l" and d.get("color") is not None:
                    p1, p2 = it[1], it[2]
                    if abs(p1.y - p2.y) < 1:
                        h += abs(p2.x - p1.x)
                    elif abs(p1.x - p2.x) < 1:
                        v += abs(p2.y - p1.y)
    return h, v, shade


def words(text):
    try:
        from pythainlp.tokenize import word_tokenize
        toks = word_tokenize(re.sub(r"\s+", " ", text), engine="newmm", keep_whitespace=False)
    except Exception:  # noqa: BLE001
        toks = text.split()
    return Counter(t for t in toks if t.strip())


def docx_problems(path):
    bad = Counter()
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if not name.endswith(".xml"):
                continue
            x = z.read(name).decode("utf8", "ignore")
            for m in re.finditer(r'\s(w:[A-Za-z]+)="(-?\d+\.\d+)"', x):
                bad["decimal " + m.group(1)] += 1
            for m in re.finditer(r'\s(w:(?:color|fill))="#', x):
                bad["#color " + m.group(1)] += 1
            if name.startswith("word/") and "/" not in name[5:]:
                try:
                    bad.update(ooxml_issues(etree.fromstring(z.read(name))))
                except etree.XMLSyntaxError:
                    bad["xml error " + name] += 1
    return bad


def _vsegs(page):
    out = []
    for d in page.get_drawings():
        for it in d["items"]:
            if it[0] == "re":
                r = fitz.Rect(it[1])
                if r.width <= 2.5 and r.height > 2.5:
                    out.append(((r.x0 + r.x1) / 2, r.y0, r.y1))
            elif it[0] == "l" and d.get("color") is not None:
                p1, p2 = it[1], it[2]
                if abs(p1.x - p2.x) < 1 and abs(p1.y - p2.y) > 2.5:
                    out.append((p1.x, min(p1.y, p2.y), max(p1.y, p2.y)))
    return out


def _line_key(t):
    return re.sub(r"\s+", "", fix_thai(t)).translate(str.maketrans("๐๑๒๓๔๕๖๗๘๙", "0123456789")) \
        .replace("\u0e33", "\u0e32").replace("\u0e4d", "")


def _text_lines(page):
    out = []
    for b in page.get_text("rawdict")["blocks"]:
        for ln in b.get("lines", []):
            chars = [c for sp in ln["spans"] for c in sp["chars"]]
            ink = [c for c in chars if not c["c"].isspace()]
            if ink:
                key = _line_key("".join(c["c"] for c in chars))
                if len(key) >= 2:
                    out.append((key, ink[0]["origin"][0], ink[0]["origin"][1]))
    return out


def line_positions(o, b):
    """(บรรทัดที่อยู่ที่เดิม, ทั้งหมด, คลาดแนวนอนสูงสุด, คลาดแนวตั้งสูงสุด) บรรทัดที่ข้อความไม่ตรงกัน = ตกบรรทัด"""
    total = ok = 0
    mx = my = 0.0
    for po, pb in zip(o, b):
        lb, used = _text_lines(pb), set()
        for key, x, y in _text_lines(po):
            total += 1
            cands = [k for k in range(len(lb)) if k not in used and lb[k][0] == key]
            if not cands:
                continue
            k = min(cands, key=lambda k: abs(lb[k][2] - y))
            used.add(k)
            dx, dy = abs(lb[k][1] - x), abs(lb[k][2] - y)
            mx, my = max(mx, dx), max(my, dy)
            ok += dx <= 1.5 and dy <= 3.0
    return ok, total, round(mx, 1), round(my, 1)


def vline_offset(o, b):
    """ระยะเลื่อนแนวนอนของเส้นตั้ง (ค่ามัธยฐาน pt) ระหว่างต้นฉบับกับไฟล์ที่แปลงกลับ"""
    dx = []
    for po, pb in zip(o, b):
        bs = _vsegs(pb)
        for x, y0, y1 in _vsegs(po):
            near = [bx - x for bx, by0, by1 in bs if abs(bx - x) < 12 and min(y1, by1) - max(y0, by0) > 2]
            if near:
                dx.append(min(near, key=abs))
    if not dx:
        return 0.0
    dx.sort()
    return round(dx[len(dx) // 2], 1)


def _schema_errors(docx):
    """จำนวนจุดที่ผิด XML Schema มาตรฐาน (None = ตรวจไม่ได้ เช่น ไม่มีอินเทอร์เน็ตตอนโหลด schema ครั้งแรก)"""
    try:
        from validate_docx import validate
        return sum(validate(docx).values())
    except Exception:  # noqa: BLE001
        return None


def to_pdf_libreoffice(docx, outdir):
    subprocess.run(["soffice", "--headless", "--norestore", f"-env:UserInstallation=file://{outdir}/lo",
                    "--convert-to", "pdf", "--outdir", outdir, docx], capture_output=True, timeout=300)
    return os.path.join(outdir, os.path.splitext(os.path.basename(docx))[0] + ".pdf")


def evaluate(pdf, font=None, keep=None):
    tmp = tempfile.mkdtemp(prefix="eval_")
    docx = os.path.join(tmp, "out.docx")
    kw = {"font": font} if font else {}
    info = convert(pdf, docx, "text", **kw)
    back = to_pdf_libreoffice(docx, tmp)
    o, b = fitz.open(pdf), fitz.open(back)
    oh, ov, osh = drawings_summary(o)
    bh, bv, bsh = drawings_summary(b)
    ow = words(fix_thai("".join(p.get_text() for p in o)).translate(str.maketrans("0123456789", "๐๑๒๓๔๕๖๗๘๙"))
               if any("IT" in f[3] for p in o for f in p.get_fonts()) else fix_thai("".join(p.get_text() for p in o)))
    bw = words("".join(p.get_text() for p in b))
    miss, extra = ow - bw, bw - ow
    total = sum(ow.values()) or 1
    res = {
        "file": os.path.basename(pdf), "pages": (len(o), len(b)),
        "h_lines": (round(oh), round(bh), round(100 * min(oh, bh) / oh) if oh else 100),
        "v_lines": (round(ov), round(bv), round(100 * min(ov, bv) / ov) if ov else 100),
        "shading": (round(sum(osh.values())), round(sum(bsh.values()))),
        "vline_dx": vline_offset(o, b),
        "lines": line_positions(o, b),
        "words_missing": sum(miss.values()), "words_extra": sum(extra.values()), "words_total": total,
        "docx_invalid": dict(docx_problems(docx)), "method": info["method"],
        "schema_errors": _schema_errors(docx),
        "top_missing": dict(miss.most_common(5)),
    }
    if keep:
        shutil.copy(docx, keep + ".docx")
        shutil.copy(back, keep + ".roundtrip.pdf")
    shutil.rmtree(tmp, ignore_errors=True)
    return res


if __name__ == "__main__":
    for f in sys.argv[1:]:
        r = evaluate(f)
        print(f"\n== {r['file']}  ({r['method']})  หน้า {r['pages'][0]} -> {r['pages'][1]}")
        print(f"   เส้นแนวนอน  ต้นฉบับ {r['h_lines'][0]} pt  หลังแปลง {r['h_lines'][1]} pt  ({r['h_lines'][2]}%)")
        print(f"   เส้นแนวตั้ง  ต้นฉบับ {r['v_lines'][0]} pt  หลังแปลง {r['v_lines'][1]} pt  ({r['v_lines'][2]}%)")
        print(f"   พื้นหลังสี   ต้นฉบับ {r['shading'][0]}  หลังแปลง {r['shading'][1]}")
        print(f"   ตำแหน่งเส้นตั้งเลื่อน {r['vline_dx']} pt")
        ok, tot, mx, my = r["lines"]
        print(f"   บรรทัดอยู่ที่เดิม {ok}/{tot} (คลาดแนวนอนสูงสุด {mx} pt, แนวตั้ง {my} pt)")
        print(f"   คำหาย {r['words_missing']} / เกิน {r['words_extra']} จาก {r['words_total']}  {r['top_missing']}")
        print(f"   ค่าที่ Word อ่านไม่ได้ใน .docx: {r['docx_invalid'] or 'ไม่มี'}")
        if r["schema_errors"] is not None:
            print(f"   ตรวจกับ XML Schema มาตรฐาน OOXML: {'ถูกต้อง' if not r['schema_errors'] else str(r['schema_errors']) + ' จุดผิด'}")
