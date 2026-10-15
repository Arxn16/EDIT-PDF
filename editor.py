"""ระบบแก้ไข PDF + Export (Flask blueprint)

เก็บเอกสารเป็นไฟล์ใน WORK_DIR/<id>/ (v0.pdf, v1.pdf, ...) ทุกการแก้ไขสร้างเวอร์ชันใหม่ ย้อนกลับได้
"""
import base64
import io
import json
import os
import re
import shutil
import subprocess
import tempfile
import time
import uuid
import zipfile

import pymupdf as fitz
from flask import Blueprint, abort, jsonify, render_template, request, send_file

from converter import DEFAULT_FONT, convert, fix_thai
from thai_pdf_text import SARABUN_LIKE, add_runs_at_baseline, add_text_box, line_metrics, pick_family, text_width

bp = Blueprint("editor", __name__)
WORK_DIR = os.environ.get("EDIT_WORK_DIR", os.path.join(tempfile.gettempdir(), "pdfedit"))
KEEP_HOURS = 12
OFFICE_EXT = {".doc", ".docx", ".odt", ".rtf", ".txt", ".xls", ".xlsx", ".ods", ".csv",
              ".ppt", ".pptx", ".odp"}
IMAGE_EXT = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp"}
THAI_DIGITS = str.maketrans("0123456789", "๐๑๒๓๔๕๖๗๘๙")


# ---------------------------------------------------------------------------
# เก็บเอกสาร / เวอร์ชัน
# ---------------------------------------------------------------------------

def _dir(doc_id):
    if not re.fullmatch(r"[a-f0-9]{32}", doc_id or ""):
        abort(404)
    d = os.path.join(WORK_DIR, doc_id)
    if not os.path.isdir(d):
        abort(404, "ไม่พบเอกสาร (อาจหมดอายุ) กรุณาเปิดไฟล์ใหม่")
    return d


def _meta(doc_id):
    with open(os.path.join(_dir(doc_id), "meta.json"), encoding="utf-8") as f:
        return json.load(f)


def _save_meta(doc_id, meta):
    with open(os.path.join(_dir(doc_id), "meta.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False)


def _current(doc_id):
    m = _meta(doc_id)
    return os.path.join(_dir(doc_id), f"v{m['version']}.pdf")


def _commit(doc_id, doc):
    """บันทึกเป็นเวอร์ชันใหม่"""
    m = _meta(doc_id)
    m["version"] += 1
    m["max_version"] = m["version"]
    path = os.path.join(_dir(doc_id), f"v{m['version']}.pdf")
    try:
        doc.subset_fonts()                          # ลดขนาดฟอนต์ที่ฝังเพิ่ม
    except Exception:  # noqa: BLE001
        pass
    doc.save(path, garbage=3, deflate=True)
    # ลบเวอร์ชันที่ถูก "redo" ทับ
    for f in os.listdir(_dir(doc_id)):
        mm = re.fullmatch(r"v(\d+)\.pdf", f)
        if mm and int(mm.group(1)) > m["version"]:
            os.remove(os.path.join(_dir(doc_id), f))
    _save_meta(doc_id, m)
    return m


def _cleanup():
    if not os.path.isdir(WORK_DIR):
        return
    now = time.time()
    for d in os.listdir(WORK_DIR):
        p = os.path.join(WORK_DIR, d)
        if os.path.isdir(p) and now - os.path.getmtime(p) > KEEP_HOURS * 3600:
            shutil.rmtree(p, ignore_errors=True)


def _info(doc_id):
    m = _meta(doc_id)
    with fitz.open(_current(doc_id)) as doc:
        pages = [{"w": p.rect.width, "h": p.rect.height} for p in doc]
    return {"id": doc_id, "name": m["name"], "version": m["version"],
            "can_undo": m["version"] > 0, "can_redo": m["version"] < m.get("max_version", m["version"]),
            "pages": pages}


# ---------------------------------------------------------------------------
# แปลงไฟล์อื่นเป็น PDF
# ---------------------------------------------------------------------------

def to_pdf(path, name):
    """คืน fitz.Document จากไฟล์ PDF / รูป / Word / Excel / PowerPoint"""
    ext = os.path.splitext(name.lower())[1]
    if ext == ".pdf":
        return fitz.open(path)
    if ext in IMAGE_EXT:
        img = fitz.open(path)
        return fitz.open("pdf", img.convert_to_pdf())
    if ext in OFFICE_EXT:
        out_dir = tempfile.mkdtemp(prefix="office_")
        src = os.path.join(out_dir, "input" + ext)
        shutil.copy(path, src)
        try:
            subprocess.run(["soffice", "--headless", "--norestore", f"-env:UserInstallation=file://{out_dir}/lo",
                            "--convert-to", "pdf", "--outdir", out_dir, src],
                           check=True, capture_output=True, timeout=180)
        except FileNotFoundError:
            raise ValueError("ไม่พบ LibreOffice สำหรับแปลงไฟล์ Office (ใช้ผ่าน Docker)")
        pdf = os.path.join(out_dir, "input.pdf")
        if not os.path.exists(pdf):
            raise ValueError("แปลงไฟล์เป็น PDF ไม่สำเร็จ")
        doc = fitz.open("pdf", open(pdf, "rb").read())
        shutil.rmtree(out_dir, ignore_errors=True)
        return doc
    raise ValueError(f"ยังไม่รองรับไฟล์ {ext}")


def _upload_to_doc():
    f = request.files.get("file")
    if not f or not f.filename:
        abort(400, "กรุณาเลือกไฟล์")
    tmp = tempfile.mkdtemp(prefix="up_")
    path = os.path.join(tmp, "in" + os.path.splitext(f.filename)[1].lower())
    f.save(path)
    try:
        doc = to_pdf(path, f.filename)
        if doc.needs_pass:
            pw = request.form.get("password", "")
            if not doc.authenticate(pw):
                abort(400, "ไฟล์นี้มีรหัสผ่าน กรุณาใส่รหัสผ่าน")
        return f.filename, fitz.open("pdf", doc.tobytes())
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _bake_rotation(doc, pno):
    """ทำให้หน้าที่หมุนแล้วกลายเป็นหน้าปกติ (rotation 0) เพื่อให้พิกัดที่แก้ไขตรงกับที่ตาเห็น"""
    if doc[pno].rotation == 0:
        return
    src = fitz.open()
    src.insert_pdf(doc, from_page=pno, to_page=pno)
    r = src[0].rect
    doc.delete_page(pno)
    np_ = doc.new_page(pno, width=r.width, height=r.height)
    np_.show_pdf_page(np_.rect, src, 0)


# ---------------------------------------------------------------------------
# หน้าเว็บ + API
# ---------------------------------------------------------------------------

@bp.get("/editor")
def editor_page():
    return render_template("editor.html")


@bp.errorhandler(400)
@bp.errorhandler(404)
def _err(e):
    return jsonify(error=getattr(e, "description", str(e))), e.code


@bp.post("/api/doc")
def open_doc():
    _cleanup()
    try:
        name, doc = _upload_to_doc()
    except ValueError as e:
        return jsonify(error=str(e)), 400
    doc_id = uuid.uuid4().hex
    os.makedirs(os.path.join(WORK_DIR, doc_id))
    with open(os.path.join(WORK_DIR, doc_id, "meta.json"), "w", encoding="utf-8") as f:
        json.dump({"name": os.path.splitext(name)[0], "version": 0, "max_version": 0}, f, ensure_ascii=False)
    for i in range(len(doc)):
        _bake_rotation(doc, i)
    doc.save(os.path.join(WORK_DIR, doc_id, "v0.pdf"), garbage=3, deflate=True)
    return jsonify(_info(doc_id))


@bp.get("/api/doc/<doc_id>")
def doc_info(doc_id):
    return jsonify(_info(doc_id))


IT_FONT = re.compile(r"IT[\?๙9]?(,|-|$)")


def _is_bold(sp):
    f = sp["font"].lower()
    return bool(sp["flags"] & 16) or any(k in f for k in ("bold", "black", "heavy", "semibold"))


def _is_italic(sp):
    f = sp["font"].lower()
    return bool(sp["flags"] & 2) or "italic" in f or "oblique" in f


def _page_lines(page):
    """บรรทัดข้อความในหน้า: ข้อความ (ซ่อมภาษาไทยแล้ว), ตำแหน่งเส้นฐาน, ส่วนตัวหนา/ธรรมดา และฟอนต์ที่ใกล้ที่สุด"""
    out = []
    for b in page.get_text("rawdict")["blocks"]:
        for ln in b.get("lines", []):
            d = ln.get("dir", (1, 0))
            if abs(d[0] - 1) > 0.01 or abs(d[1]) > 0.01:       # ข้ามข้อความที่หมุน
                continue
            chars = [ch for sp in ln["spans"] for ch in sp["chars"]]
            ink = [ch for ch in chars if not ch["c"].isspace()]
            if not ink:
                continue
            ink_box = fitz.Rect(ink[0]["bbox"])
            for ch in ink[1:]:
                ink_box |= fitz.Rect(ch["bbox"])
            first = next(sp for sp in ln["spans"] if any(not c["c"].isspace() for c in sp["chars"]))
            runs = []
            for sp in ln["spans"]:
                t = "".join(c["c"] for c in sp["chars"])
                if not t:
                    continue
                t = fix_thai(t)
                if IT_FONT.search(sp["font"]):                    # TH SarabunIT๙ แสดงเลขเป็นเลขไทย
                    t = t.translate(THAI_DIGITS)
                st = {"bold": _is_bold(sp), "italic": _is_italic(sp), "color": f"#{sp['color']:06x}",
                      "size": round(sp["size"], 2)}
                if runs and all(runs[-1][k] == v for k, v in st.items()):
                    runs[-1]["text"] += t
                else:
                    runs.append(dict(st, text=t))
            while runs and not runs[0]["text"].strip():
                runs.pop(0)
            while runs and not runs[-1]["text"].strip():
                runs.pop()
            if not runs:
                continue
            runs[0]["text"] = runs[0]["text"].lstrip()
            runs[-1]["text"] = runs[-1]["text"].rstrip()
            text = "".join(r["text"] for r in runs)
            fixed = fix_thai(text)
            if fixed != text:                                     # ซ่อมข้ามส่วน (เช่น "ส" | " าหรับ")
                runs = [dict(runs[0], text=fixed)]
                text = fixed
            size = first["size"]
            nb = sum(len(r["text"].strip()) for r in runs if r["bold"])
            bold = nb * 2 > len(text.replace(" ", ""))
            family = pick_family(first["font"], text, ink_box.width, size, bold)
            ours = sum(text_width(r["text"], r["size"], r["bold"], family) for r in runs)
            fit = ink_box.width / ours if ours > 0 and len(text) > 3 else 1.0
            lo, hi = (0.93, 1.07) if family == "sarabun" and SARABUN_LIKE.search(first["font"]) else (0.6, 1.6)
            fit = max(lo, min(hi, fit))
            out.append({
                "bbox": list(fitz.Rect(ln["bbox"]) | ink_box), "ink": list(ink_box),
                "baseline": ink[0]["origin"][1], "x": ink[0]["bbox"][0], "size": size, "fit": round(fit, 4),
                "family": family, "runs": runs, "text": text, "color": runs[0]["color"], "bold": bold,
                "font": re.sub(r"^[A-Z]{6}\+", "", first["font"]),
            })
    return out


def _remove_text(page, rects, fills=()):
    """ลบตัวอักษรเดิมออกจากไฟล์จริง (ไม่ใช่แค่ทับสีขาว)
    rects = กรอบบรรทัดที่จะแก้ (ลบเฉพาะตัวอักษร), fills = [(rect, สี)] สำหรับลบ/ปิดทับ"""
    for r in rects:
        r = fitz.Rect(r)
        pad = min(r.height * 0.2, 3)                 # กันไปโดนบรรทัดข้างเคียง
        page.add_redact_annot(fitz.Rect(r.x0 - 0.3, r.y0 + pad, r.x1 + 0.3, r.y1 - pad), fill=False)
    for r, color in fills:
        page.add_redact_annot(fitz.Rect(r), fill=_rgb(color))
    if rects or fills:
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE if not fills else fitz.PDF_REDACT_IMAGE_PIXELS,
                              graphics=fitz.PDF_REDACT_LINE_ART_NONE)
        for r, color in fills:
            page.draw_rect(fitz.Rect(r), color=None, fill=_rgb(color), overlay=True)


@bp.get("/api/doc/<doc_id>/page/<int:pno>.png")
def page_png(doc_id, pno):
    """ภาพหน้า (hide=1,5 = ซ่อนบรรทัดที่กำลังแก้ เพื่อแสดงตัวอย่างที่ตรงกับผลจริง)"""
    scale = min(max(float(request.args.get("scale", 1.5)), 0.1), 4)
    hide = [int(x) for x in request.args.get("hide", "").split(",") if x.strip().isdigit()]
    with fitz.open(_current(doc_id)) as doc:
        if not 0 <= pno < len(doc):
            abort(404)
        page = doc[pno]
        if hide:
            lines = _page_lines(page)
            _remove_text(page, [lines[i]["bbox"] for i in hide if i < len(lines)])
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
        data = pix.tobytes("png")
    resp = send_file(io.BytesIO(data), mimetype="image/png")
    resp.headers["Cache-Control"] = "private, max-age=3600"
    return resp


@bp.get("/api/doc/<doc_id>/lines/<int:pno>")
def page_lines(doc_id, pno):
    """บรรทัดข้อความในหน้า สำหรับคลิกแก้ไข"""
    with fitz.open(_current(doc_id)) as doc:
        if not 0 <= pno < len(doc):
            abort(404)
        return jsonify({"lines": _page_lines(doc[pno]), "metrics": {f: line_metrics(f) for f in ("sarabun", "noto")}})


def _rgb(hexcolor):
    h = (hexcolor or "#000000").lstrip("#")
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


def _runs_of(op, size_mult=1.0):
    base = {"bold": bool(op.get("bold")), "italic": bool(op.get("italic")),
            "underline": bool(op.get("underline")), "color": op.get("color") or "#000000",
            "size": float(op.get("size") or 14)}
    runs = op.get("runs") or [{"text": op.get("text", "")}]
    out = []
    for r in runs:
        rr = dict(base)
        rr.update({k: v for k, v in r.items() if v is not None})
        rr["size"] = float(rr.get("size") or base["size"]) * size_mult
        out.append(rr)
    return out


def _apply_ops(doc, ops):
    by_page = {}
    for op in ops:
        by_page.setdefault(int(op["page"]), []).append(op)
    for pno, pops in by_page.items():
        page = doc[pno]
        # 1) ลบตัวอักษรเดิมของบรรทัดที่แก้ และส่วนที่สั่งลบ/ปิดทับ
        _remove_text(page, [op["bbox"] for op in pops if op["type"] == "edit_text"],
                     [(op["rect"], op.get("color", "#ffffff")) for op in pops if op["type"] == "whiteout"])
        # 2) วาด/ใส่ของใหม่
        for op in pops:
            t = op["type"]
            if t == "edit_text":
                runs = _runs_of(op, float(op.get("fit", 1)))
                bb = fitz.Rect(op["bbox"])
                x = float(op.get("x", bb.x0))
                base = float(op.get("baseline", bb.y1 - (bb.height * 0.25 / 1.1)))
                if any(r["text"].strip() for r in runs):
                    add_runs_at_baseline(page, runs, x, base,
                                         float(op["size"]) * float(op.get("fit", 1)), op.get("family", "sarabun"),
                                         op.get("color", "#000000"), max_x=page.rect.width - 14)
            elif t == "text":
                runs = _runs_of(op)
                if any(r["text"].strip() for r in runs):
                    add_text_box(page, op["rect"], runs, float(op.get("size", 14)), op.get("family", "sarabun"),
                                 op.get("color", "#000000"), op.get("align", "left"))
            elif t == "image":
                data = op["data"].split(",", 1)[-1]
                page.insert_image(fitz.Rect(op["rect"]), stream=base64.b64decode(data), keep_proportion=False,
                                  overlay=True)
            elif t == "highlight":
                page.draw_rect(fitz.Rect(op["rect"]), color=None, fill=_rgb(op.get("color", "#ffeb3b")),
                               fill_opacity=float(op.get("opacity", 0.35)), overlay=True)
            elif t == "rect":
                page.draw_rect(fitz.Rect(op["rect"]), color=_rgb(op.get("color", "#e11d48")),
                               width=float(op.get("width", 1.5)), overlay=True)
            elif t == "line":
                p = op["points"]
                page.draw_line(fitz.Point(p[0]), fitz.Point(p[1]), color=_rgb(op.get("color", "#e11d48")),
                               width=float(op.get("width", 1.5)), overlay=True)
            elif t == "ink":
                shape = page.new_shape()
                for path in op["paths"]:
                    if len(path) > 1:
                        shape.draw_polyline([fitz.Point(x, y) for x, y in path])
                shape.finish(color=_rgb(op.get("color", "#1d4ed8")), width=float(op.get("width", 2)),
                             closePath=False, lineCap=1, lineJoin=1)
                shape.commit(overlay=True)


@bp.post("/api/doc/<doc_id>/apply")
def apply_edits(doc_id):
    ops = (request.get_json(silent=True) or {}).get("ops", [])
    if not ops:
        return jsonify(_info(doc_id))
    with fitz.open(_current(doc_id)) as doc:
        try:
            _apply_ops(doc, ops)
        except Exception as e:  # noqa: BLE001
            import traceback
            traceback.print_exc()
            return jsonify(error=f"บันทึกการแก้ไขไม่สำเร็จ: {e}"), 500
        _commit(doc_id, doc)
    return jsonify(_info(doc_id))


@bp.post("/api/doc/<doc_id>/pages")
def page_ops(doc_id):
    """จัดการหน้า: rotate / delete / move / duplicate / blank / order"""
    p = request.get_json(silent=True) or {}
    action = p.get("action")
    with fitz.open(_current(doc_id)) as doc:
        n = len(doc)
        if action == "rotate":
            for i in p.get("pages", []):
                doc[i].set_rotation((doc[i].rotation + int(p.get("deg", 90))) % 360)
                _bake_rotation(doc, i)
        elif action == "delete":
            pages = sorted(set(p.get("pages", [])))
            if len(pages) >= n:
                return jsonify(error="ลบทุกหน้าไม่ได้"), 400
            doc.delete_pages(pages)
        elif action == "order":
            order = p.get("order", [])
            if sorted(order) != list(range(n)):
                return jsonify(error="ลำดับหน้าไม่ถูกต้อง"), 400
            doc.select(order)
        elif action == "duplicate":
            i = int(p["page"])
            doc.fullcopy_page(i, i + 1)
        elif action == "blank":
            i = int(p.get("after", n - 1))
            r = doc[i].rect if n else fitz.paper_rect("a4")
            doc.new_page(i + 1, width=r.width, height=r.height)
        else:
            return jsonify(error="คำสั่งไม่ถูกต้อง"), 400
        _commit(doc_id, doc)
    return jsonify(_info(doc_id))


@bp.post("/api/doc/<doc_id>/merge")
def merge(doc_id):
    try:
        _, add = _upload_to_doc()
    except ValueError as e:
        return jsonify(error=str(e)), 400
    pos = request.form.get("position", "end")
    with fitz.open(_current(doc_id)) as doc:
        start = 0 if pos == "start" else len(doc)
        doc.insert_pdf(add, start_at=start)
        for i in range(start, start + len(add)):
            _bake_rotation(doc, i)
        _commit(doc_id, doc)
    return jsonify(_info(doc_id))


@bp.post("/api/doc/<doc_id>/watermark")
def watermark(doc_id):
    p = request.get_json(silent=True) or {}
    text = p.get("text", "").strip()
    if not text:
        return jsonify(error="กรุณาใส่ข้อความลายน้ำ"), 400
    size0 = float(p.get("size", 60))
    color = p.get("color", "#9ca3af")
    opacity = float(p.get("opacity", 0.3))
    angle = float(p.get("angle", 45)) if p.get("diagonal", True) else 0.0
    with fitz.open(_current(doc_id)) as doc:
        for page in doc:
            r = page.rect
            size = size0
            limit = (r.width ** 2 + r.height ** 2) ** 0.5 * 0.8 if angle else r.width * 0.9
            w = text_width(text, size, True)
            if w > limit:
                size *= limit / w
                w = limit
            c = fitz.Point(r.width / 2, r.height / 2)
            m = line_metrics("sarabun")
            h = size * (m["asc"] + m["desc"] + m["gap"])
            box = fitz.Rect(c.x - w / 2 - 4, c.y - h / 2, c.x + w / 2 + 4, c.y + h / 2)
            add_text_box(page, box, [{"text": text, "bold": True}], size, "sarabun", color, "center",
                         opacity=opacity, rotate=-angle, pivot=c)
        _commit(doc_id, doc)
    return jsonify(_info(doc_id))


@bp.post("/api/doc/<doc_id>/pagenumbers")
def page_numbers(doc_id):
    p = request.get_json(silent=True) or {}
    fmt = p.get("format", "{n}")
    pos = p.get("position", "bottom-center")
    size = float(p.get("size", 11))
    start = int(p.get("start", 1))
    thai = bool(p.get("thai_digits"))
    with fitz.open(_current(doc_id)) as doc:
        total = len(doc)
        m = line_metrics("sarabun")
        h = size * (m["asc"] + m["desc"] + m["gap"])
        for i, page in enumerate(doc):
            label = fmt.replace("{n}", str(i + start)).replace("{total}", str(total + start - 1))
            if thai:
                label = label.translate(THAI_DIGITS)
            r = page.rect
            y0 = 18 if pos.startswith("top") else r.height - 22 - h
            align = pos.split("-")[1]
            add_text_box(page, fitz.Rect(36, y0, r.width - 36, y0 + h + 2), [{"text": label}], size, "sarabun",
                         p.get("color", "#000000"), align)
        _commit(doc_id, doc)
    return jsonify(_info(doc_id))


@bp.post("/api/doc/<doc_id>/undo")
def undo(doc_id):
    m = _meta(doc_id)
    if m["version"] > 0:
        m["version"] -= 1
        _save_meta(doc_id, m)
    return jsonify(_info(doc_id))


@bp.post("/api/doc/<doc_id>/redo")
def redo(doc_id):
    m = _meta(doc_id)
    if m["version"] < m.get("max_version", m["version"]):
        m["version"] += 1
        _save_meta(doc_id, m)
    return jsonify(_info(doc_id))


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

def parse_pages(spec, n):
    """'1-3,5' -> [0,1,2,4] (ว่าง = ทุกหน้า)"""
    spec = (spec or "").replace(" ", "")
    if not spec:
        return list(range(n))
    out = []
    for part in spec.split(","):
        if "-" in part:
            a, b = part.split("-", 1)
            a = int(a) if a else 1
            b = int(b) if b else n
            out += list(range(a - 1, min(b, n)))
        elif part:
            out.append(int(part) - 1)
    out = [i for i in dict.fromkeys(out) if 0 <= i < n]
    if not out:
        raise ValueError("เลขหน้าไม่ถูกต้อง")
    return out


_XL_BAD = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _xl(v):
    return _XL_BAD.sub("", fix_thai(v)).replace("\n", " ").strip() if isinstance(v, str) else v


def _excel(doc, path):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font

    wb = Workbook()
    wb.remove(wb.active)
    th_font = Font(name="Tahoma", size=11)          # มีทั้งใน Mac และ Windows แสดงภาษาไทยได้
    for pno, page in enumerate(doc):
        tables = page.find_tables().tables
        for ti, tab in enumerate(tables):
            ws = wb.create_sheet(f"หน้า{pno + 1}-ตาราง{ti + 1}"[:31])
            # อ่านข้อความทีละช่องด้วย get_textbox (tab.extract() ทำสระล่าง ุ ู หลุดออกจากคำ)
            grid = [[_xl(page.get_textbox(fitz.Rect(c))) if c else None for c in row.cells] for row in tab.rows]
            keep = [j for j in range(max(map(len, grid), default=0))
                    if any(j < len(r) and r[j] for r in grid)]            # ตัดคอลัมน์ว่าง (จากช่องที่ผสาน)
            cols = [[r[j] if j < len(r) else None for r in grid] for j in keep]
            merged = []                                   # รวมคอลัมน์ข้างเคียงที่ไม่มีแถวไหนใช้ทั้งคู่
            for col in cols:
                if merged and not any(a and b for a, b in zip(merged[-1], col)):
                    merged[-1] = [a or b for a, b in zip(merged[-1], col)]
                else:
                    merged.append(col)
            for i in range(len(grid)):
                ws.append([col[i] for col in merged])
            for row in ws.iter_rows():
                for c in row:
                    c.font = th_font
                    c.alignment = Alignment(wrap_text=True, vertical="top")
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(
                    60, max(10, max(len(str(c.value or "")) for c in col) * 0.9))
        if not tables:
            ws = wb.create_sheet(f"หน้า{pno + 1}"[:31])
            for line in page.get_text().splitlines():
                if _xl(line):
                    ws.append([_xl(line)])
            ws.column_dimensions["A"].width = 100
            for row in ws.iter_rows():
                for c in row:
                    c.font = th_font
    wb.save(path)


@bp.get("/api/doc/<doc_id>/export")
def export(doc_id):
    fmt = request.args.get("fmt", "pdf")
    m = _meta(doc_id)
    name = m["name"]
    with fitz.open(_current(doc_id)) as full:
        try:
            pages = parse_pages(request.args.get("pages"), len(full))
        except ValueError as e:
            return jsonify(error=str(e)), 400
        doc = fitz.open()
        doc.insert_pdf(full)
        if len(pages) != len(full):
            doc.select(pages)
            name += f"_หน้า{request.args.get('pages')}"
    tmp = tempfile.mkdtemp(prefix="export_")
    try:
        if fmt == "pdf":
            out = os.path.join(tmp, "out.pdf")
            doc.save(out, garbage=4, deflate=True, clean=True)
            return send_file(out, as_attachment=True, download_name=name + ".pdf")
        if fmt == "docx":
            src = os.path.join(tmp, "in.pdf")
            doc.save(src)
            out = os.path.join(tmp, "out.docx")
            convert(src, out, request.args.get("mode", "auto"), request.args.get("font", DEFAULT_FONT))
            return send_file(out, as_attachment=True, download_name=name + ".docx")
        if fmt == "xlsx":
            out = os.path.join(tmp, "out.xlsx")
            _excel(doc, out)
            return send_file(out, as_attachment=True, download_name=name + ".xlsx")
        if fmt == "txt":
            text = "\n\n".join(f"===== หน้า {pages[i] + 1} =====\n" + fix_thai(p.get_text())
                               for i, p in enumerate(doc))
            return send_file(io.BytesIO(text.encode("utf-8")), mimetype="text/plain; charset=utf-8",
                             as_attachment=True, download_name=name + ".txt")
        if fmt in ("png", "jpg"):
            dpi = min(max(int(request.args.get("dpi", 150)), 50), 600)
            images = []
            for i, p in enumerate(doc):
                pix = p.get_pixmap(dpi=dpi, alpha=False)
                images.append((f"{name}_หน้า{pages[i] + 1}.{fmt}",
                               pix.tobytes("png") if fmt == "png" else pix.tobytes("jpg", jpg_quality=90)))
            if len(images) == 1:
                return send_file(io.BytesIO(images[0][1]), mimetype=f"image/{'png' if fmt == 'png' else 'jpeg'}",
                                 as_attachment=True, download_name=images[0][0])
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                for fn, data in images:
                    z.writestr(fn, data)
            buf.seek(0)
            return send_file(buf, mimetype="application/zip", as_attachment=True,
                             download_name=f"{name}_{fmt}.zip")
        return jsonify(error="ไม่รองรับรูปแบบนี้"), 400
    finally:
        doc.close()
