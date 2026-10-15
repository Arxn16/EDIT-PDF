"""ซ่อมไฟล์ .docx ให้ Word เปิดได้ถูกต้องทุกเครื่อง

1. sanitize_ooxml(): pdf2docx เขียนค่าที่ไม่ตรงมาตรฐาน OOXML เช่น ขนาดเส้นขอบเป็นทศนิยม (w:sz="3.2")
   และสีแบบ "#000000" — LibreOffice ยอมรับ แต่ Word ทิ้งเส้นขอบพวกนั้น (ตารางขาดเส้น)
2. normalize_ooxml(): เรียงลำดับ element ตามมาตรฐาน (Word ข้าม/ทิ้งค่าที่อยู่ผิดลำดับ เช่น tblInd หลัง tblLook),
   รวม element ที่ซ้ำ (เช่น tcMar 2 ตัวในช่องเดียว) และเปลี่ยน w:start/w:end เป็น w:left/w:right
   ซึ่งเป็นแบบที่ Word เขียนเองในไฟล์ .docx ปกติ (เส้นขอบซ้าย/ขวาของช่องตารางไม่หายในบางเวอร์ชัน)
3. fix_table_outdent(): pdf2docx คำนวณตำแหน่งตารางแบบ Word 2013+ แต่ไฟล์เป็นโหมด Word 2010
   ซึ่งดึงตารางออกไปทางซ้ายเท่าระยะขอบช่อง (5.4pt) -> ตั้งระยะขอบช่องของตารางเป็น 0 (ช่องยังเว้นเท่าเดิม)
4. embed_fonts(): ฝังฟอนต์ไทยลงในไฟล์ .docx (มาตรฐาน obfuscated font ของ Word)
   เปิดที่เครื่องที่ไม่มีฟอนต์ก็ยังแสดงถูก ไม่ขึ้น ⚠ และไม่ล้นตาราง
"""
import io
import os
import re
import uuid
import zipfile
from collections import Counter
from functools import lru_cache

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_FONT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
CT_ODTTF = "application/vnd.openxmlformats-officedocument.obfuscatedFont"
HERE = os.path.dirname(os.path.abspath(__file__))
FONT_DIRS = [os.path.join(HERE, "static", "fonts"), "/usr/share/fonts/truetype/user",
             "/usr/share/fonts/truetype/thai-compat"] + \
    [d for d in os.environ.get("EXTRA_FONT_DIRS", "").split(os.pathsep) if d]

_DEC = re.compile(r"^-?\d+\.\d+$")
_BORDER = {f"{{{W}}}{t}" for t in ("top", "left", "bottom", "right", "start", "end", "insideH", "insideV",
                                    "tl2br", "tr2bl", "between", "bar")}


def w(tag):
    return f"{{{W}}}{tag}"


# ---------------------------------------------------------------------------
# 1) ค่าให้ตรงมาตรฐาน
# ---------------------------------------------------------------------------

def sanitize_ooxml(element):
    """แก้ค่าที่ Word อ่านไม่ได้ คืนจำนวนที่แก้"""
    fixed = 0
    for e in element.iter():
        if not isinstance(e.tag, str):
            continue
        for k, v in list(e.attrib.items()):
            if not k.startswith(f"{{{W}}}"):
                continue
            local = k[len(W) + 2:]
            nv = v
            if _DEC.match(nv):                                  # ทศนิยม -> จำนวนเต็ม
                nv = str(int(round(float(nv))))
            if local in ("color", "fill") and nv.startswith("#"):
                nv = nv[1:].upper()
            if local == "sz" and e.tag in _BORDER:              # ขนาดเส้นขอบ 1/8 pt: Word รับ 2..96
                try:
                    nv = str(min(96, max(2, int(nv))))
                except ValueError:
                    pass
            if nv != v:
                e.set(k, nv)
                fixed += 1
        if e.tag == f"{{{W}}}shd":                              # w:val จำเป็นตามมาตรฐาน (ไม่มี = Word ไม่แสดงสีพื้น)
            if e.get(w("val")) is None:
                e.set(w("val"), "clear")
                fixed += 1
            if e.get(w("color")) is None:
                e.set(w("color"), "auto")
            f = e.get(w("fill"))
            if f and f.lower() != "auto":
                e.set(w("fill"), f.upper())
    return fixed


# ---------------------------------------------------------------------------
# 2) ลำดับ element ตามมาตรฐาน + ชื่อแบบที่ Word ใช้
# ---------------------------------------------------------------------------

_SEQ = {
    "pPr": ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
            "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
            "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
            "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
            "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"],
    "rPr": ["ins", "del", "moveFrom", "moveTo", "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
            "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish",
            "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr",
            "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath",
            "rPrChange"],
    "tblPr": ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW",
              "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
              "tblCaption", "tblDescription", "tblPrChange"],
    "tblPrEx": ["tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
                "tblPrExChange"],
    "tcPr": ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
             "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel", "cellMerge",
             "tcPrChange"],
    "sectPr": ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc",
               "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg",
               "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"],
    "tcBorders": ["top", "left", "start", "bottom", "right", "end", "insideH", "insideV", "tl2br", "tr2bl"],
    "tblBorders": ["top", "left", "start", "bottom", "right", "end", "insideH", "insideV"],
    "pBdr": ["top", "left", "bottom", "right", "between", "bar"],
    "pgBorders": ["top", "left", "bottom", "right"],
    "tcMar": ["top", "left", "start", "bottom", "right", "end"],
    "tblCellMar": ["top", "left", "start", "bottom", "right", "end"],
}
_MULTI = {"sectPr": {"headerReference", "footerReference"}}     # ซ้ำได้ตามมาตรฐาน
_LTR_NAMES = {"start": "left", "end": "right"}
_RENAME_IN = {"tcBorders", "tblBorders", "tcMar", "tblCellMar"}
_IND_ATTRS = (("start", "left"), ("end", "right"), ("startChars", "leftChars"), ("endChars", "rightChars"))
_PROPS_FIRST = {"p": ("pPr",), "r": ("rPr",), "tc": ("tcPr",), "tr": ("tblPrEx", "trPr")}


def _local(e):
    """ชื่อ element ใน namespace ของ Word (ไม่ใช่ของ Word = None)"""
    tag = e.tag
    if isinstance(tag, str) and tag.startswith(f"{{{W}}}"):
        return tag[len(W) + 2:]
    return None


def _merge_duplicates(e, multi=()):
    fixed, seen = 0, {}
    for c in list(e):
        name = _local(c)
        if name is None or name in multi:
            continue
        base = seen.get(name)
        if base is None:
            seen[name] = c
            continue
        for k, v in c.attrib.items():              # ตัวหลังทับตัวแรก (เหมือนที่โปรแกรมอ่านทีละตัว)
            base.set(k, v)
        for cc in list(c):
            old = base.find(cc.tag)
            if old is not None:
                base.remove(old)
            base.append(cc)
        e.remove(c)
        fixed += 1
    return fixed


def _reorder(e, seq):
    order = {n: i for i, n in enumerate(seq)}
    kids = list(e)
    keyed = sorted(enumerate(kids), key=lambda ic: (order.get(_local(ic[1]), len(seq)), ic[0]))
    new = [c for _, c in keyed]
    if new == kids:
        return 0
    for c in kids:
        e.remove(c)
    for c in new:
        e.append(c)
    return 1


def normalize_ooxml(root):
    """เรียงลำดับ/รวมซ้ำ/เปลี่ยนชื่อ start-end ให้เป็นแบบที่ Word อ่านได้แน่นอน คืนจำนวนจุดที่แก้"""
    fixed = 0
    for e in list(root.iter()):
        name = _local(e)
        if name is None:
            continue
        if name in _RENAME_IN:
            for c in list(e):
                cn = _local(c)
                if cn in _LTR_NAMES:
                    if e.find(w(_LTR_NAMES[cn])) is None:
                        c.tag = w(_LTR_NAMES[cn])
                    else:
                        e.remove(c)
                    fixed += 1
        elif name == "ind":
            for a, b in _IND_ATTRS:
                v = e.get(w(a))
                if v is not None:
                    if e.get(w(b)) is None:
                        e.set(w(b), v)
                    del e.attrib[w(a)]
                    fixed += 1
        elif name == "jc" and e.get(w("val")) in _LTR_NAMES:
            e.set(w("val"), _LTR_NAMES[e.get(w("val"))])
            fixed += 1
        seq = _SEQ.get(name)
        if seq:
            fixed += _merge_duplicates(e, _MULTI.get(name, ()))
            fixed += _reorder(e, seq)
        first = _PROPS_FIRST.get(name)
        if first:
            props = [c for n in first for c in e.findall(w(n))]
            kids = [c for c in e if isinstance(c.tag, str)]
            if props and kids[:len(props)] != props:
                for i, c in enumerate(props):
                    e.remove(c)
                    e.insert(i, c)
                fixed += 1
    return fixed


def ooxml_issues(root):
    """ตรวจสิ่งที่ Word อาจอ่านผิด (ใช้ในชุดทดสอบ)"""
    bad = Counter()
    for e in root.iter():
        name = _local(e)
        if name is None:
            continue
        seq = _SEQ.get(name)
        if seq:
            names = [_local(c) for c in e if _local(c)]
            pos = [seq.index(n) for n in names if n in seq]
            if pos != sorted(pos):
                bad["ลำดับผิดใน " + name] += 1
            multi = _MULTI.get(name, ())
            if any(names.count(n) > 1 for n in set(names) if n not in multi):
                bad["ซ้ำใน " + name] += 1
        if name in _RENAME_IN and any(_local(c) in _LTR_NAMES for c in e):
            bad["start/end ใน " + name] += 1
        if name == "ind" and any(e.get(w(a)) is not None for a, _ in _IND_ATTRS):
            bad["start/end ใน ind"] += 1
    return bad


def _margin(cm, *names):
    for n in names:
        el = cm.find(w(n)) if cm is not None else None
        if el is not None and el.get(w("type")) == "nil":
            return 0
        if el is not None and (el.get(w("type")) in (None, "dxa")):
            try:
                return int(float(el.get(w("w"), "0")))
            except ValueError:
                return None
    return None


def _style_cell_margins(styles_root, style_id=None):
    """ระยะขอบช่องซ้าย/ขวาจากสไตล์ตาราง (ไล่ตาม basedOn) ไม่พบ = ค่าเริ่มต้นของ Word 108 twips"""
    left = right = None
    if styles_root is not None:
        styles = {s.get(w("styleId")): s for s in styles_root.findall(w("style")) if s.get(w("type")) == "table"}
        sid = style_id or next((k for k, s in styles.items() if s.get(w("default")) in ("1", "true", "on")), None)
        seen = set()
        while sid and sid in styles and sid not in seen:
            seen.add(sid)
            cm = styles[sid].find(f"{w('tblPr')}/{w('tblCellMar')}")
            if left is None:
                left = _margin(cm, "left", "start")
            if right is None:
                right = _margin(cm, "right", "end")
            based = styles[sid].find(w("basedOn"))
            sid = based.get(w("val")) if based is not None else None
    return (108 if left is None else left), (108 if right is None else right)


def _set_margin(cm, side, value):
    for n in (side, "start" if side == "left" else "end"):
        old = cm.find(w(n))
        if old is not None:
            cm.remove(old)
    el = etree.SubElement(cm, w(side))
    el.set(w("w"), str(value))
    el.set(w("type"), "dxa")


def fix_table_outdent(root, styles_root=None):
    """ตารางให้อยู่ตำแหน่งเดียวกับใน PDF: ตั้งระยะขอบช่องเริ่มต้นของตารางเป็น 0
    (Word 2010 ดึงตารางออกซ้ายเท่าระยะนี้) ช่องที่ไม่ได้กำหนดระยะเองจะได้ค่าเดิมใส่ไว้ให้ ข้อความในช่องไม่ขยับ"""
    fixed = 0
    for tbl in list(root.iter(w("tbl"))):
        tblPr = tbl.find(w("tblPr"))
        if tblPr is None:
            tblPr = etree.Element(w("tblPr"))
            tbl.insert(0, tblPr)
        st = tblPr.find(w("tblStyle"))
        dl, dr = _style_cell_margins(styles_root, st.get(w("val")) if st is not None else None)
        cm = tblPr.find(w("tblCellMar"))
        v = _margin(cm, "left", "start")
        dl = dl if v is None else v
        v = _margin(cm, "right", "end")
        dr = dr if v is None else v
        if dl == 0 and dr == 0:
            continue
        for tc in tbl.iter(w("tc")):
            if next(tc.iterancestors(w("tbl")), None) is not tbl:       # ช่องของตารางซ้อน
                continue
            tcPr = tc.find(w("tcPr"))
            if tcPr is None:
                tcPr = etree.Element(w("tcPr"))
                tc.insert(0, tcPr)
            mar = tcPr.find(w("tcMar"))
            if mar is None:
                mar = etree.SubElement(tcPr, w("tcMar"))
            if _margin(mar, "left", "start") is None:
                _set_margin(mar, "left", dl)
            if _margin(mar, "right", "end") is None:
                _set_margin(mar, "right", dr)
        if cm is None:
            cm = etree.SubElement(tblPr, w("tblCellMar"))
        _set_margin(cm, "left", 0)
        _set_margin(cm, "right", 0)
        fixed += 1
    return fixed


# ---------------------------------------------------------------------------
# 3) ฝังฟอนต์
# ---------------------------------------------------------------------------

@lru_cache(maxsize=1)
def font_registry():
    """ชื่อฟอนต์ -> {regular/bold/italic/bolditalic: path} จากโฟลเดอร์ฟอนต์ของระบบนี้"""
    from fontTools.ttLib import TTFont
    reg = {}
    for d in FONT_DIRS:
        if not os.path.isdir(d):
            continue
        for fn in sorted(os.listdir(d)):
            if not fn.lower().endswith((".ttf", ".otf")):
                continue
            path = os.path.join(d, fn)
            try:
                f = TTFont(path, lazy=True)
                names = {r.nameID: r.toUnicode() for r in f["name"].names if r.nameID in (1, 2, 16, 17)}
                fam = names.get(16) or names.get(1)
                os2 = f["OS/2"]
                bold = bool(os2.fsSelection & 0x20) or os2.usWeightClass >= 600
                italic = bool(os2.fsSelection & 0x01)
                if os2.fsType & 0x0002:                        # ห้ามฝัง (restricted license)
                    continue
                if fn.lower().endswith(".otf") and "glyf" not in f:
                    continue                                   # Word ฝังได้เฉพาะ TrueType
            except Exception:  # noqa: BLE001
                continue
            style = ("bold" if bold else "") + ("italic" if italic else "") or "regular"
            reg.setdefault(fam, {}).setdefault(style, path)
    return reg


def _obfuscate(data, key_guid):
    """ECMA-376 17.8.1: XOR 32 ไบต์แรกของฟอนต์กับ GUID (อ่านจากขวาไปซ้าย)"""
    key = bytes.fromhex(key_guid.strip("{}").replace("-", ""))[::-1]
    b = bytearray(data)
    for i in range(32):
        b[i] ^= key[i % 16]
    return bytes(b)


def _used_styles(document_xml):
    x = document_xml
    bold = bool(re.search(r"<w:bCs?(?:\s+w:val=\"(?:1|true|on)\")?\s*/>", x))
    italic = bool(re.search(r"<w:iCs?(?:\s+w:val=\"(?:1|true|on)\")?\s*/>", x))
    return {"regular": True, "bold": bold, "italic": italic, "bolditalic": bold and italic}


def _used_fonts(document_xml):
    return set(re.findall(r'w:(?:ascii|hAnsi|cs|eastAsia)="([^"]+)"', document_xml))


SETTINGS_BEFORE = ["writeProtection", "view", "zoom", "removePersonalInformation", "removeDateAndTime",
                   "doNotDisplayPageBoundaries", "displayBackgroundShape", "printPostScriptOverText",
                   "printFractionalCharacterWidth", "printFormsData"]
FONT_CHILD_ORDER = ["altName", "panose1", "charset", "family", "notTrueType", "pitch", "sig",
                    "embedRegular", "embedBold", "embedItalic", "embedBoldItalic"]


def embed_fonts(docx_path, families=None):
    """ฝังฟอนต์ที่ใช้ในเอกสาร (ที่มีไฟล์อยู่ในระบบ) คืนรายชื่อฟอนต์ที่ฝัง"""
    reg = font_registry()
    with zipfile.ZipFile(docx_path) as z:
        items = {n: z.read(n) for n in z.namelist()}
    doc_xml = items.get("word/document.xml", b"").decode("utf8", "ignore")
    wanted = set(families) if families else _used_fonts(doc_xml)
    targets = [f for f in wanted if f in reg]
    if not targets or "word/fontTable.xml" not in items:
        return []
    styles = _used_styles(doc_xml)

    ft = etree.fromstring(items["word/fontTable.xml"])
    rels_name = "word/_rels/fontTable.xml.rels"
    if rels_name in items:
        rels = etree.fromstring(items[rels_name])
    else:
        rels = etree.fromstring(b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
    rel_ns = rels.nsmap.get(None, "http://schemas.openxmlformats.org/package/2006/relationships")
    used_ids = {r.get("Id") for r in rels}
    n_font = sum(1 for k in items if k.startswith("word/fonts/")) + 1
    embedded = []
    for fam in sorted(targets):
        node = next((f for f in ft.findall(w("font")) if f.get(w("name")) == fam), None)
        if node is None:
            node = etree.SubElement(ft, w("font"))
            node.set(w("name"), fam)
            for tag, val in (("charset", "DE"), ("family", "auto"), ("pitch", "variable")):
                etree.SubElement(node, w(tag)).set(w("val"), val)
        for child in list(node):
            if etree.QName(child).localname.startswith("embed"):
                node.remove(child)
        for style, tag in (("regular", "embedRegular"), ("bold", "embedBold"),
                           ("italic", "embedItalic"), ("bolditalic", "embedBoldItalic")):
            path = reg[fam].get(style)
            if not path or not styles.get(style):
                continue
            key = "{" + str(uuid.uuid4()).upper() + "}"
            part = f"fonts/font{n_font}.odttf"
            n_font += 1
            items["word/" + part] = _obfuscate(open(path, "rb").read(), key)
            rid = next(f"rIdF{i}" for i in range(1, 999) if f"rIdF{i}" not in used_ids)
            used_ids.add(rid)
            rel = etree.SubElement(rels, f"{{{rel_ns}}}Relationship")
            rel.set("Id", rid)
            rel.set("Type", REL_FONT)
            rel.set("Target", part)
            emb = etree.SubElement(node, w(tag))
            emb.set(f"{{{R}}}id", rid)
            emb.set(w("fontKey"), key)
        # เรียงลำดับลูกตามที่มาตรฐานกำหนด
        kids = sorted(list(node), key=lambda c: FONT_CHILD_ORDER.index(etree.QName(c).localname)
                      if etree.QName(c).localname in FONT_CHILD_ORDER else 99)
        for c in kids:
            node.remove(c)
            node.append(c)
        embedded.append(fam)
    if not embedded:
        return []
    items["word/fontTable.xml"] = etree.tostring(ft, xml_declaration=True, encoding="UTF-8", standalone=True)
    items[rels_name] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)

    # [Content_Types].xml: ชนิดไฟล์ .odttf
    ct = etree.fromstring(items["[Content_Types].xml"])
    ct_ns = ct.nsmap.get(None)
    if not any(d.get("Extension", "").lower() == "odttf" for d in ct.findall(f"{{{ct_ns}}}Default")):
        d = etree.Element(f"{{{ct_ns}}}Default")
        d.set("Extension", "odttf")
        d.set("ContentType", CT_ODTTF)
        ct.insert(0, d)
    items["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)

    # settings.xml: <w:embedTrueTypeFonts/> ต้องอยู่ถูกลำดับ
    if "word/settings.xml" in items:
        st = etree.fromstring(items["word/settings.xml"])
        if st.find(w("embedTrueTypeFonts")) is None:
            pos = 0
            for i, c in enumerate(st):
                if etree.QName(c).localname in SETTINGS_BEFORE:
                    pos = i + 1
            st.insert(pos, etree.Element(w("embedTrueTypeFonts")))
        items["word/settings.xml"] = etree.tostring(st, xml_declaration=True, encoding="UTF-8", standalone=True)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", items.pop("[Content_Types].xml"))
        for n, data in items.items():
            z.writestr(n, data)
    with open(docx_path, "wb") as f:
        f.write(buf.getvalue())
    return embedded
