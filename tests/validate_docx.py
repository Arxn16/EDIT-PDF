"""ตรวจไฟล์ .docx กับ XML Schema มาตรฐาน OOXML (ECMA-376 Transitional) แบบเดียวกับที่ Word ตรวจ

Word จะข้ามหรือทิ้งค่าที่ผิดมาตรฐาน (เช่น เส้นขอบตารางที่เรียงผิดลำดับ) โดยไม่เตือน
สคริปต์นี้บอกได้ว่าไฟล์มีจุดแบบนั้นหรือไม่ (0 = Word อ่านได้ครบทุกอย่าง)

รัน:  python tests/validate_docx.py ไฟล์.docx [ไฟล์2.docx ...]
ครั้งแรกจะดาวน์โหลดไฟล์ schema (~400 KB) มาเก็บไว้ที่ tests/.xsd/
"""
import collections
import os
import re
import sys
import urllib.request
import zipfile

from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
XSD_DIR = os.path.join(HERE, ".xsd")
XSD_URL = "https://raw.githubusercontent.com/python-openxml/python-docx/master/ref/xsd/"
XSD_FILES = ["wml.xsd", "dml-main.xsd", "dml-picture.xsd", "dml-wordprocessingDrawing.xsd", "dml-chart.xsd",
             "dml-chartDrawing.xsd", "dml-diagram.xsd", "dml-lockedCanvas.xsd", "shared-commonSimpleTypes.xsd",
             "shared-customXmlSchemaProperties.xsd", "shared-math.xsd", "shared-relationshipReference.xsd"]
XML_XSD = """<?xml version="1.0"?>
<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema" targetNamespace="http://www.w3.org/XML/1998/namespace">
  <xs:attribute name="space"><xs:simpleType><xs:restriction base="xs:NCName">
    <xs:enumeration value="default"/><xs:enumeration value="preserve"/></xs:restriction></xs:simpleType></xs:attribute>
  <xs:attribute name="lang" type="xs:language"/>
</xs:schema>"""
ALL_XSD = """<?xml version="1.0"?>
<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema">
  <xs:import namespace="http://www.w3.org/XML/1998/namespace" schemaLocation="xml.xsd"/>
  <xs:import namespace="http://schemas.openxmlformats.org/wordprocessingml/2006/main" schemaLocation="wml.xsd"/>
</xs:schema>"""
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
KNOWN = {
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "http://schemas.openxmlformats.org/drawingml/2006/main",
    "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "http://www.w3.org/XML/1998/namespace",
    "http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes",
}
PARTS = re.compile(r"word/(document|styles|settings|fontTable|numbering|header\d*|footer\d*)\.xml$")
_schema = None


def schema():
    global _schema
    if _schema is None:
        os.makedirs(XSD_DIR, exist_ok=True)
        for f in XSD_FILES:
            path = os.path.join(XSD_DIR, f)
            if not os.path.exists(path):
                urllib.request.urlretrieve(XSD_URL + f, path)
        for name, text in (("xml.xsd", XML_XSD), ("all.xsd", ALL_XSD)):
            with open(os.path.join(XSD_DIR, name), "w", encoding="utf8") as fh:
                fh.write(text)
        _schema = etree.XMLSchema(etree.parse(os.path.join(XSD_DIR, "all.xsd")))
    return _schema


def _ns(tag):
    return tag[1:].split("}")[0] if tag.startswith("{") else None


def _prepare(root):
    """ทำแบบที่ Word ทำก่อนอ่าน: เลือก mc:Choice และตัดส่วนขยายที่ Word เวอร์ชันใหม่เพิ่ม (w14, wps, ...)"""
    for ac in list(root.iter(f"{{{MC}}}AlternateContent")):
        parent = ac.getparent()
        if parent is None:
            continue
        idx = parent.index(ac)
        choice = ac.find(f"{{{MC}}}Choice")
        for c in reversed(list(choice) if choice is not None else []):
            parent.insert(idx, c)
        parent.remove(ac)
    for e in list(root.iter()):
        if not isinstance(e.tag, str):
            continue
        if _ns(e.tag) not in KNOWN and e.getparent() is not None:
            e.getparent().remove(e)
            continue
        for a in list(e.attrib):
            if a.startswith("{") and _ns(a) not in KNOWN:
                del e.attrib[a]
    return root


def validate(path):
    """คืน Counter ของข้อผิดพลาด (ว่าง = ถูกต้องตามมาตรฐาน)"""
    sch = schema()
    errs = collections.Counter()
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if not PARTS.match(name):
                continue
            root = _prepare(etree.fromstring(z.read(name)))
            if sch.validate(etree.ElementTree(root)):
                continue
            for e in sch.error_log:
                msg = e.message.replace("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}", "w:")
                errs[name.split("/")[-1] + ": " + msg[:170]] += 1
    return errs


if __name__ == "__main__":
    bad = 0
    for p in sys.argv[1:]:
        errs = validate(p)
        bad += sum(errs.values())
        print(f"== {os.path.basename(p)}: {'ถูกต้องตามมาตรฐาน' if not errs else str(sum(errs.values())) + ' จุดผิด'}")
        for k, v in errs.most_common(15):
            print(f"   {v:5d}  {k}")
    sys.exit(1 if bad else 0)
