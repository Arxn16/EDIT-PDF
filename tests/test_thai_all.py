"""ทดสอบตัวซ่อมภาษาไทยกับทุกพยัญชนะ ทุกสระ ทุกวรรณยุกต์ และทุกแบบความเสียหาย

รัน:  python tests/test_thai_all.py
     (ใน Docker: docker compose exec pdf2word python tests/test_thai_all.py)
"""
import os
import sys
from collections import Counter

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from converter import fix_thai  # noqa: E402
from thai_corpus import CORRUPTIONS, build_words  # noqa: E402

SPECIAL_SENTENCES = [
    "คู่มือสำหรับประชาชน : การขอรับใบอนุญาตขับรถชนิดชั่วคราว (พ.ร.บ.รถยนต์ พ.ศ. 2522)",
    "(1) มีอายุไม่ต่ำกว่า 18 ปีบริบูรณ์ แต่ถ้าเป็นผู้ขอใบอนุญาตขับรถจักรยานยนต์ส่วนบุคคลชั่วคราว",
    "(5) ไม่มีโรคประจำตัวที่ผู้ประกอบวิชาชีพเวชกรรมเห็นว่าอาจเป็นอันตรายขณะขับรถ",
    "(6) ไม่เป็นบุคคลวิกลจริตหรือจิตฟั่นเฟือน",
    "บริษัท อี.เอส.เอ็ม.โซลูชั่น จำกัด ขอเสนอราคาระบบ ERP และ Cloud Server รวมเป็นเงิน 1,250,000 บาท",
    "เจ้าหน้าที่จะตรวจสอบเอกสารหลักฐาน ทดสอบสมรรถภาพร่างกาย ได้แก่ ตาบอดสี สายตาทางลึก",
]


def main():
    words = build_words()
    print(f"คำทดสอบ {len(words):,} คำ (พยัญชนะ 44 ตัว + อักษรควบ x ทุกรูปสระ x ทุกวรรณยุกต์ + คำพิเศษ)\n")
    total_fail = 0
    for name, fn in CORRUPTIONS.items():
        fails = []
        tested = 0
        for w in words:
            bad = fn(w)
            if bad == w:
                continue
            tested += 1
            got = fix_thai(bad)
            if got != w:
                fails.append((w, bad, got))
        total_fail += len(fails)
        status = "ผ่าน" if not fails else f"พลาด {len(fails)}"
        print(f"  {name:<28} ทดสอบ {tested:>6,} คำ  -> {status}")
        for w, bad, got in fails[:8]:
            print(f"      {w!r:<14} เสียเป็น {bad!r:<22} ซ่อมได้ {got!r}")

    # คำปกติต้องไม่ถูกแก้
    untouched = [w for w in words if fix_thai(w) != w]
    print(f"\n  {'คำปกติต้องไม่ถูกเปลี่ยน':<28} ทดสอบ {len(words):>6,} คำ  -> "
          f"{'ผ่าน' if not untouched else 'พลาด ' + str(len(untouched))}")
    for w in untouched[:8]:
        print(f"      {w!r} -> {fix_thai(w)!r}")
    total_fail += len(untouched)

    # ข้อความยาวที่มีคำต่อกันหลายคำ + รหัสไทยเก่า TIS-620
    chunks = [" ".join(words[i:i + 25]) for i in range(0, len(words), 25)]
    moj_fail = sum(1 for c in chunks if fix_thai(c.encode("cp874").decode("latin-1")) != c)
    print(f"  {'รหัสไทยเก่า TIS-620':<28} ทดสอบ {len(chunks):>6,} ย่อหน้า -> "
          f"{'ผ่าน' if not moj_fail else 'พลาด ' + str(moj_fail)}")
    total_fail += moj_fail

    # แก้คำผิดจาก OCR ด้วยพจนานุกรม
    try:
        from thai_ocr_fix import fix_ocr_line
        ocr_cases = {
            "การขอรับใบอนุญาตขับรถชนิดชัวคราว": "การขอรับใบอนุญาตขับรถชนิดชั่วคราว",
            "หลกเกณฑ์": "หลักเกณฑ์", "รถยนต": "รถยนต์", "ปีบริบรณ์": "ปีบริบูรณ์",
            "จากนันอบรม": "จากนั้นอบรม", "เป็นทิเห็น": "เป็นที่เห็น",
            "บุคค ลวิกลจริต": "บุคคลวิกลจริต", "ว่าดัว ยการจราจร": "ว่าด้วยการจราจร",
        }
        keep = [w for w in SPECIAL_SENTENCES] + ["ทำ งาน", "ก็ ได้", "พ.ร.บ. พ.ศ. 2522", "สำนักงาน ขนส่งจังหวัด"]
        o_fail = [(a, fix_ocr_line(a)[0]) for a, b in ocr_cases.items() if fix_ocr_line(a)[0] != b]
        k_fail = [(a, fix_ocr_line(a)[0]) for a in keep if fix_ocr_line(a)[0] != a]
        print(f"  {'OCR: แก้คำผิดด้วยพจนานุกรม':<28} ทดสอบ {len(ocr_cases):>6,} คำ  -> "
              f"{'ผ่าน' if not o_fail else 'พลาด ' + str(len(o_fail))}")
        print(f"  {'OCR: ข้อความถูกต้องไม่ถูกแก้':<28} ทดสอบ {len(keep):>6,} ประโยค -> "
              f"{'ผ่าน' if not k_fail else 'พลาด ' + str(len(k_fail))}")
        for a, g in (o_fail + k_fail)[:8]:
            print(f"      {a!r} -> {g!r}")
        total_fail += len(o_fail) + len(k_fail)
    except ImportError:
        print("  (ข้ามการทดสอบ OCR: ไม่ได้ติดตั้ง pythainlp)")

    print("\nสรุป:", "ผ่านทั้งหมด ✔" if total_fail == 0 else f"พลาด {total_fail} กรณี")
    return 1 if total_fail else 0


if __name__ == "__main__":
    sys.exit(main())
