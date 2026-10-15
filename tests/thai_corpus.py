"""สร้างชุดทดสอบภาษาไทยครบทุกตัว: พยัญชนะ 44 ตัว + อักษรควบ x รูปสระทุกแบบ x วรรณยุกต์ทุกตัว
และจำลองความเสียหายทุกแบบที่เกิดเวลาดึงข้อความจาก PDF
"""

CONSONANTS = [chr(c) for c in range(0x0E01, 0x0E2F) if chr(c) not in "ฤฦ"]  # 44 ตัว
assert len(CONSONANTS) == 44

CLUSTERS = ["กร", "กล", "กว", "ขร", "ขล", "ขว", "คร", "คล", "คว", "ตร", "ปร", "ปล", "พร", "พล",
            "ผล", "ทร", "จร", "ซร", "สร", "ศร", "บร", "บล", "ฟร", "ฟล", "ดร",
            "หง", "หญ", "หน", "หม", "หย", "หร", "หล", "หว", "อย"]

TONES = ["", "่", "้", "๊", "๋"]

# C = พยัญชนะต้น, T = ตำแหน่งวรรณยุกต์  (ครบ 32 รูปสระ + สระลดรูป/เปลี่ยนรูป + ตัวสะกด)
PATTERNS = [
    "CTะ", "CTา", "CิT", "CีT", "CึT", "CืT", "CุT", "CูT",
    "เCTะ", "เCT", "แCTะ", "แCT", "โCTะ", "โCT", "เCTาะ", "CTอ", "เCTอะ", "เCTอ",
    "เCีTยะ", "เCีTย", "เCืTอะ", "เCืTอ", "CัTวะ", "CัTว", "CTำ", "ใCT", "ไCT", "เCTา",
    "CัTน", "CัTก", "CิTน", "CีTด", "CึTง", "CืTม", "CุTบ", "CูTก", "เCิTน", "เCีTยง",
    "เCืTอง", "CัTวร", "CTาง", "เC็ม", "แC็ง", "CTน", "CTรร", "CTาย", "CTาว", "เCTย", "CTอย",
]

# คำจริงที่มีจุดยาก (หางยาว/หางล่าง, การันต์, ฤ ฦ, เลขไทย, เครื่องหมาย)
SPECIAL = [
    "ฐาน", "ฐู", "ญาติ", "ญุ", "ฎีกา", "ฎุ", "ฏุ", "กฎหมาย", "ปฏิบัติ", "ปุ่ม", "ฝุ่น", "ฟุ้ง",
    "ปี่", "ปี๊บ", "ฝื่น", "ฟื้น", "ฬี", "จันทร์", "ศาสตร์", "ลักษณ์", "สิทธิ์", "พระพุทธรูป", "ภาพยนตร์",
    "ฤดู", "ฤๅษี", "ฦๅ", "ทฤษฎี", "อังกฤษ", "ฯลฯ", "ไป ๆ มา ๆ", "฿", "๐๑๒๓๔๕๖๗๘๙",
    "น้ำ", "ทำ", "ลำไย", "กำหนด", "ต่ำ", "ค่ำ", "ซ้ำ", "ก็", "เป็น", "ใบอนุญาต", "สำหรับ", "สำนักงาน",
    "คำขอ", "ประจำตัว", "ลูกบาศก์เซนติเมตร", "บริบูรณ์", "ผู้ประกอบวิชาชีพเวชกรรม", "ฟั่นเฟือน",
    "เพิกถอน", "เกี่ยวข้อง", "เปลี่ยนแปลง", "เครื่อง", "เพื่อ", "เบื้อง", "แก้ไข", "โต๊ะ", "เจ๊ะ", "ก๋วยเตี๋ยว",
]


def build_words():
    words = []
    for init in CONSONANTS + CLUSTERS:
        for pat in PATTERNS:
            for t in TONES:
                if "็" in pat and t:
                    continue
                words.append(pat.replace("C", init).replace("T", t))
    return list(dict.fromkeys(words + SPECIAL))


# ---------------------------------------------------------------------------
# จำลองความเสียหาย
# ---------------------------------------------------------------------------
UPPER = "ัิีึื็ํ"
LOWER = "ฺุู"
TONE_SET = "่้๊๋์"
TALL = "ปฝฟฬ"
DESC = "ฎฏ"

PUA_UPPER_LEFT = {"ั": "", "ิ": "", "ี": "", "ึ": "", "ื": "", "็": "", "ํ": ""}
PUA_TONE_DOWNLEFT = dict(zip(TONE_SET, ""))
PUA_TONE_DOWN = dict(zip(TONE_SET, ""))
PUA_TONE_LEFT = dict(zip(TONE_SET, ""))
PUA_LOWER_DOWN = dict(zip(LOWER, ""))


def corrupt_pua(w):
    """จำลองฟอนต์ที่เก็บสระ/วรรณยุกต์ตำแหน่งพิเศษเป็นรหัส PUA (เหมือน PDF ราชการหลายไฟล์)"""
    out = list(w)
    for i, c in enumerate(w):
        base_i = i - 1
        while base_i >= 0 and w[base_i] in UPPER + LOWER + TONE_SET:
            base_i -= 1
        base = w[base_i] if base_i >= 0 else ""
        has_upper = any(x in UPPER for x in w[base_i + 1:i + 2]) if base else False
        nxt_lower = i + 1 < len(w) and w[i + 1] in LOWER
        if c in UPPER and base in TALL:
            out[i] = PUA_UPPER_LEFT[c]
        elif c in TONE_SET and base:
            if base in TALL:
                out[i] = PUA_TONE_LEFT[c] if has_upper else PUA_TONE_DOWNLEFT[c]
            elif not has_upper and w[i + 1:i + 2] != "ำ":
                out[i] = PUA_TONE_DOWN[c]
        elif c in LOWER and base in DESC:
            out[i] = PUA_LOWER_DOWN[c]
        elif c == "ฐ" and nxt_lower:
            out[i] = ""
        elif c == "ญ" and nxt_lower:
            out[i] = ""
    return "".join(out)


def corrupt_spaced(w):
    """มีช่องว่างแทรกก่อนสระบน/ล่าง/วรรณยุกต์ และหลังสระหน้า"""
    out = []
    for i, c in enumerate(w):
        if c in UPPER + LOWER + TONE_SET + "ะำ" and i > 0:
            out.append(" ")
        out.append(c)
        if c in "เแโใไ" and i + 1 < len(w) and "ก" <= w[i + 1] <= "ฮ":
            out.append(" ")
    return "".join(out)


def corrupt_tone_first(w):
    """วรรณยุกต์มาก่อนสระบน/ล่าง"""
    chars = list(w)
    for i in range(len(chars) - 1):
        if chars[i] in UPPER.replace("ํ", "").replace("็", "") + LOWER and chars[i + 1] in TONE_SET:
            chars[i], chars[i + 1] = chars[i + 1], chars[i]
    return "".join(chars)


def corrupt_am_split(w):
    """สระอำแตกเป็น นิคหิต + (วรรณยุกต์) + า"""
    import re
    return re.sub("([่้๊๋]?)ำ", lambda m: "ํ" + m.group(1) + "า", w)


def corrupt_am_split2(w):
    import re
    return re.sub("([่้๊๋]?)ำ", lambda m: m.group(1) + "ํา", w)


def corrupt_lost_am(w):
    """นิคหิตหายกลายเป็นช่องว่าง: สำหรับ -> ส าหรับ"""
    import re
    return re.sub("([่้๊๋]?)ำ", lambda m: m.group(1) + " า", w)


def corrupt_dup(w):
    """สระ/วรรณยุกต์ซ้ำ: ทั้ง -> ทั้ั้ง"""
    import re
    return re.sub(f"([{UPPER}{LOWER}][{TONE_SET}]|[{TONE_SET}])", lambda m: m.group(1) * 2, w, count=1)


def corrupt_pua_spaced(w):
    return corrupt_spaced(corrupt_pua(w))


def corrupt_pua_tone_first(w):
    return corrupt_pua(corrupt_tone_first(w))


CORRUPTIONS = {
    "PUA (รหัสพิเศษ)": corrupt_pua,
    "ช่องว่างแทรก": corrupt_spaced,
    "วรรณยุกต์มาก่อนสระ": corrupt_tone_first,
    "สระอำแตก (ํ+่+า)": corrupt_am_split,
    "สระอำแตก (่+ํา)": corrupt_am_split2,
    "นิคหิตหาย (ส าหรับ)": corrupt_lost_am,
    "สระ/วรรณยุกต์ซ้ำ": corrupt_dup,
    "PUA + ช่องว่าง": corrupt_pua_spaced,
    "PUA + วรรณยุกต์มาก่อน": corrupt_pua_tone_first,
}
