"""สร้างฟอนต์ "Sarabun TH Compat" จาก Google Sarabun (OFL) ให้ขนาด/ระยะเท่า TH Sarabun New / TH SarabunIT๙

- ย่อทั้งฟอนต์ด้วยการเพิ่ม unitsPerEm (ตัวอักษรเท่า TH Sarabun ที่ pt เดียวกัน)
- ตั้งค่าความสูงบรรทัด (ascender/descender/lineGap) และความกว้างช่องว่างให้เท่า TH Sarabun
ใช้ครั้งเดียวตอนพัฒนา: python tools/build_fonts.py <โฟลเดอร์ Sarabun-*.ttf> <โฟลเดอร์ปลายทาง>
"""
import sys
from fontTools.ttLib import TTFont

# น้ำหนักเส้นที่ใกล้ TH SarabunIT๙ ที่สุด (วัดความหนาก้านตัวอักษรเทียบกับ PDF จริง 81 บรรทัด ที่ 864 dpi):
#   ตัวธรรมดา = Sarabun ExtraLight (ก้าน 0.96x, หมึก 0.93x)  [Light หนาไป: ก้าน 1.17x]
#   ตัวหนา    = Sarabun Medium     (ก้าน 1.04x, หมึก 1.02x)  [SemiBold หนาไป: ก้าน 1.15x]
SOURCE = {"Regular": "ExtraLight", "Italic": "ExtraLightItalic", "Bold": "Medium", "BoldItalic": "MediumItalic"}
# ตัวคูณ unitsPerEm ให้ความกว้างบรรทัดเท่า TH SarabunIT๙ (วัดจากความกว้างบรรทัดจริงใน PDF, แคบกว่า ~0.1% กันตัดบรรทัดเกิน)
SCALE = {"Regular": 1.4846, "Italic": 1.4846, "Bold": 1.4831, "BoldItalic": 1.4831}
TH = dict(asc=0.85, desc=0.25, gap=0.03, typo_gap=0.06)      # metric ของ TH SarabunIT๙ (หน่วย em)
SPACE = {"Regular": 0.216, "Italic": 0.216, "Bold": 0.226, "BoldItalic": 0.226}
STYLE_NAME = {"Regular": "Regular", "Bold": "Bold", "Italic": "Italic", "BoldItalic": "Bold Italic"}


def build(src_dir, dst_dir):
    for style in ["Regular", "Bold", "Italic", "BoldItalic"]:
        f = TTFont(f"{src_dir}/Sarabun-{SOURCE[style]}.ttf")
        upm = int(round(f["head"].unitsPerEm * SCALE[style]))
        f["head"].unitsPerEm = upm
        a, d, g = round(TH["asc"] * upm), round(TH["desc"] * upm), round(TH["gap"] * upm)
        f["hhea"].ascent, f["hhea"].descent, f["hhea"].lineGap = a, -d, g
        os2 = f["OS/2"]
        os2.sTypoAscender, os2.sTypoDescender, os2.sTypoLineGap = a, -d, round(TH["typo_gap"] * upm)
        os2.usWinAscent, os2.usWinDescent = a, d
        os2.fsSelection &= ~(1 << 7)                    # ไม่ใช้ USE_TYPO_METRICS (เหมือนต้นฉบับ)
        # ความกว้างช่องว่าง
        cmap = f.getBestCmap()
        for cp in (0x20, 0xA0):
            gname = cmap.get(cp)
            if gname:
                adv, lsb = f["hmtx"].metrics[gname]
                f["hmtx"].metrics[gname] = (round(SPACE[style] * upm), lsb)
        sub = STYLE_NAME[style]
        for rec in f["name"].names:
            if rec.nameID in (1, 16):
                rec.string = "Sarabun TH Compat"
            elif rec.nameID in (2, 17):
                rec.string = sub
            elif rec.nameID == 4:
                rec.string = f"Sarabun TH Compat {sub}"
            elif rec.nameID == 6:
                rec.string = f"SarabunTHCompat-{style}"
            elif rec.nameID == 3:
                rec.string = f"SarabunTHCompat-{style};derived-from-Sarabun"
        os2.usWeightClass = 700 if "Bold" in style else 400
        os2.fsSelection = (os2.fsSelection & ~0b1100001) | (0b100000 if "Bold" in style else 0) | \
            (0b1 if "Italic" in style else 0) | (0b1000000 if style == "Regular" else 0)
        f["head"].macStyle = (1 if "Bold" in style else 0) | (2 if "Italic" in style else 0)
        f.save(f"{dst_dir}/SarabunTHCompat-{style}.ttf")
        print("saved", style, "upm", upm)


if __name__ == "__main__":
    build(sys.argv[1], sys.argv[2])
