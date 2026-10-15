"""แก้คำผิดจาก OCR ภาษาไทยด้วยพจนานุกรม (PyThaiNLP)

แก้เฉพาะสิ่งที่ OCR ภาษาไทยพลาดบ่อย:
  - สระบน/ล่าง/วรรณยุกต์หาย หรืออ่านผิดตัว   เช่น ชัวคราว -> ชั่วคราว, ทีให้ -> ที่ให้, นิว -> นิ้ว
  - ช่องว่างแทรกกลางคำ                       เช่น บุคค ลวิกลจริต -> บุคคลวิกลจริต
  - บรรทัดขยะ (ตัวอักษรกระจัดกระจาย)        เช่น "๑4 Cf aA A v ส ส ๐"
แก้เฉพาะคำที่ไม่มีในพจนานุกรม และไม่เปลี่ยนพยัญชนะเลย เพื่อไม่ให้คำที่ถูกอยู่แล้วเพี้ยน
"""
import difflib
import math
import re
from functools import lru_cache

from pythainlp.corpus import thai_words
from pythainlp.tokenize import Tokenizer

WORDS = set(thai_words())
_tok = Tokenizer(WORDS, engine="newmm")

MARKS = "่้๊๋์ัิีึืุู็"          # ตัวที่ OCR มักทำหาย/อ่านสลับ
THAI_CONS = "ก-ฮ"
THAI_RUN = re.compile(r"[ก-๎]+")


try:
    from pythainlp.corpus import tnc
    _FREQ = {w: f for w, f in tnc.word_freqs()}
except Exception:  # noqa: BLE001  (ถ้าไม่มีข้อมูลความถี่ ใช้พจนานุกรมอย่างเดียว)
    _FREQ = {}
_TOTAL = float(sum(_FREQ.values()) or 1)
_UNK_LOGP = -30.0


def _logp(word: str) -> float:
    f = _FREQ.get(word)
    if f:
        return math.log(f / _TOTAL)
    if word in WORDS:                                 # คำในพจนานุกรมแต่ไม่มีสถิติ (มักเป็นคำประสม เช่น ว่าด้วย)
        return -12.0
    return _UNK_LOGP


@lru_cache(maxsize=500_000)
def _tokens(text: str):
    return tuple(_tok.word_tokenize(text))


def _is_unknown(t: str) -> bool:
    return t not in WORDS and bool(THAI_RUN.fullmatch(t))


@lru_cache(maxsize=500_000)
def _score(text: str):
    """(จำนวนตัวอักษรที่ไม่อยู่ในคำศัพท์, จำนวนคำ)"""
    toks = _tokens(text)
    return sum(len(t) for t in toks if _is_unknown(t)), len(toks)


@lru_cache(maxsize=500_000)
def _lm(text: str) -> float:
    """ความน่าจะเป็นของข้อความ (unigram) ยิ่งมากยิ่งเป็นภาษาไทยที่พบบ่อย"""
    return sum(_logp(t) for t in _tokens(text))


def _variants(s: str):
    """แก้ไข 1 จุด: เติม / ลบ / เปลี่ยน สระบน-ล่าง-วรรณยุกต์ (ไม่แตะพยัญชนะ)"""
    out = set()
    for i in range(1, len(s) + 1):
        prev = s[i - 1]
        if "\u0E01" <= prev <= "\u0E2E" or prev in MARKS:
            for m in MARKS:
                out.add(s[:i] + m + s[i:])
    for i, c in enumerate(s):
        if c in MARKS:
            out.add(s[:i] + s[i + 1:])
            for m in MARKS:
                if m != c:
                    out.add(s[:i] + m + s[i + 1:])
    out.discard(s)
    return out


EDIT_PENALTY = 3.0
MIN_NEW_WORD_LOGP = -11.0   # คำที่เกิดจากการแก้ต้องพบบ่อยพอ (กันการแก้เป็นเศษคำหายาก เช่น ขัว นุต งัน)


def _good_new_word(t: str) -> bool:
    base = re.sub("[\u0E31\u0E34-\u0E3A\u0E47-\u0E4E]", "", t)
    if len(base) < 2:                                # คำพยางค์สั้น (ปี ที่ มี) ต้องพบบ่อยมาก
        return t in _FREQ and _logp(t) >= -7.0
    if t in _FREQ:
        return _logp(t) >= MIN_NEW_WORD_LOGP
    return t in WORDS and len(base) >= 4             # คำประสมยาวในพจนานุกรม   # ยิ่งแก้หลายจุด ยิ่งต้องมั่นใจมาก


def _fix_window(w: str, max_edits: int = 3, beam: int = 12) -> str:
    """ค้นหาแบบ beam: หาข้อความที่แก้ไม่เกิน max_edits จุด แล้วทุกคำอยู่ในพจนานุกรม เลือกตัวที่พบบ่อยที่สุด"""
    frontier = [w]
    seen = {w}
    orig = set(_tokens(w))
    solutions = []
    for depth in range(1, max_edits + 1):
        cands = []
        for f in frontier:
            for v in _variants(f):
                if v in seen:
                    continue
                seen.add(v)
                unk = _score(v)[0]
                # ทุกคำในผลลัพธ์ต้องเป็นคำที่พบบ่อย (มีสถิติ) หรือเป็นคำเดิมที่ไม่ได้แก้ — กันการแก้เป็นคำหายาก
                if unk == 0 and all(t in orig or _good_new_word(t) for t in _tokens(v)):
                    solutions.append((_lm(v) - EDIT_PENALTY * depth, v))
                cands.append((unk, -_lm(v), v))
        if solutions and depth >= 2:
            break
        cands.sort()
        frontier = [v for _, _, v in cands[:beam]]
    if not solutions:
        return w
    return max(solutions)[1]


def _fix_run(run: str) -> str:
    """แก้เฉพาะช่วงรอบคำที่ไม่รู้จัก (คำแปลก + คำข้างเคียงซ้ายขวา)"""
    toks = list(_tokens(run))
    if not any(_is_unknown(t) for t in toks):
        return run
    out, i = [], 0
    while i < len(toks):
        if not _is_unknown(toks[i]):
            out.append(toks[i])
            i += 1
            continue
        j = i
        while j < len(toks) and _is_unknown(toks[j]):
            j += 1
        # ขยายหน้าต่างไปคำข้างเคียงซ้าย-ขวาอย่างละ 2 คำ
        lo = i
        while lo > 0 and i - lo < 2 and out:
            out.pop()
            lo -= 1
        hi = min(j + 2, len(toks))
        window = "".join(toks[lo:hi])
        out.append(_fix_window(window) if len(window) <= 30 else window)
        i = hi
    return "".join(out)


def _is_junk_line(line: str) -> bool:
    toks = line.split()
    if len(toks) < 2:
        return False
    if all(re.fullmatch(r"[\d,.\-()/:]+", t) for t in toks):   # ตัวเลขล้วน เช่นตาราง ไม่ใช่ขยะ
        return False
    short = sum(1 for t in toks if len(t) <= 2)
    long_thai_known = any(len(t) >= 3 and _score(t)[0] == 0 for t in THAI_RUN.findall(line))
    long_latin = any(re.fullmatch(r"[A-Za-z]{3,}", t) for t in toks)
    return short / len(toks) >= 0.7 and not long_thai_known and not long_latin


def _crosses(a: str, b: str) -> bool:
    """มีคำในพจนานุกรมคร่อมรอยต่อระหว่าง a กับ b หรือไม่ (แปลว่าช่องว่างตรงนั้นแทรกกลางคำ)"""
    pos, cut = 0, len(a)
    for tok in _tokens(a + b):
        if pos < cut < pos + len(tok):
            return tok in WORDS
        pos += len(tok)
    return False


def fix_ocr_segments(line: str):
    """คืนรายการ [(ข้อความ, ถูกแก้หรือไม่), ...] หรือ None ถ้าเป็นบรรทัดขยะ"""
    if _is_junk_line(line):
        return None
    parts = [[p, False] for p in line.split(" ")]
    # 1) รวมช่องว่างที่แทรกกลางคำ: แยกกันมีคำแปลก แต่รวมกันแล้วดีขึ้น  เช่น บุคค ลวิกลจริต
    merged = []
    for p in parts:
        if merged and THAI_RUN.fullmatch(p[0] or "-") and THAI_RUN.search(merged[-1][0][-1:] or "-"):
            a, b = merged[-1][0], p[0]
            if (_score(a)[0] or _score(b)[0]) and _crosses(a, b) \
                    and _score(a + b)[0] < _score(a)[0] + _score(b)[0]:
                merged[-1] = [a + b, True]
                continue
        merged.append(p)
    # 1b) ช่องว่างแทรก + สระหาย พร้อมกัน  เช่น ว่าดัว ยการ -> ว่าด้วยการ
    i = 0
    while i < len(merged) - 1:
        a, b = merged[i][0], merged[i + 1][0]
        if THAI_RUN.fullmatch(a or "-") and THAI_RUN.fullmatch(b or "-") and (_score(a)[0] or _score(b)[0]) \
                and (_score(_fix_run(a))[0] or _score(_fix_run(b))[0]):   # รวมเฉพาะเมื่อแก้แยกกันไม่ได้
            ab = _fix_run(a + b)
            if _score(ab)[0] == 0:
                merged[i:i + 2] = [[ab, True]]
                continue
        i += 1
    # 2) แก้สระ/วรรณยุกต์ในแต่ละช่วงตัวอักษรไทย
    segs = []
    for k, (p, was_merged) in enumerate(merged):
        if k:
            segs.append((" ", False))
        pos = 0
        for m in THAI_RUN.finditer(p):
            s_, e_ = m.span()
            if s_ > pos:
                segs.append((p[pos:s_], False))
            old = m.group(0)
            # ข้ามตัวย่อ (พ.ร.บ. / พ.ศ.) และคำสั้นมาก
            if len(old) < 3 or p[e_:e_ + 1] == "." or p[s_ - 1:s_] == ".":
                new = old
            else:
                new = _fix_run(old)
            if was_merged or new == old:
                segs.append((new, was_merged))
            else:                                       # ไฮไลต์เฉพาะคำที่เปลี่ยน
                ot, nt = list(_tokens(old)), list(_tokens(new))
                for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, ot, nt, autojunk=False).get_opcodes():
                    if j2 > j1:
                        segs.append(("".join(nt[j1:j2]), op != "equal"))
            pos = e_
        if pos < len(p):
            segs.append((p[pos:], False))
    # รวมชิ้นติดกันที่สถานะเดียวกัน
    out = []
    for t, f in segs:
        if out and out[-1][1] == f:
            out[-1] = (out[-1][0] + t, f)
        elif t:
            out.append((t, f))
    return out


def fix_ocr_line(line: str):
    """คืน (ข้อความที่แก้แล้ว, รายการชิ้นที่ถูกแก้) หรือ (None, []) ถ้าเป็นบรรทัดขยะ"""
    segs = fix_ocr_segments(line)
    if segs is None:
        return None, []
    return "".join(t for t, _ in segs), [t for t, f in segs if f]
