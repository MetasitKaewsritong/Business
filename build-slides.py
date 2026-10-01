# Build slides-group1.pptx from WORKSHOP-REVISED.md content.  Run: python build-slides.py
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from PIL import Image

FONT = "Leelawadee UI"
BLUE, YEL = RGBColor(0x00, 0x4A, 0xAD), RGBColor(0xFD, 0xCD, 0x01)
INK, MUTED, WHITE = RGBColor(0x1F, 0x29, 0x37), RGBColor(0x6B, 0x72, 0x80), RGBColor(0xFF, 0xFF, 0xFF)
TINT, YTINT, LINE = RGBColor(0xEB, 0xF1, 0xFA), RGBColor(0xFF, 0xF4, 0xCC), RGBColor(0xD1, 0xD9, 0xE6)
W, H = 13.333, 7.5
L, R = 0.6, 13.333 - 0.6          # content margins
CW = R - L

def setfont(f):
    """Latin + Thai (complex script) typeface; python-pptx only sets Latin."""
    f.name = FONT
    latin = f._rPr.find(qn("a:latin"))
    for tag in ("a:cs", "a:ea"):  # addnext in reverse keeps schema order latin, ea, cs
        el = f._rPr.makeelement(qn(tag), {"typeface": FONT})
        latin.addnext(el)


prs = Presentation()
prs.slide_width, prs.slide_height = Inches(W), Inches(H)
BLANK = prs.slide_layouts[6]


def text(s, x, y, w, h, t, size=18, bold=False, color=INK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    """t: str, or list of paragraphs; a paragraph is str or list of (text, overrides) runs."""
    tf = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)).text_frame
    tf.word_wrap, tf.vertical_anchor = True, anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, para in enumerate([t] if isinstance(t, str) else t):
        p = tf.add_paragraph() if i else tf.paragraphs[0]
        p.alignment = align
        runs = [(para, {})] if isinstance(para, str) else para
        if i:
            p.space_before = Pt(max(o.get("gap", 6) for _, o in runs))
        for run_t, o in runs:
            r = p.add_run()
            r.text = run_t
            f = r.font
            setfont(f)
            f.size = Pt(o.get("size", size))
            f.bold, f.color.rgb = o.get("bold", bold), o.get("color", color)
    return tf


def bullets(s, x, y, w, h, items, size=18, gap=10, color=INK):
    tf = text(s, x, y, w, h, [[(t, {"size": size, "gap": gap, "color": color})] for t in items])
    for p in tf.paragraphs:
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", str(Inches(0.28)))
        pPr.set("indent", str(-Inches(0.28)))
        bu = pPr.makeelement(qn("a:buChar"), {"char": "•"})
        pPr.append(bu)
    return tf


def box(s, x, y, w, h, fill=TINT, line=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE):
    sh = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.shadow.inherit = False
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line:
        sh.line.color.rgb, sh.line.width = line, Pt(1)
    else:
        sh.line.fill.background()
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        sh.adjustments[0] = min(0.12, 0.08 / min(w, h) * 1.0)
    return sh


def arrow(s, x, y, w=0.3, h=0.26, color=MUTED):
    a = box(s, x, y, w, h, color, shape=MSO_SHAPE.RIGHT_ARROW)
    return a


def pic(s, path, x, y, w=None, h=None):
    """Place image fitting inside w x h (inches), centered in that box."""
    iw, ih = Image.open(path).size
    scale = min(w / iw, h / ih) if (w and h) else (w / iw if w else h / ih)
    pw, ph = iw * scale, ih * scale
    return s.shapes.add_picture(path, Inches(x + ((w or pw) - pw) / 2), Inches(y + ((h or ph) - ph) / 2), Inches(pw), Inches(ph))


def base(kicker, notes=None):
    """Static template: banner."""
    s = prs.slides.add_slide(BLANK)
    banner = box(s, 0, 0, W, 0.84, BLUE, shape=MSO_SHAPE.RECTANGLE)
    ef = banner._element.spPr.find(qn("a:effectLst"))   # soft shadow under the banner only
    sh = ef.makeelement(qn("a:outerShdw"), {"blurRad": str(Pt(6)), "dist": str(Pt(2)), "dir": "5400000", "algn": "t", "rotWithShape": "0"})
    clr = sh.makeelement(qn("a:srgbClr"), {"val": "000000"})
    clr.append(clr.makeelement(qn("a:alpha"), {"val": "25000"}))
    sh.append(clr)
    ef.append(sh)
    pic(s, "assets/logo-buu.png", 0.25, 0.12, h=0.6)
    pic(s, "assets/logo-if.png", 0.97, 0.12, h=0.6)
    text(s, 3.0, 0, W - 6.0, 0.84, "88732065 Business Process 2569-1", 20, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, W - 3.25, 0, 3.0, 0.84, "กลุ่มที่ 1", 20, color=WHITE, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    text(s, W / 2 - 3.5, 1.5, 7.0, 0.3, kicker, 14, True, BLUE, PP_ALIGN.CENTER)   # page number goes on the right, see page_numbers()
    text(s, L, 1.5, 3.0, 0.3, "รหัสนิสิต", 14, True, BLUE)  # placeholder: presenter fills in
    if notes:
        s.notes_slide.notes_text_frame.text = notes
    return s


def content(kicker, title, notes=None):
    s = base(kicker, notes)
    text(s, L, 1.8, CW, 0.9, title, 26, True)
    return s


def tag(s, x, y, t, fill=BLUE, color=WHITE, size=12, w=None):
    w = w or 0.2 + 0.11 * len(t)
    b = box(s, x, y, w, 0.3, fill)
    text(s, x, y, w, 0.3, t, size, True, color, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    return w


def table(s, x, y, colw, rows, rowh=0.44, size=13, head=(BLUE, WHITE), stripe=True, first_blue=True, hl=(), hl_fill=YTINT):
    shp = s.shapes.add_table(len(rows), len(colw), Inches(x), Inches(y), Inches(sum(colw)), Inches(rowh * len(rows)))
    tb = shp.table
    for j, cw in enumerate(colw):
        tb.columns[j].width = Inches(cw)
    for i, row in enumerate(rows):
        tb.rows[i].height = Inches(rowh)
        for j, v in enumerate(row):
            c = tb.cell(i, j)
            c.fill.solid()
            c.fill.fore_color.rgb = head[0] if i == 0 else (hl_fill if i in hl else (TINT if stripe and i % 2 == 0 else WHITE))
            c.margin_left = c.margin_right = Inches(0.1)
            c.margin_top = c.margin_bottom = Inches(0.02)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            r = c.text_frame.paragraphs[0].add_run()
            r.text = v
            setfont(r.font)
            r.font.size = Pt(size)
            r.font.bold = i == 0 or (first_blue and j == 0)
            r.font.color.rgb = head[1] if i == 0 else (BLUE if first_blue and j == 0 else INK)
    return tb


# ── Wireframe primitives (lo-fi, grayscale + one blue for primary actions)
WFL, WFF, WFT = RGBColor(0x9C, 0xA3, 0xAF), RGBColor(0xF3, 0xF4, 0xF6), RGBColor(0x37, 0x41, 0x51)
FX, FY, FW, FH = L, 2.85, 7.4, 4.1        # wireframe frame on the slide


def wrect(s, x, y, w, h, fill=WHITE, line=WFL, rounded=False):
    return box(s, x, y, w, h, fill, line, MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE)


def wt(s, x, y, w, h, t, size=10, bold=False, color=WFT, align=PP_ALIGN.LEFT):
    return text(s, x, y, w, h, t, size, bold, color, align, MSO_ANCHOR.MIDDLE)


def field(s, x, y, w, label, value, h=0.28):
    wt(s, x, y, w, 0.22, label, 9, color=MUTED)
    wrect(s, x, y + 0.23, w, h)
    wt(s, x + 0.08, y + 0.23, w - 0.1, h, value, 10)


def btn(s, x, y, w, label, primary=True, h=0.3):
    wrect(s, x, y, w, h, BLUE if primary else WHITE, None if primary else WFL, rounded=True)
    wt(s, x, y, w, h, label, 10, True, WHITE if primary else WFT, PP_ALIGN.CENTER)


def check(s, x, y, label, on=False, radio=False):
    box(s, x, y + 0.03, 0.14, 0.14, BLUE if on else WHITE, None if on else WFL, MSO_SHAPE.OVAL if radio else MSO_SHAPE.RECTANGLE)
    wt(s, x + 0.2, y, 1.4, 0.2, label, 10)


def pill(s, x, y, label, on=True, w=0.7):
    wrect(s, x, y, w, 0.24, TINT if on else WFF, BLUE if on else WFL, rounded=True)
    wt(s, x, y, w, 0.24, label, 9, True, BLUE if on else MUTED, PP_ALIGN.CENTER)


DANGER = RGBColor(0xC6, 0x28, 0x28)


def browser(s, active, staff=False):
    wrect(s, FX, FY, FW, FH)
    wrect(s, FX, FY, FW, 0.4, WFF)
    wt(s, FX + 0.2, FY, 2.5, 0.4, "ระบบจองห้องและอุปกรณ์", 11, True)
    x = FX + FW - 0.1
    for name, w in [("ออกจากระบบ", 0.95), ("ยืม-คืน/เช็กอิน-เช็กเอาต์", 1.9), ("ตรวจสอบสถานะ", 1.15), ("ค้นหา", 0.6)]:
        if name == "ตรวจสอบสถานะ" and not staff:
            continue
        x -= w
        on = name == active
        if on:
            wrect(s, x, FY + 0.07, w, 0.26, BLUE, None, rounded=True)
        wt(s, x, FY, w, 0.4, name, 10, on, WHITE if on else (DANGER if name == "ออกจากระบบ" else WFT), PP_ALIGN.CENTER)
        x -= 0.04


RED = RGBColor(0xE5, 0x73, 0x73)


def marker(s, x, y, n):
    box(s, x, y, 0.28, 0.28, RED, None, MSO_SHAPE.OVAL)
    text(s, x, y, 0.28, 0.28, str(n), 11, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)


def region(s, x, y, w, h, n):
    """Light red frame around a wireframe area, with its number on the corner."""
    r = box(s, x, y, w, h, WHITE, RED)
    r.fill.background()
    r.line.width = Pt(1.75)
    marker(s, x - 0.14, y - 0.14, n)


def glow(sh, hexcolor="FDCD01", rad=16):
    ef = sh._element.spPr.find(qn("a:effectLst"))
    g = ef.makeelement(qn("a:glow"), {"rad": str(Pt(rad))})
    c = g.makeelement(qn("a:srgbClr"), {"val": hexcolor})
    c.append(c.makeelement(qn("a:alpha"), {"val": "75000"}))
    g.append(c)
    ef.append(g)


def line(s, pts, arrow=True, color=MUTED):
    """Polyline through pts; arrowhead on the last segment."""
    for (x1, y1), (x2, y2) in zip(pts, pts[1:]):
        c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        c.line.color.rgb, c.line.width = color, Pt(1.5)
    if arrow:
        ln = c.line._get_or_add_ln()
        ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle"}))


def no_shadows():
    """Kill shadows for every viewer: empty effect list per shape, and no theme effect reference."""
    for sl in prs.slides:
        for el in sl.shapes._spTree.iter(qn("p:spPr")):
            if el.find(qn("a:effectLst")) is None:
                ln = el.find(qn("a:ln"))
                ef = el.makeelement(qn("a:effectLst"), {})
                ln.addnext(ef) if ln is not None else el.append(ef)
        for ref in sl.shapes._spTree.iter(qn("a:effectRef")):
            ref.set("idx", "0")


def shift_up():
    """Lift content toward the banner: 8 presenters stand in front and block the bottom of the screen."""
    for n, sl in enumerate(prs.slides):
        for sh in sl.shapes:
            if sh.top < Inches(0.95):          # banner stays put
                continue
            head = sh.top < Inches(1.9)   # section label, student id, title
            sh.top -= Inches(0.55 if head else 0.7 if n == 0 else 0.85)


# รหัสนิสิตของผู้พูดแต่ละหน้า (หน้า 1-13) แก้ตรงนี้ถ้าสลับคนพูด
SIDS = ["67160403", "67160403", "67160403", "67160005", "67160003", "67160018",
        "67160042", "67160042", "67160178", "67160178", "67160025", "67160193", "67160193"]


def page_numbers():
    total = len(prs.slides)
    assert total == len(SIDS)
    for n, sl in enumerate(prs.slides):
        for sh in sl.shapes:
            if sh.has_text_frame and sh.text_frame.text == "รหัสนิสิต":
                sh.text_frame.paragraphs[0].runs[0].text = SIDS[n]
        text(sl, R - 3.0, 1.5, 3.0, 0.3, f"หน้า {n + 1}/{total}", 14, True, BLUE, PP_ALIGN.RIGHT)


def strip_theme_shadows(path):
    """Some viewers apply the theme's effect styles regardless; empty them in the saved file."""
    import re, shutil, zipfile
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("ppt/theme/"):
                xml = data.decode("utf-8")
                xml = re.sub(r"<a:effectStyle>.*?</a:effectStyle>", "<a:effectStyle><a:effectLst/></a:effectStyle>", xml, flags=re.S)
                xml = re.sub(r'effectRef idx="\d+"', 'effectRef idx="0"', xml)
                data = xml.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)


def wf_panel(s, user, points, note=None):
    rx = FX + FW + 0.4
    rw = R - rx
    text(s, rx, 2.95, rw, 0.3, "ใครใช้เป็นหลัก", 14, True, MUTED)
    text(s, rx, 3.25, rw, 0.45, user, 20, True)
    y = 3.95
    if note:
        text(s, rx, 3.7, rw, 0.3, note, 14, color=MUTED)
        y = 4.2
    for i, p in enumerate(points):
        # ponytail: line count guessed from visible Thai chars (~30 per line at 16pt); measure real text if wraps go wrong
        n = 2 if not isinstance(p, str) else max(2, -(-sum(not ("ั" == c or "ิ" <= c <= "ฺ" or "็" <= c <= "๎") for c in p) // 30))
        if not isinstance(p, str):       # explicit lines: break exactly where given
            n, p = max(2, len(p)), [[(t, {"size": 16, "gap": 0})] for t in p]
        marker(s, rx, y + 0.03, i + 1)
        text(s, rx + 0.45, y, rw - 0.45, 0.3 * n, p, 16)
        y += 0.3 * n + 0.3


# 1 ─ ปก
s = base("นำเสนอ ณ วันที่ 7 ตุลาคม 2569")
text(s, 1, 3.3, W - 2, 1.7, ["ระบบจองห้องเรียน ห้องประชุม", "และอุปกรณ์ภายในคณะ"], 40, True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
pic(s, "assets/template-laptop.png", 0.71, 5.35, w=1.54)
pic(s, "assets/template-calendar.png", 10.82, 5.35, w=1.92)

# 2 ─ สารบัญ
s = content("ภาพรวม", "สารบัญ")
toc = [("01", "ปัญหา", "การจองที่กระจัดกระจาย และผลที่ตามมา"),
       ("02", "ทางออก", "ระบบจองกลาง โครงหน้าจอ และผังงาน"),
       ("03", "ความต้องการ", "สิ่งที่ระบบช่วยได้ และใครได้ประโยชน์"),
       ("04", "สรุป", "ขั้นต่อไป")]
for i, (n, h, d) in enumerate(toc):
    y = 2.9 + i * 0.8
    box(s, L, y, CW, 0.66, TINT)
    text(s, L + 0.3, y, 0.8, 0.66, n, 24, True, BLUE, anchor=MSO_ANCHOR.MIDDLE)
    text(s, L + 1.2, y, 3.0, 0.66, h, 20, True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, L + 4.3, y, CW - 4.7, 0.66, d, 18, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

# 3 ─ ปัญหา: หลายช่องทาง
s = content("01 ปัญหา", "การจองที่กระจัดกระจาย")
for i, ch in enumerate(["แจ้งเจ้าหน้าที่โดยตรง", "โทรศัพท์", "LINE", "เอกสารกระดาษ"]):
    y = 3.0 + i * 0.95
    box(s, L, y, 3.0, 0.72, TINT)
    text(s, L, y, 3.0, 0.72, ch, 18, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    arrow(s, (L + 3.0 + 4.55) / 2 - 0.15, y + 0.23)
box(s, 4.55, 3.0, 3.6, 3.57, BLUE)
text(s, 4.75, 3.0, 3.2, 3.57, [[("เจ้าหน้าที่", {"size": 24, "bold": True, "color": WHITE})],
                               [("ไล่ตรวจตารางเรียน", {"size": 18, "color": WHITE, "gap": 10})],
                               [("ตารางประชุม และสมุดจอง", {"size": 18, "color": WHITE, "gap": 0})],
                               [("จากหลายแหล่งด้วยตนเอง", {"size": 18, "color": WHITE, "gap": 0})]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
arrow(s, 8.32, 4.66)
box(s, 8.8, 3.0, R - 8.8, 3.57, YTINT)
text(s, 9.1, 3.0, R - 9.4, 3.57, [[("เมื่อเกิดปัญหา", {"size": 24, "bold": True})],
                                 [("เจ้าหน้าที่ต้องจัดการเองทั้งหมด", {"size": 20, "gap": 10})]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 4 ─ ผลกระทบ: 3 ปัญหา
s = content("01 ปัญหา", "ปัญหาหลัก")
cards = [("เสียเวลา", "ทุกคำขอต้องถามและตรวจด้วยมือ เจ้าหน้าที่ทำงานซ้ำซาก"),
         ("ตารางชน", "ข้อมูลอยู่หลายที่ จึงจองซ้อนกับตารางเรียน การสอนสะดุด"),
         ("ผู้บริหารมองไม่เห็น", ["ข้อมูลการใช้งานอยู่ใน", "เอกสารที่ไม่เป็นดิจิทัล", "ดึงมาใช้ทันทีไม่ได้"])]
cw = (CW - 2 * 0.3) / 3
for i, (th, d) in enumerate(cards):
    x = L + i * (cw + 0.3)
    box(s, x, 2.95, cw, 2.3, TINT)
    text(s, x + 0.35, 3.25, cw - 0.7, 1.9, [[(th, {"size": 24, "bold": True, "color": BLUE})]] + [[(t, {"size": 18, "gap": 0 if j else 10})] for j, t in enumerate([d] if isinstance(d, str) else d)])

# 5 ─ ผู้เกี่ยวข้อง
s = content("01 ปัญหา", "ผู้ใช้ระบบหลัก")
why = [("นิสิต", "ถ้าใช้ยาก ระบบจะไม่ถูกใช้งานจริง"),
       ("อาจารย์", "ใช้ห้องสอนและประชุม ถ้าตารางชน การสอนสะดุด"),
       ("เจ้าหน้าที่", "อนุมัติคำขอ จ่ายและรับคืนอุปกรณ์ ถ้าระบบไม่ลดงาน ก็ซ้ำซ้อนเหมือนเดิม")]
for i, (g, d) in enumerate(why):
    y = 2.95 + i * 0.9
    box(s, L, y, 2.6, 0.76, BLUE)
    text(s, L, y, 2.6, 0.76, g, 20, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    box(s, L + 2.75, y, CW - 2.75, 0.76, WHITE, LINE)
    text(s, L + 3.05, y, CW - 3.35, 0.76, d, 18, anchor=MSO_ANCHOR.MIDDLE)
text(s, L, 5.72, CW, 0.3, "ผู้เกี่ยวข้องอื่น", 14, True, MUTED)
others = [("ผู้ดูแลห้อง/อุปกรณ์", "เตรียมและดูแลห้องกับอุปกรณ์"),
          ("ผู้บริหาร", "ใช้สถิติวางแผนทรัพยากร"),
          ("ผู้ดูแลระบบ", "ดูแลสิทธิ์และความเสถียรของระบบ")]
ow = (CW - 2 * 0.25) / 3
for i, (g, d) in enumerate(others):
    x = L + i * (ow + 0.25)
    box(s, x, 6.07, ow, 0.85, TINT)
    text(s, x + 0.25, 6.07, ow - 0.5, 0.85, [[(g, {"size": 16, "bold": True})], [(d, {"size": 14, "color": MUTED, "gap": 2})]], anchor=MSO_ANCHOR.MIDDLE)

# 6 ─ ทางออก: สะท้อน 3 ปัญหาจากหน้าผลกระทบ ลำดับเดียวกัน
s = content("02 ทางออก", "ระบบจองกลาง")
fixes = [("เสียเวลา", ["เช็กห้องว่าง", "และส่งคำขอได้เอง"]),
         ("ตารางชน", ["ระบบตรวจการชนให้", "ก่อนรับคำขอ"]),
         ("ผู้บริหารมองไม่เห็น", ["มีข้อมูลการใช้งาน", "ที่สามารถดึงได้ทันที"])]
pw = (CW - 2 * 0.3) / 3
for i, (prob, fix) in enumerate(fixes):
    x = L + i * (pw + 0.3)
    box(s, x, 2.95, pw, 1.0, TINT)
    text(s, x + 0.15, 2.95, pw - 0.3, 1.0, prob, 20, True, BLUE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    box(s, x + pw / 2 - 0.16, 4.08, 0.32, 0.36, MUTED, shape=MSO_SHAPE.DOWN_ARROW)
    box(s, x, 4.57, pw, 1.9, BLUE)
    text(s, x + 0.2, 4.57, pw - 0.4, 1.9, [[(t, {"gap": 0})] for t in fix], 24, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)

# 11 ─ โครงหน้าจอ 1: ค้นหา
s = content("02 ทางออก · โครงหน้าจอ 1/4", "หน้าค้นหา")
browser(s, "ค้นหา")
ox, oy = FX, FY
wrect(s, ox + 0.2, oy + 0.6, 2.3, 3.35)
wt(s, ox + 0.35, oy + 0.66, 2.0, 0.3, "ตัวกรอง", 11, True)
field(s, ox + 0.35, oy + 0.98, 2.0, "วันที่", "10 ต.ค. 2569")
field(s, ox + 0.35, oy + 1.55, 0.95, "เริ่ม", "09:00")
field(s, ox + 1.4, oy + 1.55, 0.95, "สิ้นสุด", "12:00")
wt(s, ox + 0.35, oy + 2.15, 2.0, 0.22, "ประเภท", 9, color=MUTED)
for i, (lb, on) in enumerate([("ห้องเรียน", True), ("ห้องประชุม", False), ("อุปกรณ์", False)]):
    check(s, ox + 0.35, oy + 2.38 + i * 0.21, lb, on, radio=True)
field(s, ox + 0.35, oy + 3.0, 2.0, "จำนวนผู้เข้าร่วม", "40 คน")
btn(s, ox + 0.35, oy + 3.58, 2.0, "ใช้ตัวกรอง", False)
wrect(s, ox + 2.7, oy + 0.56, 3.6, 0.3)
wt(s, ox + 2.8, oy + 0.56, 3.4, 0.3, "ค้นหาชื่อห้องหรืออุปกรณ์ เช่น 401", 9, color=MUTED)
btn(s, ox + 6.38, oy + 0.56, 0.82, "ค้นหา")
wt(s, ox + 2.7, oy + 0.92, 3.0, 0.25, "ผลการค้นหา 3 รายการ", 11, True)
rooms = [("ห้อง 401", "ความจุ 50 คน · โพรเจกเตอร์", True), ("ห้อง 402", "ความจุ 40 คน · จอสมาร์ต", False),
         ("ห้อง 403", "ความจุ 60 คน · โพรเจกเตอร์", True)]
for i, (nm, det, free) in enumerate(rooms):
    x, y = ox + 2.7, oy + 1.22 + i * 0.93
    wrect(s, x, y, 4.5, 0.82)
    wrect(s, x + 0.12, y + 0.11, 0.8, 0.6, WFF, None)
    wt(s, x + 1.05, y + 0.08, 2.2, 0.3, nm, 11, True)
    wt(s, x + 1.05, y + 0.4, 2.4, 0.3, det, 9, color=MUTED)
    pill(s, x + 3.6, y + 0.09, "ว่าง" if free else "ไม่ว่าง", free)
    btn(s, x + 3.4, y + 0.44, 0.95, "จอง" if free else "ดูเวลาอื่น", free, h=0.28)
region(s, ox + 0.14, oy + 0.54, 2.42, 3.47, 1)
region(s, ox + 6.0, oy + 1.25, 1.14, 0.76, 2)
region(s, ox + 6.0, oy + 2.18, 1.14, 0.76, 3)
wf_panel(s, "นิสิตและอาจารย์",
         ["ค้นหาด้วยชื่อ หรือกรองตามวัน เวลา ประเภท และจำนวนคน",
          "เห็นทันทีว่าห้องไหนว่าง ไม่ต้องโทรถาม",
          "ห้องไม่ว่างมีปุ่มดูเวลาอื่น"])

# 12 ─ โครงหน้าจอ 2: ส่งคำขอ (ต่อจากการกดจองห้อง 401 ที่ว่างในหน้าค้นหา)
s = content("02 ทางออก · โครงหน้าจอ 2/4", "หน้าส่งคำขอ")
browser(s, "ค้นหา")
wt(s, ox + 0.3, oy + 0.5, 2.5, 0.35, "ส่งคำขอจอง", 12, True)
for i, (lb, st) in enumerate([("เลือกห้อง", 2), ("กรอกคำขอ", 1), ("รอผลอนุมัติ", 0)]):
    x = ox + 0.3 + i * 1.45
    c = box(s, x, oy + 0.95, 0.24, 0.24, BLUE if st else WHITE, None if st == 2 else (BLUE if st else WFL), MSO_SHAPE.OVAL)
    wt(s, x, oy + 0.95, 0.24, 0.24, str(i + 1), 9, True, WHITE if st == 2 else (BLUE if st else MUTED), PP_ALIGN.CENTER)
    if st == 1:
        c.fill.fore_color.rgb = WHITE
    wt(s, x + 0.3, oy + 0.95, 1.1, 0.24, lb, 10, st == 1, WFT if st else MUTED)
field(s, ox + 0.3, oy + 1.33, 1.95, "ห้องหรืออุปกรณ์", "ห้อง 401")
field(s, ox + 2.4, oy + 1.33, 2.0, "วันที่", "10 ต.ค. 2569")
field(s, ox + 0.3, oy + 1.93, 0.95, "เริ่ม", "09:00")
field(s, ox + 1.35, oy + 1.93, 0.9, "สิ้นสุด", "12:00")
field(s, ox + 2.4, oy + 1.93, 2.0, "จำนวนผู้เข้าร่วม", "40 คน")
wt(s, ox + 0.3, oy + 2.53, 4.1, 0.22, "วัตถุประสงค์", 9, color=MUTED)
wrect(s, ox + 0.3, oy + 2.76, 4.1, 0.7)
wt(s, ox + 0.38, oy + 2.76, 4.0, 0.35, "ประชุมชมรม", 10)
wrect(s, ox + 4.7, oy + 1.35, 2.45, 1.95, WFF, WFL, rounded=True)
wt(s, ox + 4.85, oy + 1.43, 2.2, 0.3, "ผลการตรวจสถานะว่าง", 11, True)
pill(s, ox + 4.85, oy + 1.82, "ว่าง", True, w=0.6)
wt(s, ox + 5.55, oy + 1.82, 1.5, 0.24, "ห้อง 401", 10, True)
wt(s, ox + 4.85, oy + 2.12, 2.2, 0.25, "10 ต.ค. 2569 เวลา 09:00-12:00", 9)
wt(s, ox + 4.85, oy + 2.45, 2.2, 0.7, "ถ้าแก้ห้อง วัน หรือเวลา ระบบจะตรวจให้ใหม่ทันที", 9, color=MUTED)
btn(s, ox + 4.95, oy + 3.6, 1.0, "ยกเลิก", False)
btn(s, ox + 6.1, oy + 3.6, 1.05, "ส่งคำขอ")
region(s, ox + 0.22, oy + 0.87, 4.25, 0.4, 1)
region(s, ox + 0.22, oy + 1.29, 4.27, 1.2, 2)
region(s, ox + 4.62, oy + 1.27, 2.61, 2.11, 3)
wf_panel(s, "นิสิตและอาจารย์",
         ["เห็นว่าคำขออยู่ขั้นไหน ตั้งแต่ส่งจนได้ผล",
          "แก้ห้อง วัน และเวลาได้ ไม่ต้องค้นหาใหม่",
          "ตรวจสถานะว่างใหม่ทุกครั้งที่แก้ กันจองชนกับตารางเรียน"])

# 13 ─ โครงหน้าจอ 3: คิวคำขอ
s = content("02 ทางออก · โครงหน้าจอ 3/4", "หน้าตรวจสอบสถานะ")
browser(s, "ตรวจสอบสถานะ", staff=True)
wt(s, ox + 0.3, oy + 0.5, 3.0, 0.35, "ตรวจสอบสถานะคำขอ", 12, True)
for i, tb in enumerate(["ทั้งหมด", "รออนุมัติ", "อนุมัติแล้ว", "ปฏิเสธ"]):
    x = ox + 0.3 + i * 1.0
    wrect(s, x, oy + 0.92, 0.92, 0.28, BLUE if i == 0 else WHITE, None if i == 0 else WFL, rounded=True)
    wt(s, x, oy + 0.92, 0.92, 0.28, tb, 9, True, WHITE if i == 0 else WFT, PP_ALIGN.CENTER)
table(s, ox + 0.3, oy + 1.35, [0.8, 1.3, 1.5, 1.9, 1.3],
      [["เลขที่", "ผู้จอง", "ห้องหรืออุปกรณ์", "วันเวลา", "สถานะ"],
       ["0124", "นิสิต ก.", "ห้อง 401", "10 ต.ค. 09:00-12:00", "รออนุมัติ"],
       ["0125", "อาจารย์ ข.", "โพรเจกเตอร์ 2", "10 ต.ค. 13:00-16:00", "อนุมัติแล้ว"],
       ["0126", "เจ้าหน้าที่ ค.", "ห้องประชุม 1", "11 ต.ค. 09:00-11:00", "อนุมัติแล้ว"]],
      rowh=0.32, size=9, head=(WFF, WFT), stripe=False, first_blue=False, hl=(1,), hl_fill=TINT)
wrect(s, ox + 0.3, oy + 2.8, 6.8, 1.15, WFF, None)
wt(s, ox + 0.45, oy + 2.85, 4.0, 0.28, "รายละเอียดคำขอ 0124", 10, True)
wt(s, ox + 0.45, oy + 3.1, 4.5, 0.25, "วัตถุประสงค์: ประชุมชมรม · 40 คน", 9, color=MUTED)
wrect(s, ox + 0.45, oy + 3.42, 4.2, 0.38)
wt(s, ox + 0.53, oy + 3.42, 4.1, 0.38, "เหตุผล (ต้องกรอกเมื่อปฏิเสธ)", 9, color=MUTED)
btn(s, ox + 4.85, oy + 3.46, 1.0, "ปฏิเสธ", False)
btn(s, ox + 5.95, oy + 3.46, 1.0, "อนุมัติ")
region(s, ox + 0.22, oy + 0.85, 4.1, 0.42, 1)
region(s, ox + 0.22, oy + 1.29, 6.96, 1.38, 2)
region(s, ox + 0.37, oy + 3.36, 6.68, 0.5, 3)
wf_panel(s, "เจ้าหน้าที่",
         ["กรองตามสถานะ หาคำขอที่ต้องพิจารณาได้เร็ว",
          ["ทุกคำขอรวมอยู่ในตารางเดียว", "ไม่ต้องไล่ดูหลายช่องทาง"],
          "ปฏิเสธต้องใส่เหตุผล ผู้ขอรู้ว่าต้องแก้อะไร"])

# 16 ─ โครงหน้าจอ 4: บันทึกยืม-คืน (อุปกรณ์) และเช็กอิน-เช็กเอาต์ (ห้อง) ในหน้าเดียว
s = content("02 ทางออก · โครงหน้าจอ 4/4", "หน้าบันทึกยืม-คืน/เช็กอิน-เช็กเอาต์")
browser(s, "ยืม-คืน/เช็กอิน-เช็กเอาต์", staff=True)
wt(s, ox + 0.3, oy + 0.5, 4.0, 0.35, "บันทึกการยืม-คืน/เช็กอิน-เช็กเอาต์", 12, True)
x = ox + 0.3
for i, (tb, w) in enumerate([("ทั้งหมด", 0.8), ("รอยืม/เช็กอิน", 1.2), ("กำลังใช้", 0.9), ("เกินกำหนด", 1.0), ("คืน/เช็กเอาต์แล้ว", 1.6)]):
    wrect(s, x, oy + 0.92, w, 0.28, BLUE if i == 0 else WHITE, None if i == 0 else WFL, rounded=True)
    wt(s, x, oy + 0.92, w, 0.28, tb, 9, True, WHITE if i == 0 else WFT, PP_ALIGN.CENTER)
    x += w + 0.08
lend_cols = [0.5, 0.85, 1.1, 1.0, 1.0, 1.05, 1.3]
table(s, ox + 0.3, oy + 1.35, lend_cols,
      [["เลขที่", "ผู้จอง", "ห้องหรืออุปกรณ์", "กำหนดสิ้นสุด", "ยืม/เช็กอิน", "คืน/เช็กเอาต์", "สถานะ"],
       ["0128", "นิสิต จ.", "กล้องถ่ายวิดีโอ", "11 ต.ค. 12:00", "", "", "รอยืม/เช็กอิน"],
       ["0126", "เจ้าหน้าที่ ค.", "ห้องประชุม 1", "11 ต.ค. 11:00", "", "", "รอยืม/เช็กอิน"],
       ["0125", "อาจารย์ ข.", "โพรเจกเตอร์ 2", "10 ต.ค. 16:00", "10 ต.ค. 13:00", "", "กำลังใช้"],
       ["0127", "นิสิต ง.", "ไมโครโฟน 1", "10 ต.ค. 12:00", "10 ต.ค. 09:05", "", "เกินกำหนด"],
       ["0122", "อาจารย์ ช.", "ห้อง 402", "9 ต.ค. 12:00", "9 ต.ค. 08:55", "9 ต.ค. 12:00", "คืน/เช็กเอาต์แล้ว"],
       ["0121", "นิสิต ฉ.", "โพรเจกเตอร์ 1", "9 ต.ค. 16:00", "9 ต.ค. 13:00", "9 ต.ค. 15:40", "คืน/เช็กเอาต์แล้ว"],
       ["0119", "นิสิต ซ.", "ห้องประชุม 1", "8 ต.ค. 16:00", "8 ต.ค. 13:10", "8 ต.ค. 15:50", "คืน/เช็กเอาต์แล้ว"]],
      rowh=0.32, size=9, head=(WFF, WFT), stripe=False, first_blue=False)
x3 = 0.3 + sum(lend_cols[:4])
region(s, ox + 0.22, oy + 0.85, 5.98, 0.42, 1)
region(s, ox + 0.22, oy + 1.29, x3 - 0.27, 2.68, 2)
region(s, ox + x3 + 0.03, oy + 1.29, sum(lend_cols[4:]) + 0.05, 2.68, 3)
wf_panel(s, "เจ้าหน้าที่",
         ["แยกตามสถานะ เห็นว่าอยู่ขั้นไหน",
          "เห็นผู้จองและกำหนดสิ้นสุดของทุกรายการ",
          "เห็นเวลายืมหรือเช็กอิน และคืนหรือเช็กเอาต์ ของทุกรายการ"])

# 10 ─ ผังงานของระบบ
# สัญลักษณ์: วงรี = เริ่ม/จบ, สี่เหลี่ยมด้านขนาน = รับ/แสดงข้อมูล, สี่เหลี่ยม = ประมวลผล, ข้าวหลามตัด = ตัดสินใจ, วงกลม = จุดต่อ
# ทุกคำขอต้องผ่านเจ้าหน้าที่ ไม่ว่างกลับไปค้นหา ถูกปฏิเสธแล้วจบ
s = content("02 ทางออก", "ผังงาน")
colw = CW / 8
YT, YA, YB, YC, YD = 2.9, 3.75, 4.9, 5.9, 6.45   # YD = lane for the reject path
PW, PH, DW, DH, TW, TH, RJ = 1.3, 0.7, 1.4, 0.9, 1.0, 0.5, 0.14
hp, vp, hd, vd, ht, vt = PW / 2, PH / 2, DW / 2, DH / 2, TW / 2, TH / 2
hi = hp - 0.08      # parallelogram: slanted sides sit slightly inside the bounding box
cx = lambda i: L + colw * (i + 0.5)


def node_text(i, yc, ts, w, h, size, color):
    text(s, cx(i) - w / 2, yc - h / 2, w, h, [[(t, {"gap": 0})] for t in ts], size, True, color, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)


def proc(i, yc, ts):
    box(s, cx(i) - hp, yc - vp, PW, PH, BLUE, shape=MSO_SHAPE.RECTANGLE)
    node_text(i, yc, ts, PW - 0.06, PH, 11, WHITE)


def io(i, yc, ts):
    box(s, cx(i) - hp, yc - vp, PW, PH, TINT, BLUE, MSO_SHAPE.FLOWCHART_DATA)
    node_text(i, yc, ts, PW - 0.3, PH, 11, BLUE)


def dec(i, yc, ts):
    box(s, cx(i) - hd, yc - vd, DW, DH, YTINT, YEL, MSO_SHAPE.DIAMOND)
    node_text(i, yc, ts, 1.1, DH, 10.5, INK)


def term(i, yc, t):
    box(s, cx(i) - ht, yc - vt, TW, TH, WHITE, INK, MSO_SHAPE.FLOWCHART_TERMINATOR)
    node_text(i, yc, [t], TW, TH, 13, INK)


def joint(i, yc):
    box(s, cx(i) - RJ, yc - RJ, 2 * RJ, 2 * RJ, WHITE, INK, MSO_SHAPE.OVAL)


def lab(x, y, t):
    text(s, x, y, 0.5, 0.22, t, 11, True, MUTED)


# แถว A: ค้นหา รับคำขอ บันทึก แจ้งเจ้าหน้าที่ (ซ้ายไปขวา)
line(s, [(cx(0) + ht, YA), (cx(1) - RJ, YA)])
line(s, [(cx(1) + RJ, YA), (cx(2) - hi, YA)])
line(s, [(cx(2) + hi, YA), (cx(3) - hp, YA)])
line(s, [(cx(3) + hp, YA), (cx(4) - hd, YA)])
line(s, [(cx(4) + hd, YA), (cx(5) - hi, YA)])
lab(cx(4) + hd - 0.05, YA - 0.28, "ใช่")
line(s, [(cx(5) + hi, YA), (cx(6) - hp, YA)])
line(s, [(cx(6) + hp, YA), (cx(7) - hi, YA)])
# ไม่ว่าง: แสดงทางเลือก แล้วกลับไปรับเงื่อนไขค้นหาใหม่
line(s, [(cx(4), YA - vd), (cx(4), YT), (cx(3) + hi, YT)])
lab(cx(4) + 0.07, YA - vd - 0.26, "ไม่")
line(s, [(cx(3) - hi, YT), (cx(1), YT), (cx(1), YA - RJ)])
# แถว B: ทุกคำขอผ่านเจ้าหน้าที่ (ขวาไปซ้าย)
line(s, [(cx(7), YA + vp), (cx(7), YB - vd)])
line(s, [(cx(7) - hd, YB), (cx(6) + hi, YB)])
lab(cx(7) - hd - 0.12, YB - 0.3, "ใช่")
line(s, [(cx(6) - hi, YB), (cx(5) + hd, YB)])
# อุปกรณ์: บันทึกการยืมและคืน ผ่านเจ้าหน้าที่ (ผู้จ่ายและรับคืน) / ห้อง: บันทึกเช็กอินและเช็กเอาต์ (สองทางเลือก)
line(s, [(cx(5) - hd, YB), (cx(4) + hp, YB)])
lab(cx(5) - hd - 0.12, YB - 0.3, "ใช่")
line(s, [(cx(4) - hp, YB), (cx(1) + RJ, YB)])
line(s, [(cx(5), YB + vd), (cx(5), YC - vp)])
lab(cx(5) + 0.07, YB + vd + 0.0, "ไม่")
line(s, [(cx(5) - hp, YC), (cx(2) + RJ, YC)])
line(s, [(cx(2) - RJ, YC), (cx(1), YC), (cx(1), YB + RJ)])
line(s, [(cx(1) - RJ, YB), (cx(0) + ht, YB)])
# ปฏิเสธ: แจ้งเหตุผล แล้วไปจบ
line(s, [(cx(7), YB + vd), (cx(7), YC - vp)])
lab(cx(7) + 0.07, YB + vd + 0.0, "ไม่")
line(s, [(cx(7), YC + vp), (cx(7), YD), (cx(2), YD), (cx(2), YC + RJ)])

term(0, YA, "เริ่ม")
joint(1, YA)
io(2, YA, ["รับเงื่อนไข", "ค้นหา"])
proc(3, YA, ["ตรวจสถานะว่าง"])
dec(4, YA, ["ว่างไหม?"])
io(5, YA, ["รับคำขอจอง"])
proc(6, YA, ["บันทึกคำขอ"])
io(7, YA, ["แจ้ง", "เจ้าหน้าที่"])
io(3, YT, ["แสดง", "ทางเลือก"])
dec(7, YB, ["อนุมัติ?"])
io(6, YB, ["แจ้งผล", "การจอง"])
dec(5, YB, ["เป็นอุปกรณ์", "ไหม?"])
box(s, cx(4) - hp, YB - vp, PW, PH, BLUE, shape=MSO_SHAPE.RECTANGLE)
node_text(4, YB, ["บันทึกการยืม", "และคืน", "ผ่านเจ้าหน้าที่"], PW - 0.04, PH, 10, WHITE)
joint(1, YB)
term(0, YB, "จบ")
proc(5, YC, ["บันทึกเช็กอิน", "และเช็กเอาต์"])
io(7, YC, ["แจ้งเหตุผล", "ที่ปฏิเสธ"])
joint(2, YC)

# 8 ─ ความต้องการ: แบ่งตามผู้ได้ประโยชน์ ส่วนกลางคือได้ทั้งสองฝ่าย
s = content("03 ความต้องการ", "ระบบช่วยได้หลายอย่าง โดยมี 4 อย่างที่สำคัญที่สุด")
zones = [("นิสิตและอาจารย์", [("เห็นสถานะว่างทันที", 1), ("ค้นหาตามเงื่อนไข", 0), ("ส่งคำขอออนไลน์", 0), ("ติดตามสถานะคำขอ", 0), ("แก้ไขหรือยกเลิกได้", 0)]),
         ("ทั้งสองฝ่าย", [("ตรวจการชนอัตโนมัติ", 1), ("อนุมัติพร้อมเหตุผล", 0), ("แจ้งเตือนอัตโนมัติ", 0)]),
         ("เจ้าหน้าที่", [("บันทึกยืม-คืน/เช็กอิน-เช็กเอาต์", 1), ("ดึงข้อมูลการใช้งานได้", 1), ("จัดการห้องและอุปกรณ์", 0), ("บันทึกประวัติทุกครั้ง", 0)])]
zw, zy, zh = CW / 3, 2.95, 3.5
box(s, L, zy, 2 * zw, zh, TINT)                                   # ฝั่งผู้ใช้
box(s, L + zw, zy, 2 * zw, zh, TINT)                              # ฝั่งเจ้าหน้าที่
box(s, L + zw, zy, zw, zh, RGBColor(0xC9, 0xD9, 0xF0), shape=MSO_SHAPE.RECTANGLE)   # ส่วนที่ซ้อนกัน
for i, (hd_, items) in enumerate(zones):
    x = L + i * zw
    text(s, x + 0.35, zy + 0.2, zw - 0.7, 0.45, hd_, 20, True, BLUE)
    for j, (t, focus) in enumerate(items):
        y = zy + 0.85 + j * 0.5
        if focus:
            box(s, x + 0.25, y, zw - 0.5, 0.42, YTINT, YEL)
        text(s, x + 0.35, y, zw - 0.5, 0.42, t, 15, bool(focus), INK if focus else MUTED, anchor=MSO_ANCHOR.MIDDLE)

# 15 ─ สรุป: มีแล้ว + ยังไม่มี (คำถามจากหน้าก่อน) → ต้องไปเก็บต่อ (ข้อมูลจากหน้าก่อน)
s = content("04 สรุป", "ขั้นต่อไป")
concl = [("สิ่งที่มีแล้ว", ["ปัญหา และผู้ใช้ระบบหลัก", "สิ่งที่ระบบต้องทำ", "ผังงาน และโครงหน้าจอ"], False),
         ("คำถามที่ยังตอบไม่ได้", ["ปัญหาเกิดบ่อยแค่ไหน", "เจ้าหน้าที่เสียเวลาเท่าไร", "สำเร็จวัดจากอะไร"], False),
         ("สิ่งที่ต้องไปเก็บต่อ", ["จำนวนคำขอ และครั้งที่จองชน", "เวลาที่ใช้ต่อหนึ่งคำขอ", "ตัวเลขเป้าหมายจากผู้บริหาร"], True)]
gap = 0.6
kw = (CW - 2 * gap) / 3
for i, (h, items, nxt) in enumerate(concl):
    x = L + i * (kw + gap)
    box(s, x, 2.95, kw, 2.35, YTINT if nxt else TINT)
    text(s, x + 0.3, 3.2, kw - 0.6, 0.5, h, 22, True, BLUE)
    bullets(s, x + 0.3, 3.85, kw - 0.6, 1.4, items, 17)
    if i < 2:
        text(s, x + kw, 2.95, gap, 2.35, "+" if i == 0 else "=", 30, True, MUTED, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)

page_numbers()
shift_up()
no_shadows()
prs.save("slides-group1.pptx")
strip_theme_shadows("slides-group1.pptx")
print("saved", len(prs.slides), "slides")
