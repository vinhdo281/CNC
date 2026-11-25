import copy, shutil
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn

import os
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)  # thư mục CNC
SRC = os.path.join(ROOT, "Ref", "Thầy hùng", "CNC.pptx")   # mẫu bộ nhận diện HUST
OUT = os.path.join(ROOT, "Slide_Ga_kep_CNC.pptx")
IMG = os.path.join(HERE, "images") + os.sep

RED = RGBColor(0xC0, 0x00, 0x00)
GOLD = RGBColor(0xE0, 0x9B, 0x13)
INK = RGBColor(0x26, 0x26, 0x26)
GREY = RGBColor(0x59, 0x59, 0x59)
LIGHT = RGBColor(0xF6, 0xEC, 0xEC)
MID = RGBColor(0xD9, 0xD9, 0xD9)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

prs = Presentation(SRC)

# --- keep cover (1) and thank-you (11); drop the 9 empty template slides between
lst = prs.slides._sldIdLst
ids = list(lst)
for sid in ids[1:10]:
    prs.part.drop_rel(sid.rId)
    lst.remove(sid)
thank = list(lst)[1]

layout = next(l for l in prs.slide_layouts if l.name == "Comparison")
num_ph = next(p for p in layout.placeholders if p.placeholder_format.idx == 12)


def new_slide(title, notes):
    s = prs.slides.add_slide(layout)
    s.shapes.title.text = title
    for r in s.shapes.title.text_frame.paragraphs[0].runs:
        r.font.size = Pt(26)
        r.font.bold = True
    s.shapes._spTree.append(copy.deepcopy(num_ph._element))  # slide number like template
    s.notes_slide.notes_text_frame.text = notes
    return s


def body(slide):
    return next(p for p in slide.placeholders if p.placeholder_format.idx == 1)


def set_bullets(ph, items, x, y, w, h, size=22):
    ph.left, ph.top, ph.width, ph.height = Inches(x), Inches(y), Inches(w), Inches(h)
    tf = ph.text_frame
    tf.word_wrap = True
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if isinstance(it, tuple):
            head, rest = it
            r = p.add_run(); r.text = head
            r.font.bold = True; r.font.color.rgb = RED
            r2 = p.add_run(); r2.text = rest
            for rr in (r, r2): rr.font.size = Pt(size)
            r2.font.color.rgb = INK
        else:
            r = p.add_run(); r.text = it
            r.font.size = Pt(size); r.font.color.rgb = INK
        p.space_after = Pt(10)


def box(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, lw=1.0):
    sh = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None: sh.fill.background()
    else:
        sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if line is None: sh.line.fill.background()
    else:
        sh.line.color.rgb = line; sh.line.width = Pt(lw)
    sh.shadow.inherit = False
    return sh


def text(slide, x, y, w, h, s, size=16, bold=False, color=INK, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, shape=None):
    sh = shape or slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = sh.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.1)
    lines = s if isinstance(s, list) else [s]
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run(); r.text = ln
        r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color
    return sh


def line(slide, x1, y1, x2, y2, color=RED, w=2.5, arrow=False):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = color; c.line.width = Pt(w)
    if arrow:
        ln = c.line._get_or_add_ln()
        tail = ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"})
        ln.append(tail)
    return c


def picture(slide, name, x, y, w=None, h=None):
    kw = {}
    if w: kw["width"] = Inches(w)
    if h: kw["height"] = Inches(h)
    pic = slide.shapes.add_picture(IMG + name, Inches(x), Inches(y), **kw)
    pic.line.color.rgb = MID; pic.line.width = Pt(0.75)
    return pic


# ============ Slide A: vai trò & 4 nhiệm vụ ============
s = new_slide(
    "Gá kẹp quyết định độ chính xác của chi tiết phay CNC",
    "[~2 phút] Mở đầu bằng câu hỏi: máy chính xác 0,001 mm mà phôi kẹp lệch thì kết quả ra sao? "
    "Gá kẹp của máy phay lắp trên bàn máy. "
    "Bốn nhiệm vụ lấy từ bài giảng chương 2.3. Máy CNC tập trung nhiều nguyên công trong một lần gá nên "
    "một sai số gá sẽ lặp lại cho cả loạt chi tiết.")
s.shapes._spTree.remove(body(s)._element)
cards = [
    ("1", "Định vị chính xác", "Xác định đúng vị trí phôi trên bàn máy"),
    ("2", "Kẹp chặt", "Chống lại lực cắt khi gia công"),
    ("3", "Thay phôi nhanh", "Thao tác đơn giản, ít dừng máy"),
    ("4", "Lặp lại chính xác", "Phôi sau được gá giống phôi trước"),
]
cw, gap, x0, y0 = 2.95, 0.26, 0.45, 1.35
for i, (n, h, d) in enumerate(cards):
    x = x0 + i * (cw + gap)
    box(s, x, y0, cw, 2.9, fill=LIGHT)
    box(s, x, y0, cw, 0.09, fill=RED)
    c = box(s, x + 0.2, y0 + 0.3, 0.7, 0.7, fill=RED, shape=MSO_SHAPE.OVAL)
    text(s, 0, 0, 0, 0, n, size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, shape=c)
    text(s, x + 0.1, y0 + 1.15, cw - 0.2, 0.6, h, size=20, bold=True, color=RED)
    text(s, x + 0.1, y0 + 1.75, cw - 0.2, 1.0, d, size=16, color=INK)
box(s, 0.45, 4.6, 12.4, 0.07, fill=GOLD)
text(s, 0.45, 4.85, 12.4, 0.5, "Vị trí lắp: trên bàn máy phay",
     size=20, color=INK)
text(s, 0.45, 5.5, 12.4, 0.9,
     ["Máy CNC gia công nhiều nguyên công trong một lần gá",
      "→ gá sai một lần thì sai cả loạt chi tiết"], size=20, bold=True, color=RED)

# ============ Slide B: định vị vs kẹp chặt ============
s = new_slide(
    "Định vị và kẹp chặt là hai việc khác nhau",
    "[~2,5 phút] Slide quan trọng nhất với người nghe là giáo viên: sinh viên hay nhầm hai khái niệm. "
    "Định vị trả lời 'phôi nằm ở đâu', kẹp chặt trả lời 'phôi có đứng yên không'. "
    "Điểm không chi tiết do người lập trình chọn; trục Z theo trục chính (từ chi tiết lên dao), "
    "trục X song song mặt kẹp (chương 3.1). Giáo trình trang 42: việc kẹp cần nhanh, thông thoáng và "
    "định vị lại phải chính xác khi đổi chi tiết.")
s.shapes._spTree.remove(body(s)._element)
for k, (h, d) in enumerate([
    ("ĐỊNH VỊ", ["Xác định vị trí phôi", "so với gốc toạ độ", "→ phôi nằm ở đâu?"]),
    ("KẸP CHẶT", ["Giữ phôi không xê dịch", "khi chịu lực cắt", "→ phôi có đứng yên không?"]),
]):
    y = 1.35 + k * 2.75
    box(s, 0.45, y, 5.4, 2.5, fill=LIGHT)
    box(s, 0.45, y, 0.12, 2.5, fill=RED)
    text(s, 0.75, y + 0.15, 5.0, 0.5, h, size=22, bold=True, color=RED)
    text(s, 0.75, y + 0.75, 5.0, 1.6, d, size=19, color=INK)
# schematic (side view)
sx, sy = 6.4, 1.35
box(s, sx, sy, 6.5, 5.2, fill=None, line=MID)
text(s, sx + 0.1, sy + 0.05, 6.3, 0.4, "Sơ đồ minh hoạ (nhìn ngang)", size=14, color=GREY)
table_ = box(s, sx + 0.5, sy + 4.0, 5.5, 0.55, fill=RGBColor(0x7F, 0x7F, 0x7F))
text(s, 0, 0, 0, 0, "Bàn máy", size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
     anchor=MSO_ANCHOR.MIDDLE, shape=table_)
part = box(s, sx + 1.4, sy + 2.2, 3.0, 1.8, fill=RGBColor(0xBF, 0xD7, 0xEA), line=INK, lw=1.5)
text(s, 0, 0, 0, 0, "Phôi", size=18, bold=True, color=INK, align=PP_ALIGN.CENTER,
     anchor=MSO_ANCHOR.MIDDLE, shape=part)
box(s, sx + 1.2, sy + 2.2, 0.2, 1.8, fill=GOLD)          # mặt chuẩn (bên trái)
text(s, sx + 0.1, sy + 3.55, 1.3, 0.4, "Mặt chuẩn", size=13, bold=True, color=GOLD)
line(s, sx + 5.2, sy + 3.1, sx + 4.45, sy + 3.1, color=RED, w=4, arrow=True)   # lực kẹp ngang
line(s, sx + 2.9, sy + 1.45, sx + 2.9, sy + 2.15, color=RED, w=4, arrow=True)  # lực kẹp đứng
text(s, sx + 3.0, sy + 1.3, 2.2, 0.4, "Lực kẹp", size=15, bold=True, color=RED)
# origin + axes
ox, oy = sx + 1.4, sy + 2.2
dot = box(s, ox - 0.09, oy - 0.09, 0.18, 0.18, fill=RED, shape=MSO_SHAPE.OVAL)
line(s, ox, oy, ox + 1.0, oy - 0.0001, color=INK, w=1.75, arrow=True)  # X (song song mặt kẹp)
line(s, ox, oy, ox, oy - 0.9, color=INK, w=1.75, arrow=True)           # Z (trục chính)
text(s, ox + 0.95, oy - 0.42, 0.5, 0.35, "X", size=15, bold=True)
text(s, ox - 0.5, oy - 0.95, 0.5, 0.35, "Z", size=15, bold=True)
text(s, sx + 0.1, sy + 0.65, 3.4, 0.4, "Gốc phôi (điểm không chi tiết)", size=13, color=RED, bold=True)
line(s, sx + 0.7, sy + 1.05, ox - 0.1, oy - 0.1, color=RED, w=1.25, arrow=True)

# ============ Slide C: kẹp cơ khí ============
s = new_slide(
    "Kẹp cơ khí: ê tô cho phôi khối, đòn kẹp – gối đỡ cho phôi lớn",
    "[~3 phút] Giới thiệu hai cách kẹp cơ khí phổ biến nhất trên máy phay. Hình bên phải là hình 2.11 "
    "(giáo trình, trang 43): chỉ lần lượt vào các chi tiết đánh số – chi tiết gia công, đòn kẹp, "
    "gối đỡ, bàn máy; bulông đầu chữ T đi vào rãnh T của bàn máy. Phay đơn giản thì ê tô có mỏ kẹp điều chỉnh "
    "được là đủ; chi tiết phức tạp thì dùng đồ gá riêng hoặc đồ gá tổ hợp từ bộ phận tiêu chuẩn "
    "để gia công hoàn thiện không cần đổi gá.")
set_bullets(body(s), [
    ("Ê tô đa năng: ", "kẹp tay, hợp phôi dạng khối"),
    ("Đòn kẹp + gối đỡ + bulông đầu chữ T: ", "kẹp phôi lớn, hình dạng bất kỳ"),
    ("Phay đơn giản: ", "ê tô là đủ"),
    ("Phôi phức tạp: ", "đồ gá riêng hoặc gá tổ hợp, không cần đổi gá"),
    ("Phay nhiều mặt: ", "gá ít lần nhất"),
], 0.45, 1.3, 5.9, 5.3, size=20)
picture(s, "p-004.png", 7.0, 1.3, w=5.6)
text(s, 7.0, 3.78, 5.6, 0.35, "Hình 2.11b: đòn kẹp – gối đỡ trên bàn máy", size=12, color=GREY)
picture(s, "p-005.png", 7.0, 4.2, w=5.6)
text(s, 7.0, 6.38, 5.6, 0.35, "Hình 2.11c: gối tựa, đòn kẹp, gối đỡ (nguồn: giáo trình, trang 43)", size=12, color=GREY)

# ============ Slide D: nguồn lực kẹp ============
s = new_slide(
    "Nguồn lực kẹp quyết định lực kẹp và thao tác",
    "[~3 phút] Bảng so sánh bốn kiểu kẹp theo nguồn lực, lấy từ bài giảng chương 2.3 (máy phay CNC). "
    "Không nêu số lực cụ thể vì phụ thuộc vật liệu, kích thước phôi và chiều sâu cắt – bài giảng "
    "cũng chỉ nói lực kẹp phụ thuộc các yếu tố này. Ý chỉ dùng bàn từ cho phôi có từ tính là phần bổ sung ngoài "
    "tài liệu, cần xác nhận lại với tài liệu của thầy.")
s.shapes._spTree.remove(body(s)._element)
rows = [
    ("Kiểu kẹp", "Đặc điểm"),
    ("Ê tô kẹp tay", "Đa năng; lực kẹp bằng tay"),
    ("Ê tô khí nén", "Lực kẹp không lớn"),
    ("Ê tô thuỷ lực", "Lực kẹp lớn, điều chỉnh được; loại tự định tâm hoặc không định tâm"),
    ("Bàn từ tính", "Kẹp bằng lực từ, không cần đòn kẹp *"),
]
gt = s.shapes.add_table(len(rows), 2, Inches(0.45), Inches(1.35), Inches(8.3), Inches(4.4))
tbl = gt.table
tbl.columns[0].width = Inches(2.6)
tbl.columns[1].width = Inches(5.7)
for r, (a, b) in enumerate(rows):
    tbl.rows[r].height = Inches(0.7 if r == 0 else 0.92)
    for c, v in enumerate((a, b)):
        cell = tbl.cell(r, c)
        cell.text = v
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.fill.solid()
        cell.fill.fore_color.rgb = RED if r == 0 else (LIGHT if r % 2 else WHITE)
        for p in cell.text_frame.paragraphs:
            for run in p.runs:
                run.font.size = Pt(19 if r else 20)
                run.font.bold = (r == 0 or c == 0)
                run.font.color.rgb = WHITE if r == 0 else (RED if c == 0 else INK)
text(s, 0.45, 5.9, 8.3, 0.5, "* Bổ sung ngoài tài liệu: chỉ dùng cho phôi có từ tính", size=13, color=GREY)
box(s, 9.1, 1.35, 3.8, 4.4, fill=LIGHT)
box(s, 9.1, 1.35, 3.8, 0.09, fill=GOLD)
text(s, 9.2, 1.6, 3.6, 0.5, "Chọn lực kẹp theo", size=20, bold=True, color=RED)
for k, it in enumerate(["Trọng lượng phôi", "Vật liệu", "Kích thước phôi", "Chiều sâu cắt"]):
    text(s, 9.2, 2.3 + k * 0.8, 3.6, 0.6, "• " + it, size=21, color=INK)

# ============ Slide E: hiện đại ============
s = new_slide(
    "Hướng hiện đại: gá ngoài máy, tự động đưa vào gia công",
    "[~2,5 phút] Mục tiêu chung là giảm thời gian chuẩn bị và dừng máy – chương 1.3 nói máy CNC có thể "
    "chuẩn bị phôi độc lập và gá lắp nhanh. Hình bên phải (bài giảng chương 2.3, trang 15) là ô gia công "
    "có robot đứng riêng và bàn phôi (workpiece table); chú thích gốc bằng tiếng Anh. "
    "Giáo trình trang 43: bàn gá chuẩn (palette) ngày càng dùng nhiều vì phôi được gá ngoài vùng làm việc "
    "của người vận hành, sau đó tự động đưa vào vị trí chính xác.")
set_bullets(body(s), [
    ("Hệ gá modul: ", "ghép từ bộ phận tiêu chuẩn, không cần đổi gá"),
    ("Bàn quay hai vị trí: ", "gá phôi bên ngoài, bàn quay tự đưa vào máy"),
    ("Bàn gá chuẩn (palette): ", "chuẩn bị phôi ngoài vùng gia công"),
    ("Robot + kẹp tự động: ", "cơ cấu kẹp đóng mở bằng thuỷ lực hoặc khí nén"),
], 0.45, 1.3, 6.0, 5.3, size=21)
picture(s, "p-006.png", 6.8, 1.45, w=6.1)
text(s, 6.8, 5.5, 6.1, 0.6,
     "Ô gia công CNC có robot và bàn phôi (workpiece table). Nguồn: bài giảng Machine-tool and Tribology, chương 2.3",
     size=12, color=GREY)

# ============ Slide F: tóm tắt ============
s = new_slide(
    "Tóm tắt: chọn gá kẹp theo phôi và sản lượng",
    "[~2 phút] Nhắc lại công thức: gá kẹp = định vị + kẹp chặt. Ba cột là gợi ý chọn phương án theo ba tình huống. "
    "Dành 1–2 phút cuối cho câu hỏi thảo luận và hỏi đáp của giáo viên.")
s.shapes._spTree.remove(body(s)._element)
text(s, 0.45, 1.25, 12.4, 0.6, "Gá kẹp = Định vị + Kẹp chặt", size=28, bold=True, color=RED,
     align=PP_ALIGN.CENTER)
trio = [
    ("Phôi đơn giản", "Ê tô"),
    ("Phôi lớn, phức tạp", "Đòn kẹp – gối đỡ, đồ gá riêng hoặc gá tổ hợp"),
    ("Sản xuất loạt", "Gá modul, palette, robot"),
]
bw, bg = 3.85, 0.425
for i, (a, b) in enumerate(trio):
    x = 0.45 + i * (bw + bg)
    box(s, x, 2.15, bw, 0.8, fill=RED)
    text(s, x, 2.15, bw, 0.8, a, size=21, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE)
    box(s, x, 2.95, bw, 1.6, fill=LIGHT)
    text(s, x + 0.1, 2.95, bw - 0.2, 1.6, b, size=20, color=INK, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE)
box(s, 0.45, 5.0, 12.4, 1.5, fill=None, line=GOLD, lw=2.5)
text(s, 0.65, 5.05, 12.0, 0.45, "Câu hỏi thảo luận", size=18, bold=True, color=RED)
text(s, 0.65, 5.5, 12.0, 0.95,
     "Với một chi tiết cụ thể, nhóm chọn cách gá nào và vì sao?", size=22, color=INK)

# --- put thank-you slide last
lst.remove(thank)
lst.append(thank)

# ============ Comment (bình luận review) cho từng slide ============
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
CT = "application/vnd.openxmlformats-officedocument.presentationml."
COMMENTS = [
    "Tên đề tài và danh sách nhóm giữ nguyên từ mẫu CNC.pptx. Kiểm tra lại tên và MSSV trước khi trình bày.",
    "Nguồn: bài giảng Machine-tool and Tribology, chương 2.3 (slide 10): 4 nhiệm vụ của cơ cấu kẹp chặt. "
    "Số 0,001 mm lấy từ chương 1.3. Gợi ý mở bài bằng câu hỏi: máy chính xác thế mà gá lệch thì sao?",
    "Gốc phôi và trục X, Z lấy từ chương 3.1. Sơ đồ là hình minh hoạ tự vẽ, không phải hình trong giáo trình. "
    "Nên nhấn mạnh slide này với giáo viên vì sinh viên hay nhầm định vị với kẹp chặt.",
    "Hình 2.11b, c lấy từ bài giảng chương 2.3 (trang 14), cùng hình trong giáo trình trang 43. "
    "Chưa có hình ê tô (2.11a) vì ảnh chụp sách bị nghiêng và có bóng. Nếu có ảnh rõ hơn thì thêm vào.",
    "Ba hàng đầu theo bài giảng chương 2.3 (trang 14). Ý 'bàn từ tính chỉ dùng cho phôi có từ tính' là bổ sung "
    "ngoài tài liệu, cần xác nhận trước khi trình bày. Không đưa số lực kẹp vì tài liệu không nêu.",
    "Hệ gá modul và bàn quay hai vị trí: bài giảng chương 2.3 (trang 14). Robot + kẹp tự động: trang 15. "
    "Bàn gá chuẩn (palette): giáo trình trang 43. Chú thích trên hình gốc bằng tiếng Anh.",
    "Ba cột là gợi ý chọn phương án tổng hợp từ các slide trước, không phải quy tắc có trong tài liệu. "
    "Có thể đổi câu hỏi thảo luận cho khớp bài giảng của thầy.",
    "Nếu cần, thêm slide Tài liệu tham khảo (bài giảng Machine-tool and Tribology chương 1-3, "
    "giáo trình hình 2.10-2.11) hoặc slide Hỏi đáp trước slide này.",
]
slides_now = list(prs.slides)
assert len(slides_now) == len(COMMENTS)
n = len(COMMENTS)
authors = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<p:cmAuthorLst xmlns:p="{NS_P}"><p:cmAuthor id="0" name="Claude" initials="C" '
           f'lastIdx="{n}" clrIdx="0"/></p:cmAuthorLst>').encode("utf-8")
prs.part.relate_to(
    Part(PackURI("/ppt/commentAuthors.xml"), CT + "commentAuthors+xml", prs.part.package, authors),
    RT + "commentAuthors")
from xml.sax.saxutils import escape
for i, (sl, msg) in enumerate(zip(slides_now, COMMENTS), 1):
    xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<p:cmLst xmlns:p="{NS_P}"><p:cm authorId="0" dt="2026-11-25T15:30:00.000" idx="{i}">'
           f'<p:pos x="7200" y="60"/><p:text>{escape(msg)}</p:text></p:cm></p:cmLst>').encode("utf-8")
    sl.part.relate_to(
        Part(PackURI(f"/ppt/comments/comment{i}.xml"), CT + "comments+xml", prs.part.package, xml),
        RT + "comments")

prs.save(OUT)
print("saved", OUT, len(prs.slides._sldIdLst), "slides")
