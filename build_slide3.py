"""Slide 3 v17: redesigned 'Projekte und Projektportfoliomanagement bei STIHL'
for better board readability.

Layout:
- 4 large project-type cards (Produkt, Infrastruktur, Operations, Sonder)
  Each card: big number p.a. + KPIs + tool badges
- Cost-aggregation band: Sachkosten (Produkt + Infra) and Invest (Ops + Sonder)
- User band: 4.200 PIT for cards 1-3, Sonder has its own user counts
- Two summary bands below for IT-Demand-Mgmt and EV-Prozess
"""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

SRC = "/tmp/2026_05_18_VS_Zielbild_SNOW_SPM_v16.pptx"
DST = "/tmp/2026_05_18_VS_Zielbild_SNOW_SPM_v17.pptx"

ORANGE   = RGBColor(0xF0, 0x7F, 0x12)
ORANGE_LT= RGBColor(0xFD, 0xE6, 0xCC)
DARK     = RGBColor(0x1F, 0x29, 0x37)
GREY_TXT = RGBColor(0x4B, 0x55, 0x63)
LIGHT_BG = RGBColor(0xF7, 0xF7, 0xF7)
ROW_ALT  = RGBColor(0xEE, 0xF1, 0xF4)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
GREEN    = RGBColor(0x2E, 0x7D, 0x32)
NAVY     = RGBColor(0x2C, 0x3E, 0x50)
LIGHT_GREY = RGBColor(0xD1, 0xD5, 0xDB)
BLUE_HD  = RGBColor(0x3B, 0x6B, 0xA5)
RED_HL   = RGBColor(0xB9, 0x1C, 0x1C)
GREY_BG  = RGBColor(0xE9, 0xED, 0xF1)

U = 76200
def ex(u): return int(u * U)

prs = Presentation(SRC)
slide = prs.slides[2]

for shp in list(slide.shapes):
    sp = shp._element
    sp.getparent().remove(sp)

def add_rect(left, top, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE, line_w=1.0):
    s = slide.shapes.add_shape(shape, ex(left), ex(top), ex(w), ex(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line; s.line.width = Pt(line_w)
    s.shadow.inherit = False
    return s

def add_oval(left, top, w, h, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, ex(left), ex(top), ex(w), ex(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    s.line.fill.background()
    return s

def set_paragraphs(shape, paragraphs, *, default_size=10, default_color=DARK, default_align=PP_ALIGN.LEFT):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(30000); tf.margin_right = Emu(30000)
    tf.margin_top = Emu(10000); tf.margin_bottom = Emu(10000)
    tf.clear()
    for i, item in enumerate(paragraphs):
        if isinstance(item, str):
            text, st = item, {}
        else:
            text, st = item
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = st.get("align", default_align)
        run = p.add_run(); run.text = text
        run.font.name = "Calibri"
        run.font.size = Pt(st.get("size", default_size))
        run.font.bold = st.get("bold", False)
        run.font.italic = st.get("italic", False)
        run.font.color.rgb = st.get("color", default_color)

def add_text(left, top, w, h, paragraphs, *, anchor=MSO_ANCHOR.MIDDLE, **kw):
    tb = slide.shapes.add_textbox(ex(left), ex(top), ex(w), ex(h))
    tb.fill.background(); tb.line.fill.background()
    tb.text_frame.vertical_anchor = anchor
    set_paragraphs(tb, paragraphs, **kw)
    return tb

# ====== HEADER ======
add_rect(0, 0, 160, 6, DARK)
add_text(3, 0.4, 140, 2.6, [
    ("Projekte und Projektportfoliomanagement bei STIHL",
     dict(size=20, bold=True, color=WHITE)),
], anchor=MSO_ANCHOR.TOP)
add_text(3, 3.2, 140, 2.4, [
    ("Sechs Projekt-/Demand-Welten – Mengen, Aufwand, Nutzer und eingesetzte Software",
     dict(size=12, italic=True, color=LIGHT_GREY)),
], anchor=MSO_ANCHOR.TOP)
add_rect(150, 0, 10, 6, ORANGE)
add_text(150, 0, 10, 6, [
    ("Folie 3", dict(size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)),
])

# ====== FOUR PROJECT-TYPE CARDS ======
COL_X = 4
TOTAL_W = 152
N = 4
GAP = 1.5
CW = (TOTAL_W - (N - 1) * GAP) / N
def cardx(i): return COL_X + i * (CW + GAP)

card_top = 8
card_h = 36

projects = [
    ("PRODUKT-PROJEKTE",      "850",  "p.a.", ORANGE),
    ("INFRASTRUKTUR-PROJEKTE","76",   "p.a.", ORANGE),
    ("OPERATIONS-PROJEKTE",   "646",  "p.a.", ORANGE),
    ("SONDER-PROJEKTE",       "95",   "p.a. (davon 21 IT)", ORANGE),
]

for i, (name, num, sub, col) in enumerate(projects):
    x = cardx(i)
    # Card background
    add_rect(x, card_top, CW, card_h, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
             line=LIGHT_GREY, line_w=1.0)
    # Header band
    add_rect(x, card_top, CW, 5.5, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    # Number badge
    add_oval(x + 0.8, card_top + 0.8, 3.4, 3.4, NAVY)
    add_text(x + 0.8, card_top + 0.8, 3.4, 3.4, [
        (str(i+1), dict(size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)),
    ], anchor=MSO_ANCHOR.MIDDLE)
    # Title
    title_parts = name.split("-")
    add_text(x + 5.0, card_top + 0.6, CW - 5.5, 5.0, [
        (title_parts[0] + "-",
         dict(size=10.5, bold=True, color=WHITE)),
        (title_parts[1] if len(title_parts) > 1 else "",
         dict(size=10.5, bold=True, color=WHITE)),
    ], anchor=MSO_ANCHOR.MIDDLE)
    # Big number
    add_text(x, card_top + 7.0, CW, 7.5, [
        (num, dict(size=42, bold=True, color=DARK, align=PP_ALIGN.CENTER)),
    ], anchor=MSO_ANCHOR.MIDDLE)
    # Subtitle
    add_text(x, card_top + 14.6, CW, 2.4, [
        (sub, dict(size=10, italic=True, color=GREY_TXT, align=PP_ALIGN.CENTER)),
    ], anchor=MSO_ANCHOR.MIDDLE)

# ====== COST BAND (Sachkosten | Invest) ======
cost_y = card_top + 17.5
cost_h = 5.5

# Sachkosten spans cards 0-1
cost_x_a = cardx(0)
cost_w_a = (cardx(1) + CW) - cost_x_a
add_rect(cost_x_a, cost_y, cost_w_a, cost_h, GREY_BG,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(cost_x_a, cost_y + 0.3, cost_w_a, 2.4, [
    ("AUFWAND p.a. – SACHKOSTEN",
     dict(size=8.5, bold=True, color=GREY_TXT, align=PP_ALIGN.CENTER)),
], anchor=MSO_ANCHOR.MIDDLE)
add_text(cost_x_a, cost_y + 2.4, cost_w_a, 3.0, [
    ("~ 150 Mio. €",
     dict(size=18, bold=True, color=NAVY, align=PP_ALIGN.CENTER)),
], anchor=MSO_ANCHOR.MIDDLE)

# Invest spans cards 2-3
cost_x_b = cardx(2)
cost_w_b = (cardx(3) + CW) - cost_x_b
add_rect(cost_x_b, cost_y, cost_w_b, cost_h, GREY_BG,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(cost_x_b, cost_y + 0.3, cost_w_b, 2.4, [
    ("AUFWAND p.a. – INVEST",
     dict(size=8.5, bold=True, color=GREY_TXT, align=PP_ALIGN.CENTER)),
], anchor=MSO_ANCHOR.MIDDLE)
add_text(cost_x_b, cost_y + 2.4, cost_w_b, 3.0, [
    ("~ 420 Mio. €",
     dict(size=18, bold=True, color=NAVY, align=PP_ALIGN.CENTER)),
], anchor=MSO_ANCHOR.MIDDLE)

# ====== USER BAND ======
user_y = cost_y + cost_h + 0.6
user_h = 4.5
# 4.200 PIT spans cards 0-2
ux_a = cardx(0)
uw_a = (cardx(2) + CW) - ux_a
add_rect(ux_a, user_y, uw_a, user_h, ORANGE_LT,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(ux_a, user_y, uw_a, user_h, [
    ("# USER:  4.200 PIT  (davon 200 IT)",
     dict(size=12, bold=True, color=DARK, align=PP_ALIGN.CENTER)),
], anchor=MSO_ANCHOR.MIDDLE)
# Sonder: own user numbers
ux_b = cardx(3)
uw_b = CW
add_rect(ux_b, user_y, uw_b, user_h, ORANGE_LT,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(ux_b, user_y, uw_b, user_h, [
    ("PIT 200  ·  PowerApp 300",
     dict(size=11, bold=True, color=DARK, align=PP_ALIGN.CENTER)),
], anchor=MSO_ANCHOR.MIDDLE)

# ====== SOFTWARE BADGES ROW ======
sw_y = user_y + user_h + 0.6
sw_h = 4.8
add_text(COL_X, sw_y - 1.2, 30, 1.4, [
    ("SOFTWARE", dict(size=8.5, bold=True, color=GREY_TXT)),
])

def badge(left, top, w, h, label, fill=GREY_BG, text_color=DARK):
    add_rect(left, top, w, h, fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
             line=LIGHT_GREY, line_w=0.5)
    add_text(left, top, w, h, [
        (label, dict(size=9, bold=True, color=text_color, align=PP_ALIGN.CENTER)),
    ], anchor=MSO_ANCHOR.MIDDLE)

# Cards 0,1,2: PLANTA / PIT Produktivsystem
for i in range(3):
    x = cardx(i)
    badge(x + 1, sw_y, CW - 2, sw_h, "PLANTA · PIT Produktivsystem",
          fill=NAVY, text_color=WHITE)

# Card 3: multiple badges
x = cardx(3)
mini_w = (CW - 1.5) / 2
mini_gap = 0.5
mini_h = 2.2
badges_4 = [
    ("PowerApps",   "Power Apps"),
    ("SharePoint",  "SharePoint"),
    ("PLANTA / PIT","PLANTA · PIT"),
    ("MS Project",  "MS Project Online"),
]
for j, (_, lbl) in enumerate(badges_4):
    row = j // 2
    col = j % 2
    bx = x + 0.75 + col * (mini_w + mini_gap)
    by = sw_y + row * (mini_h + 0.4)
    badge(bx, by, mini_w, mini_h, lbl,
          fill=GREY_BG if j != 2 else NAVY,
          text_color=DARK if j != 2 else WHITE)

# ====== BOTTOM BANDS: IT-Demand + EV-Prozess ======
bot_y = card_top + card_h + 2.0
bot_h = (90 - bot_y - 6) / 2 - 1.0

# --- Band 5: IT-DEMAND-MANAGEMENT (ServiceNow) ---
band_h = bot_h
add_rect(COL_X, bot_y, TOTAL_W, band_h, LIGHT_BG,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
# Left label block
label_w = 30
add_rect(COL_X, bot_y, label_w, band_h, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_oval(COL_X + 1.5, bot_y + band_h/2 - 2.0, 4.0, 4.0, ORANGE)
add_text(COL_X + 1.5, bot_y + band_h/2 - 2.0, 4.0, 4.0, [
    ("5", dict(size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)),
], anchor=MSO_ANCHOR.MIDDLE)
add_text(COL_X + 6.5, bot_y, label_w - 7.5, band_h, [
    ("IT-DEMAND-",
     dict(size=12, bold=True, color=WHITE)),
    ("MANAGEMENT",
     dict(size=12, bold=True, color=WHITE)),
    ("(ServiceNow)",
     dict(size=9.5, italic=True, color=LIGHT_GREY)),
], anchor=MSO_ANCHOR.MIDDLE)

# Stats area on the right
stats_x = COL_X + label_w + 2
stats_w = TOTAL_W - label_w - 2
# Headline
add_text(stats_x, bot_y + 0.6, stats_w - 2, 3.0, [
    ("Ca. ",     dict(size=14, color=DARK)),
    ("1.000 Demands offen",
     dict(size=18, bold=True, color=NAVY)),
    ("  →  ",   dict(size=14, color=DARK)),
    ("450",     dict(size=18, bold=True, color=RED_HL)),
    ("  noch nicht klassifiziert",
     dict(size=14, color=DARK)),
], anchor=MSO_ANCHOR.MIDDLE)
# Progress bar
bar_y = bot_y + 5.0
bar_h = 2.0
bar_w = stats_w - 4
add_rect(stats_x + 2, bar_y, bar_w, bar_h, WHITE,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE,
         line=GREY_TXT, line_w=0.5)
# 450 of 1000 unclassified = 45%
unclass_w = bar_w * 0.45
add_rect(stats_x + 2, bar_y, unclass_w, bar_h, RED_HL,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(stats_x + 2, bar_y - 2.0, bar_w, 1.8, [
    ("Klassifizierungsstand", dict(size=8, italic=True, color=GREY_TXT)),
])
add_text(stats_x + 2, bar_y + bar_h + 0.2, bar_w, 1.8, [
    ("550 klassifiziert (55 %)",
     dict(size=8.5, color=GREEN, align=PP_ALIGN.LEFT)),
    ("", dict()),
])
add_text(stats_x + bar_w - 18, bar_y + bar_h + 0.2, 20, 1.8, [
    ("450 offen (45 %)",
     dict(size=8.5, bold=True, color=RED_HL, align=PP_ALIGN.RIGHT)),
])

# --- Band 6: EV-PROZESS (Lucom) ---
ev_y = bot_y + band_h + 1.0
add_rect(COL_X, ev_y, TOTAL_W, band_h, LIGHT_BG,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_rect(COL_X, ev_y, label_w, band_h, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_oval(COL_X + 1.5, ev_y + band_h/2 - 2.0, 4.0, 4.0, ORANGE)
add_text(COL_X + 1.5, ev_y + band_h/2 - 2.0, 4.0, 4.0, [
    ("6", dict(size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)),
], anchor=MSO_ANCHOR.MIDDLE)
add_text(COL_X + 6.5, ev_y, label_w - 7.5, band_h, [
    ("EV-PROZESS",
     dict(size=12, bold=True, color=WHITE)),
    ("(Lucom)",
     dict(size=9.5, italic=True, color=LIGHT_GREY)),
], anchor=MSO_ANCHOR.MIDDLE)

# Three-step flow
ev_step_x = COL_X + label_w + 2
ev_step_w = TOTAL_W - label_w - 2
step_w = (ev_step_w - 6) / 3
step_h = band_h - 2.0
step_y = ev_y + 1.0

steps = [
    ("1.600 EV'n",      "in der STIHL Gruppe",          NAVY),
    ("~ 1.030  (65 %)", "im Stammhaus",                  ORANGE),
    ("95 %",            "der EV'n mit Projektbezug",     GREEN),
]
for i, (big, sub, color) in enumerate(steps):
    x = ev_step_x + 1 + i * (step_w + 1.5)
    add_rect(x, step_y, step_w, step_h, WHITE,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=color, line_w=1.5)
    # Top bar
    add_rect(x, step_y, step_w, 0.5, color)
    add_text(x, step_y + 0.7, step_w, 3.4, [
        (big, dict(size=15, bold=True, color=color, align=PP_ALIGN.CENTER)),
    ], anchor=MSO_ANCHOR.MIDDLE)
    add_text(x, step_y + 4.0, step_w, step_h - 4.2, [
        (sub, dict(size=9.5, italic=True, color=GREY_TXT, align=PP_ALIGN.CENTER)),
    ], anchor=MSO_ANCHOR.TOP)
    # Arrow between
    if i < 2:
        ar = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                    ex(x + step_w + 0.1),
                                    ex(step_y + step_h/2 - 0.7),
                                    ex(1.3), ex(1.4))
        ar.fill.solid(); ar.fill.fore_color.rgb = DARK
        ar.line.fill.background()

# ====== FOOTER ======
add_rect(0, 89.3, 160, 0.7, ORANGE)
add_text(4, 86.3, 152, 2.0, [
    ("Mengenangaben aus dem PPM-Workshop;  Sondersystem-User: 4.400 PIT + 600 MSPO (IT/BR) + 300 PowerApp.",
     dict(size=8, italic=True, color=GREY_TXT))])

prs.save(DST)
print("Saved:", DST)
