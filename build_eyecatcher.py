"""New slide v20: Eyecatcher 'Strategie -> Portfolio -> Umsetzung in SNOW SPM'.
Inserted between current slide 9 (ServiceNow Steckbrief) and slide 10
(Vorprojekt SNOW - Ergebnisse 1).
"""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

SRC = "/tmp/2026_05_18_VS_Zielbild_SNOW_SPM_v19.pptx"
DST = "/tmp/2026_05_18_VS_Zielbild_SNOW_SPM_v20.pptx"

ORANGE   = RGBColor(0xF0, 0x7F, 0x12)
ORANGE_LT= RGBColor(0xFD, 0xE6, 0xCC)
DARK     = RGBColor(0x1F, 0x29, 0x37)
GREY_TXT = RGBColor(0x4B, 0x55, 0x63)
LIGHT_BG = RGBColor(0xF7, 0xF7, 0xF7)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
NAVY     = RGBColor(0x2C, 0x3E, 0x50)
NAVY_DK  = RGBColor(0x18, 0x25, 0x34)
GREEN    = RGBColor(0x2E, 0x7D, 0x32)
RED_HL   = RGBColor(0xB9, 0x1C, 0x1C)
BLUE_HD  = RGBColor(0x3B, 0x6B, 0xA5)
LIGHT_GREY = RGBColor(0xD1, 0xD5, 0xDB)
NOW_BLACK    = RGBColor(0x03, 0x2D, 0x42)
NOW_GREEN    = RGBColor(0x62, 0xD8, 0x4E)
NOW_GREEN_DK = RGBColor(0x29, 0x9B, 0x32)
NOW_GREY     = RGBColor(0xB6, 0xC0, 0xC8)

U = 76200
def ex(u): return int(u * U)

prs = Presentation(SRC)
# Use the layout of an existing slide for consistency
layout = prs.slides[8].slide_layout
new_slide = prs.slides.add_slide(layout)
slide = new_slide

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

def add_connector_line(x1, y1, x2, y2, color, weight=1.5, arrow=True):
    line = slide.shapes.add_connector(1, ex(x1), ex(y1), ex(x2), ex(y2))
    line.line.color.rgb = color
    line.line.width = Pt(weight)
    ln = line.line._get_or_add_ln()
    if arrow:
        tail = etree.SubElement(ln, qn('a:tailEnd'))
        tail.set('type', 'triangle'); tail.set('w', 'sm'); tail.set('len', 'sm')
    return line

def set_paragraphs(shape, paragraphs, *, default_size=10, default_color=DARK,
                   default_align=PP_ALIGN.LEFT):
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

# ===================== HEADER =====================
add_rect(0, 0, 160, 6, DARK)
add_text(3, 0.4, 140, 2.6, [
    ("ServiceNow SPM in Aktion – Vom Konzern-Ziel zum Projekt",
     dict(size=20, bold=True, color=WHITE))], anchor=MSO_ANCHOR.TOP)
add_text(3, 3.2, 140, 2.4, [
    ("Strategische Initiativen → Portfolios → Projekte · Produkte · Anforderungen auf einer Plattform",
     dict(size=12, italic=True, color=LIGHT_GREY))], anchor=MSO_ANCHOR.TOP)
add_rect(150, 0, 10, 6, ORANGE)
add_text(150, 0, 10, 6, [
    ("Eyecatcher", dict(size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER))])

# ===================== BACKGROUND =====================
# Slightly tinted background panel to give 'dashboard' feel
add_rect(2, 6.8, 156, 79, RGBColor(0xF3, 0xF6, 0xF9),
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)

# Column header strip
col_hdr_y = 8.0
col_hdr_h = 4.0
COL_TITLES = [
    ("STRATEGISCHE INITIATIVEN", "Konzern-Ziele · OKRs",   ORANGE),
    ("PORTFOLIOS",               "Wert- und Themen-Cluster", NAVY),
    ("UMSETZUNG",                "Projekte · Produkte · Anforderungen", NOW_GREEN_DK),
]
COL_X = [4, 56, 108]
COL_W = [50, 50, 50]
for i, (head, sub, col) in enumerate(COL_TITLES):
    x = COL_X[i]
    w = COL_W[i]
    add_rect(x, col_hdr_y, w, col_hdr_h, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(x + 2, col_hdr_y + 0.2, w - 4, 2.0, [
        (head, dict(size=11.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(x + 2, col_hdr_y + 2.0, w - 4, 1.8, [
        (sub, dict(size=9, italic=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)

# ===================== COLUMN 1: STRATEGIC INITIATIVES =====================
init_top = col_hdr_y + col_hdr_h + 1.5
init_w   = COL_W[0]
init_x   = COL_X[0]
init_card_h = 14.0
init_card_gap = 2.5

initiatives = [
    ("Operational Excellence",
     "Effiziente Prozesse · Kostendisziplin", "OKR  3 / 5", GREEN),
    ("Digitale Transformation",
     "Plattform-Strategie · Workflow-Automation", "OKR  4 / 6", ORANGE),
    ("Nachhaltigkeit & Compliance",
     "Green-IT · Energie · Reporting", "OKR  2 / 3", BLUE_HD),
]
init_centers = []  # (x_right, y_center) for arrow source

for i, (head, sub, badge, color) in enumerate(initiatives):
    y = init_top + i * (init_card_h + init_card_gap)
    add_rect(init_x, y, init_w, init_card_h, WHITE,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=color, line_w=1.5)
    # left color bar
    add_rect(init_x, y, 1.0, init_card_h, color)
    # title + subtitle
    add_text(init_x + 2, y + 0.8, init_w - 14, 4.0, [
        (head, dict(size=12, bold=True, color=DARK))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_text(init_x + 2, y + 5.0, init_w - 4, 3.5, [
        (sub, dict(size=9.5, italic=True, color=GREY_TXT))],
        anchor=MSO_ANCHOR.TOP)
    # OKR badge top right
    add_rect(init_x + init_w - 11, y + 0.8, 9, 3.0, color,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(init_x + init_w - 11, y + 0.8, 9, 3.0, [
        (badge, dict(size=9.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    # progress bar at bottom
    bar_y = y + init_card_h - 2.2
    add_rect(init_x + 2, bar_y, init_w - 4, 0.8, LIGHT_GREY,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    progress = [0.62, 0.74, 0.55][i]
    add_rect(init_x + 2, bar_y, (init_w - 4) * progress, 0.8, color,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(init_x + 2, bar_y + 1.0, init_w - 4, 1.2, [
        (f"Fortschritt: {int(progress*100)} %",
         dict(size=8, italic=True, color=GREY_TXT))],
        anchor=MSO_ANCHOR.TOP)
    init_centers.append((init_x + init_w, y + init_card_h / 2))

# ===================== COLUMN 2: PORTFOLIOS =====================
port_top = init_top
port_w   = COL_W[1]
port_x   = COL_X[1]
port_card_h = 8.5
port_card_gap = 1.5

portfolios = [
    ("PPM & Strategy",        "27 Vorhaben · 12 Mio. €",      2,  ORANGE),
    ("IT-Modernisierung",     "18 Vorhaben · 8 Mio. €",       1,  ORANGE),
    ("Production & Supply",   "33 Vorhaben · 18 Mio. €",      0,  GREEN),
    ("Green-IT & Compliance", "9 Vorhaben · 3 Mio. €",        2,  BLUE_HD),
]
port_centers_in = []  # left side connectors
port_centers_out = []  # right side connectors

for i, (head, sub, init_idx, color) in enumerate(portfolios):
    y = port_top + i * (port_card_h + port_card_gap)
    add_rect(port_x, y, port_w, port_card_h, WHITE,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=color, line_w=1.2)
    add_rect(port_x, y, 1.0, port_card_h, color)
    # title
    add_text(port_x + 2, y + 0.5, port_w - 4, 3.0, [
        (head, dict(size=11, bold=True, color=DARK))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_text(port_x + 2, y + 3.4, port_w - 4, 2.4, [
        (sub, dict(size=9, italic=True, color=GREY_TXT))],
        anchor=MSO_ANCHOR.TOP)
    # status dot
    add_oval(port_x + port_w - 4.5, y + 1.0, 1.2, 1.2, GREEN if init_idx != 1 else ORANGE)
    # mini KPI line
    add_text(port_x + 2, y + port_card_h - 2.4, port_w - 4, 2.0, [
        ("on track", dict(size=8, italic=True, color=GREEN))],
        anchor=MSO_ANCHOR.MIDDLE)
    port_centers_in.append((port_x, y + port_card_h / 2, init_idx))
    port_centers_out.append((port_x + port_w, y + port_card_h / 2))

# ===================== COLUMN 3: UMSETZUNG (Projekte/Produkte/Demands) =====================
exe_top = init_top
exe_w   = COL_W[2]
exe_x   = COL_X[2]
exe_card_h = 6.0
exe_card_gap = 1.0

executions = [
    ("SPM Rollout Phase 1",       "Projekt",     "Q4/26 Kick-off",  0, ORANGE),
    ("MSPO-Ablöse",               "Projekt",     "Go-Live 2027",    1, ORANGE),
    ("DEM-2510  PIT-Brücke",      "Demand",      "Klassifizierung", 1, NOW_GREEN_DK),
    ("Akku-Linie MS3",            "Produkt",     "Modell-Stufe 2",  2, NAVY),
    ("Supply-Chain Cockpit",      "Projekt",     "in Bewertung",    2, ORANGE),
    ("Green-Reporting EU CSRD",   "Produkt",     "Konzeption",      3, NAVY),
    ("DEM-2611  CMDB-Integration","Demand",      "Backlog",         3, NOW_GREEN_DK),
]
exe_centers_in = []  # (left, y, port_idx)

# Type color and badge label map
def type_badge(label):
    return {
        "Projekt": (ORANGE, "PRJ"),
        "Produkt": (NAVY, "PRD"),
        "Demand":  (NOW_GREEN_DK, "DMD"),
    }[label]

for i, (head, kind, status, port_idx, _) in enumerate(executions):
    y = exe_top + i * (exe_card_h + exe_card_gap)
    badge_color, badge_label = type_badge(kind)
    add_rect(exe_x, y, exe_w, exe_card_h, WHITE,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=badge_color, line_w=1.0)
    add_rect(exe_x, y, 0.8, exe_card_h, badge_color)
    # type badge top-left
    add_rect(exe_x + 1.5, y + 0.6, 5.5, 1.8, badge_color,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(exe_x + 1.5, y + 0.6, 5.5, 1.8, [
        (badge_label, dict(size=8.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    # title
    add_text(exe_x + 7.5, y + 0.4, exe_w - 8, 2.4, [
        (head, dict(size=10.5, bold=True, color=DARK))],
        anchor=MSO_ANCHOR.MIDDLE)
    # status
    add_text(exe_x + 1.5, y + 2.8, exe_w - 3, 2.6, [
        (status, dict(size=9, italic=True, color=GREY_TXT))],
        anchor=MSO_ANCHOR.MIDDLE)
    exe_centers_in.append((exe_x, y + exe_card_h / 2, port_idx))

# ===================== CONNECTING LINES =====================
# Initiative -> Portfolio (each portfolio knows which initiative)
for (px, py, init_idx) in port_centers_in:
    ix, iy = init_centers[init_idx]
    add_connector_line(ix, iy, px, py, ORANGE, weight=1.2, arrow=True)

# Portfolio -> Execution (each execution knows its portfolio)
for (ex_lx, ex_ly, port_idx) in exe_centers_in:
    px_out, py_out = port_centers_out[port_idx]
    add_connector_line(px_out, py_out, ex_lx, ex_ly,
                       NOW_GREEN_DK, weight=1.0, arrow=True)

# ===================== LEGEND =====================
legend_y = 79.5
add_text(4, legend_y, 30, 2.0, [
    ("Legende:", dict(size=9, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
items = [
    (ORANGE,       "Projekt"),
    (NAVY,         "Produkt"),
    (NOW_GREEN_DK, "Demand / Anforderung"),
]
lx = 14
for col, label in items:
    add_rect(lx, legend_y + 0.6, 1.6, 1.6, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(lx + 1.9, legend_y, 26, 2.2, [
        (label, dict(size=9, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    lx += 28
# Connection key
add_text(lx, legend_y, 60, 2.2, [
    ("→ verknüpft auf einer Plattform: Strategie → Portfolio → Umsetzung",
     dict(size=9, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.MIDDLE)

# ===================== FOOTER =====================
add_text(4, 86.3, 152, 2.0, [
    ("Beispielhafte Visualisierung der ServiceNow SPM-Plattform.  Daten illustrativ.",
     dict(size=8.5, italic=True, color=GREY_TXT))])
add_rect(0, 89.3, 160, 0.7, ORANGE)

# ===================== Position the new slide right after slide 9 =====================
sldIdLst = prs.slides._sldIdLst
slides = list(sldIdLst)
new_id = slides[-1]
sldIdLst.remove(new_id)
# Insert at index 9 (becomes slide 10), pushing old slide 10 onwards
sldIdLst.insert(9, new_id)

prs.save(DST)
print("Saved:", DST)
