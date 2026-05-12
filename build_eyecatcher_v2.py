"""Slide 10 v21: ServiceNow-styled Strategic Planning Workspace mockup.

Replaces the previous Eyecatcher with a ServiceNow look-and-feel dashboard:
- SNOW-style top navigation
- Strategic Goal cards row (with progress bars and trends)
- Portfolio Roadmap (Gantt-style timeline across 8 quarters Q4/26 - Q3/28)
  with multiple portfolio swim lanes and milestone markers
"""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

SRC = "/tmp/2026_05_18_VS_Zielbild_SNOW_SPM_v20.pptx"
DST = "/tmp/2026_05_18_VS_Zielbild_SNOW_SPM_v21.pptx"

# STIHL / generic
ORANGE   = RGBColor(0xF0, 0x7F, 0x12)
DARK     = RGBColor(0x1F, 0x29, 0x37)
GREY_TXT = RGBColor(0x4B, 0x55, 0x63)
LIGHT_BG = RGBColor(0xF7, 0xF7, 0xF7)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
NAVY     = RGBColor(0x2C, 0x3E, 0x50)
GREEN    = RGBColor(0x2E, 0x7D, 0x32)
LIGHT_GREY = RGBColor(0xD1, 0xD5, 0xDB)
GREY_BG  = RGBColor(0xE9, 0xED, 0xF1)
GRID_LN  = RGBColor(0xE5, 0xE7, 0xEB)
RED_HL   = RGBColor(0xB9, 0x1C, 0x1C)

# ServiceNow brand
NOW_BLACK    = RGBColor(0x03, 0x2D, 0x42)
NOW_BLACK_2  = RGBColor(0x05, 0x3B, 0x55)
NOW_GREEN    = RGBColor(0x62, 0xD8, 0x4E)
NOW_GREEN_DK = RGBColor(0x29, 0x9B, 0x32)
NOW_GREY     = RGBColor(0xB6, 0xC0, 0xC8)
NOW_BLUE     = RGBColor(0x29, 0x9B, 0xC6)
NOW_PURPLE   = RGBColor(0x86, 0x5D, 0xE4)
NOW_AMBER    = RGBColor(0xF5, 0xA8, 0x23)

U = 76200
def ex(u): return int(u * U)

prs = Presentation(SRC)
slide = prs.slides[9]  # current Eyecatcher slide

# Wipe everything
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

def add_oval(left, top, w, h, fill, line=None, line_w=0.0):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, ex(left), ex(top), ex(w), ex(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line; s.line.width = Pt(line_w)
    return s

def add_diamond(left, top, w, h, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.DIAMOND, ex(left), ex(top), ex(w), ex(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    s.line.fill.background()
    return s

def set_paragraphs(shape, paragraphs, *, default_size=10, default_color=DARK,
                   default_align=PP_ALIGN.LEFT):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(20000); tf.margin_right = Emu(20000)
    tf.margin_top = Emu(8000); tf.margin_bottom = Emu(8000)
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

# ===================== SLIDE HEADER =====================
add_rect(0, 0, 160, 6, DARK)
add_text(3, 0.4, 140, 2.6, [
    ("ServiceNow SPM in Aktion – Strategic Planning Workspace",
     dict(size=19, bold=True, color=WHITE))], anchor=MSO_ANCHOR.TOP)
add_text(3, 3.2, 140, 2.4, [
    ("Strategien und Portfolios auf einen Blick – auf einer Plattform",
     dict(size=12, italic=True, color=LIGHT_GREY))], anchor=MSO_ANCHOR.TOP)
add_rect(150, 0, 10, 6, ORANGE)
add_text(150, 0, 10, 6, [
    ("Eyecatcher", dict(size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER))])

# ===================== SNOW-STYLE UI FRAME =====================
ui_x = 3
ui_y = 7
ui_w = 154
ui_h = 78
add_rect(ui_x, ui_y, ui_w, ui_h, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
         line=LIGHT_GREY, line_w=0.5)
# Top navbar (SNOW black)
add_rect(ui_x, ui_y, ui_w, 4.5, NOW_BLACK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
# Cover the bottom rounded part so it sits as a flat header
add_rect(ui_x, ui_y + 2.0, ui_w, 2.6, NOW_BLACK)
# SNOW logo
add_oval(ui_x + 1.5, ui_y + 1.4, 1.7, 1.7, NOW_GREEN)
add_text(ui_x + 3.5, ui_y + 0.5, 22, 3.5, [
    ("servicenow", dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
# Title in navbar
add_text(ui_x + 26, ui_y + 0.5, 60, 3.5, [
    ("Strategic Planning Workspace",
     dict(size=11.5, bold=True, color=NOW_GREEN))], anchor=MSO_ANCHOR.MIDDLE)
# Tabs on the right side
tab_labels = ["Goals", "Roadmap", "Portfolios", "Resources"]
tab_w = 12
tab_gap = 1.0
tabs_total = len(tab_labels) * tab_w + (len(tab_labels) - 1) * tab_gap
tabs_x = ui_x + ui_w - tabs_total - 2
for i, lbl in enumerate(tab_labels):
    tx = tabs_x + i * (tab_w + tab_gap)
    active = (i == 1)  # 'Roadmap' is active
    color = NOW_GREEN if active else NOW_BLACK_2
    add_rect(tx, ui_y + 1.0, tab_w, 2.5, color,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(tx, ui_y + 1.0, tab_w, 2.5, [
        (lbl, dict(size=9.5, bold=active, color=NOW_BLACK if active else NOW_GREY,
                   align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)

# ===================== SECTION 1: STRATEGIC GOALS =====================
sec1_y = ui_y + 5.5
add_text(ui_x + 2, sec1_y, 80, 2.4, [
    ("STRATEGIC GOALS  (Konzern-OKRs)",
     dict(size=10, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
add_text(ui_x + ui_w - 30, sec1_y, 28, 2.4, [
    ("View: Konzern  ▾",
     dict(size=9, italic=True, color=GREY_TXT, align=PP_ALIGN.RIGHT))],
    anchor=MSO_ANCHOR.MIDDLE)

goal_cards_y = sec1_y + 2.4
goal_h = 11.0
goal_gap = 1.5
goals = [
    ("Operational Excellence", "Effizienz · Kosten",      62, NOW_BLUE,   "↑ 8%"),
    ("Digitale Transformation","Plattform · Workflow",    74, NOW_GREEN,  "↑ 12%"),
    ("Nachhaltigkeit",         "Green-IT · CSRD",         55, NOW_AMBER,  "↑ 4%"),
    ("Customer Experience",    "Service · Self-Service",  48, NOW_PURPLE, "↑ 6%"),
]
gw = (ui_w - 4 - (len(goals) - 1) * goal_gap) / len(goals)
for i, (head, sub, pct, color, trend) in enumerate(goals):
    x = ui_x + 2 + i * (gw + goal_gap)
    add_rect(x, goal_cards_y, gw, goal_h, WHITE,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE,
             line=LIGHT_GREY, line_w=0.7)
    # Left accent strip
    add_rect(x, goal_cards_y, 0.7, goal_h, color)
    # Title
    add_text(x + 1.5, goal_cards_y + 0.4, gw - 8, 2.4, [
        (head, dict(size=10.5, bold=True, color=DARK))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_text(x + 1.5, goal_cards_y + 2.4, gw - 3, 1.8, [
        (sub, dict(size=8.5, italic=True, color=GREY_TXT))],
        anchor=MSO_ANCHOR.MIDDLE)
    # Trend badge top right
    add_rect(x + gw - 7, goal_cards_y + 0.8, 6, 2.4, color,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(x + gw - 7, goal_cards_y + 0.8, 6, 2.4, [
        (trend, dict(size=8.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    # Big percent
    add_text(x + 1.5, goal_cards_y + 4.5, gw - 3, 3.6, [
        (f"{pct} %",
         dict(size=20, bold=True, color=color))],
        anchor=MSO_ANCHOR.MIDDLE)
    # Progress bar
    bar_y = goal_cards_y + goal_h - 2.0
    bar_w_full = gw - 3
    add_rect(x + 1.5, bar_y, bar_w_full, 1.0, GREY_BG,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(x + 1.5, bar_y, bar_w_full * (pct / 100), 1.0, color,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)

# ===================== SECTION 2: PORTFOLIO ROADMAP =====================
sec2_y = goal_cards_y + goal_h + 1.5
add_text(ui_x + 2, sec2_y, 60, 2.4, [
    ("PORTFOLIO ROADMAP  (8 Quartale)",
     dict(size=10, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
add_text(ui_x + ui_w - 50, sec2_y, 48, 2.4, [
    ("Lens: Strategic Investments ▾    Group by: Portfolio ▾",
     dict(size=8.5, italic=True, color=GREY_TXT, align=PP_ALIGN.RIGHT))],
    anchor=MSO_ANCHOR.MIDDLE)

# Timeline geometry
LANE_LBL_X = ui_x + 2
LANE_LBL_W = 32
TL_X = LANE_LBL_X + LANE_LBL_W + 1.2
TL_RIGHT = ui_x + ui_w - 2
TL_W = TL_RIGHT - TL_X
N_Q = 8  # Q4/26, Q1/27, Q2/27, Q3/27, Q4/27, Q1/28, Q2/28, Q3/28
QW = TL_W / N_Q
def qx(qi): return TL_X + qi * QW

# Quarter header
qhdr_top = sec2_y + 2.4
qhdr_h = 2.6
add_rect(LANE_LBL_X, qhdr_top, LANE_LBL_W, qhdr_h, GREY_BG,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(LANE_LBL_X + 1, qhdr_top, LANE_LBL_W - 2, qhdr_h, [
    ("Portfolio", dict(size=9.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
quarters = ["Q4 / 26", "Q1 / 27", "Q2 / 27", "Q3 / 27",
            "Q4 / 27", "Q1 / 28", "Q2 / 28", "Q3 / 28"]
for i, q in enumerate(quarters):
    x = qx(i)
    bg = GREY_BG if i in (0, 4) else WHITE  # alternate year-block
    add_rect(x, qhdr_top, QW - 0.05, qhdr_h, bg, line=GRID_LN, line_w=0.4)
    add_text(x, qhdr_top, QW - 0.05, qhdr_h, [
        (q, dict(size=8.5, bold=(i in (0, 4)), color=DARK, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
# Year separator
year_sep_x = qx(4)
add_rect(year_sep_x - 0.05, qhdr_top, 0.1, qhdr_h + 24, NAVY)

# Swim lanes
lanes = [
    ("PPM & Strategy",      "Op-Excellence",    NOW_BLUE,    0.0, 4.5,  62, "Strategy Cascade"),
    ("IT-Modernisierung",   "Digital Trans.",   NOW_GREEN,   1.0, 6.5,  74, "MSPO-Ablöse"),
    ("Production & Supply", "Op-Excellence",    NOW_AMBER,   2.0, 7.5,  55, "Supply-Chain Cockpit"),
    ("Green-IT & CSRD",     "Sustainability",   NOW_AMBER,   3.0, 8.0,  40, "CSRD-Reporting"),
    ("Customer Service",    "Customer Exp.",    NOW_PURPLE,  4.0, 8.0,  48, "Service-Portal"),
]
lane_top0 = qhdr_top + qhdr_h + 0.4
lane_h = 4.4
lane_gap = 0.4

# Today marker (vertical line at Q4/26 mid)
today_qpos = 0.4
today_x = qx(today_qpos)
total_lane_zone = len(lanes) * (lane_h + lane_gap)
add_rect(today_x - 0.06, lane_top0 - 0.3, 0.12, total_lane_zone + 0.4, RED_HL)
add_text(today_x - 6, lane_top0 + total_lane_zone + 0.3, 12, 1.5, [
    ("◆ heute (Q4 / 26)", dict(size=7.5, bold=True, color=RED_HL, align=PP_ALIGN.CENTER))],
    anchor=MSO_ANCHOR.MIDDLE)

for i, (name, parent, color, qstart, qend, pct, milestone) in enumerate(lanes):
    ly = lane_top0 + i * (lane_h + lane_gap)
    # Lane label
    add_rect(LANE_LBL_X, ly, LANE_LBL_W, lane_h, LIGHT_BG,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(LANE_LBL_X, ly, 0.6, lane_h, color)
    add_text(LANE_LBL_X + 1.2, ly + 0.2, LANE_LBL_W - 2, 2.3, [
        (name, dict(size=10, bold=True, color=DARK))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_text(LANE_LBL_X + 1.2, ly + 2.2, LANE_LBL_W - 6, 2.0, [
        (f"↑ {parent}", dict(size=8, italic=True, color=GREY_TXT))],
        anchor=MSO_ANCHOR.MIDDLE)
    # KPI badge right side
    add_rect(LANE_LBL_X + LANE_LBL_W - 6.5, ly + 0.6, 5.8, 3.0, color,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(LANE_LBL_X + LANE_LBL_W - 6.5, ly + 0.6, 5.8, 3.0, [
        (f"{pct}%", dict(size=10.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    # Lane background
    add_rect(TL_X, ly, TL_W, lane_h, WHITE, line=GRID_LN, line_w=0.4)
    # Quarter grid lines
    for qi in range(1, N_Q):
        x = qx(qi)
        if qi == 4:
            continue  # year separator handled above
        add_rect(x - 0.03, ly + 0.3, 0.06, lane_h - 0.6, GRID_LN)
    # Activity bar
    bar_x = qx(qstart) + 0.15
    bar_w = qx(qend) - bar_x - 0.15
    bar_y = ly + 0.7
    bar_h = lane_h - 1.4
    # Background of full bar (light)
    add_rect(bar_x, bar_y, bar_w, bar_h, GREY_BG,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    # Filled portion based on progress
    fill_w = bar_w * (pct / 100)
    add_rect(bar_x, bar_y, fill_w, bar_h, color,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    # Bar label
    add_text(bar_x + 0.4, bar_y, bar_w - 0.8, bar_h, [
        (milestone, dict(size=8.5, bold=True, color=WHITE if pct >= 35 else DARK))],
        anchor=MSO_ANCHOR.MIDDLE)
    # Diamond milestone at the END of the bar
    end_x = bar_x + bar_w
    add_diamond(end_x - 0.9, ly + lane_h/2 - 0.9, 1.8, 1.8, NOW_BLACK)

# ===================== FOOTER LEGEND =====================
foot_y = lane_top0 + len(lanes) * (lane_h + lane_gap) + 2.0
# Status legend chips
chips = [
    (NOW_GREEN,  "On track"),
    (NOW_AMBER,  "At risk"),
    (NOW_BLUE,   "On plan"),
    (NOW_PURPLE, "Strategic"),
]
cx = ui_x + 2
for col, lbl in chips:
    add_rect(cx, foot_y, 1.4, 1.4, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(cx + 1.7, foot_y - 0.3, 14, 2.0, [
        (lbl, dict(size=8.5, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    cx += 16

add_text(ui_x + ui_w - 80, foot_y - 0.2, 78, 2.2, [
    ("◆ Milestone   ·   ▮ Progress within scope   ·   Strategie ↔ Portfolio ↔ Demand auf einer Plattform",
     dict(size=8.5, italic=True, color=GREY_TXT, align=PP_ALIGN.RIGHT))],
    anchor=MSO_ANCHOR.MIDDLE)

# ===================== SLIDE FOOTER =====================
add_text(4, 86.3, 152, 2.0, [
    ("Beispielhafte Visualisierung der ServiceNow Strategic Planning Workspace / Portfolio Planning.  Daten illustrativ.",
     dict(size=8.5, italic=True, color=GREY_TXT))])
add_rect(0, 89.3, 160, 0.7, ORANGE)

prs.save(DST)
print("Saved:", DST)
