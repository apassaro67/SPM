"""Slide 4 v14: compact comparison table at top + TCO chart embedded
in the bottom-left box. Uses the user's v13_1 as source so all manual
edits to the table content are preserved.
"""
from pptx import Presentation
from pptx.util import Emu, Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

SRC = "/root/.claude/uploads/6f2327e1-58f0-4204-b00f-a2cbff9fd574/699f5bf9-2026_05_18_VS_Zielbild_SNOW_SPM_v13_1.pptx"
DST = "/tmp/2026_05_18_VS_Zielbild_SNOW_SPM_v14.pptx"
CHART_PNG = "/tmp/tco_chart.png"

ORANGE   = RGBColor(0xF0, 0x7F, 0x12)
ORANGE_LT= RGBColor(0xFD, 0xE6, 0xCC)
DARK     = RGBColor(0x1F, 0x29, 0x37)
GREY_TXT = RGBColor(0x4B, 0x55, 0x63)
LIGHT_BG = RGBColor(0xF7, 0xF7, 0xF7)
ROW_ALT  = RGBColor(0xEE, 0xF1, 0xF4)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
GREEN    = RGBColor(0x2E, 0x7D, 0x32)
NAVY     = RGBColor(0x2C, 0x3E, 0x50)
RED      = RGBColor(0xB9, 0x1C, 0x1C)
LIGHT_GREY = RGBColor(0xD1, 0xD5, 0xDB)

U = 76200
def ex(u): return int(u * U)

prs = Presentation(SRC)
slide = prs.slides[3]

# Wipe slide 4
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
    tf.margin_left = Emu(40000); tf.margin_right = Emu(40000)
    tf.margin_top = Emu(15000); tf.margin_bottom = Emu(15000)
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

# ========== HEADER ==========
add_rect(0, 0, 160, 6, DARK)
add_text(3, 0.4, 140, 2.6, [
    ("Bewertung der Plattformstrategie mit ServiceNow",
     dict(size=20, bold=True, color=WHITE)),
], anchor=MSO_ANCHOR.TOP)
add_text(3, 3.2, 140, 2.4, [
    ("Wirtschaftlichkeit (TCO) · Vergleich SPM-Plattform vs. Best-of-Breed · Gartner Executive Summary 2025",
     dict(size=12, italic=True, color=LIGHT_GREY)),
], anchor=MSO_ANCHOR.TOP)
add_rect(150, 0, 10, 6, ORANGE)
add_text(150, 0, 10, 6, [
    ("Folie 4", dict(size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)),
])

# ========== VERGLEICHSMATRIX (compact) ==========
COL_X = 4
# Slide ~ 160 wide. Use 152 total table width
TBL_W = 152
# Columns: Funktion | SPM (ServiceNow) | Best-of-Breed
COL_W = [42, 50, 60]
TBL_TOTAL = sum(COL_W)
def colx(i): return COL_X + sum(COL_W[:i])

# Section title
mat_top = 7.0
add_text(COL_X, mat_top, TBL_W, 2.2, [
    ("Funktions-Vergleich – SPM-Plattform vs. Best-of-Breed",
     dict(size=13, bold=True, color=DARK)),
])

# Header row
hdr_y = mat_top + 2.6
hdr_h = 3.5
add_rect(COL_X, hdr_y, COL_W[0], hdr_h, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(COL_X + 1.2, hdr_y, COL_W[0] - 2, hdr_h, [
    ("Funktion", dict(size=11, bold=True, color=WHITE)),
])
add_rect(colx(1), hdr_y, COL_W[1], hdr_h, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(colx(1) + 1, hdr_y, COL_W[1] - 2, hdr_h, [
    ("SPM-Plattform (ServiceNow)",
     dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)),
], anchor=MSO_ANCHOR.MIDDLE)
add_rect(colx(2), hdr_y, COL_W[2], hdr_h, GREY_TXT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(colx(2) + 1, hdr_y, COL_W[2] - 2, hdr_h, [
    ("Best-of-Breed (Einzeltools)",
     dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)),
], anchor=MSO_ANCHOR.MIDDLE)

# Rows (matches user's edits in v13_1)
rows = [
    ("Strategie (Ziele, KPI)",
     "integriert",
     "Workpath"),
    ("Demand & Portfolio Management",
     "integriert",
     "ServiceNow / SharePoint-Listen / Power-Automate"),
    ("Workflow Management",
     "integriert",
     "SharePoint-Listen / Power-Automate"),
    ("Projekt-Mgmt. – Ressourcen-Planung",
     "integriert",
     "Planisware"),
    ("Aufgaben-Management",
     "integriert",
     "MS Planner · Jira"),
    ("AI-Fähigkeiten",
     "integriert",
     "Je Tool separat, unterschiedliche Reife"),
    ("Reporting",
     "integriert + Power BI",
     "Power BI"),
]

row_y = hdr_y + hdr_h + 0.3
row_h = 3.2  # compact
row_gap = 0.2

for i, (krit, snow, bob) in enumerate(rows):
    y = row_y + i * (row_h + row_gap)
    bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
    add_rect(COL_X, y, TBL_TOTAL, row_h, bg)
    # Funktion (navy strip + label)
    add_rect(COL_X, y, COL_W[0], row_h, NAVY)
    add_text(COL_X + 1.2, y, COL_W[0] - 2, row_h, [
        (krit, dict(size=10.5, bold=True, color=WHITE)),
    ], anchor=MSO_ANCHOR.MIDDLE)
    # SNOW – orange accent + checkmark style
    add_rect(colx(1), y, 0.6, row_h, ORANGE)
    is_integ = "integ" in snow.lower()
    if is_integ:
        # Green checkmark badge
        add_oval(colx(1) + 1.2, y + row_h/2 - 1.0, 2.0, 2.0, GREEN)
        add_text(colx(1) + 1.2, y + row_h/2 - 1.0, 2.0, 2.0, [
            ("✓", dict(size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)),
        ], anchor=MSO_ANCHOR.MIDDLE)
        add_text(colx(1) + 3.8, y, COL_W[1] - 4.5, row_h, [
            (snow, dict(size=10.5, bold=True, color=DARK)),
        ], anchor=MSO_ANCHOR.MIDDLE)
    else:
        add_text(colx(1) + 1.2, y, COL_W[1] - 2, row_h, [
            (snow, dict(size=10.5, color=DARK)),
        ], anchor=MSO_ANCHOR.MIDDLE)
    # BoB
    add_text(colx(2) + 1.2, y, COL_W[2] - 2, row_h, [
        (bob, dict(size=10.5, color=GREY_TXT)),
    ], anchor=MSO_ANCHOR.MIDDLE)

table_bot = row_y + len(rows) * (row_h + row_gap)
print(f"Table bottom Y unit = {table_bot:.2f}")

# ========== BOTTOM SPLIT: TCO chart + Gartner ==========
bot_top = table_bot + 1.0
left_w = 76
right_w = TBL_W - left_w - 2

# --- Left: TCO chart panel ---
tco_x = COL_X
# Header bar
add_rect(tco_x, bot_top, left_w, 3.6, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(tco_x + 1.5, bot_top, left_w - 3, 3.6, [
    ("Wirtschaftlichkeitsbetrachtung – TCO-Analyse 5 Jahre",
     dict(size=12, bold=True, color=WHITE)),
], anchor=MSO_ANCHOR.MIDDLE)

# Body
chart_top = bot_top + 4.0
chart_body_h = 86 - chart_top - 4  # leave room for footer
add_rect(tco_x, chart_top, left_w, chart_body_h, LIGHT_BG,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)

# Embed the chart image
pic_pad = 1.0
pic_left = tco_x + pic_pad
pic_top = chart_top + pic_pad
pic_w = left_w - 2 * pic_pad
pic_h = chart_body_h - 2 * pic_pad

slide.shapes.add_picture(
    CHART_PNG,
    ex(pic_left), ex(pic_top),
    width=ex(pic_w), height=ex(pic_h)
)

# --- Right: Gartner Executive Summary ---
gar_x = COL_X + left_w + 2
add_rect(gar_x, bot_top, right_w, 3.6, NAVY,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(gar_x + 1.5, bot_top, right_w - 3, 3.6, [
    ("Gartner Executive Summary 2025  (Pure-Play vs. Platform)",
     dict(size=12, bold=True, color=WHITE)),
], anchor=MSO_ANCHOR.MIDDLE)

gar_top = bot_top + 4.0
gar_h = chart_body_h
add_rect(gar_x, gar_top, right_w, gar_h, LIGHT_BG,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)

statements = [
    ("Integration schlägt Funktionsbreite",
     "Nur eine Plattform ermöglicht echte End-to-End-Governance über Strategie, Projekte, Produkte, IT & Finance hinweg."),
    ("TCO + Governance-Vorteile",
     "Weniger Tools · weniger Schnittstellen · konsistente KPIs."),
    ("Zukunftssicherheit",
     "Plattformen wie ServiceNow erweitern SPM nativ um AI und Workflow-Automatisierung."),
]
sb_h = (gar_h - 2.4) / 3
for i, (head, body) in enumerate(statements):
    y = gar_top + 0.8 + i * sb_h
    add_rect(gar_x + 1.5, y, right_w - 3, sb_h - 0.8, WHITE,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE,
             line=ORANGE, line_w=1.2)
    add_oval(gar_x + 2.5, y + 0.9, 2.6, 2.6, NAVY)
    add_text(gar_x + 2.5, y + 0.9, 2.6, 2.6, [
        (str(i + 1), dict(size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)),
    ], anchor=MSO_ANCHOR.MIDDLE)
    add_text(gar_x + 6.0, y + 0.5, right_w - 8, 2.8, [
        (head, dict(size=11, bold=True, color=DARK)),
    ], anchor=MSO_ANCHOR.MIDDLE)
    add_text(gar_x + 6.0, y + 3.0, right_w - 8, sb_h - 3.8, [
        (body, dict(size=9.5, color=GREY_TXT)),
    ], anchor=MSO_ANCHOR.TOP)

add_text(gar_x + 2, gar_top + gar_h - 2.0, right_w - 4, 1.8, [
    ("Quelle: Gartner 2025 – Pure-Play Versus Platform, Comparing SPM and APMR Solutions",
     dict(size=8, italic=True, color=GREY_TXT)),
], anchor=MSO_ANCHOR.MIDDLE)

# ========== FOOTER ==========
add_rect(0, 89.3, 160, 0.7, ORANGE)
add_text(4, 86.3, 152, 2.0, [
    ("Konsolidiert aus 'Bewertung_der_Plattformstrategie' + TCO_Vergleich_5J_SN_vs_BoB.xlsx (Tab: Auswertung).",
     dict(size=8, italic=True, color=GREY_TXT)),
])

prs.save(DST)
print("Saved:", DST)
