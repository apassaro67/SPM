"""Replace slide 4 with consolidated 'Bewertung der Plattformstrategie mit
ServiceNow' content from the upload (TCO + Vergleichsmatrix + Gartner)."""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

SRC = "/tmp/2026_05_18_VS_Zielbild_SNOW_SPM_v12.pptx"
DST = "/tmp/2026_05_18_VS_Zielbild_SNOW_SPM_v13.pptx"

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
NOW_GREEN= RGBColor(0x62, 0xD8, 0x4E)

U = 76200
def ex(u): return int(u * U)

prs = Presentation(SRC)
slide = prs.slides[3]  # slide 4

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
        if "space_after" in st:
            p.space_after = Pt(st["space_after"])

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

# ========== VERGLEICHSMATRIX ==========
mat_top = 7.5
mat_label_h = 2.2
add_text(4, mat_top, 152, mat_label_h, [
    ("Vergleichsmatrix – SPM-Plattform vs. Best-of-Breed (Gartner 2025)",
     dict(size=12.5, bold=True, color=DARK)),
], anchor=MSO_ANCHOR.MIDDLE)

# Table header
hdr_y = mat_top + mat_label_h + 0.3
hdr_h = 3.5
COL_X = 4
COL_W = [38, 56, 58]   # Kriterium | ServiceNow | Best-of-Breed
TOTAL_W = sum(COL_W)
def colx(i): return COL_X + sum(COL_W[:i])

# Header row
add_rect(COL_X, hdr_y, COL_W[0], hdr_h, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(COL_X + 1, hdr_y, COL_W[0] - 2, hdr_h, [
    ("Kriterium", dict(size=11, bold=True, color=WHITE)),
])
add_rect(colx(1), hdr_y, COL_W[1], hdr_h, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(colx(1) + 1, hdr_y, COL_W[1] - 2, hdr_h, [
    ("SPM-Plattform (ServiceNow)",
     dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)),
], anchor=MSO_ANCHOR.MIDDLE)
add_rect(colx(2), hdr_y, COL_W[2], hdr_h, GREY_TXT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(colx(2) + 1, hdr_y, COL_W[2] - 2, hdr_h, [
    ("Best-of-Breed (z. B. Planisware, Meisterplan)",
     dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)),
], anchor=MSO_ANCHOR.MIDDLE)

# Rows
rows = [
    ("Integration",
     "Nativ integriert über Strategie, Portfolio, Projekte, Produkte, IT, Finance",
     "Hohe Integrationskosten, oft Schnittstellen-Risiken"),
    ("Datenmodell &\nGovernance",
     "Einheitlich, plattformweit, KPI-konsistent",
     "Unterschiedliche Datenmodelle, Risiko für Schatten-Reporting"),
    ("Collaboration &\nWorkflow",
     "Durchgängige Workflows über Teams hinweg",
     "Mehrere Tools – Medienbrüche"),
    ("TCO / Betrieb",
     "Geringer Footprint (eine Plattform)",
     "Höhere Betriebs-/Lizenz-Kosten (mehrere Produkte + Integrationen)"),
    ("Transparenz &\nReporting",
     "Single Source of Truth",
     "Aggregation aufwendig"),
    ("AI-Fähigkeiten",
     "Plattformweit nutzbar, MQ-Leader bei ITSM-AI",
     "Jeder Vendor separat, unterschiedliche Reife"),
    ("Fit für STIHL",
     "Sehr hoch (bestehende Plattform-Strategie)",
     "Niedriger (erfordert Umbauten und Mehr-System-Standards)"),
]

row_y = hdr_y + hdr_h + 0.3
row_h = 3.5
for i, (krit, snow, bob) in enumerate(rows):
    y = row_y + i * (row_h + 0.2)
    bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
    add_rect(COL_X, y, TOTAL_W, row_h, bg)
    # Kriterium
    add_rect(COL_X, y, COL_W[0], row_h, NAVY)
    add_text(COL_X + 1, y, COL_W[0] - 2, row_h, [
        (krit, dict(size=10, bold=True, color=WHITE)),
    ], anchor=MSO_ANCHOR.MIDDLE)
    # SNOW
    add_rect(colx(1), y, 0.6, row_h, ORANGE)
    add_text(colx(1) + 1.0, y, COL_W[1] - 1.2, row_h, [
        (snow, dict(size=10, color=DARK)),
    ], anchor=MSO_ANCHOR.MIDDLE)
    # B-o-B
    add_text(colx(2) + 1.0, y, COL_W[2] - 1.2, row_h, [
        (bob, dict(size=10, color=GREY_TXT)),
    ], anchor=MSO_ANCHOR.MIDDLE)

# ========== BOTTOM SPLIT: TCO + Gartner ==========
bot_top = row_y + 7 * (row_h + 0.2) + 1.0  # ~37.6
left_w = 70
right_w = TOTAL_W - left_w - 2

# --- Left: TCO-Dimensionen ---
tco_x = COL_X
add_rect(tco_x, bot_top, left_w, 4.0, ORANGE,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(tco_x + 1.5, bot_top, left_w - 3, 4.0, [
    ("Wirtschaftlichkeitsbetrachtung – TCO-Dimensionen",
     dict(size=12, bold=True, color=WHITE)),
], anchor=MSO_ANCHOR.MIDDLE)

tco_body_top = bot_top + 4.5
tco_body_h = 38
add_rect(tco_x, tco_body_top, left_w, tco_body_h, LIGHT_BG,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)

tco_items = [
    "Lizenzkosten",
    "Implementierung & Integration",
    "Customizing / Configuration",
    "Schnittstellen (APIs) · Datenkonsistenz · Pflege",
    "Betrieb & Wartung",
    "Release-Management",
    "Support-Strukturen (intern FTE & extern)",
]
for i, item in enumerate(tco_items):
    y = tco_body_top + 1.6 + i * 4.6
    add_oval(tco_x + 2, y + 1.2, 1.6, 1.6, ORANGE)
    add_text(tco_x + 2, y + 1.2, 1.6, 1.6, [
        (str(i + 1), dict(size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)),
    ], anchor=MSO_ANCHOR.MIDDLE)
    add_text(tco_x + 4.5, y, left_w - 5.5, 4.0, [
        (item, dict(size=10.5, color=DARK)),
    ], anchor=MSO_ANCHOR.MIDDLE)

# Footnote inside TCO box
add_text(tco_x + 2, tco_body_top + tco_body_h - 3.0, left_w - 4, 2.5, [
    ("TCO basiert auf Annahmen / Schätzungen", dict(size=8.5, italic=True, color=GREY_TXT)),
], anchor=MSO_ANCHOR.MIDDLE)

# --- Right: Gartner Executive Summary ---
gar_x = COL_X + left_w + 2
add_rect(gar_x, bot_top, right_w, 4.0, NAVY,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(gar_x + 1.5, bot_top, right_w - 3, 4.0, [
    ("Gartner Executive Summary 2025  (Pure-Play vs. Platform)",
     dict(size=12, bold=True, color=WHITE)),
], anchor=MSO_ANCHOR.MIDDLE)

gar_body_top = bot_top + 4.5
gar_body_h = 38
add_rect(gar_x, gar_body_top, right_w, gar_body_h, LIGHT_BG,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)

statements = [
    ("Integration schlägt Funktionsbreite",
     "Nur eine Plattform ermöglicht echte End-to-End-Governance über Strategie, Projekte, Produkte, IT & Finance hinweg."),
    ("TCO + Governance-Vorteile",
     "Weniger Tools · weniger Schnittstellen · konsistente KPIs."),
    ("Zukunftssicherheit",
     "Plattformen wie ServiceNow erweitern SPM nativ um AI und Workflow-Automatisierung."),
]
sb_h = (gar_body_h - 2) / 3
for i, (head, body) in enumerate(statements):
    y = gar_body_top + 1.0 + i * sb_h
    add_rect(gar_x + 1.5, y, right_w - 3, sb_h - 1.0, WHITE,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE,
             line=ORANGE, line_w=1.2)
    # Number badge
    add_oval(gar_x + 2.5, y + 1.0, 3.0, 3.0, NAVY)
    add_text(gar_x + 2.5, y + 1.0, 3.0, 3.0, [
        (str(i + 1), dict(size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)),
    ], anchor=MSO_ANCHOR.MIDDLE)
    add_text(gar_x + 6.5, y + 0.6, right_w - 9, 3.0, [
        (head, dict(size=11, bold=True, color=DARK)),
    ], anchor=MSO_ANCHOR.MIDDLE)
    add_text(gar_x + 6.5, y + 3.4, right_w - 9, sb_h - 4.5, [
        (body, dict(size=9.5, color=GREY_TXT)),
    ], anchor=MSO_ANCHOR.TOP)

# Footnote inside Gartner box
add_text(gar_x + 2, gar_body_top + gar_body_h - 2.4, right_w - 4, 2.0, [
    ("Quelle: Gartner 2025 – Pure-Play Versus Platform, Comparing SPM and APMR Solutions",
     dict(size=8, italic=True, color=GREY_TXT)),
], anchor=MSO_ANCHOR.MIDDLE)

# ========== FOOTER ==========
add_rect(0, 89.3, 160, 0.7, ORANGE)
add_text(4, 86.3, 152, 2.0, [
    ("Konsolidiert aus 'Bewertung_der_Plattformstrategie_mit_ServiceNow.pptx' (TCO · Vergleichsmatrix · Gartner Executive Summary).",
     dict(size=8, italic=True, color=GREY_TXT)),
])

prs.save(DST)
print("Saved:", DST)
