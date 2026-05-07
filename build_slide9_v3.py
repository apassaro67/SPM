"""Slide 9 v3: top section becomes a ServiceNow-branded 'Steckbrief'
(Now Black + Now Green, key facts + coarse platform capabilities).
Reference customers and STIHL section unchanged.
"""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

SRC = "/tmp/2026_05_18_VS_Zielbild_SNOW_SPM_v6.pptx"
DST = "/tmp/2026_05_18_VS_Zielbild_SNOW_SPM_v7.pptx"

# STIHL / generic
ORANGE   = RGBColor(0xF0, 0x7F, 0x12)
DARK     = RGBColor(0x1F, 0x29, 0x37)
GREY_TXT = RGBColor(0x4B, 0x55, 0x63)
LIGHT_BG = RGBColor(0xF7, 0xF7, 0xF7)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
GREEN    = RGBColor(0x2E, 0x7D, 0x32)
LIGHT_GREY = RGBColor(0xD1, 0xD5, 0xDB)

# ServiceNow brand
NOW_BLACK    = RGBColor(0x03, 0x2D, 0x42)
NOW_BLACK_2  = RGBColor(0x05, 0x3B, 0x55)
NOW_GREEN    = RGBColor(0x62, 0xD8, 0x4E)
NOW_GREEN_DK = RGBColor(0x29, 0x9B, 0x32)
NOW_GREY     = RGBColor(0xB6, 0xC0, 0xC8)

U = 76200
def ex(u): return int(u * U)

prs = Presentation(SRC)
slide = prs.slides[8]

# Wipe
for shp in list(slide.shapes):
    sp = shp._element
    sp.getparent().remove(sp)

def add_rect(left, top, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE, line_w=1.2):
    s = slide.shapes.add_shape(shape, ex(left), ex(top), ex(w), ex(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line; s.line.width = Pt(line_w)
    s.shadow.inherit = False
    return s

def add_oval(left, top, w, h, fill, line=None):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, ex(left), ex(top), ex(w), ex(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
    return s

def set_paragraphs(shape, paragraphs, *, default_size=10, default_color=DARK, default_align=PP_ALIGN.LEFT):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(50000); tf.margin_right = Emu(50000)
    tf.margin_top = Emu(20000); tf.margin_bottom = Emu(20000)
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

def add_text(left, top, w, h, paragraphs, *, anchor=MSO_ANCHOR.TOP, **kw):
    tb = slide.shapes.add_textbox(ex(left), ex(top), ex(w), ex(h))
    tb.fill.background(); tb.line.fill.background()
    tb.text_frame.vertical_anchor = anchor
    set_paragraphs(tb, paragraphs, **kw)
    return tb

# ============== HEADER (slide-level, unchanged style) ==============
add_rect(0, 0, 160, 6, DARK)
add_text(3, 0.4, 140, 2.6, [
    ("ServiceNow als Tool – Reichweite, Referenzkunden und STIHL-Kontext",
     dict(size=20, bold=True, color=WHITE)),
])
add_text(3, 3.2, 140, 2.4, [
    ("Bewährter Industriestandard, große deutsche Referenzbasis und seit 2014 bei STIHL etabliert",
     dict(size=12, italic=True, color=LIGHT_GREY)),
])
add_rect(150, 0, 10, 6, ORANGE)
tb = slide.shapes.add_textbox(ex(150), ex(0), ex(10), ex(6))
tb.fill.background(); tb.line.fill.background()
tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
set_paragraphs(tb, [("Folie 9", dict(size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER))])

# ============== STECKBRIEF (ServiceNow-branded card) ==============
sb_x = 4
sb_y = 7.5
sb_w = 152
sb_h = 21.5

# Background frame in Now Black
add_rect(sb_x, sb_y, sb_w, sb_h, NOW_BLACK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
# Green accent strip on the left
add_rect(sb_x, sb_y, 0.9, sb_h, NOW_GREEN)

# --- Branded header (logo + tagline) ---
hdr_y = sb_y + 1.0
# Logo block: green dot + "servicenow"
add_oval(sb_x + 2.8, hdr_y + 0.9, 1.6, 1.6, NOW_GREEN)
add_text(sb_x + 5.0, hdr_y, 60, 3.5, [
    ("servicenow",
     dict(size=22, bold=True, color=WHITE)),
], anchor=MSO_ANCHOR.MIDDLE)
# Brand tagline
add_text(sb_x + 38, hdr_y, 110, 3.5, [
    ("Die Now Platform – eine Plattform für alle Geschäftsprozesse",
     dict(size=13, italic=True, color=NOW_GREY, align=PP_ALIGN.RIGHT)),
], anchor=MSO_ANCHOR.MIDDLE)

# --- Key facts strip (2 rows × 3 columns) ---
kf_top = sb_y + 4.8
kf_h = 4.0
add_rect(sb_x + 2, kf_top, sb_w - 4, kf_h, NOW_BLACK_2, shape=MSO_SHAPE.ROUNDED_RECTANGLE)

facts = [
    ("Gegründet", "2003"),
    ("HQ", "Santa Clara, CA"),
    ("Mitarbeitende", "~ 26.000"),
    ("Fortune 500-Kunden", "> 85 %"),
    ("Gartner", "Leader SPM · ITSM · ITOM"),
]
fact_w = (sb_w - 4) / len(facts)
for i, (label, value) in enumerate(facts):
    fx = sb_x + 2 + i * fact_w
    add_text(fx, kf_top + 0.4, fact_w, 1.6, [
        (label.upper(),
         dict(size=8, bold=True, color=NOW_GREEN, align=PP_ALIGN.CENTER)),
    ], anchor=MSO_ANCHOR.BOTTOM)
    add_text(fx, kf_top + 1.9, fact_w, 1.9, [
        (value,
         dict(size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)),
    ], anchor=MSO_ANCHOR.TOP)
    if i < len(facts) - 1:
        # subtle vertical divider
        add_rect(fx + fact_w - 0.05, kf_top + 0.6, 0.1, kf_h - 1.2,
                 RGBColor(0x18, 0x4F, 0x6E))

# --- Section label "PLATTFORM-FUNKTIONEN" ---
cap_top = kf_top + kf_h + 1.5
add_text(sb_x + 3, cap_top, 100, 2.0, [
    ("PLATTFORM-FUNKTIONEN",
     dict(size=9, bold=True, color=NOW_GREEN)),
])

# --- 5 capability cards (coarse functions, not modules) ---
cards = [
    ("Workflow-Plattform",
     "Prozesse end-to-end digitalisieren – statt Excel, E-Mail und Insellösungen."),
    ("KI / Now Assist",
     "Generative KI eingebaut: Risikofrüherkennung, Vorschläge, Auto-Reporting."),
    ("Low-Code App Engine",
     "Eigene Apps und Workflows ohne klassische Entwicklung – durch Fachbereich."),
    ("Integration Hub",
     "200+ Standard-Konnektoren zu SAP, M365, Cloud-Diensten und Custom-APIs."),
    ("Eine Datenbasis & CMDB",
     "Gemeinsames Datenmodell für IT, HR, Service, Risiko und Finanzen."),
]
c_top = cap_top + 2.4
c_h = 7.5
c_gap = 1.2
inner_w = sb_w - 6
cw = (inner_w - (len(cards) - 1) * c_gap) / len(cards)

for i, (head, body) in enumerate(cards):
    x = sb_x + 3 + i * (cw + c_gap)
    add_rect(x, c_top, cw, c_h, NOW_BLACK_2, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
             line=NOW_GREEN, line_w=0.75)
    # green top bar
    add_rect(x, c_top, cw, 0.4, NOW_GREEN)
    add_text(x, c_top + 0.8, cw, 2.4, [
        (head, dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)),
    ], anchor=MSO_ANCHOR.MIDDLE)
    add_text(x + 0.5, c_top + 3.3, cw - 1, c_h - 3.6, [
        (body, dict(size=9, color=NOW_GREY, align=PP_ALIGN.CENTER)),
    ], anchor=MSO_ANCHOR.TOP)

# ============== MIDDLE: Reference customers ==============
ref_top = 30.5
add_text(4, ref_top, 152, 2.4, [
    ("Große deutsche Referenzkunden – ServiceNow im Einsatz in IT und Geschäftsprozessen",
     dict(size=13, bold=True, color=DARK)),
])

customers = [
    ("SIEMENS",          "Industrie / Tech",       RGBColor(0x00, 0x9D, 0x9D), WHITE),
    ("SAP",              "Software",               RGBColor(0x00, 0x3A, 0x5D), RGBColor(0x07, 0xC8, 0xF0)),
    ("Allianz",          "Versicherung",           RGBColor(0x00, 0x37, 0x81), WHITE),
    ("Deutsche Bank",    "Banking",                RGBColor(0x00, 0x18, 0xA8), WHITE),
    ("T",                "Deutsche Telekom",       RGBColor(0xE2, 0x00, 0x74), WHITE),
    ("BOSCH",            "Industrie / Automotive", RGBColor(0xC8, 0x10, 0x2E), WHITE),
    ("Mercedes-Benz",    "Automotive",             RGBColor(0x00, 0x00, 0x00), WHITE),
    ("Volkswagen",       "Automotive",             RGBColor(0x00, 0x1E, 0x50), WHITE),
]
total = 152
box_top = 33.5
box_h = 11.5
n = len(customers)
gap_c = 0.8
box_w = (total - (n - 1) * gap_c) / n

for i, (name, branche, bg, fg) in enumerate(customers):
    x = 4 + i * (box_w + gap_c)
    add_rect(x, box_top, box_w, box_h - 4, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    name_size = 16 if len(name) <= 4 else (13 if len(name) <= 8 else 11)
    add_text(x, box_top, box_w, box_h - 4, [
        (name, dict(size=name_size, bold=True, color=fg, align=PP_ALIGN.CENTER)),
    ], anchor=MSO_ANCHOR.MIDDLE)
    add_text(x, box_top + box_h - 3.6, box_w, 3.4, [
        (branche, dict(size=8.5, italic=True, color=GREY_TXT, align=PP_ALIGN.CENTER)),
    ], anchor=MSO_ANCHOR.TOP)

add_text(4, 46.5, total, 2.0, [
    ("Stilisierte Markenfarben als Logo-Platzhalter – echte Logos können in der finalen Vorstandsfassung eingesetzt werden  |  Quelle: ServiceNow Customer References, öffentliche Quellen",
     dict(size=8, italic=True, color=LIGHT_GREY)),
])
add_rect(4, 49.5, total, 0.15, ORANGE)

# ============== BOTTOM: STIHL-spezifisch ==============
bot_top = 51.0
add_text(4, bot_top, total, 2.4, [
    ("ServiceNow @ STIHL – die Plattform ist bereits etabliert",
     dict(size=14, bold=True, color=DARK)),
])

stihl_boxes = [
    ("Im Einsatz seit 2014",
     "Globaler IT Service Management-Standard – inklusive China.\n"
     "Etabliertes Betriebsteam, eingespielte Lieferantenbeziehung.",
     "12 Jahre Betrieb"),
    ("DSGVO-konformes Hosting in Deutschland",
     "Hosting in den ServiceNow-Datacentern Frankfurt am Main und Düsseldorf.\n"
     "Vertragliche und technische DSGVO-Konformität gegeben.",
     "Frankfurt + Düsseldorf"),
    ("Anschluss-/Erweiterungsfähigkeit",
     "ITSM-CMDB als Andockpunkt für SPM (Apps, Services, Verträge).\n"
     "Plattform-Erweiterung statt Tool-Neubeschaffung – Synergien bei Lizenz, Betrieb, Schulung.",
     "Plattform-Synergie"),
]
sb_top2 = 53.5
sb_h2 = 21
gap_s = 2
sw2 = (total - 2 * gap_s) / 3
for i, (head, body, badge) in enumerate(stihl_boxes):
    x = 4 + i * (sw2 + gap_s)
    add_rect(x, sb_top2, sw2, sb_h2, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(x, sb_top2, sw2, 1.0, GREEN)
    badge_w = 16
    add_rect(x + sw2 - badge_w - 0.8, sb_top2 + 1.6, badge_w, 2.4, GREEN,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(x + sw2 - badge_w - 0.8, sb_top2 + 1.6, badge_w, 2.4, [
        (badge, dict(size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)),
    ], anchor=MSO_ANCHOR.MIDDLE)
    add_text(x + 1.5, sb_top2 + 1.8, sw2 - badge_w - 3.0, 4.5, [
        (head, dict(size=12, bold=True, color=DARK)),
    ])
    body_paras = [(line, dict(size=10, color=DARK, space_after=4))
                  for line in body.split("\n") if line]
    add_text(x + 1.5, sb_top2 + 7.0, sw2 - 3, sb_h2 - 8, body_paras)

# ============== FOOTER ==============
add_text(4, 86.3, total, 2.2, [
    ("Quellen: ServiceNow Produkt- und Kunden-Referenzen (servicenow.com); STIHL IT (Bestand seit 2014, Hosting Frankfurt + Düsseldorf, DSGVO-konform).",
     dict(size=8.5, italic=True, color=GREY_TXT)),
])
add_rect(0, 89.3, 160, 0.7, ORANGE)

prs.save(DST)
print("Saved:", DST)
