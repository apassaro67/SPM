"""Erzeugt die CIO-Kurzfassung 'IT-Ablauforganisation - Homebase & Mannschaft'.

Zielbild fuer die Ablauforganisation der neuen IT-Organisation: Operations-Klammer
(SIAM), Value-Stream-Schnitte auf Basis des Org-Charts, Rollen, Spielregeln der
Matrix (RACI) und Ritual-Kalender. Stil an die bestehenden SPM-Decks angelehnt.
"""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

DST = "2026_07_08_IT_Ablauforganisation_Homebase_Mannschaft.pptx"

# ---- Hausfarben (aus build_slide4.py) + zwei Bedeutungsfarben ----
ORANGE    = RGBColor(0xF0, 0x7F, 0x12)   # Delivery / Mannschaft
ORANGE_LT = RGBColor(0xFD, 0xE6, 0xCC)
DARK      = RGBColor(0x1F, 0x29, 0x37)
GREY_TXT  = RGBColor(0x4B, 0x55, 0x63)
LIGHT_BG  = RGBColor(0xF7, 0xF7, 0xF7)
ROW_ALT   = RGBColor(0xEE, 0xF1, 0xF4)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
NAVY      = RGBColor(0x2C, 0x3E, 0x50)
HOMEBASE  = RGBColor(0x2E, 0x5A, 0x8F)   # Linie / Struktur
HOMEBASE_LT = RGBColor(0xE4, 0xEC, 0xF4)
PURPLE    = RGBColor(0x6B, 0x4F, 0xA0)   # Community
GREEN     = RGBColor(0x2E, 0x7D, 0x32)
RED       = RGBColor(0xB9, 0x1C, 0x1C)
LINE_GREY = RGBColor(0xD1, 0xD5, 0xDB)
# RACI
R_C = RGBColor(0x2E, 0x7D, 0x5B); A_C = RGBColor(0xB4, 0x53, 0x1A)
C_C = RGBColor(0x4A, 0x6B, 0x9E); I_C = RGBColor(0x8A, 0x93, 0x9D)

U = 76200                     # 1 Rastereinheit = 1/12 Zoll
def ex(u): return int(u * U)  # Slide = 160 x 90 Einheiten (13.33 x 7.5")

prs = Presentation()
prs.slide_width = ex(160)
prs.slide_height = ex(90)
BLANK = prs.slide_layouts[6]

FONT = "Calibri"

# --------------------------------------------------------------------------
def add_rect(slide, left, top, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE,
             line_w=1.0, round_=False):
    shp = MSO_SHAPE.ROUNDED_RECTANGLE if round_ else shape
    s = slide.shapes.add_shape(shp, ex(left), ex(top), ex(w), ex(h))
    if round_:
        try: s.adjustments[0] = 0.08
        except Exception: pass
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line; s.line.width = Pt(line_w)
    s.shadow.inherit = False
    return s

def connect(slide, x1, y1, x2, y2, color, w=1.25, dash=False):
    cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, ex(x1), ex(y1), ex(x2), ex(y2))
    cn.line.color.rgb = color; cn.line.width = Pt(w)
    cn.shadow.inherit = False
    if dash:
        from pptx.oxml.ns import qn
        ln = cn.line._get_or_add_ln()
        pd = ln.makeelement(qn('a:prstDash'), {'val': 'dash'}); ln.append(pd)
    return cn

def set_paragraphs(shape, paragraphs, *, default_size=11, default_color=DARK,
                   default_align=PP_ALIGN.LEFT, line_spacing=None):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(45000); tf.margin_right = Emu(45000)
    tf.margin_top = Emu(15000); tf.margin_bottom = Emu(15000)
    tf.clear()
    for i, item in enumerate(paragraphs):
        text, st = (item, {}) if isinstance(item, str) else item
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = st.get("align", default_align)
        if line_spacing: p.line_spacing = line_spacing
        if "space_after" in st: p.space_after = Pt(st["space_after"])
        if "space_before" in st: p.space_before = Pt(st["space_before"])
        # Unterstuetzt mehrere Runs pro Absatz via Liste von (text, style)
        runs = st.get("runs")
        if runs:
            for rt, rs in runs:
                run = p.add_run(); run.text = rt
                run.font.name = FONT
                run.font.size = Pt(rs.get("size", st.get("size", default_size)))
                run.font.bold = rs.get("bold", False)
                run.font.italic = rs.get("italic", False)
                run.font.color.rgb = rs.get("color", st.get("color", default_color))
        else:
            run = p.add_run(); run.text = text
            run.font.name = FONT
            run.font.size = Pt(st.get("size", default_size))
            run.font.bold = st.get("bold", False)
            run.font.italic = st.get("italic", False)
            run.font.color.rgb = st.get("color", default_color)

def add_text(slide, left, top, w, h, paragraphs, *, anchor=MSO_ANCHOR.TOP, **kw):
    tb = slide.shapes.add_textbox(ex(left), ex(top), ex(w), ex(h))
    tb.fill.background(); tb.line.fill.background()
    tb.text_frame.vertical_anchor = anchor
    set_paragraphs(tb, paragraphs, **kw)
    return tb

def slide_header(slide, kicker, title, page):
    add_rect(slide, 0, 0, 160, 90, WHITE)
    add_rect(slide, 0, 0, 1.6, 90, ORANGE)                 # linke Akzentkante
    add_text(slide, 7, 5, 130, 4,
             [(kicker, {"size": 11, "bold": True, "color": ORANGE})])
    add_text(slide, 7, 8.5, 146, 9,
             [(title, {"size": 24, "bold": True, "color": DARK})])
    add_rect(slide, 7, 17.5, 146, 0.12, LINE_GREY)
    add_text(slide, 150, 84.5, 8, 4,
             [(str(page), {"size": 10, "color": LINE_GREY, "align": PP_ALIGN.RIGHT})])
    add_text(slide, 7, 84.5, 90, 4,
             [("IT-Ablauforganisation · Zielbild", {"size": 9, "color": LINE_GREY})])

def legend_chip(slide, left, top, color, label):
    add_rect(slide, left, top, 1.4, 1.4, color, round_=False)
    add_text(slide, left + 1.9, top - 0.55, 40, 2.6,
             [(label, {"size": 9.5, "color": GREY_TXT})], anchor=MSO_ANCHOR.MIDDLE)

# ==========================================================================
# SLIDE 1 -- Titel / Kernbotschaft
# ==========================================================================
s = prs.slides.add_slide(BLANK)
add_rect(s, 0, 0, 160, 90, DARK)
add_rect(s, 0, 0, 160, 2.4, ORANGE)
add_text(s, 12, 20, 136, 5,
         [("ZIELBILD ABLAUFORGANISATION · IT / IPS", {"size": 13, "bold": True, "color": ORANGE})])
add_text(s, 12, 26, 138, 20,
         [("Die Einheit ist die Homebase.", {"size": 40, "bold": True, "color": WHITE}),
          ("Gespielt wird in Mannschaften.", {"size": 40, "bold": True, "color": ORANGE})])
add_text(s, 12, 52, 132, 8,
         [("Betriebs- und Zusammenarbeitsmodell für eine IT-Organisation mit verteiltem "
           "Operations – fachliche Heimat erhalten, Wertschöpfung End-to-End über "
           "Einheitsgrenzen organisieren.",
           {"size": 14, "color": RGBColor(0xC7, 0xCD, 0xD4)})],
         line_spacing=1.25)
# zwei Karten
add_rect(s, 12, 64, 66, 16, RGBColor(0x27, 0x31, 0x40), round_=True)
add_rect(s, 12, 64, 0.9, 16, HOMEBASE)
add_text(s, 15, 66, 60, 3, [("HOMEBASE – LINIENORGANISATION", {"size": 10, "bold": True, "color": RGBColor(0x8F, 0xB4, 0xDB)})])
add_text(s, 15, 69.5, 60, 9, [("Fachliche Heimat, Skill-Aufbau, Kapazität, Personalentwicklung. Die Org-Einheiten aus dem Chart.", {"size": 11.5, "color": RGBColor(0xC7, 0xCD, 0xD4)})], line_spacing=1.15)
add_rect(s, 82, 64, 66, 16, RGBColor(0x27, 0x31, 0x40), round_=True)
add_rect(s, 82, 64, 0.9, 16, ORANGE)
add_text(s, 85, 66, 60, 3, [("MANNSCHAFT – WERTSTROM / SERVICE", {"size": 10, "bold": True, "color": RGBColor(0xF2, 0xB2, 0x77)})])
add_text(s, 85, 69.5, 60, 9, [("End-to-End-Verantwortung für Services & Produkte, besetzt aus mehreren Homebases. Hier wird geliefert.", {"size": 11.5, "color": RGBColor(0xC7, 0xCD, 0xD4)})], line_spacing=1.15)

# ==========================================================================
# SLIDE 2 -- Operations-Klammer (SIAM)
# ==========================================================================
s = prs.slides.add_slide(BLANK)
slide_header(s, "OPERATIONS", "Operations nicht zentralisieren, sondern integrieren", 2)
add_text(s, 7, 19.5, 146, 6,
         [("Betrieb liegt heute an vier sinnvollen Stellen – nah an der Technologie. Die Klammer entsteht "
           "durch eine SIAM-Integrationsschicht statt struktureller Zusammenlegung.",
           {"size": 12, "color": GREY_TXT})], line_spacing=1.2)

# vier Homebase-Ops-Domaenen
dom = [("Platform & Infra Ops", "Base ITSM Ops, Service Desk, Datacenter, Netzwerk, Compute"),
       ("Application / Business Ops", "Solution Operations, SAP AMS, Release & Deployment"),
       ("Local IT Operations", "Local Service Desk, Device- & Server-Support vor Ort"),
       ("Security Operations", "SOC, Threat Detection, Incident Response")]
x = 7; w = 35.2; gap = 1.6
for i, (t, d) in enumerate(dom):
    lx = x + i * (w + gap)
    add_rect(s, lx, 27, w, 15, HOMEBASE_LT, round_=True)
    add_rect(s, lx, 27, w, 1.0, HOMEBASE)
    add_text(s, lx + 0.4, 28.4, w - 0.8, 3, [("HOMEBASE", {"size": 8.5, "bold": True, "color": HOMEBASE})])
    add_text(s, lx + 0.4, 31, w - 0.8, 4, [(t, {"size": 12, "bold": True, "color": DARK})])
    add_text(s, lx + 0.4, 35, w - 0.8, 7, [(d, {"size": 9.5, "color": GREY_TXT})], line_spacing=1.1)

# SIAM-Layer
add_rect(s, 7, 46, 146, 13.5, ORANGE_LT, round_=True)
add_rect(s, 7, 46, 146, 1.0, ORANGE)
add_text(s, 9, 47.4, 100, 3.5, [("DIE KLAMMER – SERVICE-INTEGRATION-LAYER (SIAM)", {"size": 11, "bold": True, "color": RGBColor(0xB4, 0x53, 0x1A)})])
add_text(s, 9, 50.8, 142, 8,
         [("", {"runs": [
             ("Service Integrator", {"bold": True, "color": DARK, "size": 11}),
             ("  · orchestriert interne Ops + Dienstleister als ein Liefernetzwerk       ", {"color": GREY_TXT, "size": 11}),
             ("End-to-End Service Owner", {"bold": True, "color": DARK, "size": 11}),
             ("  · je Service verantwortlich über alle Einheiten", {"color": GREY_TXT, "size": 11})]}),
          ("", {"space_before": 3, "runs": [
             ("Gemeinsame ITSM-Prozesse & ServiceNow", {"bold": True, "color": DARK, "size": 11}),
             ("  · ein Incident/Problem/Change/Request       ", {"color": GREY_TXT, "size": 11}),
             ("Operations Command Layer", {"bold": True, "color": DARK, "size": 11}),
             ("  · Major-Incident-Mgmt + Ops-Board", {"color": GREY_TXT, "size": 11})]})],
         line_spacing=1.15)

add_rect(s, 7, 62, 146, 9, LIGHT_BG, round_=True)
add_text(s, 9, 63.2, 142, 7,
         [("", {"runs": [
             ("Kernprinzip:  ", {"bold": True, "color": ORANGE, "size": 12}),
             ("Betrieb bleibt fachlich in den Einheiten (Homebase). Die SIAM-Schicht macht aus vier "
              "getrennten Betrieben ", {"color": DARK, "size": 12}),
             ("ein", {"bold": True, "italic": True, "color": DARK, "size": 12}),
             (" Betriebssystem – Dienstleister inklusive, gekoppelt über OLAs/UCs an dieselbe "
              "Prozess- und SLA-Logik.", {"color": DARK, "size": 12})]})],
         anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.2)

# ==========================================================================
# SLIDE 3 -- Value-Stream-Schnitte (konkret auf Org-Chart)
# ==========================================================================
s = prs.slides.add_slide(BLANK)
slide_header(s, "MANNSCHAFTEN", "Value-Stream-Schnitte – konkret auf Ihre Organisation", 3)
add_text(s, 7, 19.5, 146, 4,
         [("Sieben End-to-End-Services als dauerhafte Mannschaften. Jede besetzt aus mehreren Homebases; "
           "je ein Service Owner ist accountable.", {"size": 11.5, "color": GREY_TXT})], line_spacing=1.15)

# Tabelle: Value Stream | Owner-Heimat | Beitragende Homebases
rows = [
    ("Modern Workplace & Collaboration", "Technical Sol. › Workplace Solutions",
     "Local IT Ops · Cyber (Endpoint) · Business Sol. (Collab-Apps)"),
    ("Network & Connectivity", "Infra & Platforms › Network & Security Infra",
     "Local Network Support · Cyber (Network Sec.)"),
    ("Compute, Cloud & Datacenter", "Infra & Platforms › Compute Platform Mgmt",
     "Local Server Support · SAP Basis · Cyber"),
    ("Identity & Access (IAM)", "Infra & Platforms › Authentication / IDM",
     "Cyber (Identity Sec.) · Business Solutions"),
    ("SAP & Business Applications", "Business Sol. › Solution Operations (AMS)",
     "Technical Sol. (SAP Basis, Release & Deployment)"),
    ("Data, Analytics & KI", "Business Sol. › Reporting & Analytics & KI",
     "EA & Innovation · Master Data Management"),
    ("Digital Service Desk / ITSM", "Technical Sol. › Base ITSM Ops / Service Desk",
     "Local Service Desk · alle L2/L3 · SIAM"),
]
tx, ty, tw = 7, 25, 146
c1, c2, c3 = 46, 46, 54
hh = 4.4; rh = 6.0
# Kopf
add_rect(s, tx, ty, tw, hh, DARK)
add_text(s, tx + 0.8, ty, c1 - 1, hh, [("Value Stream (Mannschaft)", {"size": 10.5, "bold": True, "color": WHITE})], anchor=MSO_ANCHOR.MIDDLE)
add_text(s, tx + c1 + 0.8, ty, c2 - 1, hh, [("Service-Owner-Heimat", {"size": 10.5, "bold": True, "color": WHITE})], anchor=MSO_ANCHOR.MIDDLE)
add_text(s, tx + c1 + c2 + 0.8, ty, c3 - 1, hh, [("Beitragende Homebases", {"size": 10.5, "bold": True, "color": WHITE})], anchor=MSO_ANCHOR.MIDDLE)
for i, (a, b, c) in enumerate(rows):
    ry = ty + hh + i * rh
    add_rect(s, tx, ry, tw, rh, WHITE if i % 2 == 0 else ROW_ALT)
    add_rect(s, tx, ry, 0.7, rh, ORANGE)
    add_text(s, tx + 1.2, ry, c1 - 1.5, rh, [(a, {"size": 10.5, "bold": True, "color": DARK})], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, tx + c1 + 0.8, ry, c2 - 1, rh, [(b, {"size": 9.5, "color": HOMEBASE, "bold": True})], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, tx + c1 + c2 + 0.8, ry, c3 - 1, rh, [(c, {"size": 9.3, "color": GREY_TXT})], anchor=MSO_ANCHOR.MIDDLE)
by = ty + hh + len(rows) * rh
add_rect(s, tx, ty, tw, hh + len(rows) * rh, None, line=LINE_GREY, line_w=1.0)
add_text(s, 7, by + 1.2, 146, 4,
         [("", {"runs": [
             ("Quer dazu: ", {"bold": True, "color": PURPLE, "size": 10.5}),
             ("Cyber-Security-Services (SOC) werden von allen Streams konsumiert; Owner bleibt Cyber Defense. "
              "Optional OT / Shopfloor-IT als eigener Stream, falls relevant.", {"color": GREY_TXT, "size": 10.5})]})],
         line_spacing=1.1)

# ==========================================================================
# SLIDE 4 -- Virtuelle Teams + Rollen
# ==========================================================================
s = prs.slides.add_slide(BLANK)
slide_header(s, "ZUSAMMENARBEIT", "Drei Typen virtueller Teams – und wer sie führt", 4)
cards = [
    (ORANGE, "VERTIKAL · WERTSTROM", "Value-Stream- / Service-Teams",
     "Liefern das End-to-End-Ergebnis. Dauerhaft, cross-funktional aus mehreren Homebases.",
     "Rolle: Service / Product Owner  +  Delivery Lead („Kapitän“)"),
    (PURPLE, "HORIZONTAL · HANDWERK", "Chapters / Communities of Practice",
     "Bauen Standards, Skills, Wiederverwendung. Gleiche Disziplin über alle Teams.",
     "Rolle: Chapter Lead / CoP Lead   (z. B. Ops, Automation, Security)"),
    (HOMEBASE, "GOVERNANCE · ENTSCHEIDUNG", "Boards / Gremien",
     "Treffen übergreifende Entscheidungen & Priorisierung. Entscheider aus allen Bereichen.",
     "Rolle: Board-Chair   (Portfolio, CAB, ARB, SIAM-Board)"),
]
cw = 47.3; cx = 7; cgap = 2.5
for i, (col, tag, title, desc, role) in enumerate(cards):
    lx = cx + i * (cw + cgap)
    add_rect(s, lx, 22, cw, 34, WHITE, line=LINE_GREY, line_w=1.0, round_=True)
    add_rect(s, lx, 22, cw, 1.1, col)
    add_text(s, lx + 1.6, 24, cw - 3, 3, [(tag, {"size": 9, "bold": True, "color": col})])
    add_text(s, lx + 1.6, 27, cw - 3, 6, [(title, {"size": 13.5, "bold": True, "color": DARK})])
    add_text(s, lx + 1.6, 34, cw - 3, 11, [(desc, {"size": 11, "color": GREY_TXT})], line_spacing=1.2)
    add_rect(s, lx + 1.6, 48.5, cw - 3.2, 6, LIGHT_BG, round_=True)
    add_text(s, lx + 2.2, 48.5, cw - 4, 6, [(role, {"size": 9.8, "bold": True, "color": DARK})], anchor=MSO_ANCHOR.MIDDLE)

# Schlüsselrollen-Leiste
add_text(s, 7, 58.5, 146, 4, [("Schlüsselrollen, die das Modell tragen", {"size": 13, "bold": True, "color": DARK})])
key = [
    (ORANGE, "Service Owner", "accountable für Ergebnis, Kosten, SLA – über alle Einheiten"),
    (ORANGE, "Delivery Lead", "führt die Mannschaft durch die Lieferung"),
    (HOMEBASE, "Homebase / People Lead", "Menschen, Skill, Kapazität – sagt Kapazität zu"),
    (PURPLE, "Chapter Lead", "Standards & Handwerk über Teams hinweg"),
    (PURPLE, "SIAM Lead", "Betrieb & Provider als ein Netzwerk"),
    (HOMEBASE, "Capacity Manager", "vermittelt Angebot & Nachfrage – Schiedsrichter"),
]
kw = 47.3; ky0 = 63
for i, (col, t, d) in enumerate(key):
    row = i // 3; coln = i % 3
    lx = 7 + coln * (kw + 2.5); ly = ky0 + row * 10.5
    add_rect(s, lx, ly, kw, 9, WHITE, line=LINE_GREY, line_w=0.75, round_=True)
    add_rect(s, lx, ly, 0.8, 9, col)
    add_text(s, lx + 1.6, ly + 0.6, kw - 2.5, 3.5, [(t, {"size": 11, "bold": True, "color": DARK})])
    add_text(s, lx + 1.6, ly + 4.0, kw - 2.5, 5, [(d, {"size": 9, "color": GREY_TXT})], line_spacing=1.05)

# ==========================================================================
# SLIDE 5 -- Homebase vs Mannschaft + RACI
# ==========================================================================
s = prs.slides.add_slide(BLANK)
slide_header(s, "SPIELREGELN DER MATRIX", "Homebase entscheidet über Menschen – Mannschaft über Ergebnis", 5)

# zwei Spalten oben
add_rect(s, 7, 20.5, 71, 15, HOMEBASE_LT, round_=True)
add_rect(s, 7, 20.5, 0.9, 15, HOMEBASE)
add_text(s, 9.2, 21.4, 66, 3, [("HOMEBASE ENTSCHEIDET ÜBER …", {"size": 10, "bold": True, "color": HOMEBASE})])
add_text(s, 9.2, 24.6, 67, 11,
         [("•  Einstellung / Entwicklung / Beförderung", {"size": 10.5, "color": DARK, "space_after": 2}),
          ("•  Skill-Profile & fachliche Standards", {"size": 10.5, "color": DARK, "space_after": 2}),
          ("•  Verfügbare Kapazität (Zusage an Mannschaften)", {"size": 10.5, "color": DARK, "space_after": 2}),
          ("•  Personalthemen, Feedback, Disziplinarik", {"size": 10.5, "color": DARK})])
add_rect(s, 82, 20.5, 71, 15, ORANGE_LT, round_=True)
add_rect(s, 82, 20.5, 0.9, 15, ORANGE)
add_text(s, 84.2, 21.4, 66, 3, [("MANNSCHAFT ENTSCHEIDET ÜBER …", {"size": 10, "bold": True, "color": RGBColor(0xB4, 0x53, 0x1A)})])
add_text(s, 84.2, 24.6, 67, 11,
         [("•  Was priorisiert und geliefert wird", {"size": 10.5, "color": DARK, "space_after": 2}),
          ("•  Wie die tägliche Arbeit im Team abläuft", {"size": 10.5, "color": DARK, "space_after": 2}),
          ("•  Service-/Produkt-Ergebnis, SLA, Roadmap", {"size": 10.5, "color": DARK, "space_after": 2}),
          ("•  Fachliche Zuweisung im Sprint/Betrieb", {"size": 10.5, "color": DARK})])

# RACI-Tabelle
add_text(s, 7, 37.5, 100, 3.5, [("RACI – wer verantwortet was", {"size": 13, "bold": True, "color": DARK})])
# Legende
leg = [("R", R_C, "Responsible"), ("A", A_C, "Accountable"), ("C", C_C, "Consulted"), ("I", I_C, "Informed")]
lx = 60
for code, col, lab in leg:
    add_rect(s, lx, 37.8, 2.6, 2.6, col, round_=True)
    add_text(s, lx + 0.0, 38.0, 2.6, 2.2, [(code, {"size": 10, "bold": True, "color": WHITE, "align": PP_ALIGN.CENTER})], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, lx + 3.0, 37.6, 20, 3, [(lab, {"size": 9.5, "color": GREY_TXT})], anchor=MSO_ANCHOR.MIDDLE)
    lx += 24

raci_cols = ["Homebase\nLead", "Service\nOwner", "Delivery\nLead", "Chapter\nLead", "SIAM /\nCapacity", "CIO /\nPortfolio"]
raci_rows = [
    ("Einstellung & Personalentwicklung", ["A", "C", "C", "C", "I", "I"]),
    ("Kapazitäts-Zusage an Mannschaften", ["R", "C", "I", "I", "A", "I"]),
    ("Service-/Produkt-Priorisierung", ["I", "A", "R", "C", "I", "C"]),
    ("Betrieb & SLA (End-to-End)", ["C", "A", "R", "I", "R", "I"]),
    ("Provider-/Dienstleister-Steuerung", ["I", "C", "I", "I", "A", "I"]),
    ("Major Incident (P1/P2)", ["I", "C", "R", "I", "A", "I"]),
    ("Ressourcen-Konflikt zw. Teams", ["C", "C", "I", "I", "R", "A"]),
]
tx, ty = 7, 42
lblw = 46; cw2 = 100.0 / 6
hh2 = 4.6; rh2 = 4.7
add_rect(s, tx, ty, lblw, hh2, DARK)
add_text(s, tx + 0.8, ty, lblw - 1, hh2, [("Entscheidung / Aktivität", {"size": 10, "bold": True, "color": WHITE})], anchor=MSO_ANCHOR.MIDDLE)
for j, cn in enumerate(raci_cols):
    cxx = tx + lblw + j * cw2
    add_rect(s, cxx, ty, cw2, hh2, DARK)
    add_text(s, cxx, ty, cw2, hh2, [(cn.replace("\n", " "), {"size": 8.3, "bold": True, "color": WHITE, "align": PP_ALIGN.CENTER})], anchor=MSO_ANCHOR.MIDDLE)
raci_map = {"R": R_C, "A": A_C, "C": C_C, "I": I_C}
for i, (lab, vals) in enumerate(raci_rows):
    ry = ty + hh2 + i * rh2
    add_rect(s, tx, ry, lblw, rh2, WHITE if i % 2 == 0 else ROW_ALT)
    add_text(s, tx + 0.8, ry, lblw - 1, rh2, [(lab, {"size": 9.5, "color": DARK})], anchor=MSO_ANCHOR.MIDDLE)
    for j, v in enumerate(vals):
        cxx = tx + lblw + j * cw2
        add_rect(s, cxx, ry, cw2, rh2, WHITE if i % 2 == 0 else ROW_ALT)
        bw = 3.0
        add_rect(s, cxx + cw2/2 - bw/2, ry + rh2/2 - 1.35, bw, 2.7, raci_map[v], round_=True)
        add_text(s, cxx + cw2/2 - bw/2, ry + rh2/2 - 1.4, bw, 2.8, [(v, {"size": 9.5, "bold": True, "color": WHITE, "align": PP_ALIGN.CENTER})], anchor=MSO_ANCHOR.MIDDLE)
add_rect(s, tx, ty, lblw + 6 * cw2, hh2 + len(raci_rows) * rh2, None, line=LINE_GREY, line_w=1.0)
add_text(s, 7, 84.0, 146, 3.5,
         [("", {"runs": [
             ("Der Kern: ", {"bold": True, "color": ORANGE, "size": 10}),
             ("Kapazität = Homebase accountable, Priorität = Service Owner accountable. "
              "Genau ein A pro Zeile.", {"color": GREY_TXT, "size": 10})]})])

# ==========================================================================
# SLIDE 6 -- Ritual-Kalender + Next Steps
# ==========================================================================
s = prs.slides.add_slide(BLANK)
slide_header(s, "BETRIEBSRHYTHMUS", "Ritual-Kalender – vier Ebenen, ein Takt", 6)
rit = [
    ("QUARTAL", PURPLE, [
        ("IPS Steering / Portfolio-Board", "Priorisierung, Budget & Kapazität → Roadmap"),
        ("Architecture Review Board", "Standards & Design-Reviews → ADR"),
        ("Quarterly Business Review", "Wertbeitrag & OKR je Wertstrom")]),
    ("MONAT", HOMEBASE, [
        ("Demand & Sponsoring", "Bewertung neuer Anforderungen → Backlog"),
        ("Service Management Review", "SLA/OLA, Provider-Steuerung → Maßnahmen")]),
    ("WOCHE", ORANGE, [
        ("CAB – Change Advisory Board", "Risikobewertung & Freigabe von Changes"),
        ("Value-Stream Planning (2-wöch.)", "Increment-Planung → Team-Backlog"),
        ("Provider Ops Jour Fixe", "Operative Provider-Steuerung")]),
    ("TÄGLICH", GREEN, [
        ("Ops Sync / Service Desk Standup", "Lage, Incidents, Eskalationen"),
        ("Squad Daily", "Fortschritt & Hindernisse der Mannschaft"),
        ("Major Incident Mgmt (bei Bedarf)", "P1/P2 → Wiederherstellung + Post-Mortem")]),
]
cx = 7; cw = 36.0; cgap = 1.33
for i, (freq, col, items) in enumerate(rit):
    lx = cx + i * (cw + cgap)
    add_rect(s, lx, 21, cw, 42, LIGHT_BG, round_=True)
    add_rect(s, lx, 21, cw, 4.2, col)
    add_text(s, lx, 21, cw, 4.2, [(freq, {"size": 11, "bold": True, "color": WHITE, "align": PP_ALIGN.CENTER})], anchor=MSO_ANCHOR.MIDDLE)
    yy = 26.5
    for t, d in items:
        add_text(s, lx + 1.4, yy, cw - 2.6, 4, [(t, {"size": 10, "bold": True, "color": DARK})], line_spacing=1.0)
        add_text(s, lx + 1.4, yy + 3.4, cw - 2.6, 5, [(d, {"size": 8.7, "color": GREY_TXT})], line_spacing=1.02)
        yy += 8.9

# Next steps
add_text(s, 7, 65.5, 146, 3.5, [("Einführung in vier Schritten", {"size": 13, "bold": True, "color": DARK})])
steps = [
    ("1", "Klammer setzen", "5–8 Kern-Services definieren, Service Owner benennen, SIAM verankern"),
    ("2", "Prozesse vereinheitlichen", "Ein ITSM-Modell auf ServiceNow, OLAs/UCs mit Providern"),
    ("3", "Rhythmus starten", "Wenige Rituale einführen – jedes mit einem Output"),
    ("4", "Matrix leben", "1–2 Pilot-Streams, %-Kapazitätszusagen, Capacity-Board"),
]
sw = 36.0
for i, (n, t, d) in enumerate(steps):
    lx = 7 + i * (sw + 1.33)
    add_rect(s, lx, 70, sw, 12, WHITE, line=LINE_GREY, line_w=0.75, round_=True)
    add_rect(s, lx, 70, 0.8, 12, ORANGE)
    add_text(s, lx + 1.8, 70.8, 6, 5, [(n, {"size": 20, "bold": True, "color": ORANGE_LT if False else ORANGE})])
    add_text(s, lx + 7, 71.0, sw - 8, 3.5, [(t, {"size": 10.5, "bold": True, "color": DARK})])
    add_text(s, lx + 7, 74.4, sw - 8, 7, [(d, {"size": 8.7, "color": GREY_TXT})], line_spacing=1.05)

# ==========================================================================
# SLIDE 7 -- SLA / OLA / UC Kaskade + Service-Tiers
# ==========================================================================
s = prs.slides.add_slide(BLANK)
slide_header(s, "SERVICE-LEVEL", "SLA → OLA → UC – eine Zusage, back-to-back zerlegt", 7)
add_text(s, 7, 19.5, 146, 4,
         [("Der Service Owner gibt dem Business eine SLA. Diese wird intern in OLAs (Homebases) und "
           "extern in Underpinning Contracts (Dienstleister) zerlegt – ohne Lücke.",
           {"size": 11.5, "color": GREY_TXT})], line_spacing=1.15)

# Kaskade links (drei gestapelte Ebenen)
casc = [(ORANGE, "SLA", "Service Owner → Business", "Zugesagte Verfügbarkeit & Reaktionszeit je Service"),
        (HOMEBASE, "OLA", "Service Owner ↔ Homebases", "Interne Betriebsvereinbarungen je beitragende Einheit"),
        (PURPLE, "UC", "SIAM → Dienstleister", "Underpinning Contracts – Ziele strenger als OLA")]
cy = 26
for col, code, who, desc in casc:
    add_rect(s, 7, cy, 70, 9.5, WHITE, line=LINE_GREY, line_w=1.0, round_=True)
    add_rect(s, 7, cy, 11, 9.5, col, round_=False)
    add_text(s, 7, cy, 11, 9.5, [(code, {"size": 17, "bold": True, "color": WHITE, "align": PP_ALIGN.CENTER})], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 20, cy + 1.1, 55, 3.5, [(who, {"size": 11, "bold": True, "color": DARK})])
    add_text(s, 20, cy + 4.6, 56, 4.5, [(desc, {"size": 9.5, "color": GREY_TXT})], line_spacing=1.05)
    cy += 11
add_rect(s, 7, 59.5, 70, 8, LIGHT_BG, round_=True)
add_text(s, 8.6, 59.5, 67, 8,
         [("", {"runs": [
             ("Goldene Regel:  ", {"bold": True, "color": ORANGE, "size": 10.5}),
             ("UC-Ziel ≤ OLA-Ziel ≤ SLA-Ziel.", {"bold": True, "color": DARK, "size": 10.5}),
             (" Jede Provider-Zusage ist strenger als die interne, jede interne strenger als die "
              "Business-Zusage – so entstehen keine Deckungslücken.", {"color": GREY_TXT, "size": 10.5})]})],
         anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.15)

# Rechts: Tier-Tabellen
add_text(s, 82, 25, 71, 3.5, [("Verfügbarkeitsklassen", {"size": 12, "bold": True, "color": DARK})])
tiers = [("Platinum", "99,9 %", "24×7", "IAM, Netzwerk, Compute, SAP-Prod"),
         ("Gold", "99,5 %", "5×11 + Rufb.", "Workplace, Data, Service Desk"),
         ("Silver", "99,0 %", "5×11", "nachrangige Services")]
ty = 29
add_rect(s, 82, ty, 71, 3.6, DARK)
for cx0, w0, txt in [(82, 15, "Klasse"), (97, 12, "Verfügb."), (109, 16, "Kernzeit"), (125, 28, "Beispiele")]:
    add_text(s, cx0, ty, w0, 3.6, [(txt, {"size": 9, "bold": True, "color": WHITE})], anchor=MSO_ANCHOR.MIDDLE)
for i, (a, b, c, d) in enumerate(tiers):
    ry = ty + 3.6 + i * 4.4
    add_rect(s, 82, ry, 71, 4.4, WHITE if i % 2 == 0 else ROW_ALT)
    add_text(s, 82.6, ry, 15, 4.4, [(a, {"size": 9.5, "bold": True, "color": ORANGE})], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 97, ry, 12, 4.4, [(b, {"size": 9.5, "color": DARK})], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 109, ry, 16, 4.4, [(c, {"size": 9, "color": DARK})], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 125, ry, 28, 4.4, [(d, {"size": 8.5, "color": GREY_TXT})], anchor=MSO_ANCHOR.MIDDLE)
add_rect(s, 82, ty, 71, 3.6 + 3 * 4.4, None, line=LINE_GREY, line_w=1.0)

add_text(s, 82, 47.5, 71, 3.5, [("Prioritäten – Reaktion / Wiederherstellung", {"size": 12, "bold": True, "color": DARK})])
prio = [("P1 kritisch", "15 min", "≤ 4 h"), ("P2 hoch", "30 min", "≤ 8 h"),
        ("P3 mittel", "4 h", "2 AT"), ("P4 niedrig", "8 h", "5 AT"), ("Service Request", "–", "3 AT")]
py = 51.5
add_rect(s, 82, py, 71, 3.6, DARK)
for cx0, w0, txt in [(82, 27, "Priorität"), (109, 22, "Reaktion"), (131, 22, "Wiederherst.")]:
    add_text(s, cx0, py, w0, 3.6, [(txt, {"size": 9, "bold": True, "color": WHITE})], anchor=MSO_ANCHOR.MIDDLE)
for i, (a, b, c) in enumerate(prio):
    ry = py + 3.6 + i * 3.5
    add_rect(s, 82, ry, 71, 3.5, WHITE if i % 2 == 0 else ROW_ALT)
    col = RED if i == 0 else (RGBColor(0xB4, 0x53, 0x1A) if i == 1 else DARK)
    add_text(s, 82.6, ry, 27, 3.5, [(a, {"size": 9.3, "bold": i < 2, "color": col})], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 109, ry, 22, 3.5, [(b, {"size": 9.3, "color": DARK})], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 131, ry, 22, 3.5, [(c, {"size": 9.3, "color": DARK})], anchor=MSO_ANCHOR.MIDDLE)
add_rect(s, 82, py, 71, 3.6 + 5 * 3.5, None, line=LINE_GREY, line_w=1.0)

# ==========================================================================
# SLIDE 8 -- SLA / OLA je Value Stream
# ==========================================================================
s = prs.slides.add_slide(BLANK)
slide_header(s, "SERVICE-LEVEL", "SLA- & OLA-Skizze je Value Stream", 8)
sla_rows = [
    ("Modern Workplace", "Gold\n99,5 %", "5×11 +Rufb.", "4 h",
     "Client Platform Ops (WS) · Endpoint-Sec (Cyber) · Collab-Apps (BusSol) · Vor-Ort (Local)",
     "Field-/Deskside-Services · MDM-Provider · Print-Services"),
    ("Network & Connectivity", "Platinum\n99,9 %", "24×7", "2 h",
     "Network Ops (Infra) · Local Network (Local) · Network-Sec (Cyber)",
     "WAN-Carrier · SD-WAN Managed Service"),
    ("Compute, Cloud & DC", "Platinum\n99,9 %", "24×7", "2 h",
     "Platform Ops (Infra) · Server Vor-Ort (Local) · Cloud-Sec (Cyber)",
     "Hyperscaler · Colocation-DC · Backup-Provider"),
    ("Identity & Access", "Platinum\n99,9 %", "24×7", "1 h",
     "IDM Ops (Infra) · Identity-Security (Cyber)",
     "IAM-SaaS (Entra/Okta) · PKI-Provider"),
    ("SAP & Business Apps", "Platinum\n99,9 %", "6×14", "2 h",
     "AMS (BusSol) · SAP Basis (AppTech) · Release & Deploy",
     "SAP-AMS-Dienstleister · Hosting (RISE/Hyperscaler)"),
    ("Data, Analytics & KI", "Gold\n99,5 %", "5×11", "8 h",
     "BI Ops (BusSol) · Data Platform (EA/Infra) · MDM",
     "Cloud-Analytics-Provider"),
    ("Service Desk / ITSM", "Gold\n99,5 %", "24×7 (SD)", "Disp. <15m",
     "Base ITSM Ops · Local Service Desk · ServiceNow-Plattform",
     "Service-Desk-Dienstleister (1st Level) · ServiceNow (SaaS)"),
]
tx, ty = 7, 21
cols = [("Value Stream", 25), ("Ziel-SLA", 15), ("Kernzeit", 14), ("P1", 12),
        ("OLA-Bausteine (intern)", 42), ("Underpinning Contracts (Provider)", 38)]
hh = 5.5
add_rect(s, tx, ty, sum(w for _, w in cols), hh, DARK)
cxp = tx
for name, w in cols:
    add_text(s, cxp + 0.5, ty, w - 0.8, hh, [(name, {"size": 9, "bold": True, "color": WHITE})], anchor=MSO_ANCHOR.MIDDLE)
    cxp += w
rh = 8.4
for i, row in enumerate(sla_rows):
    ry = ty + hh + i * rh
    add_rect(s, tx, ry, sum(w for _, w in cols), rh, WHITE if i % 2 == 0 else ROW_ALT)
    add_rect(s, tx, ry, 0.6, rh, ORANGE)
    cxp = tx
    tier_col = HOMEBASE if "Platinum" in row[1] else RGBColor(0xB4, 0x53, 0x1A)
    styles = [
        {"size": 9.8, "bold": True, "color": DARK},
        {"size": 9, "bold": True, "color": tier_col, "align": PP_ALIGN.CENTER},
        {"size": 8.8, "color": DARK, "align": PP_ALIGN.CENTER},
        {"size": 8.8, "color": RED, "bold": True, "align": PP_ALIGN.CENTER},
        {"size": 8.5, "color": GREY_TXT},
        {"size": 8.5, "color": PURPLE},
    ]
    for j, (name, w) in enumerate(cols):
        val = row[j]
        add_text(s, cxp + 0.6, ry, w - 1.0, rh, [(val, styles[j])], anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.02)
        cxp += w
add_rect(s, tx, ty, sum(w for _, w in cols), hh + len(sla_rows) * rh, None, line=LINE_GREY, line_w=1.0)
add_text(s, 7, 84.0, 146, 3.5,
         [("", {"runs": [
             ("Quer: ", {"bold": True, "color": PURPLE, "size": 9.5}),
             ("SOC (Cyber Defense) – 24×7-Monitoring, SecInc-Reaktion 30 min, UC mit MDR-Provider; "
              "als Sicherheits-OLA in jeden Stream eingebunden.", {"color": GREY_TXT, "size": 9.5})]})])

# ==========================================================================
# SLIDE 9 -- Kapazitaets-Allokationsmatrix (Heatmap)
# ==========================================================================
s = prs.slides.add_slide(BLANK)
slide_header(s, "STEUERUNG", "Kapazitäts-Allokationsmatrix – Homebase × Mannschaft", 9)
add_text(s, 7, 19.5, 146, 4,
         [("Wie viel Kapazität sagt jede Homebase welcher Mannschaft zu? Editierbares Steuerungs-"
           "Template fürs Capacity-Board (Excel beiliegend). Beispielwerte in %.",
           {"size": 11, "color": GREY_TXT})], line_spacing=1.1)

hb_short = ["Workplace Solutions", "Infrastructure & Platforms", "Local IT Operations",
            "Business Solutions", "Application Technology", "Cyber Defense/SecOps",
            "Cyber GRC", "EA & Innovation"]
hb_fte = [18, 30, 22, 26, 14, 12, 8, 6]
col_short = ["M.Work", "Network", "Compute", "IAM", "SAP", "Data", "Svc Desk", "Chapter", "Run&Line"]
matrix = [
    [70, 0, 0, 0, 0, 0, 5, 15, 10],
    [0, 25, 30, 20, 0, 0, 0, 10, 15],
    [30, 10, 0, 0, 0, 0, 40, 5, 15],
    [0, 0, 0, 5, 40, 25, 0, 15, 15],
    [0, 0, 0, 0, 60, 10, 0, 15, 15],
    [5, 5, 5, 10, 0, 0, 5, 20, 50],
    [0, 0, 0, 10, 0, 0, 0, 30, 60],
    [0, 0, 0, 0, 0, 20, 0, 40, 40],
]
def heat(v):
    if v == 0:
        return RGBColor(0xF4, 0xF6, 0xF8)
    t = min(v, 70) / 70.0
    r = int(0xFD + (0xF0 - 0xFD) * t); g = int(0xE6 + (0x7F - 0xE6) * t); b = int(0xCC + (0x12 - 0xCC) * t)
    return RGBColor(r, g, b)

gx, gy = 7, 26
lblw, ftew = 33, 8
cellw = (152 - gx - lblw - ftew) / len(col_short)
hh = 6.5
# Kopf
add_text(s, gx, gy, lblw, hh, [("Homebase", {"size": 9, "bold": True, "color": WHITE})], anchor=MSO_ANCHOR.MIDDLE)
add_rect(s, gx, gy, lblw, hh, DARK); add_text(s, gx + 0.6, gy, lblw - 1, hh, [("Homebase", {"size": 9, "bold": True, "color": WHITE})], anchor=MSO_ANCHOR.MIDDLE)
add_rect(s, gx + lblw, gy, ftew, hh, DARK); add_text(s, gx + lblw, gy, ftew, hh, [("FTE", {"size": 8.5, "bold": True, "color": WHITE, "align": PP_ALIGN.CENTER})], anchor=MSO_ANCHOR.MIDDLE)
for j, cn in enumerate(col_short):
    cxx = gx + lblw + ftew + j * cellw
    add_rect(s, cxx, gy, cellw, hh, DARK)
    add_text(s, cxx, gy, cellw, hh, [(cn, {"size": 7.6, "bold": True, "color": WHITE, "align": PP_ALIGN.CENTER})], anchor=MSO_ANCHOR.MIDDLE)
rh = 5.4
for i, hb in enumerate(hb_short):
    ry = gy + hh + i * rh
    is_ms = False
    add_rect(s, gx, ry, lblw, rh, WHITE if i % 2 == 0 else ROW_ALT)
    add_rect(s, gx, ry, 0.7, rh, HOMEBASE)
    add_text(s, gx + 1.2, ry, lblw - 1.5, rh, [(hb, {"size": 8.3, "bold": True, "color": DARK})], anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, gx + lblw, ry, ftew, rh, WHITE if i % 2 == 0 else ROW_ALT)
    add_text(s, gx + lblw, ry, ftew, rh, [(str(hb_fte[i]), {"size": 8.5, "bold": True, "color": HOMEBASE, "align": PP_ALIGN.CENTER})], anchor=MSO_ANCHOR.MIDDLE)
    for j, v in enumerate(matrix[i]):
        cxx = gx + lblw + ftew + j * cellw
        add_rect(s, cxx, ry, cellw, rh, heat(v))
        if v > 0:
            tc = WHITE if v >= 45 else DARK
            add_text(s, cxx, ry, cellw, rh, [(str(v), {"size": 8.3, "bold": v >= 40, "color": tc, "align": PP_ALIGN.CENTER})], anchor=MSO_ANCHOR.MIDDLE)
gh = hh + len(hb_short) * rh
add_rect(s, gx, gy, lblw + ftew + len(col_short) * cellw, gh, None, line=LINE_GREY, line_w=1.0)

# Angebot/Nachfrage-Leiste
add_text(s, 7, 76.5, 146, 3.5, [("Angebot vs. Nachfrage (FTE) – aus % × FTE berechnet", {"size": 11, "bold": True, "color": DARK})])
sn = [("Modern Workplace", 19.8, 22), ("Network", 10.3, 10), ("Compute", 9.6, 10),
      ("IAM", 9.3, 9), ("SAP", 18.8, 18), ("Data", 9.1, 10), ("Service Desk", 10.3, 10)]
bw = 20.6
for i, (nm, sup, dem) in enumerate(sn):
    lx = 7 + i * (bw + 0.3)
    gap = sup - dem
    col = GREEN if gap >= 0 else RED
    bg = RGBColor(0xE3, 0xF2, 0xEA) if gap >= 0 else RGBColor(0xFD, 0xE2, 0xE2)
    add_rect(s, lx, 80, bw, 7.5, bg, round_=True)
    add_text(s, lx + 0.4, 80.3, bw - 0.8, 3, [(nm, {"size": 7.6, "bold": True, "color": DARK})])
    add_text(s, lx + 0.4, 83.2, bw - 0.8, 4,
             [("", {"runs": [
                 (f"{sup:.1f}", {"size": 9, "bold": True, "color": HOMEBASE}),
                 (" / ", {"size": 8, "color": GREY_TXT}),
                 (f"{dem:.0f}", {"size": 9, "color": DARK}),
                 (f"  ({gap:+.1f})", {"size": 8.5, "bold": True, "color": col})]})])

# ==========================================================================
# SLIDE 10 -- Status quo: Support-Struktur & Probleme (Ist)
# ==========================================================================
s = prs.slides.add_slide(BLANK)
slide_header(s, "OPERATIONS · STATUS QUO", "Heute koordiniert jeder Service Owner seinen Support selbst", 10)
add_text(s, 7, 19.3, 90, 4,
         [("Beispiel BSS: ein Service = Applikation + DXC (AMS) + TSS-Services (Windows Server) + "
           "Service Desk. Alle Koordinationslinien laufen beim Service Owner zusammen.",
           {"size": 10.5, "color": GREY_TXT})], line_spacing=1.12)

# --- Diagramm links (Service Owner strahlt auf verstreute Einheiten) ---
so_cx, so_y, so_w, so_h = 21, 26, 26, 8.5
add_rect(s, so_cx, so_y, so_w, so_h, ORANGE, round_=True)
add_text(s, so_cx, so_y, so_w, so_h, [("Service Owner (BSS)", {"size": 10.5, "bold": True, "color": WHITE, "align": PP_ALIGN.CENTER}),
                                       ("koordiniert ALLES selbst", {"size": 8.5, "color": RGBColor(0xFD,0xE6,0xCC), "align": PP_ALIGN.CENTER})], anchor=MSO_ANCHOR.MIDDLE)
so_bottom = (so_cx + so_w / 2, so_y + so_h)

nodes = [
    ("Business-Applikation\n(BSS)", HOMEBASE_LT, HOMEBASE, 3, 40),
    ("AMS · DXC\nSAP L1–L3", RGBColor(0xEB,0xE6,0xF5), PURPLE, 27.5, 40),
    ("weitere AMS-\nDienstleister", RGBColor(0xEB,0xE6,0xF5), PURPLE, 52, 40),
    ("TSS · Windows Server\nInfra-Operations", HOMEBASE_LT, HOMEBASE, 3, 52.5),
    ("Zentraler\nService Desk (L1)", HOMEBASE_LT, HOMEBASE, 27.5, 52.5),
    ("Regional\nSupport Center", HOMEBASE_LT, HOMEBASE, 52, 52.5),
]
nw, nh = 20, 9
for txt, bg, edge, nx, ny in nodes:
    connect(s, so_bottom[0], so_bottom[1], nx + nw / 2, ny, RED, 1.25)
for txt, bg, edge, nx, ny in nodes:
    add_rect(s, nx, ny, nw, nh, bg, line=edge, line_w=1.0, round_=True)
    lines = txt.split("\n")
    add_text(s, nx, ny, nw, nh, [(lines[0], {"size": 8.6, "bold": True, "color": DARK, "align": PP_ALIGN.CENTER})]
             + [(l, {"size": 7.8, "color": GREY_TXT, "align": PP_ALIGN.CENTER}) for l in lines[1:]],
             anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.0)
add_text(s, 3, 62.5, 66, 4, [("Jede Linie = manuelle, individuelle Koordination — keine Klammer, kein Standard.",
                              {"size": 8.5, "italic": True, "color": RED, "align": PP_ALIGN.CENTER})])

# --- Probleme rechts ---
add_text(s, 74, 25, 79, 4, [("Probleme, die daraus entstehen", {"size": 13, "bold": True, "color": DARK})])
probs = [
    ("Rollenüberladung Service Owner", "Service fachlich verantworten UND alle Service Levels koordinieren = zwei Jobs."),
    ("Keine Klammerfunktion", "Niemand ist übergeordnet für alle Supportaktivitäten zuständig."),
    ("Fragmentierte Provider-Steuerung", "Jeder SO steuert DXC & Co. einzeln — keine Skaleneffekte, uneinheitliche SLAs."),
    ("Kein End-to-End-SLA über die Kette", "App (DXC) + Infra (TSS) + Desk je eigene Zusage — niemand verantwortet das Ganze."),
    ("Ungeregelte Schnittstellen", "An Übergaben BSS↔TSS↔AMS↔SD versanden Tickets & Eskalationen."),
    ("Personen- statt Prozessabhängigkeit", "„SO kennt Service am besten“ → Key-Person-Risiko, keine Skalierung, keine Transparenz."),
]
py = 30
for i, (t, d) in enumerate(probs):
    add_rect(s, 74, py, 79, 8.6, WHITE, line=LINE_GREY, line_w=0.75, round_=True)
    add_rect(s, 74, py, 0.7, 8.6, RED)
    add_text(s, 75.4, py + 0.5, 77, 3.4, [("", {"runs": [(f"P{i+1}  ", {"bold": True, "color": RED, "size": 8.5}),
             (t, {"bold": True, "color": DARK, "size": 10})]})])
    add_text(s, 75.4, py + 4.0, 77, 4.2, [(d, {"size": 8.5, "color": GREY_TXT})], line_spacing=1.02)
    py += 9.3
add_text(s, 7, 84.0, 146, 3.5,
         [("", {"runs": [("Kernbefund:  ", {"bold": True, "color": RED, "size": 10}),
             ("Die IPS hat die richtigen Bausteine, aber keinen Dirigenten — Koordination ist Privatsache "
              "des Service Owners. Genau hier fordert BSS die „dedizierte Ressource“.", {"color": GREY_TXT, "size": 10})]})])

# ==========================================================================
# SLIDE 11 -- Zielmodell: Operations-Klammer (Soll)
# ==========================================================================
s = prs.slides.add_slide(BLANK)
slide_header(s, "OPERATIONS · ZIELMODELL", "Eine Klammer trennt „WAS“ vom „WIE“ und bündelt die Koordination", 11)
add_text(s, 7, 19.3, 146, 4,
         [("Der Service Owner bleibt fürs Ergebnis (WAS) accountable. Die operative Koordination über "
           "alle Provider und internen Einheiten (WIE) übernimmt ein Service Delivery Manager.",
           {"size": 10.5, "color": GREY_TXT})], line_spacing=1.12)

# --- Fluss links: SO -> SDM -> Liefereinheiten ---
add_rect(s, 7, 25, 42, 8, ORANGE, round_=True)
add_text(s, 7, 25, 42, 8, [("Service Owner", {"size": 10.5, "bold": True, "color": WHITE, "align": PP_ALIGN.CENTER}),
                           ("Ergebnis · WAS", {"size": 8.3, "color": RGBColor(0xFD,0xE6,0xCC), "align": PP_ALIGN.CENTER})], anchor=MSO_ANCHOR.MIDDLE)
connect(s, 28, 33, 28, 36, GREEN, 1.5)
add_rect(s, 7, 36, 42, 8.5, ORANGE_LT, line=PURPLE, line_w=1.25, round_=True)
add_text(s, 7, 36, 42, 8.5, [("Service Delivery Manager", {"size": 10.5, "bold": True, "color": RGBColor(0x7a,0x3d,0x05), "align": PP_ALIGN.CENTER}),
                             ("Koordination · WIE   (neu)", {"size": 8.3, "color": RGBColor(0xB4,0x53,0x1A), "align": PP_ALIGN.CENTER})], anchor=MSO_ANCHOR.MIDDLE)
add_text(s, 30, 32.4, 19, 3, [("Ziele, SLA →", {"size": 7.5, "italic": True, "color": GREEN})])

# Liefereinheiten-Reihe
deliv = [("AMS · DXC", PURPLE), ("weitere AMS /\nOps-Provider", PURPLE),
         ("TSS · Infra-Ops", HOMEBASE), ("Service Desk\n(L1)", HOMEBASE), ("Regional\nSupport Center", HOMEBASE)]
dw = 8.0; dgap = 0.5; dx0 = 7; dy = 49
for i, (t, edge) in enumerate(deliv):
    dx = dx0 + i * (dw + dgap)
    connect(s, 28, 44.5, dx + dw / 2, dy, GREEN, 1.0)
for i, (t, edge) in enumerate(deliv):
    dx = dx0 + i * (dw + dgap)
    bg = RGBColor(0xEB,0xE6,0xF5) if edge == PURPLE else HOMEBASE_LT
    add_rect(s, dx, dy, dw, 8.5, bg, line=edge, line_w=0.9, round_=True)
    lines = t.split("\n")
    add_text(s, dx, dy, dw, 8.5, [(lines[0], {"size": 7.0, "bold": True, "color": DARK, "align": PP_ALIGN.CENTER})]
             + [(l, {"size": 6.5, "color": GREY_TXT, "align": PP_ALIGN.CENTER}) for l in lines[1:]],
             anchor=MSO_ANCHOR.MIDDLE, line_spacing=0.95)

# --- Klammer rechts: SIAM im IPS MO ---
add_rect(s, 84, 25, 69, 22.5, RGBColor(0xE3,0xF2,0xEA), line=GREEN, line_w=1.25, round_=True)
add_text(s, 85.5, 26, 66, 6, [("KLAMMER: Service Integration & Operations (SIAM)", {"size": 10, "bold": True, "color": RGBColor(0x1c,0x4a,0x35)}),
                              ("verankert im IPS Management Office", {"size": 8.5, "italic": True, "color": GREEN})], line_spacing=1.05)
gov = ["Practice Supplier Mgmt", "Service Level Mgmt", "Major Incident Mgmt", "Prozess-Owner (Inc/Prob/Change)"]
for i, g in enumerate(gov):
    gx = 85.5 + (i % 2) * 33.5; gy = 33 + (i // 2) * 6.7
    add_rect(s, gx, gy, 32, 5.8, WHITE, line=GREEN, line_w=0.75, round_=True)
    add_text(s, gx, gy, 32, 5.8, [(g, {"size": 8.3, "bold": True, "color": RGBColor(0x1c,0x4a,0x35), "align": PP_ALIGN.CENTER})], anchor=MSO_ANCHOR.MIDDLE)
connect(s, 84, 40, 49, 40.2, GREEN, 1.25, dash=True)
add_text(s, 50, 44.6, 34, 3.5, [("Standards · Verträge · Prozesse", {"size": 7.5, "italic": True, "color": GREEN})])

# --- Unten: neue Rollen + Verortung ---
add_text(s, 7, 60, 146, 3.5, [("Zusätzlich benötigt — zwei neue Rollen, ein neuer Ort", {"size": 12.5, "bold": True, "color": DARK})])
newroles = [
    (ORANGE, "Service Delivery Manager (neu)", "Die „dedizierte Ressource“, die BSS fordert: koordiniert alle AMS-Provider + interne Support-Einheiten (TSS, Service Desk, RSC) je Service-Cluster."),
    (GREEN, "SIAM / Service Integration Lead (neu)", "Führt die Klammer: End-to-End-Prozesse, Provider-Integration, konsolidierte Service- & Kostensicht."),
    (HOMEBASE, "Verortung: IPS Management Office", "Neutral über BSS/TSS/Cyber; Andockpunkte (Governance, ITSM, Practice Supplier Mgmt, ServiceNow) sind dort schon vorhanden."),
]
cw = 47.3
for i, (col, t, d) in enumerate(newroles):
    lx = 7 + i * (cw + 2.5)
    add_rect(s, lx, 64.5, cw, 17, WHITE, line=LINE_GREY, line_w=0.9, round_=True)
    add_rect(s, lx, 64.5, cw, 0.9, col)
    add_text(s, lx + 1.6, 66, cw - 3, 4, [(t, {"size": 10.3, "bold": True, "color": DARK})], line_spacing=1.05)
    add_text(s, lx + 1.6, 71.5, cw - 3, 9, [(d, {"size": 9, "color": GREY_TXT})], line_spacing=1.12)

add_text(s, 7, 83.5, 146, 4,
         [("", {"runs": [("Fazit:  ", {"bold": True, "color": GREEN, "size": 10}),
             ("Keine neue Säule — eine Klammer. SIAM im IPS MO mit dem SDM als operativem Koordinator "
              "entlastet die Service Owner, bündelt die Provider-Steuerung und schließt die Schnittstellen.",
              {"color": GREY_TXT, "size": 10})]})])

prs.save(DST)
print("saved", DST, "with", len(prs.slides._sldIdLst), "slides")
