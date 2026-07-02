"""Business Case SEPM v2 - exakt im STIHL-Master-Format.
Layout 'Nur Titel' + nur Titel-Placeholder + Content ab T=1.05" wie
in echten STIHL-Slides.
"""
from copy import deepcopy
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

SRC_TEMPLATE = "/root/.claude/uploads/e26c03e3-000c-5c76-aa93-1101c0a1be8f/7360a6f1-2026_07_01_SPEM_PPB.pptx"
DST_PPTX = "/tmp/2026_07_01_Business_Case_SEPM_STIHL_v2.pptx"

# STIHL colors
ORANGE   = RGBColor(0xF0, 0x7F, 0x12)
ORANGE_LT= RGBColor(0xFD, 0xE6, 0xCC)
ORANGE2  = RGBColor(0xE9, 0x6C, 0x0C)
DARK     = RGBColor(0x1F, 0x29, 0x37)
GREY_TXT = RGBColor(0x4B, 0x55, 0x63)
LIGHT_BG = RGBColor(0xF7, 0xF7, 0xF7)
ROW_ALT  = RGBColor(0xEE, 0xF1, 0xF4)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
NAVY     = RGBColor(0x2C, 0x3E, 0x50)
GREEN    = RGBColor(0x2E, 0x7D, 0x32)
GREEN_LT = RGBColor(0xDD, 0xF1, 0xDE)
BLUE_HD  = RGBColor(0x3B, 0x6B, 0xA5)
RED_HL   = RGBColor(0xB9, 0x1C, 0x1C)
YELLOW   = RGBColor(0xF5, 0xC2, 0x14)
LIGHT_GREY = RGBColor(0xD1, 0xD5, 0xDB)
GREY_BG  = RGBColor(0xE9, 0xED, 0xF1)

# Positions in INCHES (native STIHL scale)
INCH = 914400

# Open STIHL template and remove existing content slides
prs = Presentation(SRC_TEMPLATE)
sldIdLst = prs.slides._sldIdLst
for sld in list(sldIdLst):
    rId = sld.get(qn("r:id"))
    sldIdLst.remove(sld)
    try:
        prs.part.drop_rel(rId)
    except Exception:
        pass

# Layouts
LAYOUT_COVER = prs.slide_masters[0].slide_layouts[2]   # 'Title slide empty'
LAYOUT_TITLE = prs.slide_masters[0].slide_layouts[4]   # 'Nur Titel'
LAYOUT_CONCL = prs.slide_masters[0].slide_layouts[18]  # 'Conclusion'

# Content area (inches) - matches real STIHL slide 5
CONTENT_TOP  = 1.05   # inches (right below title area ending at 0.88")
CONTENT_BOT  = 7.05   # inches (above STIHL footer at 7.21")
CONTENT_LEFT = 0.33   # inches (matches slide 5)
CONTENT_RIGHT = 13.00 # inches
CONTENT_W = CONTENT_RIGHT - CONTENT_LEFT   # 12.67
CONTENT_H = CONTENT_BOT - CONTENT_TOP      # 6.00

def add_rect(slide, left_in, top_in, w_in, h_in, fill, line=None,
             shape=MSO_SHAPE.RECTANGLE, line_w=1.0):
    s = slide.shapes.add_shape(shape, int(left_in * INCH), int(top_in * INCH),
                               int(w_in * INCH), int(h_in * INCH))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line; s.line.width = Pt(line_w)
    s.shadow.inherit = False
    return s

def add_oval(slide, left_in, top_in, w_in, h_in, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                               int(left_in * INCH), int(top_in * INCH),
                               int(w_in * INCH), int(h_in * INCH))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    s.line.fill.background()
    return s

def add_diamond(slide, left_in, top_in, w_in, h_in, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.DIAMOND,
                               int(left_in * INCH), int(top_in * INCH),
                               int(w_in * INCH), int(h_in * INCH))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    s.line.fill.background()
    return s

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
        run.font.name = "Calibri"  # STIHL uses Calibri as fallback in Textplatzhalter
        run.font.size = Pt(st.get("size", default_size))
        run.font.bold = st.get("bold", False)
        run.font.italic = st.get("italic", False)
        run.font.color.rgb = st.get("color", default_color)

def add_text(slide, left_in, top_in, w_in, h_in, paragraphs, *,
             anchor=MSO_ANCHOR.MIDDLE, **kw):
    tb = slide.shapes.add_textbox(int(left_in * INCH), int(top_in * INCH),
                                  int(w_in * INCH), int(h_in * INCH))
    tb.fill.background(); tb.line.fill.background()
    tb.text_frame.vertical_anchor = anchor
    set_paragraphs(tb, paragraphs, **kw)
    return tb

def set_placeholder_text(slide, ph_idx, text):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == ph_idx:
            ph.text_frame.text = text
            return

def add_slide_stihl(title):
    """Create a STIHL 'Nur Titel' slide, populate title only (no subline)."""
    s = prs.slides.add_slide(LAYOUT_TITLE)
    set_placeholder_text(s, 0, title)
    return s

# ------------------------------------------------------------
# Assumptions
# ------------------------------------------------------------
LICENSE_Y1 = 720
LICENSE_Y_STEADY = 850
IMPL_ONE = 1300
INT_FTE_IMPL = 4
CHANGE_ONE = 250
SUPPORT_INT_FTE_STEADY = 3.0
SUPPORT_EXT = 180
FTE_RATE = 130

CONSOL_TOOLS = [
    ("PIT / PLANTA",         "Projekt-/Ressourcen-Mgmt",  380),
    ("MS Project Online",    "Klass. Projekt-Mgmt",       120),
    ("Workpath",             "Strategie / OKR",           150),
    ("Power Apps",           "Custom-Formulare",           90),
    ("SharePoint-Listen",    "Demand-Listen",              70),
    ("Lucom (EV, teilweise)","EV-Prozess",                140),
]
TOTAL_TOOL_SAVINGS = sum(x[2] for x in CONSOL_TOOLS)
INT_SAVINGS = 220

VALUE_CATEGORIES = [
    ("Effizienz PM/PPM",       "Weniger manuelle Excels/Berichte", 820),
    ("Time-to-Market",         "Schnellere Idee->Projekt-Start",   480),
    ("Governance / Rework",    "Weniger Doppelarbeit",             360),
    ("Ressourcen-Optimierung", "Skill-Match, Auslastungs-Steuerung",520),
    ("Compliance / Audit",     "Nachvollziehbarkeit",              140),
]
TOTAL_VALUE = sum(x[2] for x in VALUE_CATEGORIES)

FLANKING = [
    ("Massnahmen-Management Zielkostenmgmt Produktentwicklung",
     "Zentraler Massnahmen-Backlog", 380),
    ("Innovations- / Vorhaben-Trichter (LH/CRD -> Steckbrief -> Projekt)",
     "Durchgaengige Pipeline", 220),
    ("Application Portfolio Mgmt (APM)",
     "Lifecycle-Management",         160),
    ("Benefits Realization Tracking",
     "Ist-/Soll-Nutzenabgleich",     120),
    ("Compliance / EV-Prozess-Digitalisierung",
     "EV-Workflow in SNOW",          180),
]
TOTAL_FLANKING = sum(x[2] for x in FLANKING)

SNOW_HISTORY = [
    (2014, "Erstvertrag ITSM",       "Standard Package", 180),
    (2016, "Erweiterung Global",     "Ausrollung inkl. China", 260),
    (2018, "Enterprise-Vertrag",     "Volumen-Diskont, HR-SD add-on", 340),
    (2020, "ITOM + CMDB",            "Discovery, Event Mgmt", 410),
    (2022, "Contract Renewal 3J",    "Preisanpassung +5% p.a.", 490),
    (2024, "Now Assist Preview",     "GenAI eingefuehrt", 560),
    (2026, "SPM Pro (geplant)",      "Neuer Modul-Baustein", 820),
]

RISKS = [
    ("Vendor Lock-in", 4, "Standard-Konfiguration, Migrations-Klauseln, Marktscreening"),
    ("Preis-Eskalation", 3, "Mehrjahresvertrag mit Preisgarantie, Volumen-Rabatte"),
    ("Ausfall/Betrieb", 2, "SLA 99,9% + DC Frankfurt/Duesseldorf, DSGVO"),
    ("Roadmap-Divergenz", 3, "Standard first, halbjaehrl. Review"),
    ("Datenhoheit / DSGVO", 2, "DE-Hosting, ADV-Vertrag, verschluesselt"),
    ("Team-Kompetenzaufbau", 3, "Ausbau ITSM-Team +1,5-2 FTE, Ausbildung"),
]

# ============================================================
# COVER
# ============================================================
def slide_cover():
    s = prs.slides.add_slide(LAYOUT_COVER)
    set_placeholder_text(s, 0, "Business Case SEPM")
    set_placeholder_text(s, 13, "Detaillierte Wirtschaftlichkeits-Betrachtung  ·  TCO · ROI · Business Value · Szenarien · Risiken")
    set_placeholder_text(s, 14, "Alex Passaro · Torsten Zahn")
    set_placeholder_text(s, 15, "01.07.2026")
slide_cover()

# ============================================================
# SLIDE 2: Management Summary
# ============================================================
def slide_management_summary():
    s = add_slide_stihl("Management Summary")
    kpis = [
        ("5-Jahres-TCO Vorteil",   "~ 2,2 M€",  "vs Best-of-Breed", GREEN),
        ("Payback",                "~ 3,5 J.",  "nach Go-Live Phase 1", ORANGE),
        ("Jaehrl. Nutzen (steady)","~ 4,3 M€",  "Konsolid.+Value+Flanking", NAVY),
        ("Empfohlenes Szenario",   "A · 24 Mo", "Full Scope Phase 1+2", BLUE_HD),
    ]
    top = CONTENT_TOP
    kpi_h = 1.55
    gap = 0.10
    w = (CONTENT_W - 3 * gap) / 4
    for i, (label, big, sub, col) in enumerate(kpis):
        x = CONTENT_LEFT + i * (w + gap)
        add_rect(s, x, top, w, kpi_h, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=col, line_w=1.5)
        add_rect(s, x, top, w, 0.30, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x + 0.05, top + 0.02, w - 0.10, 0.26, [
            (label, dict(size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.05, top + 0.40, w - 0.10, 0.70, [
            (big, dict(size=24, bold=True, color=col, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.05, top + 1.15, w - 0.10, 0.35, [
            (sub, dict(size=9.5, italic=True, color=GREY_TXT, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    km_top = top + kpi_h + 0.20
    add_text(s, CONTENT_LEFT, km_top, CONTENT_W, 0.30, [
        ("Kernbotschaften", dict(size=13, bold=True, color=ORANGE))])
    msgs = [
        ("Konsolidierung von >6 PPM-Tools", "spart ~950 k€ / Jahr an Lizenz + Wartung + FTE"),
        ("Business Value SPM (steady state)", "~2,3 M€ / Jahr durch Effizienz, Time-to-Market, Governance"),
        ("Flankierende Prozesse", "zusaetzlich ~1,0 M€ / Jahr Einsparpotenzial"),
        ("SNOW-Plattform-Risiko", "Vendor Lock-in adressiert durch Standard-first"),
        ("Vertragshistorie 2014-heute", "SNOW als bewaehrter STIHL-Partner (ITSM seit 12 J.)"),
    ]
    mb_top = km_top + 0.35
    row_h = 0.55
    for i, (head, body) in enumerate(msgs):
        y = mb_top + i * row_h
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        add_rect(s, CONTENT_LEFT, y, CONTENT_W, row_h - 0.04, bg,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, CONTENT_LEFT, y, 0.07, row_h - 0.04, ORANGE)
        add_text(s, CONTENT_LEFT + 0.15, y + 0.02, 5.5, row_h - 0.08, [
            (head, dict(size=11, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, CONTENT_LEFT + 5.70, y + 0.02, 7.0, row_h - 0.08, [
            (body, dict(size=10, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)
slide_management_summary()

# ============================================================
# SLIDE 3: Business Case Methodik
# ============================================================
def slide_methodik():
    s = add_slide_stihl("Business Case Methodik")
    top = CONTENT_TOP
    h = 2.85
    gap = 0.15
    w = (CONTENT_W - gap) / 2
    quads = [
        ("TCO (5 Jahre)", "Lizenzen · Implementierung · Support · Change",
         ["Lizenzen SNOW SPM Pro", "Implementierung",
          "Change & Schulung", "Support / Wartung"], ORANGE, "€"),
        ("Nutzen / Business Value", "Effizienz · Time-to-Market · Governance",
         ["Effizienz PPM", "Time-to-Market",
          "Ressourcen-Steuerung", "Governance & Compliance"], GREEN, "*"),
        ("Konsolidierung", "Tool-Abloese · Schnittstellen · flankierende Prozesse",
         ["PIT / MSPO / Workpath", "Schnittstellen",
          "Massnahmen-Mgmt", "Application Portfolio"], NAVY, ">>"),
        ("Risiko / Sensitivitaet", "Vendor · Preis · Betrieb · DSGVO",
         ["Vendor Lock-in", "Preis-Eskalation",
          "Ausfall / Betrieb", "Sensitivity"], RED_HL, "!"),
    ]
    for i, (head, sub, items, color, icon) in enumerate(quads):
        row = i // 2
        col = i % 2
        x = CONTENT_LEFT + col * (w + gap)
        y = top + row * (h + gap)
        add_rect(s, x, y, w, h, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=color, line_w=1.5)
        add_rect(s, x, y, w, 0.42, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_oval(s, x + 0.08, y + 0.06, 0.30, 0.30, WHITE)
        add_text(s, x + 0.08, y + 0.06, 0.30, 0.30, [
            (icon, dict(size=13, bold=True, color=color, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.45, y + 0.02, w - 0.55, 0.38, [
            (head, dict(size=13, bold=True, color=WHITE))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.15, y + 0.52, w - 0.30, 0.30, [
            (sub, dict(size=10, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)
        for j, it in enumerate(items):
            iy = y + 0.90 + j * 0.42
            add_oval(s, x + 0.20, iy + 0.13, 0.08, 0.08, color)
            add_text(s, x + 0.35, iy, w - 0.50, 0.36, [
                (it, dict(size=10.5, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
slide_methodik()

# ============================================================
# SLIDE 4: Alternative PM-Plattform (Punkt 3 Folie 4)
# ============================================================
def slide_alternative_pm():
    s = add_slide_stihl("Alternative PM-Plattform - Bewertung  (Punkt 3 Folie 4)")
    left_w = (CONTENT_W - 0.15) / 2   # 6.26"
    right_x = CONTENT_LEFT + left_w + 0.15
    top1 = CONTENT_TOP
    top1_h = 2.90
    add_rect(s, CONTENT_LEFT, top1, left_w, top1_h, LIGHT_BG,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, CONTENT_LEFT, top1, left_w, 0.35, NAVY,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.10, top1, left_w - 0.20, 0.35, [
        ("Warum Alternative bewerten?",
         dict(size=11.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    rat = [
        ("Risiko-Absicherung", "Falls SNOW fuer VEW/VPM nicht tief genug."),
        ("Verhandlungsposition", "Best-of-Breed-Angebot senkt SNOW-Preis."),
        ("Engineering-Spezifika", "PLM-nahe Prozesse ggf. besser in Fach-Tool."),
        ("EWW-Bedenken", "Explizite Alternative senkt Widerstand."),
    ]
    for i, (h, b) in enumerate(rat):
        y = top1 + 0.55 + i * 0.56
        add_oval(s, CONTENT_LEFT + 0.15, y + 0.12, 0.22, 0.22, NAVY)
        add_text(s, CONTENT_LEFT + 0.15, y + 0.12, 0.22, 0.22, [
            (str(i + 1), dict(size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, CONTENT_LEFT + 0.45, y, left_w - 0.55, 0.25, [
            (h, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, CONTENT_LEFT + 0.45, y + 0.26, left_w - 0.55, 0.24, [
            (b, dict(size=9, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)

    # Right column: Shortlist
    add_rect(s, right_x, top1, left_w, top1_h, LIGHT_BG,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, right_x, top1, left_w, 0.35, NAVY,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, right_x + 0.10, top1, left_w - 0.20, 0.35, [
        ("Shortlist Alternative PM-Plattformen",
         dict(size=11.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    alts = [
        ("Planisware Enterprise", "Marktfuehrer Engineering-PPM"),
        ("SAP EPPM", "SAP-Oekosystem-Integration"),
        ("Sciforma", "Mittelstands-Standard, hybrid"),
    ]
    for i, (n, sub) in enumerate(alts):
        y = top1 + 0.50 + i * 0.60
        add_rect(s, right_x + 0.15, y, left_w - 0.30, 0.50, WHITE,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=NAVY, line_w=0.5)
        add_rect(s, right_x + 0.15, y, 0.07, 0.50, NAVY)
        add_text(s, right_x + 0.30, y + 0.02, left_w - 1.90, 0.22, [
            (n, dict(size=10.5, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, right_x + 0.30, y + 0.24, left_w - 1.90, 0.24, [
            (sub, dict(size=9, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.TOP)
        add_rect(s, right_x + left_w - 1.60, y + 0.10, 1.40, 0.30, ORANGE,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, right_x + left_w - 1.60, y + 0.10, 1.40, 0.30, [
            ("Lead Jonas", dict(size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, right_x + 0.10, top1 + 2.42, left_w - 0.20, 0.30, [
        ("Zeitraum: 11.06. - 30.08.2026  ·  Ergebnis: klare Empfehlung mit TCO-Vergleich",
         dict(size=9.5, italic=True, color=GREY_TXT))],
        anchor=MSO_ANCHOR.MIDDLE)

    # Bottom: comparison table
    top2 = top1 + top1_h + 0.15
    top2_h = CONTENT_BOT - top2
    add_rect(s, CONTENT_LEFT, top2, CONTENT_W, top2_h, WHITE,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=LIGHT_GREY, line_w=0.7)
    add_rect(s, CONTENT_LEFT, top2, CONTENT_W, 0.35, DARK,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.10, top2, CONTENT_W - 0.20, 0.35, [
        ("Bewertungskriterien und Zwischenergebnis",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    criteria = [
        ("Funktions-Deckung",             "70%",  "80%",  "60%",  "65%"),
        ("Integration STIHL-Landschaft",  "80%",  "50%",  "70%",  "40%"),
        ("TCO 5 Jahre (rel.)",            "100%", "115%", "125%", "95%"),
        ("Change-Aufwand",                "Mittel","Hoch", "Mittel","Niedrig"),
        ("Zukunftssicherheit AI",         "Hoch", "Mittel","Hoch",  "Niedrig"),
    ]
    col_x = [CONTENT_LEFT+0.10, CONTENT_LEFT+5.20, CONTENT_LEFT+7.10, CONTENT_LEFT+9.00, CONTENT_LEFT+10.90]
    col_w = [5.00, 1.85, 1.85, 1.85, 1.65]
    headers = ["Kriterium", "SNOW SPM", "Planisware", "SAP EPPM", "Sciforma"]
    hy = top2 + 0.42
    for i, hd in enumerate(headers):
        col = ORANGE if i == 1 else NAVY
        if i == 0:
            add_text(s, col_x[i], hy, col_w[i], 0.30, [
                (hd, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        else:
            add_rect(s, col_x[i], hy, col_w[i], 0.30, col,
                     shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            add_text(s, col_x[i], hy, col_w[i], 0.30, [
                (hd, dict(size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
                anchor=MSO_ANCHOR.MIDDLE)
    for j, (k, *vals) in enumerate(criteria):
        y = top2 + 0.80 + j * 0.35
        bg = LIGHT_BG if j % 2 == 0 else ROW_ALT
        add_rect(s, CONTENT_LEFT + 0.10, y, CONTENT_W - 0.20, 0.32, bg)
        add_text(s, col_x[0], y, col_w[0], 0.32, [
            (k, dict(size=10, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        for i, v in enumerate(vals):
            add_text(s, col_x[i + 1], y, col_w[i + 1], 0.32, [
                (v, dict(size=10, color=DARK, align=PP_ALIGN.CENTER))],
                anchor=MSO_ANCHOR.MIDDLE)
slide_alternative_pm()

# ============================================================
# SLIDE 5: Komplexitäts-Reduzierung (Punkt 4 Folie 4)
# ============================================================
def slide_komplex_reduzierung():
    s = add_slide_stihl("Komplexitaets-Reduzierung ProjektControlling  (Punkt 4 Folie 4)")
    left_w = (CONTENT_W - 0.15) / 2
    right_x = CONTENT_LEFT + left_w + 0.15
    top = CONTENT_TOP
    h = CONTENT_BOT - top
    # HEUTE
    add_rect(s, CONTENT_LEFT, top, left_w, h, LIGHT_BG,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, CONTENT_LEFT, top, left_w, 0.35, RED_HL,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.10, top, left_w - 0.20, 0.35, [
        ("HEUTE  -  Komplexitaetstreiber",
         dict(size=11.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    hoy = [
        ("Multi-Tool-Landschaft PPM", "PIT · MSPO · Excel · SharePoint · Power Apps",
         "~ 2,5 FTE Betriebs- & Integrations-Aufwand"),
        ("Manuelle Daten-Aggregation", "Excel-Konsolidierung fuer Reportings",
         "~ 2 Wochen/Monat Reporting-Aufwand"),
        ("Uneinheitliche Datenmodelle", "Jede Insel eigener Struktur",
         "~ 1 FTE Datenqualitaets-Arbeit"),
        ("Doppelte Datenpflege", "Projekte in PIT + Excel + SharePoint",
         "~ 0,8 FTE Doppelerfassung"),
        ("Fehlende Traceability", "Idee-Portfolio-Projekt-Kosten nicht durchgaengig",
         "Governance-Aufwand hoch, Audits schwierig"),
    ]
    for i, (t, sub, imp) in enumerate(hoy):
        y = top + 0.50 + i * 1.05
        add_rect(s, CONTENT_LEFT + 0.15, y, left_w - 0.30, 0.95, WHITE,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=RED_HL, line_w=0.5)
        add_rect(s, CONTENT_LEFT + 0.15, y, 0.06, 0.95, RED_HL)
        add_text(s, CONTENT_LEFT + 0.30, y + 0.04, left_w - 0.50, 0.28, [
            (t, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, CONTENT_LEFT + 0.30, y + 0.30, left_w - 0.50, 0.30, [
            (sub, dict(size=9, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, CONTENT_LEFT + 0.30, y + 0.62, left_w - 0.50, 0.30, [
            (imp, dict(size=9.5, bold=True, color=RED_HL))], anchor=MSO_ANCHOR.MIDDLE)
    # MIT SEPM
    add_rect(s, right_x, top, left_w, h, LIGHT_BG,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, right_x, top, left_w, 0.35, GREEN,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, right_x + 0.10, top, left_w - 0.20, 0.35, [
        ("MIT SEPM  -  Konsolidierungs-Hebel",
         dict(size=11.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    mit = [
        ("Eine Plattform, ein Datenmodell", "SNOW SPM als Single Source of Truth",
         "- 6 Tools · - 0,5 FTE Betrieb"),
        ("Automatische Reports & Dashboards", "Live-Dashboards ersetzen Excel",
         "- 1,5 Wochen/Monat Reporting"),
        ("Standard-Datenmodell", "Vorhaben-Steckbrief einheitlich",
         "- 0,7 FTE Datenqualitaets-Arbeit"),
        ("Erfassung einmal - Nutzung ueberall", "Demand-Portfolio-Projekt in einem System",
         "- 0,6 FTE Doppelerfassung"),
        ("Durchgaengige Traceability", "OKR-Portfolio-Projekt-Kosten-Nutzen",
         "Audit-faehig, Compliance-sicher"),
    ]
    for i, (t, sub, imp) in enumerate(mit):
        y = top + 0.50 + i * 1.05
        add_rect(s, right_x + 0.15, y, left_w - 0.30, 0.95, WHITE,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=GREEN, line_w=0.5)
        add_rect(s, right_x + 0.15, y, 0.06, 0.95, GREEN)
        add_text(s, right_x + 0.30, y + 0.04, left_w - 0.50, 0.28, [
            (t, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, right_x + 0.30, y + 0.30, left_w - 0.50, 0.30, [
            (sub, dict(size=9, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, right_x + 0.30, y + 0.62, left_w - 0.50, 0.30, [
            (imp, dict(size=9.5, bold=True, color=GREEN))], anchor=MSO_ANCHOR.MIDDLE)
slide_komplex_reduzierung()

# ============================================================
# SLIDE 6: Business Value
# ============================================================
def slide_business_value():
    s = add_slide_stihl("Business Value Analyse - Nutzen-Kategorien Steady State p.a.")
    add_rect(s, CONTENT_LEFT, CONTENT_TOP, CONTENT_W, 0.70, NAVY,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.15, CONTENT_TOP, 5.5, 0.70, [
        ("Gesamt-Nutzen SEPM (Steady State p.a.)",
         dict(size=12.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, CONTENT_LEFT + 5.7, CONTENT_TOP, 5, 0.70, [
        (f"~ {TOTAL_VALUE:,} k€ / Jahr".replace(",", "."),
         dict(size=22, bold=True, color=ORANGE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    max_val = max(v[2] for v in VALUE_CATEGORIES)
    row_top = CONTENT_TOP + 0.90
    row_h = 0.90
    for i, (name, desc, val) in enumerate(VALUE_CATEGORIES):
        y = row_top + i * (row_h + 0.06)
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        add_rect(s, CONTENT_LEFT, y, CONTENT_W, row_h, bg,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, CONTENT_LEFT, y, 0.07, row_h, ORANGE)
        add_text(s, CONTENT_LEFT + 0.15, y + 0.04, 4.5, row_h - 0.10, [
            (name, dict(size=11, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, CONTENT_LEFT + 0.15, y + row_h - 0.34, 4.5, 0.28, [
            (desc, dict(size=9, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.TOP)
        bar_x = CONTENT_LEFT + 4.80
        bar_w_max = 6.30
        bar_w_val = bar_w_max * (val / max_val)
        add_rect(s, bar_x, y + 0.30, bar_w_max, 0.34, LIGHT_GREY,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, bar_x, y + 0.30, bar_w_val, 0.34, ORANGE,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, CONTENT_LEFT + 11.20, y, 1.40, row_h, [
            (f"{val:,} k€".replace(",", "."),
             dict(size=13, bold=True, color=NAVY, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
slide_business_value()

# ============================================================
# SLIDE 7: TCO 5 Jahre
# ============================================================
def slide_tco():
    s = add_slide_stihl("TCO - Total Cost of Ownership (5 Jahre)")
    total_5y = IMPL_ONE + CHANGE_ONE + 5 * (LICENSE_Y_STEADY +
                                             SUPPORT_INT_FTE_STEADY * FTE_RATE + SUPPORT_EXT)
    add_rect(s, CONTENT_LEFT, CONTENT_TOP, CONTENT_W, 0.70, DARK,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.15, CONTENT_TOP, 6, 0.70, [
        ("5-Jahres-TCO SEPM (Base Case)",
         dict(size=13, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, CONTENT_LEFT + 6.2, CONTENT_TOP, 5, 0.70, [
        (f"~ {total_5y/1000:.1f} Mio. €",
         dict(size=22, bold=True, color=ORANGE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, CONTENT_LEFT + 11.20, CONTENT_TOP, 1.60, 0.70, [
        (f"(~ {total_5y/5/1000:.2f} M€ p.a. Ø)",
         dict(size=10, italic=True, color=LIGHT_GREY, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    left_w = (CONTENT_W - 0.15) / 2
    right_x = CONTENT_LEFT + left_w + 0.15
    tt = CONTENT_TOP + 0.85
    h = CONTENT_BOT - tt
    add_rect(s, CONTENT_LEFT, tt, left_w, h, LIGHT_BG,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, CONTENT_LEFT, tt, left_w, 0.35, NAVY,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.10, tt, left_w - 0.20, 0.35, [
        ("Einmalige Kosten (Y1-Y2)",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    einmalig = [
        ("Implementierungs-Partner (SI)", IMPL_ONE),
        ("Interne FTE (Phase 1+2)", INT_FTE_IMPL * FTE_RATE * 2),
        ("Change / Schulung", CHANGE_ONE),
        ("Contingency (10%)", int(0.1 * (IMPL_ONE + INT_FTE_IMPL * FTE_RATE * 2 + CHANGE_ONE))),
    ]
    for i, (n, v) in enumerate(einmalig):
        y = tt + 0.50 + i * 0.75
        add_rect(s, CONTENT_LEFT + 0.15, y, left_w - 0.30, 0.65, WHITE,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=NAVY, line_w=0.5)
        add_text(s, CONTENT_LEFT + 0.30, y, 3.8, 0.65, [
            (n, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, CONTENT_LEFT + left_w - 1.70, y, 1.50, 0.65, [
            (f"{v:,} k€".replace(",", "."),
             dict(size=12.5, bold=True, color=NAVY, align=PP_ALIGN.RIGHT))],
            anchor=MSO_ANCHOR.MIDDLE)
    total_e = sum(x[1] for x in einmalig)
    add_text(s, CONTENT_LEFT + 0.10, tt + 3.75, left_w - 0.20, 0.40, [
        (f"Summe einmalig: {total_e:,} k€".replace(",", "."),
         dict(size=12, bold=True, color=NAVY, align=PP_ALIGN.RIGHT))],
        anchor=MSO_ANCHOR.MIDDLE)
    # Laufend
    add_rect(s, right_x, tt, left_w, h, LIGHT_BG,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, right_x, tt, left_w, 0.35, ORANGE,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, right_x + 0.10, tt, left_w - 0.20, 0.35, [
        ("Laufende Kosten p.a. (Steady State)",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    laufend = [
        ("Lizenzen SNOW SPM Pro", LICENSE_Y_STEADY),
        ("Interne Betriebs-FTE", int(SUPPORT_INT_FTE_STEADY * FTE_RATE)),
        ("Externer Support / Wartung", SUPPORT_EXT),
        ("Weiterentwicklung", 220),
    ]
    for i, (n, v) in enumerate(laufend):
        y = tt + 0.50 + i * 0.75
        add_rect(s, right_x + 0.15, y, left_w - 0.30, 0.65, WHITE,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=ORANGE, line_w=0.5)
        add_text(s, right_x + 0.30, y, 3.8, 0.65, [
            (n, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, right_x + left_w - 1.70, y, 1.50, 0.65, [
            (f"{v:,} k€".replace(",", "."),
             dict(size=12.5, bold=True, color=ORANGE, align=PP_ALIGN.RIGHT))],
            anchor=MSO_ANCHOR.MIDDLE)
    total_l = sum(x[1] for x in laufend)
    add_text(s, right_x + 0.10, tt + 3.75, left_w - 0.20, 0.40, [
        (f"Summe p.a.: {total_l:,} k€".replace(",", "."),
         dict(size=12, bold=True, color=ORANGE, align=PP_ALIGN.RIGHT))],
        anchor=MSO_ANCHOR.MIDDLE)
slide_tco()

# ============================================================
# SLIDE 8: Support-Kosten
# ============================================================
def slide_support():
    s = add_slide_stihl("Support-Kosten - Detailbetrachtung")
    cols = [
        ("Interner Betrieb", ORANGE,
         [("SNOW-Plattform-Admin", "1,0 FTE", 130),
          ("SPM Business Analyst", "1,0 FTE", 130),
          ("Integration / Schnittstellen", "0,5 FTE", 65),
          ("Data Governance",     "0,3 FTE", 39),
          ("User Support 2nd Level", "0,2 FTE", 26)]),
        ("Externer Support", NAVY,
         [("Wartungsvertrag SNOW",  "Standard 22%", 70),
          ("Managed-Service Partner","Retainer + T&M", 90),
          ("Externes Consulting", "20-30 Tage/y", 60),
          ("Zertifikate & Schulungen", "5-8 Personen", 30)]),
        ("Weiterentwicklung", GREEN,
         [("Feature-Backlog (Business)", "~10 PT/Mo", 100),
          ("Release-Upgrade-Anpassungen","2× / Jahr", 50),
          ("Custom-Reports / Dashboards","4-6 PT/Mo", 40),
          ("Innovation Sprints",         "1× Q", 30)]),
    ]
    top = CONTENT_TOP
    w = (CONTENT_W - 0.30) / 3
    gap = 0.15
    for i, (title, color, items) in enumerate(cols):
        x = CONTENT_LEFT + i * (w + gap)
        col_h = CONTENT_BOT - top
        add_rect(s, x, top, w, col_h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, x, top, w, 0.42, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x + 0.10, top + 0.03, w - 0.20, 0.36, [
            (title, dict(size=13, bold=True, color=WHITE))],
            anchor=MSO_ANCHOR.MIDDLE)
        total = 0
        for j, (n, note, val) in enumerate(items):
            y = top + 0.50 + j * 0.65
            add_rect(s, x + 0.10, y, w - 0.20, 0.60, WHITE,
                     shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=color, line_w=0.5)
            add_text(s, x + 0.20, y + 0.02, w - 1.60, 0.28, [
                (n, dict(size=10.5, bold=True, color=DARK))],
                anchor=MSO_ANCHOR.MIDDLE)
            add_text(s, x + 0.20, y + 0.30, w - 1.60, 0.28, [
                (note, dict(size=9, italic=True, color=GREY_TXT))],
                anchor=MSO_ANCHOR.MIDDLE)
            add_text(s, x + w - 1.50, y + 0.02, 1.35, 0.56, [
                (f"{val} k€",
                 dict(size=12, bold=True, color=color, align=PP_ALIGN.RIGHT))],
                anchor=MSO_ANCHOR.MIDDLE)
            total += val
        add_rect(s, x + 0.10, top + col_h - 0.55, w - 0.20, 0.35, color,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x + 0.10, top + col_h - 0.55, w - 0.20, 0.35, [
            (f"Summe: {total} k€ / Jahr",
             dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
slide_support()

# ============================================================
# SLIDE 9: Konsolidierung Tools
# ============================================================
def slide_konsolidierung_tools():
    s = add_slide_stihl("Konsolidierungseinsparungen - Tool-Abloese")
    add_rect(s, CONTENT_LEFT, CONTENT_TOP, CONTENT_W, 0.70, GREEN,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.15, CONTENT_TOP, 8, 0.70, [
        ("Gesamt Tool-Abloese-Ersparnis (Steady State p.a.)",
         dict(size=12.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, CONTENT_LEFT + 8.2, CONTENT_TOP, 4.4, 0.70, [
        (f"~ {TOTAL_TOOL_SAVINGS} k€ / Jahr",
         dict(size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    max_val = max(x[2] for x in CONSOL_TOOLS)
    row_top = CONTENT_TOP + 0.90
    row_h = 0.72
    for i, (tool, cat, saving) in enumerate(CONSOL_TOOLS):
        y = row_top + i * (row_h + 0.05)
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        add_rect(s, CONTENT_LEFT, y, CONTENT_W, row_h, bg,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, CONTENT_LEFT, y, 0.07, row_h, GREEN)
        add_text(s, CONTENT_LEFT + 0.15, y + 0.03, 3.5, row_h - 0.06, [
            (tool, dict(size=11.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, CONTENT_LEFT + 0.15, y + row_h - 0.28, 3.5, 0.22, [
            (cat, dict(size=8.5, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.TOP)
        bar_x = CONTENT_LEFT + 3.90; bar_w_max = 7.30
        bar_w_val = bar_w_max * (saving / max_val)
        add_rect(s, bar_x, y + 0.28, bar_w_max, 0.24, LIGHT_GREY,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, bar_x, y + 0.28, bar_w_val, 0.24, GREEN,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, CONTENT_LEFT + 11.30, y, 1.30, row_h, [
            (f"- {saving} k€",
             dict(size=13, bold=True, color=GREEN, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
slide_konsolidierung_tools()

# ============================================================
# SLIDE 10: Konsolidierung Schnittstellen
# ============================================================
def slide_konsolidierung_interfaces():
    s = add_slide_stihl("Konsolidierungseinsparungen - Schnittstellen")
    add_rect(s, CONTENT_LEFT, CONTENT_TOP, CONTENT_W, 0.70, GREEN,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.15, CONTENT_TOP, 8, 0.70, [
        ("Ersparnis Schnittstellen & Integration p.a.",
         dict(size=12.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, CONTENT_LEFT + 8.2, CONTENT_TOP, 4.4, 0.70, [
        (f"~ {INT_SAVINGS} k€ / Jahr",
         dict(size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    left_w = (CONTENT_W - 0.15) / 2
    right_x = CONTENT_LEFT + left_w + 0.15
    tt = CONTENT_TOP + 0.85
    h = CONTENT_BOT - tt
    add_rect(s, CONTENT_LEFT, tt, left_w, h, LIGHT_BG,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, CONTENT_LEFT, tt, left_w, 0.35, RED_HL,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.10, tt, left_w - 0.20, 0.35, [
        ("HEUTE - Fragmentierte Integration",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    heute_ints = [
        "PIT -> SAP FI (Kostentransfer)",
        "MSPO -> SAP HCM (Personalstamm)",
        "SharePoint -> Excel (Datenexport)",
        "Workpath -> Excel (OKR-Reporting)",
        "PowerApps -> Dataverse",
        "Lucom EV -> SAP",
        "Excel-Konsolidierungen (div.)",
    ]
    for i, txt in enumerate(heute_ints):
        y = tt + 0.50 + i * 0.55
        add_rect(s, CONTENT_LEFT + 0.15, y, left_w - 0.30, 0.50, WHITE,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=RED_HL, line_w=0.5)
        add_rect(s, CONTENT_LEFT + 0.15, y, 0.06, 0.50, RED_HL)
        add_text(s, CONTENT_LEFT + 0.30, y, left_w - 0.50, 0.50, [
            (txt, dict(size=10, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    # MIT SEPM
    add_rect(s, right_x, tt, left_w, h, LIGHT_BG,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, right_x, tt, left_w, 0.35, GREEN,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, right_x + 0.10, tt, left_w - 0.20, 0.35, [
        ("MIT SEPM - Native SNOW-Integration",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    mit_ints = [
        ("SAP FI Konnektor",     "Standard, Integration Hub"),
        ("SAP HCM Konnektor",    "Standard, Integration Hub"),
        ("Interne Datenmodelle", "keine Excel-Export-Ketten"),
        ("OKR-Modul nativ",      "kein Workpath-Bridge"),
        ("App Engine (Low-Code)","Custom-Data auf gleicher Plattform"),
        ("EV-Prozess nativ",     "kein Lucom-Bridge"),
        ("Live-Reporting",       "Excel-Konsolidierung entfaellt"),
    ]
    for i, (h_, sub) in enumerate(mit_ints):
        y = tt + 0.50 + i * 0.55
        add_rect(s, right_x + 0.15, y, left_w - 0.30, 0.50, WHITE,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=GREEN, line_w=0.5)
        add_rect(s, right_x + 0.15, y, 0.06, 0.50, GREEN)
        add_text(s, right_x + 0.30, y, 2.6, 0.50, [
            (h_, dict(size=10, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, right_x + 3.0, y, left_w - 3.20, 0.50, [
            (sub, dict(size=9, italic=True, color=GREEN))], anchor=MSO_ANCHOR.MIDDLE)
slide_konsolidierung_interfaces()

# ============================================================
# SLIDE 11: Flankierende Prozesse
# ============================================================
def slide_flanking():
    s = add_slide_stihl("Flankierende Prozess-Potentiale")
    add_rect(s, CONTENT_LEFT, CONTENT_TOP, CONTENT_W, 0.70, ORANGE,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.15, CONTENT_TOP, 8.5, 0.70, [
        ("Zusaetzliches Einspar-Potenzial (SPM-flankierend, p.a.)",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, CONTENT_LEFT + 8.7, CONTENT_TOP, 3.9, 0.70, [
        (f"~ {TOTAL_FLANKING:,} k€ / Jahr".replace(",", "."),
         dict(size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    max_val = max(x[2] for x in FLANKING)
    row_top = CONTENT_TOP + 0.90
    row_h = 0.86
    for i, (name, desc, val) in enumerate(FLANKING):
        y = row_top + i * (row_h + 0.06)
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        add_rect(s, CONTENT_LEFT, y, CONTENT_W, row_h, bg,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, CONTENT_LEFT, y, 0.07, row_h, ORANGE)
        add_oval(s, CONTENT_LEFT + 0.15, y + 0.25, 0.34, 0.34, ORANGE)
        add_text(s, CONTENT_LEFT + 0.15, y + 0.25, 0.34, 0.34, [
            (str(i + 1), dict(size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, CONTENT_LEFT + 0.65, y + 0.04, 8.5, 0.34, [
            (name, dict(size=11, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, CONTENT_LEFT + 0.65, y + 0.40, 8.5, 0.42, [
            (desc, dict(size=9.5, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.TOP)
        bar_x = CONTENT_LEFT + 9.30; bar_w_max = 1.80
        bar_w_val = bar_w_max * (val / max_val)
        add_rect(s, bar_x, y + 0.30, bar_w_max, 0.26, LIGHT_GREY,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, bar_x, y + 0.30, bar_w_val, 0.26, ORANGE,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, CONTENT_LEFT + 11.30, y, 1.30, row_h, [
            (f"{val} k€",
             dict(size=13, bold=True, color=NAVY, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
slide_flanking()

# ============================================================
# SLIDE 12: Szenarien-Uebersicht
# ============================================================
def slide_szenarien_uebersicht():
    s = add_slide_stihl("Einfuehrungs-Szenarien im Vergleich")
    top = CONTENT_TOP
    h = CONTENT_BOT - top
    w = (CONTENT_W - 0.30) / 3
    gap = 0.15
    scenarios = [
        ("A - Full Scope\n24 Monate", 24, ORANGE, "Empfohlen",
         [("Scope", "Alle Module Phase 1 + 2"),
          ("Ramp Up", "Standard-Timeline"),
          ("Payback", "~ 3,5 Jahre"),
          ("Business Value", "Voll ab Jahr 3"),
          ("Risiko", "Mittel")],
         True),
        ("B - Extended Ramp\n36 Monate", 36, BLUE_HD, "Konservativ",
         [("Scope", "Alle Module, verteilter"),
          ("Ramp Up", "Reduziertes Tempo"),
          ("Payback", "~ 4,5 Jahre"),
          ("Business Value", "Voll ab Jahr 4"),
          ("Risiko", "Niedrig")],
         False),
        ("C - MVP-First\n12 Mo + Ausbau", 30, GREEN, "Agil",
         [("Scope", "Strategy + Portfolio zuerst"),
          ("Ramp Up", "MVP live in 12 Mo"),
          ("Payback", "~ 3 Jahre auf Teil-Nutzen"),
          ("Business Value", "Teilnutzen frueh"),
          ("Risiko", "Mittel")],
         False),
    ]
    for i, (name, months, color, badge, items, recommend) in enumerate(scenarios):
        x = CONTENT_LEFT + i * (w + gap)
        add_rect(s, x, top, w, h, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=color, line_w=2 if recommend else 1)
        add_rect(s, x, top, w, 0.66, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        if recommend:
            add_rect(s, x + w - 1.60, top, 1.55, 0.66, DARK,
                     shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            add_text(s, x + w - 1.60, top, 1.55, 0.66, [
                ("★ EMPFOHLEN",
                 dict(size=9, bold=True, color=ORANGE, align=PP_ALIGN.CENTER))],
                anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.08, top + 0.04, w - 1.75 if recommend else w - 0.15, 0.58, [
            (name, dict(size=13, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, x + 0.15, top + 0.75, w - 0.30, 0.30, LIGHT_BG,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x + 0.15, top + 0.75, w - 0.30, 0.30, [
            (badge, dict(size=10, bold=True, color=color, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.15, top + 1.15, w - 0.30, 0.24, [
            (f"Dauer: {months} Monate",
             dict(size=9.5, italic=True, color=GREY_TXT, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        max_m = 36
        bar_w = (w - 0.30) * (months / max_m)
        add_rect(s, x + 0.15, top + 1.40, w - 0.30, 0.18, LIGHT_GREY,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, x + 0.15, top + 1.40, bar_w, 0.18, color,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        for j, (k, v) in enumerate(items):
            y = top + 1.75 + j * 0.72
            add_rect(s, x + 0.15, y, w - 0.30, 0.66, LIGHT_BG,
                     shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            add_text(s, x + 0.25, y + 0.02, w - 0.45, 0.26, [
                (k, dict(size=9, bold=True, color=GREY_TXT))], anchor=MSO_ANCHOR.MIDDLE)
            add_text(s, x + 0.25, y + 0.30, w - 0.45, 0.34, [
                (v, dict(size=10, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
slide_szenarien_uebersicht()

# ============================================================
# SLIDE 13: Szenario A Detail (compact)
# ============================================================
def slide_szenario_A():
    s = add_slide_stihl("Szenario A - Full Scope 24 Monate (Empfohlen)")
    add_text(s, CONTENT_LEFT, CONTENT_TOP, CONTENT_W, 0.25, [
        ("Umsetzungs-Timeline",
         dict(size=12.5, bold=True, color=ORANGE))])
    axis_y = CONTENT_TOP + 0.30
    n_m = 24
    lbl_w = 2.5
    left_x = CONTENT_LEFT + lbl_w
    right_x = CONTENT_LEFT + CONTENT_W
    m_w = (right_x - left_x) / n_m
    def mx(m): return left_x + m * m_w
    add_rect(s, left_x, axis_y, 12 * m_w, 0.28, ORANGE_LT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, left_x, axis_y, 12 * m_w, 0.28, [
        ("Jahr 1", dict(size=10.5, bold=True, color=DARK, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, left_x + 12 * m_w, axis_y, 12 * m_w, 0.28, ORANGE2,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, left_x + 12 * m_w, axis_y, 12 * m_w, 0.28, [
        ("Jahr 2", dict(size=10.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    phases = [
        ("Konzeption",              NAVY,     0,  3),
        ("Impl. Phase 1",           ORANGE,   3, 12),
        ("Go-Live Phase 1",         GREEN,   12, 13),
        ("Impl. Phase 2",           ORANGE2, 12, 23),
        ("Go-Live Phase 2",         GREEN,   23, 24),
        ("Change / Rollout",        BLUE_HD,  3, 24),
    ]
    lane_top = axis_y + 0.35
    lane_h = 0.36
    lane_gap = 0.05
    for i, (name, color, ms, me) in enumerate(phases):
        y = lane_top + i * (lane_h + lane_gap)
        add_text(s, CONTENT_LEFT, y, lbl_w - 0.10, lane_h, [
            (name, dict(size=9.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, left_x, y, n_m * m_w, lane_h, LIGHT_BG,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        bar_x = mx(ms); bar_w = (me - ms) * m_w - 0.02
        add_rect(s, bar_x + 0.01, y + 0.02, bar_w, lane_h - 0.04, color,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    curve_top = lane_top + len(phases) * (lane_h + lane_gap) + 0.25
    add_text(s, CONTENT_LEFT, curve_top, CONTENT_W, 0.25, [
        ("Investitions- und Nutzenprofil (schematisch)",
         dict(size=12, bold=True, color=ORANGE))])
    ct = curve_top + 0.30
    ch = CONTENT_BOT - ct
    add_rect(s, CONTENT_LEFT, ct, CONTENT_W, ch, LIGHT_BG,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    year_w = CONTENT_W / 5
    for yi in range(5):
        x = CONTENT_LEFT + yi * year_w
        col = WHITE if yi % 2 == 0 else GREY_BG
        add_rect(s, x, ct, year_w - 0.01, ch, col)
        add_text(s, x, ct + ch - 0.28, year_w, 0.24, [
            (f"Jahr {yi + 1}",
             dict(size=9, color=GREY_TXT, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    inv_vals = [2.0, 2.2, 1.2, 1.0, 1.0]
    ben_vals = [0.3, 1.5, 3.0, 4.0, 4.3]
    max_v = 4.5
    for yi in range(5):
        cx = CONTENT_LEFT + yi * year_w + year_w / 2
        h_inv = (inv_vals[yi] / max_v) * (ch - 0.5)
        add_rect(s, cx - 0.28, ct + ch - 0.35 - h_inv, 0.24, h_inv, RED_HL,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, cx - 0.55, ct + ch - 0.30 - h_inv - 0.20, 0.60, 0.20, [
            (f"{inv_vals[yi]:.1f}", dict(size=8, bold=True, color=RED_HL, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        h_ben = (ben_vals[yi] / max_v) * (ch - 0.5)
        add_rect(s, cx + 0.04, ct + ch - 0.35 - h_ben, 0.24, h_ben, GREEN,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, cx - 0.16, ct + ch - 0.30 - h_ben - 0.20, 0.60, 0.20, [
            (f"{ben_vals[yi]:.1f}", dict(size=8, bold=True, color=GREEN, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
slide_szenario_A()

# ============================================================
# SLIDE 14: Szenario B - Vergleich
# ============================================================
def slide_szenario_B():
    s = add_slide_stihl("Szenario B - Extended Ramp-Up 36 Monate")
    left_w = (CONTENT_W - 0.20) / 2
    right_x = CONTENT_LEFT + left_w + 0.20
    top = CONTENT_TOP
    h_top = 2.60
    for k, (title, color, months, txt) in enumerate([
        ("Szenario A - 24 Monate", ORANGE, 24,
         ["Standard-Tempo, 2 Jahre", "Hoeherer Change-Druck",
          "Nutzen ab Jahr 3 vollstaendig", "Payback ~ 3,5 Jahre",
          "Empfohlen bei ausr. Ressourcen"]),
        ("Szenario B - 36 Monate", BLUE_HD, 36,
         ["Verteiltes Tempo, 3 Jahre", "Weniger Change-Druck",
          "Nutzen ab Jahr 4 vollstaendig", "Payback ~ 4,5 Jahre",
          "Empfohlen bei knappen Ress."]),
    ]):
        x = CONTENT_LEFT if k == 0 else right_x
        add_rect(s, x, top, left_w, h_top, WHITE,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=color, line_w=1.5)
        add_rect(s, x, top, left_w, 0.42, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x, top, left_w, 0.42, [
            (title, dict(size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.15, top + 0.50, left_w - 0.30, 0.24, [
            (f"Dauer: {months} Monate",
             dict(size=10, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, x + 0.15, top + 0.78, left_w - 0.30, 0.20, LIGHT_GREY,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, x + 0.15, top + 0.78, (left_w - 0.30) * months / 36, 0.20, color,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        for j, tx in enumerate(txt):
            y = top + 1.05 + j * 0.30
            add_oval(s, x + 0.15, y + 0.10, 0.12, 0.12, color)
            add_text(s, x + 0.35, y, left_w - 0.50, 0.30, [
                (tx, dict(size=10, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    # Comparison table
    tt = top + h_top + 0.20
    add_text(s, CONTENT_LEFT, tt, CONTENT_W, 0.25, [
        ("Kosten- & Nutzen-Delta",
         dict(size=12.5, bold=True, color=ORANGE))])
    rows = [
        ("Einmalige Kosten",     "3,0 M€",  "2,7 M€",   "-0,3 M€"),
        ("Laufend p.a. steady",  "0,85 M€", "0,80 M€",  "-0,05 M€"),
        ("5-Jahres-TCO",         "7,0 M€",  "6,7 M€",   "-0,3 M€"),
        ("Nutzen Y3 (kumul.)",   "3,5 M€",  "1,8 M€",   "-1,7 M€"),
        ("Nutzen Y5 (kumul.)",   "16 M€",   "12 M€",    "-4 M€"),
        ("Payback",              "~ 3,5 J.","~ 4,5 J.", "+1 J."),
    ]
    hdr_y = tt + 0.30
    add_rect(s, CONTENT_LEFT, hdr_y, CONTENT_W, 0.28, DARK,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    col_positions = [CONTENT_LEFT, CONTENT_LEFT + 4.5, CONTENT_LEFT + 7.4,
                     CONTENT_LEFT + 10.2]
    col_widths = [4.5, 2.9, 2.8, 2.5]
    for i, hh in enumerate(["Kennzahl", "Szenario A", "Szenario B", "Δ B vs A"]):
        add_text(s, col_positions[i], hdr_y, col_widths[i], 0.28, [
            (hh, dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    for j, (k, va, vb, d) in enumerate(rows):
        y = hdr_y + 0.30 + j * 0.28
        bg = LIGHT_BG if j % 2 == 0 else ROW_ALT
        add_rect(s, CONTENT_LEFT, y, CONTENT_W, 0.26, bg)
        add_text(s, col_positions[0] + 0.15, y, col_widths[0] - 0.15, 0.26, [
            (k, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, col_positions[1], y, col_widths[1], 0.26, [
            (va, dict(size=10.5, color=ORANGE, bold=True, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, col_positions[2], y, col_widths[2], 0.26, [
            (vb, dict(size=10.5, color=BLUE_HD, bold=True, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, col_positions[3], y, col_widths[3], 0.26, [
            (d, dict(size=10.5, color=RED_HL if "-" in d and "M€" in d else DARK,
                     bold=True, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
slide_szenario_B()

# ============================================================
# SLIDE 15: Budgetplanung (Punkt 6)
# ============================================================
def slide_budget():
    s = add_slide_stihl("Budgetplanung SEPM  (Punkt 6 Folie 4)")
    years = [2027, 2028, 2029, 2030, 2031]
    invs = [(2200, "Phase 1 Impl. + Lizenz"),
            (2200, "Phase 2 Impl. + Lizenz"),
            (1200, "Steady + Weiterentw."),
            (1000, "Steady State"),
            (1000, "Steady State")]
    total = sum(v[0] for v in invs)
    add_rect(s, CONTENT_LEFT, CONTENT_TOP, CONTENT_W, 0.70, NAVY,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.15, CONTENT_TOP, 8, 0.70, [
        ("Gesamt-Investition SEPM 2027-2031 (Base Case)",
         dict(size=12.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, CONTENT_LEFT + 8.2, CONTENT_TOP, 4.4, 0.70, [
        (f"~ {total/1000:.1f} Mio. €",
         dict(size=22, bold=True, color=ORANGE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    chart_top = CONTENT_TOP + 1.0
    chart_h = 3.6
    chart_x = CONTENT_LEFT + 0.30; chart_w = CONTENT_W - 0.60
    n = len(years)
    slot_w = chart_w / n
    max_v = max(v[0] for v in invs)
    for i, (yr, (val, label)) in enumerate(zip(years, invs)):
        cx = chart_x + i * slot_w + slot_w / 2
        h = (val / max_v) * (chart_h - 0.6)
        color = ORANGE if i < 2 else (ORANGE2 if i < 3 else ORANGE_LT)
        add_rect(s, cx - 0.75, chart_top + chart_h - h - 0.35, 1.50, h, color,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, cx - 1.0, chart_top + chart_h - h - 0.55, 2.0, 0.20, [
            (f"{val/1000:.1f} M€",
             dict(size=10.5, bold=True, color=DARK, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, cx - 1.0, chart_top + chart_h - 0.28, 2.0, 0.24, [
            (str(yr), dict(size=11, bold=True, color=DARK, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, cx - 1.4, chart_top + chart_h - 0.02, 2.8, 0.24, [
            (label, dict(size=8, italic=True, color=GREY_TXT, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, chart_x, chart_top + chart_h - 0.35, chart_w, 0.02, DARK)
    ms_top = chart_top + chart_h + 0.30
    add_text(s, CONTENT_LEFT, ms_top, CONTENT_W, 0.24, [
        ("Milestones Jahresplanung",
         dict(size=12, bold=True, color=ORANGE))])
    ms = [
        ("Q3 2026: Jahresplanung", "Beschluss 2027 & Rahmen 2028"),
        ("Q3 2027: Jahresplanung", "Freigabe Phase 2"),
        ("Q3 2028: Review", "Uebergang Steady State"),
    ]
    ms_y = ms_top + 0.28
    m_w = (CONTENT_W - 0.30) / 3
    for i, (t, sub) in enumerate(ms):
        x = CONTENT_LEFT + i * (m_w + 0.15)
        add_rect(s, x, ms_y, m_w, 0.75, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=NAVY, line_w=0.8)
        add_diamond(s, x + 0.10, ms_y + 0.22, 0.28, 0.28, RED_HL)
        add_text(s, x + 0.50, ms_y + 0.04, m_w - 0.60, 0.30, [
            (t, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.50, ms_y + 0.36, m_w - 0.60, 0.35, [
            (sub, dict(size=9, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.MIDDLE)
slide_budget()

# ============================================================
# SLIDE 16: Risiken
# ============================================================
def slide_risk():
    s = add_slide_stihl("Risiken der ServiceNow-Plattform-Abhaengigkeit")
    add_text(s, CONTENT_LEFT, CONTENT_TOP, CONTENT_W, 0.25, [
        ("Risiko-Matrix", dict(size=12.5, bold=True, color=ORANGE))])
    hdr_y = CONTENT_TOP + 0.30
    add_rect(s, CONTENT_LEFT, hdr_y, CONTENT_W, 0.28, DARK,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.15, hdr_y, 3.3, 0.28, [
        ("Risiko", dict(size=11, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, CONTENT_LEFT + 3.6, hdr_y, 1.5, 0.28, [
        ("Schwere", dict(size=11, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, CONTENT_LEFT + 5.3, hdr_y, 7.5, 0.28, [
        ("Mitigation", dict(size=11, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    for i, (risk, sev, mit) in enumerate(RISKS):
        y = hdr_y + 0.32 + i * 0.60
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        add_rect(s, CONTENT_LEFT, y, CONTENT_W, 0.56, bg,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, CONTENT_LEFT + 0.15, y, 3.3, 0.56, [
            (risk, dict(size=10.5, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        for k in range(5):
            filled = k < sev
            color = (RED_HL if sev >= 4 else (YELLOW if sev == 3 else GREEN)) if filled else LIGHT_GREY
            add_oval(s, CONTENT_LEFT + 3.60 + k * 0.28, y + 0.22, 0.18, 0.18, color)
        add_text(s, CONTENT_LEFT + 5.30, y, 7.5, 0.56, [
            (mit, dict(size=10, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    sy = hdr_y + 0.32 + len(RISKS) * 0.60 + 0.10
    sh_h = CONTENT_BOT - sy
    add_rect(s, CONTENT_LEFT, sy, CONTENT_W, sh_h, GREEN_LT,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, CONTENT_LEFT, sy, 0.07, sh_h, GREEN)
    add_text(s, CONTENT_LEFT + 0.15, sy + 0.04, CONTENT_W - 0.30, 0.30, [
        ("Fazit Risiko-Bewertung",
         dict(size=12, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, CONTENT_LEFT + 0.15, sy + 0.35, CONTENT_W - 0.30, sh_h - 0.40, [
        ("Kein Risiko ist ein Show-Stopper. Alle Risiken sind mit Standard-Vertraegen "
         "(Preisgarantie, DSGVO-ADV, SLA) und Standard-First-Governance beherrschbar.",
         dict(size=10, color=DARK))], anchor=MSO_ANCHOR.TOP)
slide_risk()

# ============================================================
# SLIDE 17: SNOW History
# ============================================================
def slide_snow_history():
    s = add_slide_stihl("ServiceNow bei STIHL - Vertragshistorie 2014-2026")
    add_text(s, CONTENT_LEFT, CONTENT_TOP, CONTENT_W, 0.25, [
        ("Vertragsstufen und jaehrliche Lizenzausgaben (indikativ)",
         dict(size=12, bold=True, color=ORANGE))])
    ct = CONTENT_TOP + 0.30
    ch = 3.0
    add_rect(s, CONTENT_LEFT, ct, CONTENT_W, ch, LIGHT_BG,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    def xp(yr): return CONTENT_LEFT + 0.4 + (yr - 2014) / (2026 - 2014) * (CONTENT_W - 0.8)
    add_rect(s, CONTENT_LEFT + 0.4, ct + ch - 0.4, CONTENT_W - 0.8, 0.02, DARK)
    max_val = max(x[3] for x in SNOW_HISTORY)
    for yr, event, note, val in SNOW_HISTORY:
        x = xp(yr)
        h = (val / max_val) * (ch - 0.9)
        y = ct + ch - 0.4 - h
        add_rect(s, x - 0.12, y, 0.24, h, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_oval(s, x - 0.10, y - 0.10, 0.20, 0.20, NAVY)
        add_text(s, x - 0.85, y - 0.40, 1.70, 0.22, [
            (f"{val} k€",
             dict(size=9, bold=True, color=NAVY, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x - 0.85, ct + ch - 0.32, 1.70, 0.24, [
            (str(yr), dict(size=10, bold=True, color=DARK, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    ev_top = ct + ch + 0.15
    add_text(s, CONTENT_LEFT, ev_top, CONTENT_W, 0.22, [
        ("Meilensteine der Zusammenarbeit",
         dict(size=11.5, bold=True, color=ORANGE))])
    ev_row_top = ev_top + 0.30
    card_w = (CONTENT_W - 0.30) / 3
    for i, (yr, event, note, val) in enumerate(SNOW_HISTORY):
        col = i % 3
        row = i // 3
        x = CONTENT_LEFT + col * (card_w + 0.15)
        y = ev_row_top + row * 0.60
        add_rect(s, x, y, card_w, 0.55, WHITE,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=ORANGE, line_w=0.5)
        add_rect(s, x, y, 0.06, 0.55, ORANGE)
        add_text(s, x + 0.15, y + 0.02, card_w - 0.20, 0.24, [
            (f"{yr} · {event}",
             dict(size=10, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.15, y + 0.26, card_w - 0.20, 0.28, [
            (note, dict(size=9, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)
slide_snow_history()

# ============================================================
# SLIDE 18: ROI & Empfehlung
# ============================================================
def slide_roi_recommendation():
    s = add_slide_stihl("ROI-Kalkulation & Empfehlung")
    kpis = [
        ("NPV 5 Jahre",   "~ +8,5 M€", "Base Case, WACC 6 %", GREEN),
        ("IRR",           "~ 32 %",    "ueber Investitionshuerde", GREEN),
        ("Payback",       "~ 3,5 J.",  "nach Kickoff", ORANGE),
        ("TCO 5 Jahre",   "~ 7,0 M€",  "einmalig + laufend",  NAVY),
    ]
    top = CONTENT_TOP
    kpi_h = 1.20
    gap = 0.12
    w = (CONTENT_W - 3 * gap) / 4
    for i, (label, big, sub, col) in enumerate(kpis):
        x = CONTENT_LEFT + i * (w + gap)
        add_rect(s, x, top, w, kpi_h, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=col, line_w=1.5)
        add_rect(s, x, top, w, 0.30, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x, top, w, 0.30, [
            (label, dict(size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x, top + 0.35, w, 0.50, [
            (big, dict(size=22, bold=True, color=col, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x, top + 0.90, w, 0.28, [
            (sub, dict(size=9, italic=True, color=GREY_TXT, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    tt_y = top + kpi_h + 0.20
    add_text(s, CONTENT_LEFT, tt_y, CONTENT_W, 0.25, [
        ("Kosten-Nutzen-Uebersicht (5 Jahre kumuliert, Base Case)",
         dict(size=12.5, bold=True, color=ORANGE))])
    rows = [
        ("Kosten (TCO)",                  "- 7,0 M€",  RED_HL),
        ("Konsolidierung Tools (5 J.)",   "+ 4,7 M€", GREEN),
        ("Konsolidierung Schnittstellen", "+ 1,1 M€", GREEN),
        ("Business Value",                "+ 11,6 M€",GREEN),
        ("Flankierende Prozesse",         "+ 5,3 M€", GREEN),
        ("Netto-Effekt 5 Jahre",          "+ 15,7 M€",DARK),
    ]
    tt = tt_y + 0.30
    for i, (name, val, color) in enumerate(rows):
        y = tt + i * 0.38
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        is_total = (i == len(rows) - 1)
        add_rect(s, CONTENT_LEFT, y, CONTENT_W, 0.34,
                 bg if not is_total else DARK,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, CONTENT_LEFT + 0.15, y + 0.02, 8.0, 0.30, [
            (name, dict(size=11, bold=is_total,
                        color=DARK if not is_total else WHITE))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, CONTENT_LEFT + 8.5, y + 0.02, CONTENT_W - 8.65, 0.30, [
            (val, dict(size=13 if is_total else 12, bold=True,
                       color=color if not is_total else ORANGE, align=PP_ALIGN.RIGHT))],
            anchor=MSO_ANCHOR.MIDDLE)
    rec_y = tt + len(rows) * 0.38 + 0.15
    rec_h = CONTENT_BOT - rec_y
    add_rect(s, CONTENT_LEFT, rec_y, CONTENT_W, rec_h, ORANGE,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.20, rec_y, 3.5, rec_h, [
        ("★ EMPFEHLUNG", dict(size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, CONTENT_LEFT + 3.8, rec_y + 0.10, CONTENT_W - 4.0, rec_h - 0.20, [
        ("Beschluss zur Umsetzung von Szenario A (Full Scope 24 Monate) und "
         "Freigabe Detail-Planung / RFP-Prozess bis Q3 2026 empfohlen. "
         "NPV +8,5 M€, IRR 32 %, Payback ~3,5 J. - deutlich ueber Investitions-Huerde.",
         dict(size=11, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
slide_roi_recommendation()

# ============================================================
# SLIDE 19: Vielen Dank
# ============================================================
def slide_thanks():
    s = prs.slides.add_slide(LAYOUT_CONCL)
    for ph in s.placeholders:
        ph.text_frame.text = "Vielen Dank"
        break
slide_thanks()

prs.save(DST_PPTX)
print(f"Saved: {DST_PPTX}  ({len(prs.slides)} slides)")
