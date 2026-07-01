"""Build Business Case SEPM PowerPoint (STIHL format, ~18 slides)
plus Business_Case_SEPM.xlsx template for data collection."""
from pptx import Presentation
from pptx.util import Inches, Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DST_PPTX = "/tmp/2026_07_01_Business_Case_SEPM.pptx"
DST_XLSX = "/tmp/2026_07_01_Business_Case_SEPM.xlsx"

# ------------ Palette ------------
ORANGE   = RGBColor(0xF0, 0x7F, 0x12)
ORANGE_LT= RGBColor(0xFD, 0xE6, 0xCC)
ORANGE2  = RGBColor(0xE9, 0x6C, 0x0C)
ORANGE3  = RGBColor(0xC9, 0x53, 0x00)
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

# Slide is 13.333" x 7.5" ; grid 160x90; U = 76200 EMU
U = 76200
def ex(u): return int(u * U)

prs = Presentation()
prs.slide_width  = 12192000
prs.slide_height = 6858000
blank_layout = prs.slide_layouts[6]

# ============================================================
# Shape helpers
# ============================================================
def new_slide():
    s = prs.slides.add_slide(blank_layout)
    for shp in list(s.shapes):
        sp = shp._element
        sp.getparent().remove(sp)
    return s

def add_rect(slide, left, top, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE, line_w=1.0):
    s = slide.shapes.add_shape(shape, ex(left), ex(top), ex(w), ex(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line; s.line.width = Pt(line_w)
    s.shadow.inherit = False
    return s

def add_oval(slide, left, top, w, h, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, ex(left), ex(top), ex(w), ex(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    s.line.fill.background()
    return s

def add_diamond(slide, left, top, w, h, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.DIAMOND, ex(left), ex(top), ex(w), ex(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    s.line.fill.background()
    return s

def set_paragraphs(shape, paragraphs, *, default_size=10, default_color=DARK,
                   default_align=PP_ALIGN.LEFT):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(30000); tf.margin_right = Emu(30000)
    tf.margin_top = Emu(12000); tf.margin_bottom = Emu(12000)
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

def add_text(slide, left, top, w, h, paragraphs, *, anchor=MSO_ANCHOR.MIDDLE, **kw):
    tb = slide.shapes.add_textbox(ex(left), ex(top), ex(w), ex(h))
    tb.fill.background(); tb.line.fill.background()
    tb.text_frame.vertical_anchor = anchor
    set_paragraphs(tb, paragraphs, **kw)
    return tb

def add_slide_header(slide, title, subtitle, slide_num=None, page_of=None):
    add_rect(slide, 0, 0, 160, 6, DARK)
    add_text(slide, 3, 0.4, 140, 2.6, [
        (title, dict(size=20, bold=True, color=WHITE))], anchor=MSO_ANCHOR.TOP)
    add_text(slide, 3, 3.2, 140, 2.4, [
        (subtitle, dict(size=12, italic=True, color=LIGHT_GREY))],
        anchor=MSO_ANCHOR.TOP)
    add_rect(slide, 150, 0, 10, 6, ORANGE)
    if slide_num is not None:
        num_txt = f"Folie {slide_num}" if page_of is None else f"{slide_num} / {page_of}"
        add_text(slide, 150, 0, 10, 6, [
            (num_txt, dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER))])

def add_slide_footer(slide, text):
    add_text(slide, 4, 86.3, 152, 2.0, [
        (text, dict(size=8, italic=True, color=GREY_TXT))])
    add_rect(slide, 0, 89.3, 160, 0.7, ORANGE)

def add_stihl_brand(slide):
    """Small STIHL wordmark upper right."""
    pass  # keeping header clean

# ============================================================
# ASSUMPTIONS (referenced in slides + Excel)
# ============================================================
LICENSE_Y1 = 720   # k€/y year 1 (illustrative, TBD)
LICENSE_Y_STEADY = 850  # k€/y after scale-up
IMPL_ONE = 1300   # k€ one-off implementation (Beratung + SI)
INT_FTE_IMPL = 4  # FTE internal during implementation
CHANGE_ONE = 250  # k€ one-off change/training
SUPPORT_INT_FTE_STEADY = 3.0  # FTE p.a.
SUPPORT_EXT = 180  # k€ p.a. external support contracts
FTE_RATE = 130    # k€ per FTE fully loaded

# Consolidation - tool ablöse (illustrative)
CONSOL_TOOLS = [
    ("PIT / PLANTA",         "Projekt-/Ressourcen-Mgmt",  380),
    ("MS Project Online",    "Klass. Projekt-Mgmt",       120),
    ("Workpath",             "Strategie / OKR",           150),
    ("Power Apps",           "Custom-Formulare",           90),
    ("SharePoint-Listen",    "Demand-Listen",              70),
    ("Lucom (EV, teilweise)","EV-Prozess",                140),
]
TOTAL_TOOL_SAVINGS = sum(x[2] for x in CONSOL_TOOLS)  # 950 k€/y

# Interface / integration savings
INT_SAVINGS = 220  # k€/y (schnittstellen, custom-APIs, betrieb)

# Business Value benefits categories (mid-range yearly steady-state)
VALUE_CATEGORIES = [
    ("Effizienz PM/PPM",       "Weniger manuelle Excels/Berichte, Ressourcen-Steuerung", 820),
    ("Time-to-Market",         "Schnellere Idee→Projekt-Start (Gates in System)",         480),
    ("Governance / Rework",    "Weniger Doppelarbeit, klare Priorisierung",               360),
    ("Ressourcen-Optimierung", "Skill-Match, Auslastungs-Steuerung",                      520),
    ("Compliance / Audit",     "Nachvollziehbarkeit, Business Cases im System",           140),
]
TOTAL_VALUE = sum(x[2] for x in VALUE_CATEGORIES)  # 2320 k€/y

# Flanking processes (SPM adjacent) — quantified savings
FLANKING = [
    ("Maßnahmen-Management im Zielkostenmanagement der Produktentwicklung",
     "Zentraler Maßnahmen-Backlog, Verknüpfung mit Portfolios/Cost-Targets", 380),
    ("Innovations- / Vorhaben-Trichter (LH/CRD → Steckbrief → Projekt)",
     "Durchgängige Pipeline; weniger Reibungsverluste",                     220),
    ("Application Portfolio Mgmt (APM)",
     "Lifecycle-Management der Anwendungslandschaft",                       160),
    ("Benefits Realization Tracking",
     "Ist-/Soll-Nutzenabgleich nach Projektabschluss",                      120),
    ("Compliance / EV-Prozess-Digitalisierung",
     "EV-Workflow in SNOW; weniger manuelle Freigabeschleifen",             180),
]
TOTAL_FLANKING = sum(x[2] for x in FLANKING)  # 1060 k€/y

# Scenario overview (24 months full vs 36 months extended)
SCENARIOS = [
    ("A – Full Scope 24 Monate",     24, "Standard",   "Full",      1300, 850,  100),
    ("B – Extended Ramp 36 Monate",  36, "Reduziert",  "Full",      1050, 800,   80),
    ("C – MVP-First 12 Monate + Ausbau", 30, "MVP",    "Sukzessiv",  650, 500,   60),
]

# Contract history SNOW 2014 to today - illustrative
SNOW_HISTORY = [
    (2014, "Erstvertrag ITSM",       "Standard Package, 500 Fulfiller", 180),
    (2016, "Erweiterung Global",     "Ausrollung inkl. China",           260),
    (2018, "Enterprise-Vertrag",     "Volumen-Diskont, HR-SD add-on",    340),
    (2020, "ITOM + CMDB",            "Discovery, Event Mgmt",            410),
    (2022, "Contract Renewal 3J",    "Preisanpassung +5% p.a.",          490),
    (2024, "Now Assist Preview",     "GenAI eingeführt",                 560),
    (2026, "SPM Pro (geplant)",      "Neuer Modul-Baustein",             820),
]

# Risk items with severity 1-5 and mitigation
RISKS = [
    ("Vendor Lock-in", 4, "Standard-Konfiguration, Migrations-Klauseln, jährl. Marktscreening"),
    ("Preis-Eskalation", 3, "Mehrjahresvertrag mit Preisgarantie, Volumen-Rabatte"),
    ("Ausfall/Betrieb", 2, "SLA 99,9% + regionaler DC (Frankfurt/Düsseldorf), DSGVO"),
    ("Roadmap-Divergenz", 3, "Standard first, keine tiefen Anpassungen; halbjährl. Review"),
    ("Datenhoheit / DSGVO", 2, "DE-Hosting, ADV-Vertrag, verschlüsselte Datenhaltung"),
    ("Team-Kompetenzaufbau", 3, "Ausbau ITSM-Team +1,5-2 FTE, Ausbildungsplan, Partner-Support"),
]

# ============================================================
# SLIDE 1: COVER
# ============================================================
def slide_cover():
    s = new_slide()
    add_rect(s, 0, 0, 160, 90, DARK)
    add_rect(s, 0, 0, 6, 90, ORANGE)
    add_rect(s, 154, 0, 6, 90, ORANGE)
    # Big STIHL brand
    add_text(s, 20, 15, 120, 6, [
        ("STIHL", dict(size=28, bold=True, color=ORANGE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    # Main title
    add_text(s, 15, 32, 130, 20, [
        ("Business Case SEPM",
         dict(size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 15, 46, 130, 8, [
        ("Detaillierte Wirtschaftlichkeits-Betrachtung",
         dict(size=20, italic=True, color=ORANGE_LT, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 15, 55, 130, 6, [
        ("TCO · ROI · Business Value · Szenarien · Risiken",
         dict(size=14, color=LIGHT_GREY, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, 55, 65, 50, 0.4, ORANGE)
    add_text(s, 15, 68, 130, 5, [
        ("Stand: 01.07.2026",
         dict(size=13, color=LIGHT_GREY, align=PP_ALIGN.CENTER))])
    add_text(s, 15, 74, 130, 4, [
        ("Autoren: Alex Passaro · Torsten Zahn",
         dict(size=11, italic=True, color=LIGHT_GREY, align=PP_ALIGN.CENTER))])

slide_cover()

# ============================================================
# SLIDE 2: MANAGEMENT SUMMARY
# ============================================================
def slide_management_summary():
    s = new_slide()
    add_slide_header(s, "Management Summary",
                     "Kernaussagen des Business Case auf einen Blick",
                     slide_num=2)
    # Big KPI tiles (4)
    kpis = [
        ("5-Jahres-TCO Vorteil",   "~ 2,2 M€",  "vs Best-of-Breed", GREEN),
        ("Payback",                "~ 3,5 J.",  "nach Go-Live Phase 1", ORANGE),
        ("Jährl. Nutzen (steady)", "~ 4,3 M€",  "Konsolidierung + Value + Flanking", NAVY),
        ("Empfohlenes Szenario",   "A · 24 Mo", "Full Scope Phase 1+2", BLUE_HD),
    ]
    top = 8
    gap = 1.5
    w = (152 - 3 * gap) / 4
    for i, (label, big, sub, col) in enumerate(kpis):
        x = 4 + i * (w + gap)
        add_rect(s, x, top, w, 20, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=col, line_w=1.5)
        add_rect(s, x, top, w, 4, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x + 0.5, top + 0.4, w - 1, 3.2, [
            (label, dict(size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.5, top + 6, w - 1, 8, [
            (big, dict(size=28, bold=True, color=col, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.5, top + 15, w - 1, 4, [
            (sub, dict(size=10, italic=True, color=GREY_TXT, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    # Key messages block
    km_top = 32
    add_text(s, 4, km_top, 152, 2.4, [
        ("Kernbotschaften", dict(size=14, bold=True, color=DARK))])
    msgs = [
        ("Konsolidierung von >6 PPM-Tools", "spart ~950 k€ / Jahr an Lizenz + Wartung + FTE"),
        ("Business Value SPM (steady state)", "~2,3 M€ / Jahr durch Effizienz, Time-to-Market, Governance"),
        ("Flankierende Prozesse (Maßnahmen-Mgmt etc.)", "zusätzlich ~1,0 M€ / Jahr Einsparpotenzial"),
        ("SNOW-Plattform-Risiko", "Vendor Lock-in adressiert durch Standard-first und Marktscreening"),
        ("Vertragshistorie 2014–heute", "SNOW als bewährter STIHL-Partner (ITSM seit 12 J.)"),
    ]
    mb_top = km_top + 3.0
    row_h = 5.0
    for i, (head, body) in enumerate(msgs):
        y = mb_top + i * row_h
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        add_rect(s, 4, y, 152, row_h - 0.4, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, 4, y, 0.7, row_h - 0.4, ORANGE)
        add_text(s, 6, y + 0.2, 60, row_h - 0.6, [
            (head, dict(size=11.5, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 68, y + 0.2, 87, row_h - 0.6, [
            (body, dict(size=10.5, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)
    add_slide_footer(s,
        "Alle Zahlen illustrativ.  Endgültige Erhebung erfolgt in der Vorlage Business_Case_SEPM.xlsx.")

slide_management_summary()

# ============================================================
# SLIDE 3: BUSINESS CASE METHODIK
# ============================================================
def slide_methodik():
    s = new_slide()
    add_slide_header(s, "Business Case Methodik",
                     "Vier Dimensionen der Wirtschaftlichkeits-Betrachtung",
                     slide_num=3)
    # Big 4-quadrant framework
    top = 9
    h = 32
    gap = 2
    w = (152 - gap) / 2
    quads = [
        ("TCO (5 Jahre)",
         "Lizenzen · Implementierung · Support · Change · Betrieb",
         ["Lizenzen SNOW SPM Pro", "Implementierung (Berater + interne FTE)",
          "Change & Schulung", "Support / Wartung (intern + extern)"],
         ORANGE, "€"),
        ("Nutzen / Business Value",
         "Effizienz · Time-to-Market · Governance · Ressourcen",
         ["Effizienz PPM (weniger Excels)", "Time-to-Market Ideen → Projekte",
          "Ressourcen-Steuerung", "Governance & Compliance"],
         GREEN, "★"),
        ("Konsolidierung",
         "Tool-Ablöse · Schnittstellen · flankierende Prozesse",
         ["PIT / MSPO / Workpath / PowerApps", "Schnittstellen (APIs, ETL)",
          "Maßnahmen-Mgmt Produktentw.", "Application Portfolio Mgmt"],
         NAVY, "⇨"),
        ("Risiko / Sensitivität",
         "Vendor · Preis · Betrieb · Roadmap · DSGVO",
         ["Vendor Lock-in", "Preis-Eskalation",
          "Ausfall / Betriebs-SLA", "Best/Base/Worst Sensitivity"],
         RED_HL, "⚠"),
    ]
    for i, (head, sub, items, color, icon) in enumerate(quads):
        row = i // 2
        col = i % 2
        x = 4 + col * (w + gap)
        y = top + row * (h + gap)
        add_rect(s, x, y, w, h, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=color, line_w=1.5)
        # Header band
        add_rect(s, x, y, w, 5, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        # Big icon
        add_oval(s, x + 1, y + 0.8, 3.4, 3.4, WHITE)
        add_text(s, x + 1, y + 0.8, 3.4, 3.4, [
            (icon, dict(size=15, bold=True, color=color, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 5, y + 0.4, w - 6, 4.5, [
            (head, dict(size=13, bold=True, color=WHITE))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 1.5, y + 6, w - 3, 3.2, [
            (sub, dict(size=10, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)
        # Items
        for j, it in enumerate(items):
            iy = y + 10 + j * 4.5
            add_oval(s, x + 2, iy + 1.4, 1.0, 1.0, color)
            add_text(s, x + 3.6, iy, w - 5, 4.0, [
                (it, dict(size=10.5, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    add_slide_footer(s,
        "Framework orientiert an Gartner Business Case Toolkit + STIHL Investitionsrichtlinie.")

slide_methodik()

# ============================================================
# SLIDE 4: ALTERNATIVE PM-PLATTFORM (Punkt 3 Folie 4)
# ============================================================
def slide_alternative_pm():
    s = new_slide()
    add_slide_header(s, "Alternative PM-Plattform – Bewertung",
                     "Punkt 3 Folie 4: PM-Plattform VEW / VPM als Absicherung (Hedge)",
                     slide_num=4)
    # Left: rationale
    add_rect(s, 4, 8, 74, 40, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, 4, 8, 74, 4, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 5, 8, 72, 4, [
        ("Warum Alternative bewerten?",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    rat = [
        ("Risiko-Absicherung", "Falls SNOW SPM für VEW/VPM nicht ausreichend tief."),
        ("Verhandlungsposition", "Best-of-Breed-Angebot senkt SNOW-Verhandlungspreis."),
        ("Engineering-Spezifika", "PLM-nahe Prozesse ggf. besser in Fach-Tool."),
        ("Vorstand EWW-Bedenken", "Explizite Alternative senkt Widerstand."),
    ]
    for i, (h, b) in enumerate(rat):
        y = 13.5 + i * 8.2
        add_oval(s, 6, y + 1.5, 2.4, 2.4, NAVY)
        add_text(s, 6, y + 1.5, 2.4, 2.4, [
            (str(i+1), dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 9.5, y + 0.3, 66, 3.2, [
            (h, dict(size=11, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 9.5, y + 3.5, 66, 3.6, [
            (b, dict(size=9.5, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)

    # Right: shortlist of alternatives + criteria
    add_rect(s, 82, 8, 74, 40, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, 82, 8, 74, 4, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 83, 8, 72, 4, [
        ("Shortlist Alternative PM-Plattformen",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    alts = [
        ("Planisware Enterprise", "Marktführer Engineering-PPM", "Lead Jonas"),
        ("SAP Enterprise Portfolio & Project Mgmt (EPPM)", "SAP-Ökosystem-Integration", "Lead Jonas"),
        ("Sciforma", "Mittelstands-Standard, hybrid", "Lead Jonas"),
    ]
    for i, (n, sub, own) in enumerate(alts):
        y = 13.5 + i * 6.5
        add_rect(s, 83.5, y, 71, 5.5, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=NAVY, line_w=0.5)
        add_rect(s, 83.5, y, 0.6, 5.5, NAVY)
        add_text(s, 85, y + 0.2, 45, 2.4, [
            (n, dict(size=11, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 85, y + 2.5, 45, 3.0, [
            (sub, dict(size=9, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.TOP)
        add_rect(s, 132, y + 1.0, 20, 3.5, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, 132, y + 1.0, 20, 3.5, [
            (own, dict(size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 83, 34, 72, 3.4, [
        ("Zeitraum: 11.06. – 30.08.2026  ·  Ergebnis: klare Empfehlung mit TCO-Vergleich",
         dict(size=10, italic=True, color=GREY_TXT))],
        anchor=MSO_ANCHOR.MIDDLE)

    # Bottom: comparison teaser
    add_rect(s, 4, 51, 152, 30, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
             line=LIGHT_GREY, line_w=0.7)
    add_rect(s, 4, 51, 152, 4, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 5, 51, 150, 4, [
        ("Bewertungskriterien und Zwischenergebnis",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    criteria = [
        ("Funktions-Deckung", "70%", "80%", "60%", "65%"),
        ("Integration STIHL-Landschaft", "80%", "50%", "70%", "40%"),
        ("TCO 5 Jahre (rel.)", "100%", "115%", "125%", "95%"),
        ("Change-Aufwand", "Mittel", "Hoch", "Mittel", "Niedrig"),
        ("Zukunftssicherheit AI", "Hoch", "Mittel", "Hoch", "Niedrig"),
    ]
    col_x = [5, 66, 89, 112, 135]
    col_w = [60, 22, 22, 22, 20]
    headers = ["Kriterium", "SNOW SPM", "Planisware", "SAP EPPM", "Sciforma"]
    hy = 55.6
    for i, hd in enumerate(headers):
        col = ORANGE if i == 1 else NAVY
        if i == 0:
            add_text(s, col_x[i], hy, col_w[i], 3.2, [
                (hd, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        else:
            add_rect(s, col_x[i], hy, col_w[i], 3.2, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            add_text(s, col_x[i], hy, col_w[i], 3.2, [
                (hd, dict(size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
                anchor=MSO_ANCHOR.MIDDLE)
    for j, (k, *vals) in enumerate(criteria):
        y = 59 + j * 4.0
        bg = LIGHT_BG if j % 2 == 0 else ROW_ALT
        add_rect(s, 5, y, 150, 3.6, bg)
        add_text(s, 6, y, col_w[0], 3.6, [
            (k, dict(size=10, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        for i, v in enumerate(vals):
            add_text(s, col_x[i+1], y, col_w[i+1], 3.6, [
                (v, dict(size=10, color=DARK, align=PP_ALIGN.CENTER))],
                anchor=MSO_ANCHOR.MIDDLE)
    add_slide_footer(s,
        "Detaillierte Bewertung erfolgt bis 30.08.2026 (siehe Timeline Folie 5, Vorgang 3).")

slide_alternative_pm()

# ============================================================
# SLIDE 5: KOMPLEXITÄTS-REDUZIERUNG PROJEKTCONTROLLING (Punkt 4)
# ============================================================
def slide_komplex_reduzierung():
    s = new_slide()
    add_slide_header(s, "Komplexitäts-Reduzierung ProjektControlling",
                     "Punkt 4 Folie 4: Was heute Aufwand verursacht – und was SEPM konsolidiert",
                     slide_num=5)
    # Left column: HEUTE
    add_rect(s, 4, 8, 74, 74, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, 4, 8, 74, 4, RED_HL, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 5, 8, 72, 4, [
        ("HEUTE  –  Komplexitätstreiber",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    hoy = [
        ("Multi-Tool-Landschaft PPM",
         "PIT · MSPO · Excel · SharePoint · Power Apps · Workpath",
         "~ 2,5 FTE Betriebs- & Integrations-Aufwand"),
        ("Manuelle Daten-Aggregation",
         "Excel-Konsolidierung für Reportings, Steckbriefe",
         "~ 2 Wochen/Monat Reporting-Aufwand"),
        ("Uneinheitliche Datenmodelle",
         "Jede Insel eigener Struktur – Vergleichbarkeit begrenzt",
         "~ 1 FTE Datenqualitäts-Arbeit"),
        ("Doppelte Datenpflege",
         "Projekte in PIT + Excel + SharePoint erfasst",
         "~ 0,8 FTE Doppelerfassung"),
        ("Fehlende Traceability",
         "Idee ↔ Portfolio ↔ Projekt ↔ Kosten nicht durchgängig",
         "Governance-Aufwand hoch, Audits schwierig"),
    ]
    for i, (h, sub, imp) in enumerate(hoy):
        y = 13.5 + i * 12
        add_rect(s, 5.5, y, 71, 11, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=RED_HL, line_w=0.6)
        add_rect(s, 5.5, y, 0.6, 11, RED_HL)
        add_text(s, 7, y + 0.4, 68, 3.2, [
            (h, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 7, y + 3.4, 68, 3.4, [
            (sub, dict(size=9, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 7, y + 7.2, 68, 3.4, [
            (imp, dict(size=9.5, bold=True, color=RED_HL))],
            anchor=MSO_ANCHOR.MIDDLE)
    # Right column: MIT SEPM
    add_rect(s, 82, 8, 74, 74, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, 82, 8, 74, 4, GREEN, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 83, 8, 72, 4, [
        ("MIT SEPM  –  Konsolidierungs-Hebel",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    mit = [
        ("Eine Plattform, ein Datenmodell",
         "SNOW SPM als Single Source of Truth",
         "− 6 Tools · − 0,5 FTE Betrieb"),
        ("Automatische Reports & Dashboards",
         "Live-Dashboards ersetzen manuelles Excel",
         "− 1,5 Wochen/Monat Reporting"),
        ("Standard-Datenmodell",
         "Vorhaben-Steckbrief einheitlich, konsistente KPIs",
         "− 0,7 FTE Datenqualitäts-Arbeit"),
        ("Erfassung einmal – Nutzung überall",
         "Demand → Portfolio → Projekt in einem System",
         "− 0,6 FTE Doppelerfassung"),
        ("Durchgängige Traceability",
         "OKR → Portfolio → Projekt → Kosten → Nutzen",
         "Audit-fähig, Compliance-sicher"),
    ]
    for i, (h, sub, imp) in enumerate(mit):
        y = 13.5 + i * 12
        add_rect(s, 83.5, y, 71, 11, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=GREEN, line_w=0.6)
        add_rect(s, 83.5, y, 0.6, 11, GREEN)
        add_text(s, 85, y + 0.4, 68, 3.2, [
            (h, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 85, y + 3.4, 68, 3.4, [
            (sub, dict(size=9, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 85, y + 7.2, 68, 3.4, [
            (imp, dict(size=9.5, bold=True, color=GREEN))],
            anchor=MSO_ANCHOR.MIDDLE)
    add_slide_footer(s,
        "Gesamt-Effekt Komplexitäts-Reduzierung: ~ 2,3 FTE Einsparung + Reporting-Zeit halbiert.")

slide_komplex_reduzierung()

# ============================================================
# SLIDE 6: BUSINESS VALUE ANALYSE
# ============================================================
def slide_business_value():
    s = new_slide()
    add_slide_header(s, "Business Value Analyse",
                     "Nutzen-Kategorien und quantifizierter Steady-State p.a.",
                     slide_num=6)
    # KPI at top
    add_rect(s, 4, 8, 152, 8, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 6, 8, 60, 8, [
        ("Gesamt-Nutzen SEPM (Steady State p.a.)",
         dict(size=13, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 66, 8, 60, 8, [
        (f"~ {TOTAL_VALUE:,} k€ / Jahr".replace(",", "."),
         dict(size=22, bold=True, color=ORANGE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 126, 8, 30, 8, [
        (f"davon quantifiziert  100 %",
         dict(size=10, italic=True, color=LIGHT_GREY, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)

    # Value categories table with bar visualization
    add_text(s, 4, 18, 152, 2.4, [
        ("Detail – Nutzen-Kategorien",
         dict(size=13, bold=True, color=DARK))])
    # Table
    max_val = max(v[2] for v in VALUE_CATEGORIES)
    row_top = 22
    row_h = 10
    for i, (name, desc, val) in enumerate(VALUE_CATEGORIES):
        y = row_top + i * (row_h + 0.5)
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        add_rect(s, 4, y, 152, row_h, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, 4, y, 0.7, row_h, ORANGE)
        add_text(s, 6, y + 0.4, 50, row_h - 0.8, [
            (name, dict(size=11, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 6, y + row_h - 3.5, 50, 3.0, [
            (desc, dict(size=9, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.TOP)
        # Value bar
        bar_x = 58
        bar_w_max = 78
        bar_w_val = bar_w_max * (val / max_val)
        add_rect(s, bar_x, y + 3.0, bar_w_max, 4, LIGHT_GREY,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, bar_x, y + 3.0, bar_w_val, 4, ORANGE,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        # Value label
        add_text(s, 138, y, 18, row_h, [
            (f"{val:,} k€".replace(",", "."),
             dict(size=13, bold=True, color=NAVY, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    add_slide_footer(s,
        "Alle Werte Steady State p.a. nach Erreichung Phase 2 (~ ab Jahr 3 nach Kickoff).")

slide_business_value()

# ============================================================
# SLIDE 7: TCO DETAIL
# ============================================================
def slide_tco():
    s = new_slide()
    add_slide_header(s, "TCO – Total Cost of Ownership (5 Jahre)",
                     "Gesamt-Kosten SEPM: einmalig + laufend, aufgeschlüsselt",
                     slide_num=7)
    # Big number
    total_5y = IMPL_ONE + CHANGE_ONE + 5 * (LICENSE_Y_STEADY + SUPPORT_INT_FTE_STEADY * FTE_RATE + SUPPORT_EXT)
    add_rect(s, 4, 8, 152, 8, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 6, 8, 70, 8, [
        ("5-Jahres-TCO SEPM (Base Case)",
         dict(size=13, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 76, 8, 60, 8, [
        (f"~ {total_5y/1000:.1f} Mio. €",
         dict(size=24, bold=True, color=ORANGE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 136, 8, 20, 8, [
        (f"(~ {total_5y/5/1000:.2f} M€ p.a. Ø)",
         dict(size=10, italic=True, color=LIGHT_GREY, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)

    # Breakdown two columns: Einmalig | Laufend p.a.
    left_x = 4; right_x = 82; w = 74
    top = 18
    add_rect(s, left_x, top, w, 60, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, left_x, top, w, 4, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, left_x + 1, top, w - 2, 4, [
        ("Einmalige Kosten (Year 1)",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    einmalig = [
        ("Implementierungs-Partner (SI)", IMPL_ONE, "SNOW-Zertifiziert"),
        ("Interne FTE (Phase 1+2)", INT_FTE_IMPL * FTE_RATE * 2, f"{INT_FTE_IMPL} FTE × 2 J. Ø"),
        ("Change / Schulung", CHANGE_ONE, "Anwender + Multiplikatoren"),
        ("Contingency (10 %)", int(0.1 * (IMPL_ONE + INT_FTE_IMPL * FTE_RATE * 2 + CHANGE_ONE)), "Puffer"),
    ]
    for i, (name, val, note) in enumerate(einmalig):
        y = top + 6 + i * 11
        add_rect(s, left_x + 1.5, y, w - 3, 10, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=NAVY, line_w=0.5)
        add_text(s, left_x + 2.5, y + 0.4, 45, 3.6, [
            (name, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, left_x + 2.5, y + 4.0, 45, 5.0, [
            (note, dict(size=8.5, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.TOP)
        add_text(s, left_x + w - 20, y, 18, 10, [
            (f"{val:,} k€".replace(",", "."),
             dict(size=13, bold=True, color=NAVY, align=PP_ALIGN.RIGHT))],
            anchor=MSO_ANCHOR.MIDDLE)
    total_einmal = sum(v[1] for v in einmalig)
    add_text(s, left_x + 1, top + 54, w - 2, 3.5, [
        (f"Summe einmalig: {total_einmal:,} k€".replace(",", "."),
         dict(size=12, bold=True, color=NAVY, align=PP_ALIGN.RIGHT))],
        anchor=MSO_ANCHOR.MIDDLE)

    # Laufend
    add_rect(s, right_x, top, w, 60, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, right_x, top, w, 4, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, right_x + 1, top, w - 2, 4, [
        ("Laufende Kosten p.a. (Steady State)",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    laufend = [
        ("Lizenzen SNOW SPM Pro", LICENSE_Y_STEADY, "~ 350-450 Named User"),
        ("Interne Betriebs-FTE", int(SUPPORT_INT_FTE_STEADY * FTE_RATE),
         f"~ {SUPPORT_INT_FTE_STEADY} FTE (IT + Business)"),
        ("Externer Support / Wartung", SUPPORT_EXT, "Retainer + Ticket-Volumen"),
        ("Weiterentwicklung (Backlog)", 220, "kontinuierliche Feature-Erweiterung"),
    ]
    for i, (name, val, note) in enumerate(laufend):
        y = top + 6 + i * 11
        add_rect(s, right_x + 1.5, y, w - 3, 10, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=ORANGE, line_w=0.5)
        add_text(s, right_x + 2.5, y + 0.4, 45, 3.6, [
            (name, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, right_x + 2.5, y + 4.0, 45, 5.0, [
            (note, dict(size=8.5, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.TOP)
        add_text(s, right_x + w - 20, y, 18, 10, [
            (f"{val:,} k€".replace(",", "."),
             dict(size=13, bold=True, color=ORANGE, align=PP_ALIGN.RIGHT))],
            anchor=MSO_ANCHOR.MIDDLE)
    total_laufend = sum(v[1] for v in laufend)
    add_text(s, right_x + 1, top + 54, w - 2, 3.5, [
        (f"Summe p.a.: {total_laufend:,} k€".replace(",", "."),
         dict(size=12, bold=True, color=ORANGE, align=PP_ALIGN.RIGHT))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_slide_footer(s,
        "Alle Werte Base Case. Sensitivity ±20 % in Szenario-Analyse (Folie 15). Detail siehe Excel-Vorlage Tab 'TCO_Detail'.")

slide_tco()

# ============================================================
# SLIDE 8: SUPPORT-KOSTEN DETAIL
# ============================================================
def slide_support():
    s = new_slide()
    add_slide_header(s, "Support-Kosten – Detailbetrachtung",
                     "Interner + externer Support, laufender Betrieb, Weiterentwicklung",
                     slide_num=8)
    # Three columns: Intern, Extern, Weiterentwicklung
    cols = [
        ("Interner Betrieb", ORANGE,
         [("SNOW-Plattform-Admin", "1,0 FTE", 130),
          ("SPM Business Analyst", "1,0 FTE", 130),
          ("Integration / Schnittstellen", "0,5 FTE",  65),
          ("Data Governance",     "0,3 FTE",  39),
          ("User Support (2nd Level)", "0,2 FTE", 26)]),
        ("Externer Support", NAVY,
         [("Wartungsvertrag SNOW",  "Standard 22%",  70),
          ("Managed-Service Partner","Retainer + T&M", 90),
          ("Externes Consulting (Retainer)", "20-30 Tage/y", 60),
          ("Zertifikate & Schulungen", "5-8 Personen",   30)]),
        ("Weiterentwicklung", GREEN,
         [("Feature-Backlog (Business)", "~10 PT/Mo",  100),
          ("Release-Upgrade-Anpassungen","2× / Jahr",   50),
          ("Custom-Reports / Dashboards","4-6 PT/Mo",   40),
          ("Innovation Sprints",         "1× Q",        30)]),
    ]
    top = 8
    w = (152 - 4) / 3
    gap = 2
    for i, (title, color, items) in enumerate(cols):
        x = 4 + i * (w + gap)
        add_rect(s, x, top, w, 60, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, x, top, w, 5, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x + 1, top + 0.4, w - 2, 4.4, [
            (title, dict(size=13, bold=True, color=WHITE))],
            anchor=MSO_ANCHOR.MIDDLE)
        total = 0
        for j, (name, note, val) in enumerate(items):
            y = top + 6 + j * 8.5
            add_rect(s, x + 1, y, w - 2, 8, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                     line=color, line_w=0.5)
            add_text(s, x + 2, y + 0.3, w - 20, 3.4, [
                (name, dict(size=10.5, bold=True, color=DARK))],
                anchor=MSO_ANCHOR.MIDDLE)
            add_text(s, x + 2, y + 3.4, w - 20, 4.2, [
                (note, dict(size=9, italic=True, color=GREY_TXT))],
                anchor=MSO_ANCHOR.MIDDLE)
            add_text(s, x + w - 18, y + 0.3, 16, 7.4, [
                (f"{val} k€".replace(",", "."),
                 dict(size=12, bold=True, color=color, align=PP_ALIGN.RIGHT))],
                anchor=MSO_ANCHOR.MIDDLE)
            total += val
        # Column total
        add_rect(s, x + 1, top + 54, w - 2, 4, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x + 1, top + 54, w - 2, 4, [
            (f"Summe: {total} k€ / Jahr".replace(",", "."),
             dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    # Overall total bar
    add_rect(s, 4, 71, 152, 8, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 6, 71, 90, 8, [
        ("Gesamt Support & Weiterentwicklung",
         dict(size=13, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    total_support = 390 + 250 + 220  # from data above roughly
    add_text(s, 96, 71, 60, 8, [
        ("~ 860 k€ / Jahr",
         dict(size=22, bold=True, color=ORANGE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_slide_footer(s,
        "Zahlen indikativ, endgültige Erhebung in Excel-Vorlage Tab 'Support_Kosten'. FTE-Satz Ø 130 k€ p.a.")

slide_support()

# ============================================================
# SLIDE 9: KONSOLIDIERUNG - TOOL-ABLÖSE
# ============================================================
def slide_konsolidierung_tools():
    s = new_slide()
    add_slide_header(s, "Konsolidierungseinsparungen – Tool-Ablöse",
                     "Wegfall bisheriger Tools durch SNOW SPM Konsolidierung",
                     slide_num=9)
    # KPI banner
    add_rect(s, 4, 8, 152, 8, GREEN, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 6, 8, 90, 8, [
        ("Gesamt Tool-Ablöse-Ersparnis (Steady State p.a.)",
         dict(size=13, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 96, 8, 60, 8, [
        (f"~ {TOTAL_TOOL_SAVINGS} k€ / Jahr".replace(",", "."),
         dict(size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    # Table of tools
    add_text(s, 4, 18, 152, 2.4, [
        ("Detail – Tool-für-Tool-Ablöse",
         dict(size=13, bold=True, color=DARK))])
    max_val = max(x[2] for x in CONSOL_TOOLS)
    row_top = 22
    row_h = 8.5
    for i, (tool, cat, saving) in enumerate(CONSOL_TOOLS):
        y = row_top + i * (row_h + 0.4)
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        add_rect(s, 4, y, 152, row_h, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, 4, y, 0.7, row_h, GREEN)
        # Cross-out styled tool name
        add_text(s, 6, y + 0.3, 40, row_h - 0.6, [
            (tool, dict(size=11.5, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 6, y + row_h - 3.0, 40, 2.5, [
            (cat, dict(size=8.5, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.TOP)
        # Bar
        bar_x = 48
        bar_w_max = 88
        bar_w_val = bar_w_max * (saving / max_val)
        add_rect(s, bar_x, y + 3, bar_w_max, 3, LIGHT_GREY,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, bar_x, y + 3, bar_w_val, 3, GREEN,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        # Value
        add_text(s, 138, y, 18, row_h, [
            (f"− {saving} k€",
             dict(size=13, bold=True, color=GREEN, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 4, 78, 152, 3.4, [
        (f"Basis-Annahme: Ablöse durch SEPM ab Jahr 3 vollständig realisierbar. Übergangsphase 12-18 Monate.",
         dict(size=10, italic=True, color=GREY_TXT))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_slide_footer(s,
        "Datenerhebung Tab 'Konsolidierung_Tools' in Business_Case_SEPM.xlsx.")

slide_konsolidierung_tools()

# ============================================================
# SLIDE 10: KONSOLIDIERUNG - SCHNITTSTELLEN
# ============================================================
def slide_konsolidierung_interfaces():
    s = new_slide()
    add_slide_header(s, "Konsolidierungseinsparungen – Schnittstellen",
                     "Wegfall komplexer Integrationen durch eine Plattform",
                     slide_num=10)
    # KPI
    add_rect(s, 4, 8, 152, 8, GREEN, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 6, 8, 90, 8, [
        ("Ersparnis Schnittstellen & Integration p.a.",
         dict(size=13, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 96, 8, 60, 8, [
        (f"~ {INT_SAVINGS} k€ / Jahr",
         dict(size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    # Left: heute (viele APIs)
    add_rect(s, 4, 18, 74, 62, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, 4, 18, 74, 4, RED_HL, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 5, 18, 72, 4, [
        ("HEUTE – Fragmentierte Integration",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    heute_ints = [
        "PIT ↔ SAP FI (Kostentransfer)",
        "MSPO ↔ SAP HCM (Personalstamm)",
        "SharePoint ↔ Excel (Datenexport)",
        "Workpath ↔ Excel (OKR-Reporting)",
        "PowerApps ↔ Dataverse (Custom-Data)",
        "Lucom EV ↔ SAP (Freigabestatus)",
        "Diverse Excel-Konsolidierungen",
    ]
    for i, txt in enumerate(heute_ints):
        y = 24 + i * 6.8
        add_rect(s, 6, y, 70, 5.6, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=RED_HL, line_w=0.5)
        add_rect(s, 6, y, 0.5, 5.6, RED_HL)
        add_text(s, 8, y, 68, 5.6, [
            (txt, dict(size=10, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    # Right: mit SEPM (native)
    add_rect(s, 82, 18, 74, 62, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, 82, 18, 74, 4, GREEN, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 83, 18, 72, 4, [
        ("MIT SEPM – Native SNOW-Integration",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    mit_ints = [
        ("SAP FI Konnektor", "Standard, im Integration Hub"),
        ("SAP HCM Konnektor", "Standard, im Integration Hub"),
        ("Interne Datenmodelle", "keine Excel-Export-Ketten mehr"),
        ("OKR-Modul nativ", "kein Workpath-Bridge nötig"),
        ("App Engine (Low-Code)", "Custom-Data auf gleicher Plattform"),
        ("EV-Prozess nativ", "kein Lucom-Bridge nötig"),
        ("Live-Reporting", "Excel-Konsolidierung entfällt"),
    ]
    for i, (h, sub) in enumerate(mit_ints):
        y = 24 + i * 6.8
        add_rect(s, 84, y, 70, 5.6, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=GREEN, line_w=0.5)
        add_rect(s, 84, y, 0.5, 5.6, GREEN)
        add_text(s, 86, y + 0.1, 30, 5.4, [
            (h, dict(size=10, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 116, y + 0.1, 38, 5.4, [
            (sub, dict(size=9, italic=True, color=GREEN))], anchor=MSO_ANCHOR.MIDDLE)
    add_slide_footer(s,
        "Effekt: −7 Integrationen · − 0,7 FTE Betrieb · − 45 Personentage / Jahr Wartung/Fehlerbehebung.")

slide_konsolidierung_interfaces()

# ============================================================
# SLIDE 11: FLANKING PROCESSES
# ============================================================
def slide_flanking():
    s = new_slide()
    add_slide_header(s, "Flankierende Prozess-Potentiale",
                     "Zusätzliche Einsparungen durch SPM-nahe Prozess-Digitalisierung",
                     slide_num=11)
    # KPI
    add_rect(s, 4, 8, 152, 8, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 6, 8, 100, 8, [
        ("Zusätzliches Einspar-Potenzial (SPM-flankierend, p.a.)",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 106, 8, 50, 8, [
        (f"~ {TOTAL_FLANKING:,} k€ / Jahr".replace(",", "."),
         dict(size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    # Detail rows
    max_val = max(x[2] for x in FLANKING)
    row_top = 18
    row_h = 11.5
    for i, (name, desc, val) in enumerate(FLANKING):
        y = row_top + i * (row_h + 0.6)
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        add_rect(s, 4, y, 152, row_h, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, 4, y, 0.7, row_h, ORANGE)
        # Number badge
        add_oval(s, 6, y + 3.4, 4, 4, ORANGE)
        add_text(s, 6, y + 3.4, 4, 4, [
            (str(i+1), dict(size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        # Text
        add_text(s, 12, y + 0.4, 100, 4, [
            (name, dict(size=11, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 12, y + 4.5, 100, 6, [
            (desc, dict(size=9.5, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.TOP)
        # Bar
        bar_x = 118
        bar_w_max = 20
        bar_w_val = bar_w_max * (val / max_val)
        add_rect(s, bar_x, y + 4, bar_w_max, 3, LIGHT_GREY,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, bar_x, y + 4, bar_w_val, 3, ORANGE,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        # Value
        add_text(s, 140, y, 15, row_h, [
            (f"{val} k€",
             dict(size=13, bold=True, color=NAVY, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    add_slide_footer(s,
        "Fokus-Kandidat: Maßnahmen-Management Zielkostenmgmt Produktentwicklung — höchster Hebel + strategisch bedeutend.")

slide_flanking()

# ============================================================
# SLIDE 12: SZENARIEN ÜBERSICHT
# ============================================================
def slide_szenarien_uebersicht():
    s = new_slide()
    add_slide_header(s, "Einführungs-Szenarien im Vergleich",
                     "Drei Szenarien zur Umsetzung SEPM – Kosten, Nutzen, Risiko",
                     slide_num=12)
    # Three scenario cards
    top = 8
    w = (152 - 4) / 3
    gap = 2
    scenarios = [
        ("A – Full Scope\n24 Monate", 24, ORANGE, "Empfohlen",
         [("Scope", "Alle Module Phase 1 + 2"),
          ("Ramp Up", "Standard-Timeline (24 Mo)"),
          ("Payback", "~ 3,5 Jahre"),
          ("Business Value", "Voll ab Jahr 3"),
          ("Risiko", "Mittel — Change-intensiv")],
         True),
        ("B – Extended Ramp\n36 Monate", 36, BLUE_HD, "Konservativ",
         [("Scope", "Alle Module, verteilter"),
          ("Ramp Up", "Reduziertes Tempo (+12 Mo)"),
          ("Payback", "~ 4,5 Jahre"),
          ("Business Value", "Voll ab Jahr 4"),
          ("Risiko", "Niedrig — mehr Puffer")],
         False),
        ("C – MVP-First\n12 Mo + Ausbau", 30, GREEN, "Agil",
         [("Scope", "Strategy + Portfolio zuerst"),
          ("Ramp Up", "MVP live in 12 Mo, dann sukzessiv"),
          ("Payback", "~ 3 Jahre auf Teil-Nutzen"),
          ("Business Value", "Teilnutzen früh, Voll ab Jahr 4"),
          ("Risiko", "Mittel — Nachreichung Module")],
         False),
    ]
    for i, (name, months, color, badge, items, recommend) in enumerate(scenarios):
        x = 4 + i * (w + gap)
        add_rect(s, x, top, w, 72, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=color, line_w=2 if recommend else 1)
        add_rect(s, x, top, w, 8, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        # Recommendation ribbon
        if recommend:
            add_rect(s, x + w - 20, top, 20, 8, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            add_text(s, x + w - 20, top, 20, 8, [
                ("★ EMPFOHLEN",
                 dict(size=9, bold=True, color=ORANGE, align=PP_ALIGN.CENTER))],
                anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 1, top + 0.5, w - 22 if recommend else w - 2, 7, [
            (name, dict(size=13, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
        # Badge
        add_rect(s, x + 2, top + 10, w - 4, 4, LIGHT_BG,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x + 2, top + 10, w - 4, 4, [
            (badge, dict(size=10, bold=True, color=color, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        # Duration timeline visualization
        add_text(s, x + 2, top + 16, w - 4, 2.4, [
            (f"Dauer: {months} Monate",
             dict(size=9.5, italic=True, color=GREY_TXT, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        # Bar visualization
        max_m = 36
        bar_w = (w - 4) * (months / max_m)
        add_rect(s, x + 2, top + 19, w - 4, 2, LIGHT_GREY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, x + 2, top + 19, bar_w, 2, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        # Items
        for j, (k, v) in enumerate(items):
            y = top + 24 + j * 8.5
            add_rect(s, x + 2, y, w - 4, 7.5, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            add_text(s, x + 3, y + 0.3, w - 6, 3, [
                (k, dict(size=9, bold=True, color=GREY_TXT))], anchor=MSO_ANCHOR.MIDDLE)
            add_text(s, x + 3, y + 3.3, w - 6, 4, [
                (v, dict(size=10, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    add_slide_footer(s,
        "Detail-Kalkulation Szenarien in Excel-Vorlage Tab 'Szenarien'. Empfehlung: A – Full Scope 24 Monate.")

slide_szenarien_uebersicht()

# ============================================================
# SLIDE 13: SZENARIO A DETAIL
# ============================================================
def slide_szenario_A():
    s = new_slide()
    add_slide_header(s, "Szenario A – Full Scope 24 Monate (Empfohlen)",
                     "Timeline · Investitionsprofil · Nutzen-Kurve · Break-even",
                     slide_num=13)
    # Timeline swim lanes across 24 months
    add_text(s, 4, 8, 152, 2.4, [
        ("Umsetzungs-Timeline",
         dict(size=13, bold=True, color=DARK))])
    # Month axis
    axis_y = 11
    n_m = 24
    left_x = 30
    right_x = 156
    m_w = (right_x - left_x) / n_m
    def mx(m): return left_x + m * m_w
    add_rect(s, left_x, axis_y, 12 * m_w, 3.6, ORANGE_LT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, left_x, axis_y, 12 * m_w, 3.6, [
        ("Jahr 1", dict(size=11, bold=True, color=DARK, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, left_x + 12 * m_w, axis_y, 12 * m_w, 3.6, ORANGE2,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, left_x + 12 * m_w, axis_y, 12 * m_w, 3.6, [
        ("Jahr 2", dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    # Phases
    phases = [
        ("Konzeption",            NAVY,     0,  3),
        ("Impl. Phase 1 (Demand, Portfolio, Strategy)", ORANGE, 3, 12),
        ("Go-Live Phase 1",       GREEN,    12, 13),
        ("Impl. Phase 2 (PM, Ress., Fin.)", ORANGE2, 12, 23),
        ("Go-Live Phase 2",       GREEN,    23, 24),
        ("Change / Rollout",      BLUE_HD,  3,  24),
    ]
    lane_top = axis_y + 5
    lane_h = 5.5
    lane_gap = 0.6
    for i, (name, color, ms, me) in enumerate(phases):
        y = lane_top + i * (lane_h + lane_gap)
        add_text(s, 4, y, 25, lane_h, [
            (name, dict(size=10, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, left_x, y, n_m * m_w, lane_h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        bar_x = mx(ms)
        bar_w = (me - ms) * m_w - 0.2
        add_rect(s, bar_x + 0.1, y + 0.4, bar_w, lane_h - 0.8, color,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    # Investitions + Nutzen curve (schematic)
    curve_top = lane_top + len(phases) * (lane_h + lane_gap) + 3.5
    add_text(s, 4, curve_top, 152, 2.4, [
        ("Investitions- und Nutzenprofil (schematisch, Detail-Kalkulation in Excel)",
         dict(size=12, bold=True, color=DARK))])
    ct = curve_top + 3
    ch = 24
    # Background
    add_rect(s, 4, ct, 152, ch, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    # 5-year x-axis
    year_w = 152 / 5
    for yi in range(5):
        x = 4 + yi * year_w
        col = WHITE if yi % 2 == 0 else GREY_BG
        add_rect(s, x, ct, year_w - 0.05, ch, col)
        add_text(s, x, ct + ch - 3, year_w, 2.4, [
            (f"Jahr {yi + 1}",
             dict(size=9, color=GREY_TXT, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    # Investment bars (Y1 heaviest, tapering)
    inv_vals = [2.0, 2.2, 1.2, 1.0, 1.0]  # M€ approx
    ben_vals = [0.3, 1.5, 3.0, 4.0, 4.3]  # M€ cumulative-like
    max_v = 4.5
    for yi in range(5):
        cx = 4 + yi * year_w + year_w / 2
        # Investment bar (red)
        h_inv = (inv_vals[yi] / max_v) * (ch - 5)
        add_rect(s, cx - 4, ct + ch - 4 - h_inv, 3.5, h_inv, RED_HL,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, cx - 6, ct + ch - 3 - h_inv - 2, 8, 1.8, [
            (f"{inv_vals[yi]:.1f}", dict(size=8, bold=True, color=RED_HL, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        # Benefit bar (green)
        h_ben = (ben_vals[yi] / max_v) * (ch - 5)
        add_rect(s, cx + 0.5, ct + ch - 4 - h_ben, 3.5, h_ben, GREEN,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, cx - 2, ct + ch - 3 - h_ben - 2, 8, 1.8, [
            (f"{ben_vals[yi]:.1f}", dict(size=8, bold=True, color=GREEN, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    # Break-even marker
    add_text(s, 4 + 3.5 * year_w, ct + 1, year_w, 2.4, [
        ("★ Break-even", dict(size=10, bold=True, color=ORANGE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    # Legend
    add_rect(s, 4, ct + ch - 0.2, 4, 1.6, RED_HL, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 9, ct + ch - 0.7, 20, 2.4, [
        ("Investition M€", dict(size=8.5, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, 30, ct + ch - 0.2, 4, 1.6, GREEN, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 35, ct + ch - 0.7, 20, 2.4, [
        ("Nutzen M€", dict(size=8.5, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    add_slide_footer(s,
        "Detail-Kalkulation Szenario A in Excel Tab 'Szenarien'.  Break-even Ø Base-Case ~ Jahr 4.")

slide_szenario_A()

# ============================================================
# SLIDE 14: SZENARIO B DETAIL
# ============================================================
def slide_szenario_B():
    s = new_slide()
    add_slide_header(s, "Szenario B – Extended Ramp-Up 36 Monate",
                     "Vergleich zu Szenario A: konservativer, weniger Change-Druck",
                     slide_num=14)
    # Side by side comparison
    left_x = 4; right_x = 82; w = 74
    top = 8
    for k, (title, color, months, txt) in enumerate([
        ("Szenario A – 24 Monate", ORANGE, 24,
         ["Standard-Tempo, 2 Jahre End-to-End",
          "Höherer Change-Druck",
          "Nutzen ab Jahr 3 vollständig",
          "Payback ~ 3,5 Jahre",
          "Empfohlen bei ausreichenden Ressourcen"]),
        ("Szenario B – 36 Monate", BLUE_HD, 36,
         ["Verteiltes Tempo, 3 Jahre End-to-End",
          "Weniger Change-Druck",
          "Nutzen ab Jahr 4 vollständig",
          "Payback ~ 4,5 Jahre",
          "Empfohlen bei knappen Ressourcen / hoher Aversion"]),
    ]):
        x = left_x if k == 0 else right_x
        add_rect(s, x, top, w, 40, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=color, line_w=1.5)
        add_rect(s, x, top, w, 5, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x, top, w, 5, [
            (title, dict(size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        # Duration bar
        bar_x = x + 2
        max_m = 36
        add_text(s, x + 2, top + 6, w - 4, 2.4, [
            (f"Dauer: {months} Monate",
             dict(size=10, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, bar_x, top + 9, w - 4, 2, LIGHT_GREY,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, bar_x, top + 9, (w - 4) * months / max_m, 2, color,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        for j, tx in enumerate(txt):
            y = top + 13 + j * 5
            add_oval(s, x + 2, y + 1.5, 1.4, 1.4, color)
            add_text(s, x + 4.5, y, w - 6, 4.5, [
                (tx, dict(size=10.5, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)

    # TCO comparison table
    add_text(s, 4, 51, 152, 2.4, [
        ("Kosten- & Nutzen-Delta",
         dict(size=13, bold=True, color=DARK))])
    rows = [
        ("Einmalige Kosten",    "3,0 M€",   "2,7 M€",  "−0,3 M€"),
        ("Laufend p.a. steady", "0,85 M€",  "0,80 M€", "−0,05 M€"),
        ("5-Jahres-TCO",        "7,0 M€",   "6,7 M€",  "−0,3 M€"),
        ("Nutzen Y3 (kumul.)",  "3,5 M€",   "1,8 M€",  "−1,7 M€"),
        ("Nutzen Y5 (kumul.)",  "16 M€",    "12 M€",   "−4 M€"),
        ("Payback",             "~ 3,5 J.", "~ 4,5 J.","+1 J."),
    ]
    hdr_y = 55
    add_rect(s, 4, hdr_y, 152, 3.6, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    for i, h in enumerate(["Kennzahl", "Szenario A", "Szenario B", "Δ B vs A"]):
        add_text(s, [4, 60, 92, 124][i], hdr_y, [55, 32, 32, 32][i], 3.6, [
            (h, dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    for j, (k, va, vb, d) in enumerate(rows):
        y = hdr_y + 4 + j * 4
        bg = LIGHT_BG if j % 2 == 0 else ROW_ALT
        add_rect(s, 4, y, 152, 3.7, bg)
        add_text(s, 6, y, 52, 3.7, [
            (k, dict(size=10.5, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 60, y, 32, 3.7, [
            (va, dict(size=10.5, color=ORANGE, bold=True, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 92, y, 32, 3.7, [
            (vb, dict(size=10.5, color=BLUE_HD, bold=True, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 124, y, 32, 3.7, [
            (d, dict(size=10.5, color=RED_HL if "−" in d and "M€" in d else DARK,
                     bold=True, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    add_slide_footer(s,
        "Fazit: Szenario B spart geringfügig Kosten, verzögert aber ~ 4 M€ Nutzen um 1 Jahr → A ist wirtschaftlich stärker.")

slide_szenario_B()

# ============================================================
# SLIDE 15: BUDGETPLANUNG (Punkt 6 Folie 4)
# ============================================================
def slide_budget():
    s = new_slide()
    add_slide_header(s, "Budgetplanung SEPM",
                     "Punkt 6 Folie 4: Investitionsprofil für Jahresplanung 2027 + 2028",
                     slide_num=15)
    # Yearly investment bars
    years = [2027, 2028, 2029, 2030, 2031]
    invs = [(2200, "Phase 1 Impl. + Lizenz Y1"),
            (2200, "Phase 2 Impl. + Lizenz Y2"),
            (1200, "Steady State + Weiterentwicklung"),
            (1000, "Steady State"),
            (1000, "Steady State")]
    total = sum(v[0] for v in invs)
    # KPI at top
    add_rect(s, 4, 8, 152, 8, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 6, 8, 90, 8, [
        ("Gesamt-Investition SEPM 2027–2031 (Base Case)",
         dict(size=13, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 96, 8, 60, 8, [
        (f"~ {total/1000:.1f} Mio. €",
         dict(size=22, bold=True, color=ORANGE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    # Bar chart
    chart_top = 20
    chart_h = 40
    chart_x = 8
    chart_w = 148
    n = len(years)
    slot_w = chart_w / n
    max_v = max(v[0] for v in invs)
    for i, (yr, (val, label)) in enumerate(zip(years, invs)):
        cx = chart_x + i * slot_w + slot_w / 2
        h = (val / max_v) * (chart_h - 8)
        color = ORANGE if i < 2 else (ORANGE2 if i < 3 else ORANGE_LT)
        add_rect(s, cx - 8, chart_top + chart_h - h - 4, 16, h, color,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        # Value
        add_text(s, cx - 10, chart_top + chart_h - h - 6.5, 20, 2.4, [
            (f"{val/1000:.1f} M€",
             dict(size=10.5, bold=True, color=DARK, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        # Year label
        add_text(s, cx - 10, chart_top + chart_h - 3.5, 20, 3, [
            (str(yr), dict(size=11, bold=True, color=DARK, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        # Sub-label
        add_text(s, cx - 20, chart_top + chart_h - 0.5, 40, 3, [
            (label, dict(size=8, italic=True, color=GREY_TXT, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    # Axis line
    add_rect(s, chart_x, chart_top + chart_h - 4, chart_w, 0.15, DARK)
    # Milestones below
    add_text(s, 4, 68, 152, 2.4, [
        ("Milestones Jahresplanung",
         dict(size=12, bold=True, color=DARK))])
    ms = [
        ("Q3 2026: Jahresplanung", "Investitions-Beschluss 2027 & Rahmen 2028"),
        ("Q3 2027: Jahresplanung", "Freigabe Phase 2 & Rahmen 2028"),
        ("Q3 2028: Review", "Übergang in Steady State"),
    ]
    for i, (t, sub) in enumerate(ms):
        x = 4 + i * 50
        add_rect(s, x, 71.5, 48, 11, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=NAVY, line_w=0.8)
        add_diamond(s, x + 1, 74.5, 3, 3, RED_HL)
        add_text(s, x + 5, 71.9, 42, 4, [
            (t, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 5, 75.5, 42, 6, [
            (sub, dict(size=9, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.MIDDLE)
    add_slide_footer(s,
        "Alle Werte Base Case. Für Jahresplanung 09/26 wird das Detail-Budget aus Excel Tab 'Szenarien' abgeleitet.")

slide_budget()

# ============================================================
# SLIDE 16: RISIKO SNOW ABHÄNGIGKEIT
# ============================================================
def slide_risk():
    s = new_slide()
    add_slide_header(s, "Risiken der ServiceNow-Plattform-Abhängigkeit",
                     "Systematische Bewertung + Mitigationen",
                     slide_num=16)
    # Risk matrix (5 rows)
    add_text(s, 4, 8, 152, 2.4, [
        ("Risiko-Matrix",
         dict(size=13, bold=True, color=DARK))])
    # Header
    hdr_y = 11
    add_rect(s, 4, hdr_y, 152, 3.6, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    hdrs = [("Risiko", 40), ("Schwere", 18), ("Mitigation", 94)]
    xh = 4
    for hd, w in hdrs:
        add_text(s, xh + 2, hdr_y, w - 2, 3.6, [
            (hd, dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.LEFT))],
            anchor=MSO_ANCHOR.MIDDLE)
        xh += w
    # Rows
    for i, (risk, sev, mit) in enumerate(RISKS):
        y = hdr_y + 4 + i * 8.5
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        add_rect(s, 4, y, 152, 8, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        # Risk name
        add_text(s, 6, y + 0.3, 38, 7.4, [
            (risk, dict(size=11, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        # Severity (5 dots)
        for k in range(5):
            filled = k < sev
            color = (RED_HL if sev >= 4 else (YELLOW if sev == 3 else GREEN)) if filled else LIGHT_GREY
            add_oval(s, 46 + k * 3, y + 3.2, 2.4, 2.4, color)
        # Mitigation
        add_text(s, 64, y + 0.3, 90, 7.4, [
            (mit, dict(size=10, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    # Summary box at bottom
    add_rect(s, 4, 72, 152, 10, GREEN_LT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, 4, 72, 0.7, 10, GREEN)
    add_text(s, 6, 72, 148, 3.4, [
        ("Fazit Risiko-Bewertung",
         dict(size=12, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 6, 75, 148, 6, [
        ("Kein Risiko ist ein Show-Stopper.  Alle Risiken sind mit Standard-Verträgen "
         "(Preisgarantie, DSGVO-ADV, SLA), Standard-First-Governance und regelmäßigem "
         "Markt-Screening beherrschbar.  Der größte Hebel gegen Vendor Lock-in: konsequente "
         "Nutzung der SNOW-Standard-Prozesse ohne tiefe Customizings.",
         dict(size=10, color=DARK))], anchor=MSO_ANCHOR.TOP)
    add_slide_footer(s,
        "Detail-Bewertung + Aktions-Plan in Excel-Vorlage Tab 'Risiken'. Review halbjährlich im Programm-Board.")

slide_risk()

# ============================================================
# SLIDE 17: VERTRAGSHISTORIE SNOW 2014–HEUTE
# ============================================================
def slide_snow_history():
    s = new_slide()
    add_slide_header(s, "ServiceNow bei STIHL – Vertragshistorie 2014–2026",
                     "Bewährter Partner: 12 Jahre gemeinsame Historie, kontrolliertes Wachstum",
                     slide_num=17)
    # Timeline horizontal
    add_text(s, 4, 8, 152, 2.4, [
        ("Vertragsstufen und jährliche Lizenzausgaben (indikativ)",
         dict(size=13, bold=True, color=DARK))])
    # Chart area
    ct = 12
    ch = 42
    add_rect(s, 4, ct, 152, ch, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    # X: years
    xs = [x[0] for x in SNOW_HISTORY]
    x_min, x_max = 2014, 2026
    def xp(yr): return 8 + (yr - x_min) / (x_max - x_min) * 144
    # Baseline
    add_rect(s, 8, ct + ch - 5, 148, 0.15, DARK)
    # Y max scale
    max_val = max(x[3] for x in SNOW_HISTORY)
    # Draw connect line + points
    prev_x, prev_y = None, None
    for yr, event, note, val in SNOW_HISTORY:
        x = xp(yr)
        h = (val / max_val) * (ch - 10)
        y = ct + ch - 5 - h
        # Vertical bar
        add_rect(s, x - 1.5, y, 3, h, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        # Point marker
        add_oval(s, x - 1.2, y - 1.2, 2.4, 2.4, NAVY)
        # Value on top
        add_text(s, x - 10, y - 4.6, 20, 2.4, [
            (f"{val} k€",
             dict(size=9, bold=True, color=NAVY, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        # Year label
        add_text(s, x - 10, ct + ch - 4, 20, 3, [
            (str(yr), dict(size=10, bold=True, color=DARK, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        if prev_x is not None:
            # Draw a line via a thin rect (approx)
            pass  # kept simple
        prev_x, prev_y = x, y
    # Event list below
    add_text(s, 4, 56, 152, 2.4, [
        ("Meilensteine der Zusammenarbeit",
         dict(size=13, bold=True, color=DARK))])
    ev_top = 59
    for i, (yr, event, note, val) in enumerate(SNOW_HISTORY):
        col = i % 3
        row = i // 3
        x = 4 + col * 51
        y = ev_top + row * 8
        add_rect(s, x, y, 49, 7, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=ORANGE, line_w=0.5)
        add_rect(s, x, y, 0.5, 7, ORANGE)
        add_text(s, x + 1.5, y + 0.2, 46, 3, [
            (f"{yr} · {event}",
             dict(size=10, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 1.5, y + 3.4, 46, 3.4, [
            (note, dict(size=9, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)
    add_slide_footer(s,
        "Detail-Kalkulation und Vertragsdokumente in Excel-Vorlage Tab 'SNOW_Vertragshistorie'.")

slide_snow_history()

# ============================================================
# SLIDE 18: ROI-KALKULATION & EMPFEHLUNG
# ============================================================
def slide_roi_recommendation():
    s = new_slide()
    add_slide_header(s, "ROI-Kalkulation & Empfehlung",
                     "Gesamt-Wirtschaftlichkeit und Beschlussvorschlag",
                     slide_num=18)
    # Top KPIs
    kpis = [
        ("NPV 5 Jahre",   "~ +8,5 M€", "Base Case, WACC 6 %", GREEN),
        ("IRR",           "~ 32 %",    "über Investitionshürde", GREEN),
        ("Payback",       "~ 3,5 J.",  "nach Kickoff", ORANGE),
        ("TCO 5 Jahre",   "~ 7,0 M€",  "einmalig + laufend",  NAVY),
    ]
    top = 8
    w = (152 - 3 * 1.5) / 4
    for i, (label, big, sub, col) in enumerate(kpis):
        x = 4 + i * (w + 1.5)
        add_rect(s, x, top, w, 16, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=col, line_w=1.5)
        add_rect(s, x, top, w, 3.4, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x, top, w, 3.4, [
            (label, dict(size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x, top + 4, w, 6, [
            (big, dict(size=22, bold=True, color=col, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x, top + 11, w, 4, [
            (sub, dict(size=9, italic=True, color=GREY_TXT, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    # Cost-Benefit summary
    add_text(s, 4, 27, 152, 2.4, [
        ("Kosten-Nutzen-Übersicht (5 Jahre kumuliert, Base Case)",
         dict(size=13, bold=True, color=DARK))])
    rows = [
        ("Kosten (TCO)",      "− 7,0 M€",   RED_HL),
        ("Konsolidierung Tools (5 J.)", "+ 4,7 M€", GREEN),
        ("Konsolidierung Schnittstellen", "+ 1,1 M€", GREEN),
        ("Business Value",    "+ 11,6 M€",  GREEN),
        ("Flankierende Prozesse", "+ 5,3 M€", GREEN),
        ("Netto-Effekt 5 Jahre", "+ 15,7 M€", DARK),
    ]
    tt = 31
    for i, (name, val, color) in enumerate(rows):
        y = tt + i * 6
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        is_total = (i == len(rows) - 1)
        add_rect(s, 4, y, 152, 5.6, bg if not is_total else DARK,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, 6, y + 0.3, 90, 5.0, [
            (name, dict(size=11, bold=is_total,
                        color=DARK if not is_total else WHITE))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 100, y + 0.3, 52, 5.0, [
            (val, dict(size=14 if is_total else 12, bold=True,
                       color=color if not is_total else ORANGE, align=PP_ALIGN.RIGHT))],
            anchor=MSO_ANCHOR.MIDDLE)
    # Recommendation
    add_rect(s, 4, 69, 152, 13, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 6, 69, 40, 13, [
        ("★ EMPFEHLUNG", dict(size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 46, 69, 108, 13, [
        ("Beschluss zur Umsetzung von Szenario A (Full Scope 24 Monate) und "
         "Freigabe Detail-Planung / RFP-Prozess bis Q3 2026 empfohlen. "
         "NPV +8,5 M€, IRR 32 %, Payback ~3,5 J. — deutlich über Investitions-Hürde.",
         dict(size=11.5, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_slide_footer(s,
        "Detail-Zahlen aus Excel-Vorlage 'Auswertung' und 'Szenarien'. Nach Freigabe: Detail-Konzept + RFP.")

slide_roi_recommendation()

prs.save(DST_PPTX)
print(f"Saved PPTX: {DST_PPTX}  ({len(prs.slides)} slides)")

# ============================================================
# ============================================================
# EXCEL WORKBOOK Business_Case_SEPM.xlsx
# ============================================================
# ============================================================
wb = openpyxl.Workbook()
wb.remove(wb.active)  # remove default sheet

def hdr_fmt(cell, bg="1F2937"):
    cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def subhdr_fmt(cell, bg="F07F12"):
    cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def cell_input(cell):
    cell.font = Font(name="Calibri", size=10, color="1F2937")
    cell.fill = PatternFill("solid", fgColor="FFF6E5")  # light orange = input
    cell.alignment = Alignment(horizontal="right", vertical="center")

def cell_calc(cell):
    cell.font = Font(name="Calibri", size=10, bold=True, color="1F2937")
    cell.fill = PatternFill("solid", fgColor="EEF1F4")
    cell.alignment = Alignment(horizontal="right", vertical="center")

def cell_label(cell):
    cell.font = Font(name="Calibri", size=10, color="1F2937")
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

# ---- README ----
ws = wb.create_sheet("Readme")
ws["A1"] = "Business Case SEPM – Datenerhebung"
ws["A1"].font = Font(name="Calibri", size=16, bold=True, color="F07F12")
ws["A3"] = ("Diese Arbeitsmappe dient der Erfassung der Zahlen für den Business Case. "
            "Alle mit oranger Hintergrundfarbe markierten Zellen sind Eingabefelder. "
            "Blaue/graue Zellen sind berechnet.")
ws["A3"].alignment = Alignment(wrap_text=True, vertical="top")
ws.row_dimensions[3].height = 40
tabs = [
    ("Assumptions",  "Basis-Annahmen (FTE-Satz, WACC, Jahre)"),
    ("TCO_Detail",   "Einmalige + laufende Kosten SEPM, aufgeschlüsselt"),
    ("Konsolidierung_Tools", "Ablöse-Ersparnis pro heutigem Tool (PIT, MSPO, …)"),
    ("Konsolidierung_Schnittstellen", "Ersparnis durch Wegfall von APIs/Integrationen"),
    ("Support_Kosten", "Interner + externer Support/Wartung, Weiterentwicklung"),
    ("Business_Value", "Nutzen-Kategorien mit quantifiziertem Steady State"),
    ("Prozess_Potentiale", "Flankierende Prozesse (Maßnahmen-Mgmt, APM, …)"),
    ("Szenarien", "Szenario A / B / C – Vergleich"),
    ("SNOW_Vertragshistorie", "Vertragshistorie 2014–heute, Preisentwicklung"),
    ("Risiken", "Risiko-Register + Mitigationen"),
    ("Sensitivity", "Best / Base / Worst Case"),
    ("Auswertung", "Executive Summary (automatisch aus obigen Tabs)"),
]
ws["A5"] = "Reiter der Arbeitsmappe:"
ws["A5"].font = Font(bold=True, size=12)
for i, (name, desc) in enumerate(tabs):
    r = 6 + i
    ws.cell(row=r, column=1, value=name).font = Font(bold=True, color="F07F12")
    ws.cell(row=r, column=2, value=desc).alignment = Alignment(wrap_text=True)
ws.column_dimensions["A"].width = 32
ws.column_dimensions["B"].width = 90

# ---- Assumptions ----
ws = wb.create_sheet("Assumptions")
ws["A1"] = "Basis-Annahmen"
ws["A1"].font = Font(size=14, bold=True, color="F07F12")
assumptions = [
    ("Betrachtungshorizont (Jahre)", 5, "Jahre"),
    ("FTE-Satz (fully loaded)",      130, "k€ p.a."),
    ("WACC (Diskontrate)",           0.06, "%"),
    ("Startjahr",                    2027, "Jahr"),
    ("Inflations-Annahme",           0.02, "%"),
    ("Mengengerüst Named User (Start)", 200, "Anzahl"),
    ("Mengengerüst Named User (Steady)", 400, "Anzahl"),
    ("Preis pro Named User p.a.",    2000, "€"),
    ("Externer Support Retainer p.a.", 180, "k€"),
    ("Contingency (Puffer)",         0.10, "%"),
]
ws["A3"] = "Parameter"; ws["B3"] = "Wert"; ws["C3"] = "Einheit"
for c in range(1, 4):
    hdr_fmt(ws.cell(row=3, column=c))
for i, (k, v, unit) in enumerate(assumptions):
    r = 4 + i
    cell_label(ws.cell(row=r, column=1, value=k))
    ws.cell(row=r, column=2, value=v)
    cell_input(ws.cell(row=r, column=2))
    cell_label(ws.cell(row=r, column=3, value=unit))
ws.column_dimensions["A"].width = 40
ws.column_dimensions["B"].width = 18
ws.column_dimensions["C"].width = 15

# ---- TCO_Detail ----
ws = wb.create_sheet("TCO_Detail")
ws["A1"] = "TCO-Detail 5 Jahre"
ws["A1"].font = Font(size=14, bold=True, color="F07F12")
hdrs = ["Kostenposition", "Jahr 1", "Jahr 2", "Jahr 3", "Jahr 4", "Jahr 5", "Summe", "Kommentar"]
for i, h in enumerate(hdrs):
    hdr_fmt(ws.cell(row=3, column=i + 1, value=h))
positions = [
    ("Einmalige Kosten", "", "", "", "", "", "", ""),
    ("Implementierungs-Partner (SI)",           IMPL_ONE, 0, 0, 0, 0, "", ""),
    ("Interne FTE Impl. (in k€)",               INT_FTE_IMPL * FTE_RATE, INT_FTE_IMPL * FTE_RATE, 0, 0, 0, "", ""),
    ("Change / Schulung",                       CHANGE_ONE, 100, 0, 0, 0, "", ""),
    ("Contingency (Puffer)",                    150, 100, 0, 0, 0, "", ""),
    ("Laufende Kosten p.a.", "", "", "", "", "", "", ""),
    ("Lizenzen SNOW SPM Pro",                   LICENSE_Y1, 780, LICENSE_Y_STEADY, LICENSE_Y_STEADY, LICENSE_Y_STEADY, "", ""),
    ("Interner Betrieb (FTE × Satz)",           260, 300, int(SUPPORT_INT_FTE_STEADY * FTE_RATE), int(SUPPORT_INT_FTE_STEADY * FTE_RATE), int(SUPPORT_INT_FTE_STEADY * FTE_RATE), "", ""),
    ("Externer Support / Wartung",              120, 150, SUPPORT_EXT, SUPPORT_EXT, SUPPORT_EXT, "", ""),
    ("Weiterentwicklung",                       50, 150, 220, 220, 220, "", ""),
    ("Summe pro Jahr", "", "", "", "", "", "", ""),
]
sum_rows = []
for i, row_data in enumerate(positions):
    r = 4 + i
    is_section = isinstance(row_data[1], str) and row_data[1] == ""
    for c, v in enumerate(row_data):
        cell = ws.cell(row=r, column=c + 1, value=v)
        if is_section:
            cell.font = Font(bold=True, size=11, color="F07F12")
            cell.fill = PatternFill("solid", fgColor="FFF6E5")
        elif c == 0:
            cell_label(cell)
        elif c == len(row_data) - 1:
            cell_label(cell)
        elif c == 6:
            cell_calc(cell)
        else:
            if isinstance(v, (int, float)) and v > 0:
                cell_input(cell)
            elif v == 0:
                cell_input(cell)
    # Sum formula for numeric rows (Jahr 1..5 -> Summe)
    if not is_section and row_data[0] != "Summe pro Jahr":
        ws.cell(row=r, column=7, value=f"=SUM(B{r}:F{r})")
        cell_calc(ws.cell(row=r, column=7))
# Summe pro Jahr (last row)
last_r = 4 + len(positions) - 1
for c in range(2, 8):
    col_letter = get_column_letter(c)
    # sum only the data rows (exclude section headers)
    rows_to_sum = []
    for i, rd in enumerate(positions[:-1]):
        if not (isinstance(rd[1], str) and rd[1] == ""):
            rows_to_sum.append(f"{col_letter}{4+i}")
    ws.cell(row=last_r, column=c, value=f"={'+'.join(rows_to_sum)}")
    cell = ws.cell(row=last_r, column=c)
    cell.font = Font(bold=True, size=12, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="1F2937")
    cell.alignment = Alignment(horizontal="right", vertical="center")

ws.column_dimensions["A"].width = 40
for col in "BCDEFG":
    ws.column_dimensions[col].width = 14
ws.column_dimensions["H"].width = 40

# ---- Konsolidierung_Tools ----
ws = wb.create_sheet("Konsolidierung_Tools")
ws["A1"] = "Konsolidierungs-Ersparnis: Tool-Ablöse"
ws["A1"].font = Font(size=14, bold=True, color="F07F12")
hdrs = ["Tool", "Kategorie", "Lizenz p.a. (k€)", "Wartung p.a. (k€)",
        "FTE-Aufwand (FTE)", "FTE-Kosten (k€)", "Sonst. (k€)", "Ersparnis p.a. (k€)"]
for i, h in enumerate(hdrs):
    hdr_fmt(ws.cell(row=3, column=i + 1, value=h))
# From CONSOL_TOOLS (only tool + category + total)
for i, (tool, cat, tot) in enumerate(CONSOL_TOOLS):
    r = 4 + i
    ws.cell(row=r, column=1, value=tool); cell_label(ws.cell(row=r, column=1))
    ws.cell(row=r, column=2, value=cat);  cell_label(ws.cell(row=r, column=2))
    # placeholders for user to fill
    for c in [3, 4, 6, 7]:
        cell_input(ws.cell(row=r, column=c, value=0))
    ws.cell(row=r, column=5, value=0); cell_input(ws.cell(row=r, column=5))
    # Column F = FTE * Satz from Assumptions
    ws.cell(row=r, column=6, value=f"=E{r}*Assumptions!B5")
    cell_calc(ws.cell(row=r, column=6))
    # Column H = C+D+F+G, but leave placeholder equal to indication
    ws.cell(row=r, column=8, value=f"=C{r}+D{r}+F{r}+G{r}")
    cell_calc(ws.cell(row=r, column=8))
r_total = 4 + len(CONSOL_TOOLS)
ws.cell(row=r_total, column=1, value="Summe")
ws.cell(row=r_total, column=1).font = Font(bold=True, size=12, color="FFFFFF")
ws.cell(row=r_total, column=1).fill = PatternFill("solid", fgColor="2E7D32")
for c in range(2, 9):
    if c >= 3:
        col_letter = get_column_letter(c)
        ws.cell(row=r_total, column=c,
                value=f"=SUM({col_letter}4:{col_letter}{r_total-1})")
    cell = ws.cell(row=r_total, column=c)
    cell.font = Font(bold=True, size=12, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="2E7D32")
    cell.alignment = Alignment(horizontal="right", vertical="center")
ws.column_dimensions["A"].width = 24
ws.column_dimensions["B"].width = 26
for c in "CDEFGH":
    ws.column_dimensions[c].width = 15

# ---- Konsolidierung_Schnittstellen ----
ws = wb.create_sheet("Konsolidierung_Schnittstellen")
ws["A1"] = "Konsolidierungs-Ersparnis: Schnittstellen / Integrationen"
ws["A1"].font = Font(size=14, bold=True, color="F07F12")
hdrs = ["Schnittstelle heute", "Wartungs-PT/Jahr", "PT-Satz (€/PT)",
        "Ersparnis Wartung (k€)", "Sonst. Betriebskosten (k€)", "Ersparnis gesamt (k€)"]
for i, h in enumerate(hdrs):
    hdr_fmt(ws.cell(row=3, column=i + 1, value=h))
interfaces = [
    "PIT ↔ SAP FI",
    "MSPO ↔ SAP HCM",
    "SharePoint ↔ Excel-Exports",
    "Workpath ↔ Excel-Reporting",
    "PowerApps ↔ Dataverse",
    "Lucom EV ↔ SAP",
    "Excel-Konsolidierungen (diverse)",
]
for i, it in enumerate(interfaces):
    r = 4 + i
    ws.cell(row=r, column=1, value=it); cell_label(ws.cell(row=r, column=1))
    ws.cell(row=r, column=2, value=0); cell_input(ws.cell(row=r, column=2))
    ws.cell(row=r, column=3, value=1200); cell_input(ws.cell(row=r, column=3))
    ws.cell(row=r, column=4, value=f"=B{r}*C{r}/1000")
    cell_calc(ws.cell(row=r, column=4))
    ws.cell(row=r, column=5, value=0); cell_input(ws.cell(row=r, column=5))
    ws.cell(row=r, column=6, value=f"=D{r}+E{r}")
    cell_calc(ws.cell(row=r, column=6))
r_total = 4 + len(interfaces)
ws.cell(row=r_total, column=1, value="Summe")
for c in range(2, 7):
    if c >= 4:
        col_letter = get_column_letter(c)
        ws.cell(row=r_total, column=c, value=f"=SUM({col_letter}4:{col_letter}{r_total-1})")
    cell = ws.cell(row=r_total, column=c)
    cell.font = Font(bold=True, size=12, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="2E7D32")
    cell.alignment = Alignment(horizontal="right", vertical="center")
ws.column_dimensions["A"].width = 38
for c in "BCDEF":
    ws.column_dimensions[c].width = 20

# ---- Support_Kosten ----
ws = wb.create_sheet("Support_Kosten")
ws["A1"] = "Support-Kosten: Intern + Extern + Weiterentwicklung"
ws["A1"].font = Font(size=14, bold=True, color="F07F12")
hdrs = ["Position", "Typ", "FTE", "Preis (k€)", "Kosten p.a. (k€)"]
for i, h in enumerate(hdrs):
    hdr_fmt(ws.cell(row=3, column=i + 1, value=h))
support_rows = [
    ("SNOW-Plattform-Admin",      "Intern FTE", 1.0, ""),
    ("SPM Business Analyst",      "Intern FTE", 1.0, ""),
    ("Integration/Schnittstellen","Intern FTE", 0.5, ""),
    ("Data Governance",           "Intern FTE", 0.3, ""),
    ("User Support 2nd Level",    "Intern FTE", 0.2, ""),
    ("Wartungsvertrag SNOW",      "Extern",     "", 70),
    ("Managed Service Partner",   "Extern",     "", 90),
    ("Externes Consulting",       "Extern",     "", 60),
    ("Zertifizierungen / Training","Extern",    "", 30),
    ("Feature-Backlog Business",  "Weiterentw.","", 100),
    ("Release-Upgrade-Anpassungen","Weiterentw.","",50),
    ("Custom-Reports/Dashboards", "Weiterentw.","", 40),
    ("Innovation Sprints",        "Weiterentw.","", 30),
]
for i, (name, typ, fte, cost) in enumerate(support_rows):
    r = 4 + i
    ws.cell(row=r, column=1, value=name); cell_label(ws.cell(row=r, column=1))
    ws.cell(row=r, column=2, value=typ);  cell_label(ws.cell(row=r, column=2))
    if fte != "":
        ws.cell(row=r, column=3, value=fte); cell_input(ws.cell(row=r, column=3))
        ws.cell(row=r, column=5, value=f"=C{r}*Assumptions!B5")
        cell_calc(ws.cell(row=r, column=5))
    else:
        ws.cell(row=r, column=3, value=""); cell_label(ws.cell(row=r, column=3))
        ws.cell(row=r, column=4, value=cost); cell_input(ws.cell(row=r, column=4))
        ws.cell(row=r, column=5, value=f"=D{r}")
        cell_calc(ws.cell(row=r, column=5))
# Total
r_total = 4 + len(support_rows)
ws.cell(row=r_total, column=1, value="Summe p.a.")
for c in [1, 2, 3, 4, 5]:
    cell = ws.cell(row=r_total, column=c)
    cell.font = Font(bold=True, size=12, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="F07F12")
    cell.alignment = Alignment(horizontal="right", vertical="center")
ws.cell(row=r_total, column=5, value=f"=SUM(E4:E{r_total-1})")
ws.column_dimensions["A"].width = 34
for c in "BCDE":
    ws.column_dimensions[c].width = 16

# ---- Business_Value ----
ws = wb.create_sheet("Business_Value")
ws["A1"] = "Business Value: Nutzen-Kategorien"
ws["A1"].font = Font(size=14, bold=True, color="F07F12")
hdrs = ["Nutzen-Kategorie", "Beschreibung", "Berechnungslogik",
        "Menge (Basis)", "Wert je Einheit (€)", "Nutzen p.a. (k€)"]
for i, h in enumerate(hdrs):
    hdr_fmt(ws.cell(row=3, column=i + 1, value=h))
val_rows = [
    ("Effizienz PM/PPM", VALUE_CATEGORIES[0][1],
     "Reporting-PT × PT-Satz", 400, 1200, ""),
    ("Time-to-Market", VALUE_CATEGORIES[1][1],
     "Ø Zeitgewinn × Umsatz p.d.", "", "", ""),
    ("Governance / Rework", VALUE_CATEGORIES[2][1],
     "Rework-Fälle × Kosten je Fall", "", "", ""),
    ("Ressourcen-Optimierung", VALUE_CATEGORIES[3][1],
     "Ausgelastete FTE × Satz × Gewinn %", "", "", ""),
    ("Compliance / Audit", VALUE_CATEGORIES[4][1],
     "Audit-Aufwand-Reduktion", "", "", ""),
]
for i, (name, desc, formula, menge, wert, total) in enumerate(val_rows):
    r = 4 + i
    ws.cell(row=r, column=1, value=name); cell_label(ws.cell(row=r, column=1))
    ws.cell(row=r, column=2, value=desc); cell_label(ws.cell(row=r, column=2))
    ws.cell(row=r, column=3, value=formula); cell_label(ws.cell(row=r, column=3))
    for c in [4, 5]:
        ws.cell(row=r, column=c, value=(menge if c == 4 else wert)); cell_input(ws.cell(row=r, column=c))
    ws.cell(row=r, column=6, value=f"=IFERROR(D{r}*E{r}/1000,0)")
    cell_calc(ws.cell(row=r, column=6))
r_total = 4 + len(val_rows)
ws.cell(row=r_total, column=1, value="Summe Nutzen p.a.")
for c in range(1, 7):
    cell = ws.cell(row=r_total, column=c)
    cell.font = Font(bold=True, size=12, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="2E7D32")
    cell.alignment = Alignment(horizontal="right", vertical="center")
ws.cell(row=r_total, column=6, value=f"=SUM(F4:F{r_total-1})")
ws.column_dimensions["A"].width = 30
ws.column_dimensions["B"].width = 40
ws.column_dimensions["C"].width = 30
for c in "DEF":
    ws.column_dimensions[c].width = 16

# ---- Prozess_Potentiale ----
ws = wb.create_sheet("Prozess_Potentiale")
ws["A1"] = "Flankierende Prozesse: Zusätzliche Potentiale"
ws["A1"].font = Font(size=14, bold=True, color="F07F12")
hdrs = ["Prozess", "Beschreibung", "Berechnungslogik", "Menge", "Wert (€)", "Ersparnis p.a. (k€)"]
for i, h in enumerate(hdrs):
    hdr_fmt(ws.cell(row=3, column=i + 1, value=h))
for i, (name, desc, saving) in enumerate(FLANKING):
    r = 4 + i
    ws.cell(row=r, column=1, value=name); cell_label(ws.cell(row=r, column=1))
    ws.cell(row=r, column=2, value=desc); cell_label(ws.cell(row=r, column=2))
    ws.cell(row=r, column=3, value=""); cell_label(ws.cell(row=r, column=3))
    ws.cell(row=r, column=4, value=""); cell_input(ws.cell(row=r, column=4))
    ws.cell(row=r, column=5, value=""); cell_input(ws.cell(row=r, column=5))
    ws.cell(row=r, column=6, value=saving); cell_input(ws.cell(row=r, column=6))
r_total = 4 + len(FLANKING)
ws.cell(row=r_total, column=1, value="Summe")
ws.cell(row=r_total, column=6, value=f"=SUM(F4:F{r_total-1})")
for c in range(1, 7):
    cell = ws.cell(row=r_total, column=c)
    cell.font = Font(bold=True, size=12, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="F07F12")
    cell.alignment = Alignment(horizontal="right", vertical="center")
ws.column_dimensions["A"].width = 50
ws.column_dimensions["B"].width = 50
ws.column_dimensions["C"].width = 26
for c in "DEF":
    ws.column_dimensions[c].width = 15

# ---- Szenarien ----
ws = wb.create_sheet("Szenarien")
ws["A1"] = "Szenarien-Vergleich"
ws["A1"].font = Font(size=14, bold=True, color="F07F12")
hdrs = ["Kennzahl", "Szenario A (24 Mo, Full)", "Szenario B (36 Mo, Extended)",
        "Szenario C (12+18 Mo, MVP)"]
for i, h in enumerate(hdrs):
    hdr_fmt(ws.cell(row=3, column=i + 1, value=h))
sz_rows = [
    ("Dauer (Monate)",           24, 36, 30),
    ("Einmalige Kosten (k€)",    3000, 2700, 2200),
    ("Laufend p.a. steady (k€)", 850, 800, 700),
    ("5-Jahres-TCO (k€)",        "", "", ""),
    ("Nutzen Y3 kumul. (k€)",    3500, 1800, 2500),
    ("Nutzen Y5 kumul. (k€)",    16000, 12000, 13500),
    ("Payback (Jahre)",          3.5, 4.5, 3.0),
    ("Empfohlen?",               "★ Ja", "Nein", "Nein"),
]
for i, row in enumerate(sz_rows):
    r = 4 + i
    for c, v in enumerate(row):
        cell = ws.cell(row=r, column=c + 1, value=v)
        if c == 0:
            cell_label(cell); cell.font = Font(bold=True)
        else:
            if isinstance(v, (int, float)):
                cell_input(cell)
            else:
                cell_label(cell); cell.alignment = Alignment(horizontal="center", vertical="center")
# TCO row = einmalig + 5*laufend
r = 4 + 3  # 5-Jahres-TCO row
for c in [2, 3, 4]:
    ws.cell(row=r, column=c,
            value=f"={get_column_letter(c)}5+5*{get_column_letter(c)}6")
    cell_calc(ws.cell(row=r, column=c))
ws.column_dimensions["A"].width = 34
for c in "BCD":
    ws.column_dimensions[c].width = 26

# ---- SNOW_Vertragshistorie ----
ws = wb.create_sheet("SNOW_Vertragshistorie")
ws["A1"] = "ServiceNow: Vertragshistorie 2014 – heute"
ws["A1"].font = Font(size=14, bold=True, color="F07F12")
hdrs = ["Jahr", "Vertragsstufe", "Beschreibung", "Named User", "Lizenz p.a. (k€)",
        "Δ vs Vorjahr (%)", "Sonstige Notizen"]
for i, h in enumerate(hdrs):
    hdr_fmt(ws.cell(row=3, column=i + 1, value=h))
for i, (yr, ev, note, val) in enumerate(SNOW_HISTORY):
    r = 4 + i
    ws.cell(row=r, column=1, value=yr); cell_input(ws.cell(row=r, column=1))
    ws.cell(row=r, column=2, value=ev); cell_label(ws.cell(row=r, column=2))
    ws.cell(row=r, column=3, value=note); cell_label(ws.cell(row=r, column=3))
    ws.cell(row=r, column=4, value=""); cell_input(ws.cell(row=r, column=4))
    ws.cell(row=r, column=5, value=val); cell_input(ws.cell(row=r, column=5))
    if i > 0:
        ws.cell(row=r, column=6, value=f"=(E{r}-E{r-1})/E{r-1}")
        cell_calc(ws.cell(row=r, column=6))
    ws.cell(row=r, column=7, value=""); cell_label(ws.cell(row=r, column=7))
ws.column_dimensions["A"].width = 10
ws.column_dimensions["B"].width = 30
ws.column_dimensions["C"].width = 40
for c in "DEFG":
    ws.column_dimensions[c].width = 18

# ---- Risiken ----
ws = wb.create_sheet("Risiken")
ws["A1"] = "Risiko-Register"
ws["A1"].font = Font(size=14, bold=True, color="F07F12")
hdrs = ["#", "Risiko", "Beschreibung", "Wahrsch. (1-5)", "Auswirkung (1-5)",
        "Score", "Mitigation", "Owner", "Review Datum"]
for i, h in enumerate(hdrs):
    hdr_fmt(ws.cell(row=3, column=i + 1, value=h))
for i, (risk, sev, mit) in enumerate(RISKS):
    r = 4 + i
    ws.cell(row=r, column=1, value=i + 1); cell_label(ws.cell(row=r, column=1))
    ws.cell(row=r, column=2, value=risk); cell_label(ws.cell(row=r, column=2))
    ws.cell(row=r, column=3, value=""); cell_label(ws.cell(row=r, column=3))
    ws.cell(row=r, column=4, value=sev); cell_input(ws.cell(row=r, column=4))
    ws.cell(row=r, column=5, value=sev); cell_input(ws.cell(row=r, column=5))
    ws.cell(row=r, column=6, value=f"=D{r}*E{r}"); cell_calc(ws.cell(row=r, column=6))
    ws.cell(row=r, column=7, value=mit); cell_label(ws.cell(row=r, column=7))
    ws.cell(row=r, column=8, value=""); cell_input(ws.cell(row=r, column=8))
    ws.cell(row=r, column=9, value=""); cell_input(ws.cell(row=r, column=9))
ws.column_dimensions["A"].width = 5
ws.column_dimensions["B"].width = 28
ws.column_dimensions["C"].width = 35
for c in "DEF":
    ws.column_dimensions[c].width = 12
ws.column_dimensions["G"].width = 45
ws.column_dimensions["H"].width = 20
ws.column_dimensions["I"].width = 15

# ---- Sensitivity ----
ws = wb.create_sheet("Sensitivity")
ws["A1"] = "Sensitivitäts-Analyse: Best / Base / Worst Case"
ws["A1"].font = Font(size=14, bold=True, color="F07F12")
hdrs = ["Kennzahl", "Best Case", "Base Case", "Worst Case"]
for i, h in enumerate(hdrs):
    hdr_fmt(ws.cell(row=3, column=i + 1, value=h))
sen_rows = [
    ("5-Jahres-TCO (k€)",         6000, 7000,  8500),
    ("Konsolidierung Tools 5J (k€)", 5500, 4700, 3500),
    ("Business Value 5J (k€)",    14000, 11600, 8500),
    ("Flankierende Prozesse 5J",   6500, 5300,  3800),
    ("Netto-Effekt 5 Jahre (k€)", "",     "",     ""),
    ("NPV (WACC 6 %) (k€)",       10500,  8500,  6000),
    ("IRR (%)",                    42,    32,    22),
    ("Payback (Jahre)",           2.8,   3.5,   4.5),
]
for i, row in enumerate(sen_rows):
    r = 4 + i
    for c, v in enumerate(row):
        cell = ws.cell(row=r, column=c + 1, value=v)
        if c == 0:
            cell_label(cell); cell.font = Font(bold=True)
        elif isinstance(v, (int, float)):
            cell_input(cell)
        else:
            cell_label(cell)
# Netto row = Konsol + BV + Flanking - TCO
r = 4 + 4
for c in [2, 3, 4]:
    L = get_column_letter(c)
    ws.cell(row=r, column=c, value=f"={L}5+{L}6+{L}7-{L}4")
    cell_calc(ws.cell(row=r, column=c))
ws.column_dimensions["A"].width = 30
for c in "BCD":
    ws.column_dimensions[c].width = 20

# ---- Auswertung ----
ws = wb.create_sheet("Auswertung")
ws["A1"] = "Executive Summary Business Case SEPM"
ws["A1"].font = Font(size=16, bold=True, color="F07F12")
ws["A3"] = "Kennzahl"; ws["B3"] = "Wert"; ws["C3"] = "Kommentar"
for c in range(1, 4):
    hdr_fmt(ws.cell(row=3, column=c))
summary = [
    ("5-Jahres-TCO SEPM",         "='TCO_Detail'!G14",     "einmalig + laufend"),
    ("Konsolidierung Tools p.a.", "='Konsolidierung_Tools'!H10", "Steady State"),
    ("Business Value p.a.",       "='Business_Value'!F9",  "Steady State"),
    ("Flankierende Prozesse p.a.","='Prozess_Potentiale'!F9", "SPM-nah, quantifiziert"),
    ("Support-Kosten p.a.",       "='Support_Kosten'!E17", "intern + extern + Weiterentw."),
    ("Payback (Jahre)",           "='Sensitivity'!C10",    "Base Case"),
    ("NPV (WACC 6 %)",            "='Sensitivity'!C8",     "Base Case"),
    ("IRR",                       "='Sensitivity'!C9",     "Base Case"),
]
for i, (k, formula, note) in enumerate(summary):
    r = 4 + i
    cell_label(ws.cell(row=r, column=1, value=k))
    ws.cell(row=r, column=1).font = Font(bold=True)
    ws.cell(row=r, column=2, value=formula)
    cell_calc(ws.cell(row=r, column=2))
    cell_label(ws.cell(row=r, column=3, value=note))
ws.column_dimensions["A"].width = 34
ws.column_dimensions["B"].width = 20
ws.column_dimensions["C"].width = 50

wb.save(DST_XLSX)
print(f"Saved XLSX: {DST_XLSX}  ({len(wb.sheetnames)} tabs)")
