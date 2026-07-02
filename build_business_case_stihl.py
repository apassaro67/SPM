"""Rebuild Business Case SEPM PowerPoint using STIHL template layouts.
Uses Layout 2 ('Title slide empty') for cover and Layout 4 ('Nur Titel')
for content slides. Custom content sits below the STIHL title+subline area.
"""
from copy import deepcopy
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

SRC_TEMPLATE = "/root/.claude/uploads/e26c03e3-000c-5c76-aa93-1101c0a1be8f/7360a6f1-2026_07_01_SPEM_PPB.pptx"
DST_PPTX = "/tmp/2026_07_01_Business_Case_SEPM_STIHL.pptx"

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

U = 76200
def ex(u): return int(u * U)

# Open STIHL template
prs = Presentation(SRC_TEMPLATE)

# Remove all existing slides (keep only layouts + master)
sldIdLst = prs.slides._sldIdLst
for sld in list(sldIdLst):
    rId = sld.get(qn("r:id"))
    sldIdLst.remove(sld)
    try:
        prs.part.drop_rel(rId)
    except Exception:
        pass

# Layout access
LAYOUT_COVER  = prs.slide_masters[0].slide_layouts[2]   # 'Title slide empty'
LAYOUT_TITLE  = prs.slide_masters[0].slide_layouts[4]   # 'Nur Titel'
LAYOUT_CONCL  = prs.slide_masters[0].slide_layouts[18]  # 'Conclusion'

# ------------------------------------------------------------
# Shape helpers
# ------------------------------------------------------------
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
        run.font.name = "Stihl Contraface Text"  # STIHL default; fallback Calibri
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

def set_placeholder_text(slide, placeholder_idx, text, *, size=None, bold=None, color=None):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == placeholder_idx:
            tf = ph.text_frame
            tf.text = text
            if size is not None or bold is not None or color is not None:
                for para in tf.paragraphs:
                    for run in para.runs:
                        if size is not None:
                            run.font.size = Pt(size)
                        if bold is not None:
                            run.font.bold = bold
                        if color is not None:
                            run.font.color.rgb = color
            return
    print(f"WARN: placeholder idx {placeholder_idx} not found")

def add_slide_titled(title, subline):
    """Create a content slide with STIHL 'Nur Titel' layout, populate
    title (idx=0) and subline (idx=13)."""
    s = prs.slides.add_slide(LAYOUT_TITLE)
    set_placeholder_text(s, 0, title)
    set_placeholder_text(s, 13, subline)
    return s

# ------------------------------------------------------------
# Data (same as previous build)
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
    ("Effizienz PM/PPM",       "Weniger manuelle Excels/Berichte, Ressourcen-Steuerung", 820),
    ("Time-to-Market",         "Schnellere Idee→Projekt-Start (Gates in System)",         480),
    ("Governance / Rework",    "Weniger Doppelarbeit, klare Priorisierung",               360),
    ("Ressourcen-Optimierung", "Skill-Match, Auslastungs-Steuerung",                      520),
    ("Compliance / Audit",     "Nachvollziehbarkeit, Business Cases im System",           140),
]
TOTAL_VALUE = sum(x[2] for x in VALUE_CATEGORIES)

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
TOTAL_FLANKING = sum(x[2] for x in FLANKING)

SNOW_HISTORY = [
    (2014, "Erstvertrag ITSM",       "Standard Package, 500 Fulfiller", 180),
    (2016, "Erweiterung Global",     "Ausrollung inkl. China",           260),
    (2018, "Enterprise-Vertrag",     "Volumen-Diskont, HR-SD add-on",    340),
    (2020, "ITOM + CMDB",            "Discovery, Event Mgmt",            410),
    (2022, "Contract Renewal 3J",    "Preisanpassung +5% p.a.",          490),
    (2024, "Now Assist Preview",     "GenAI eingeführt",                 560),
    (2026, "SPM Pro (geplant)",      "Neuer Modul-Baustein",             820),
]

RISKS = [
    ("Vendor Lock-in", 4, "Standard-Konfiguration, Migrations-Klauseln, jährl. Marktscreening"),
    ("Preis-Eskalation", 3, "Mehrjahresvertrag mit Preisgarantie, Volumen-Rabatte"),
    ("Ausfall/Betrieb", 2, "SLA 99,9% + regionaler DC (Frankfurt/Düsseldorf), DSGVO"),
    ("Roadmap-Divergenz", 3, "Standard first, keine tiefen Anpassungen; halbjährl. Review"),
    ("Datenhoheit / DSGVO", 2, "DE-Hosting, ADV-Vertrag, verschlüsselte Datenhaltung"),
    ("Team-Kompetenzaufbau", 3, "Ausbau ITSM-Team +1,5-2 FTE, Ausbildungsplan, Partner-Support"),
]

# Content area available (below STIHL subline at y~20 to bottom y~86)
CY_TOP = 20   # units - start of content
CY_BOT = 84   # units - bottom of content

# ============================================================
# COVER  (Layout 'Title slide empty')
# ============================================================
def slide_cover():
    s = prs.slides.add_slide(LAYOUT_COVER)
    # Title at layout idx=0 (0.62, 2.31, 9.45, 3.54)
    set_placeholder_text(s, 0, "Business Case SEPM")
    # Subline idx=13: what is this project about
    set_placeholder_text(s, 13, "Detaillierte Wirtschaftlichkeits-Betrachtung  ·  TCO · ROI · Business Value · Szenarien · Risiken")
    # Name idx=14
    set_placeholder_text(s, 14, "Alex Passaro · Torsten Zahn")
    # Datum idx=15
    set_placeholder_text(s, 15, "01.07.2026")

slide_cover()

# ============================================================
# SLIDE 2: Management Summary
# ============================================================
def slide_management_summary():
    s = add_slide_titled(
        "Management Summary",
        "Kernaussagen des Business Case auf einen Blick")
    kpis = [
        ("5-Jahres-TCO Vorteil",   "~ 2,2 M€",  "vs Best-of-Breed", GREEN),
        ("Payback",                "~ 3,5 J.",  "nach Go-Live Phase 1", ORANGE),
        ("Jährl. Nutzen (steady)", "~ 4,3 M€",  "Konsolidierung + Value + Flanking", NAVY),
        ("Empfohlenes Szenario",   "A · 24 Mo", "Full Scope Phase 1+2", BLUE_HD),
    ]
    top = CY_TOP
    gap = 1.5
    w = (152 - 3 * gap) / 4
    for i, (label, big, sub, col) in enumerate(kpis):
        x = 4 + i * (w + gap)
        add_rect(s, x, top, w, 20, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=col, line_w=1.5)
        add_rect(s, x, top, w, 4, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x + 0.5, top + 0.3, w - 1, 3.4, [
            (label, dict(size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.5, top + 5.8, w - 1, 8, [
            (big, dict(size=26, bold=True, color=col, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.5, top + 14.5, w - 1, 4.5, [
            (sub, dict(size=10, italic=True, color=GREY_TXT, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    km_top = top + 22
    add_text(s, 4, km_top, 152, 2.4, [
        ("Kernbotschaften", dict(size=13, bold=True, color=ORANGE))])
    msgs = [
        ("Konsolidierung von >6 PPM-Tools", "spart ~950 k€ / Jahr an Lizenz + Wartung + FTE"),
        ("Business Value SPM (steady state)", "~2,3 M€ / Jahr durch Effizienz, Time-to-Market, Governance"),
        ("Flankierende Prozesse (Maßnahmen-Mgmt etc.)", "zusätzlich ~1,0 M€ / Jahr Einsparpotenzial"),
        ("SNOW-Plattform-Risiko", "Vendor Lock-in adressiert durch Standard-first und Marktscreening"),
        ("Vertragshistorie 2014–heute", "SNOW als bewährter STIHL-Partner (ITSM seit 12 J.)"),
    ]
    mb_top = km_top + 3.0
    row_h = 5.5
    for i, (head, body) in enumerate(msgs):
        y = mb_top + i * row_h
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        add_rect(s, 4, y, 152, row_h - 0.4, bg)
        add_rect(s, 4, y, 0.7, row_h - 0.4, ORANGE)
        add_text(s, 6, y + 0.2, 60, row_h - 0.6, [
            (head, dict(size=11, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 68, y + 0.2, 87, row_h - 0.6, [
            (body, dict(size=10, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)

slide_management_summary()

# ============================================================
# SLIDE 3: Business Case Methodik
# ============================================================
def slide_methodik():
    s = add_slide_titled(
        "Business Case Methodik",
        "Vier Dimensionen der Wirtschaftlichkeits-Betrachtung")
    top = CY_TOP
    h = 30
    gap = 2
    w = (152 - gap) / 2
    quads = [
        ("TCO (5 Jahre)", "Lizenzen · Implementierung · Support · Change · Betrieb",
         ["Lizenzen SNOW SPM Pro", "Implementierung (Berater + interne FTE)",
          "Change & Schulung", "Support / Wartung"], ORANGE, "€"),
        ("Nutzen / Business Value", "Effizienz · Time-to-Market · Governance · Ressourcen",
         ["Effizienz PPM", "Time-to-Market",
          "Ressourcen-Steuerung", "Governance & Compliance"], GREEN, "★"),
        ("Konsolidierung", "Tool-Ablöse · Schnittstellen · flankierende Prozesse",
         ["PIT / MSPO / Workpath / PowerApps", "Schnittstellen (APIs, ETL)",
          "Maßnahmen-Mgmt Produktentw.", "Application Portfolio Mgmt"],
         NAVY, "⇨"),
        ("Risiko / Sensitivität", "Vendor · Preis · Betrieb · Roadmap · DSGVO",
         ["Vendor Lock-in", "Preis-Eskalation",
          "Ausfall / Betrieb", "Best/Base/Worst Sensitivity"], RED_HL, "⚠"),
    ]
    for i, (head, sub, items, color, icon) in enumerate(quads):
        row = i // 2
        col = i % 2
        x = 4 + col * (w + gap)
        y = top + row * (h + gap)
        add_rect(s, x, y, w, h, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=color, line_w=1.5)
        add_rect(s, x, y, w, 5, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_oval(s, x + 1, y + 0.8, 3.4, 3.4, WHITE)
        add_text(s, x + 1, y + 0.8, 3.4, 3.4, [
            (icon, dict(size=14, bold=True, color=color, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 5, y + 0.4, w - 6, 4.5, [
            (head, dict(size=13, bold=True, color=WHITE))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 1.5, y + 6, w - 3, 3.2, [
            (sub, dict(size=10, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)
        for j, it in enumerate(items):
            iy = y + 10 + j * 4.5
            add_oval(s, x + 2, iy + 1.4, 1.0, 1.0, color)
            add_text(s, x + 3.6, iy, w - 5, 4.0, [
                (it, dict(size=10.5, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)

slide_methodik()

# ============================================================
# SLIDE 4: Alternative PM-Plattform (Punkt 3 Folie 4)
# ============================================================
def slide_alternative_pm():
    s = add_slide_titled(
        "Alternative PM-Plattform – Bewertung",
        "Punkt 3 Folie 4: PM-Plattform VEW / VPM als Absicherung (Hedge)")
    add_rect(s, 4, CY_TOP, 74, 32, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, 4, CY_TOP, 74, 4, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 5, CY_TOP, 72, 4, [
        ("Warum Alternative bewerten?",
         dict(size=11.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    rat = [
        ("Risiko-Absicherung", "Falls SNOW für VEW/VPM nicht ausreichend tief."),
        ("Verhandlungsposition", "Best-of-Breed-Angebot senkt SNOW-Verhandlungspreis."),
        ("Engineering-Spezifika", "PLM-nahe Prozesse ggf. besser in Fach-Tool."),
        ("EWW-Bedenken", "Explizite Alternative senkt Widerstand."),
    ]
    for i, (h, b) in enumerate(rat):
        y = CY_TOP + 5.5 + i * 6.5
        add_oval(s, 6, y + 1.4, 2.4, 2.4, NAVY)
        add_text(s, 6, y + 1.4, 2.4, 2.4, [
            (str(i + 1), dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 9.5, y + 0.2, 66, 2.8, [
            (h, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 9.5, y + 3.0, 66, 2.8, [
            (b, dict(size=9, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)

    # Right column
    add_rect(s, 82, CY_TOP, 74, 32, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, 82, CY_TOP, 74, 4, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 83, CY_TOP, 72, 4, [
        ("Shortlist Alternative PM-Plattformen",
         dict(size=11.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    alts = [
        ("Planisware Enterprise", "Marktführer Engineering-PPM", "Lead Jonas"),
        ("SAP EPPM", "SAP-Ökosystem-Integration", "Lead Jonas"),
        ("Sciforma", "Mittelstands-Standard, hybrid", "Lead Jonas"),
    ]
    for i, (n, sub, own) in enumerate(alts):
        y = CY_TOP + 5.5 + i * 7
        add_rect(s, 83.5, y, 71, 6, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=NAVY, line_w=0.5)
        add_rect(s, 83.5, y, 0.6, 6, NAVY)
        add_text(s, 85, y + 0.2, 45, 2.5, [
            (n, dict(size=11, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 85, y + 2.7, 45, 3.3, [
            (sub, dict(size=9, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.TOP)
        add_rect(s, 132, y + 1.2, 20, 3.6, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, 132, y + 1.2, 20, 3.6, [
            (own, dict(size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 83, CY_TOP + 27, 72, 3.4, [
        ("Zeitraum: 11.06. – 30.08.2026  ·  Ergebnis: klare Empfehlung mit TCO-Vergleich",
         dict(size=9.5, italic=True, color=GREY_TXT))],
        anchor=MSO_ANCHOR.MIDDLE)

    # Comparison teaser
    ty = CY_TOP + 34
    add_rect(s, 4, ty, 152, 30, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
             line=LIGHT_GREY, line_w=0.7)
    add_rect(s, 4, ty, 152, 4, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 5, ty, 150, 4, [
        ("Bewertungskriterien und Zwischenergebnis",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    criteria = [
        ("Funktions-Deckung",             "70%",  "80%",  "60%",  "65%"),
        ("Integration STIHL-Landschaft",  "80%",  "50%",  "70%",  "40%"),
        ("TCO 5 Jahre (rel.)",            "100%", "115%", "125%", "95%"),
        ("Change-Aufwand",                "Mittel","Hoch", "Mittel","Niedrig"),
        ("Zukunftssicherheit AI",         "Hoch", "Mittel","Hoch",  "Niedrig"),
    ]
    col_x = [5, 66, 89, 112, 135]
    col_w = [60, 22, 22, 22, 20]
    headers = ["Kriterium", "SNOW SPM", "Planisware", "SAP EPPM", "Sciforma"]
    hy = ty + 4.6
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
        y = ty + 8 + j * 4
        bg = LIGHT_BG if j % 2 == 0 else ROW_ALT
        add_rect(s, 5, y, 150, 3.7, bg)
        add_text(s, 6, y, col_w[0], 3.7, [
            (k, dict(size=10, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        for i, v in enumerate(vals):
            add_text(s, col_x[i + 1], y, col_w[i + 1], 3.7, [
                (v, dict(size=10, color=DARK, align=PP_ALIGN.CENTER))],
                anchor=MSO_ANCHOR.MIDDLE)

slide_alternative_pm()

# ============================================================
# SLIDE 5: Komplexitäts-Reduzierung ProjektControlling (Punkt 4)
# ============================================================
def slide_komplex_reduzierung():
    s = add_slide_titled(
        "Komplexitäts-Reduzierung ProjektControlling",
        "Punkt 4 Folie 4: Was heute Aufwand verursacht – und was SEPM konsolidiert")
    # Left: HEUTE
    add_rect(s, 4, CY_TOP, 74, 64, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, 4, CY_TOP, 74, 4, RED_HL, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 5, CY_TOP, 72, 4, [
        ("HEUTE  –  Komplexitätstreiber",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    hoy = [
        ("Multi-Tool-Landschaft PPM",
         "PIT · MSPO · Excel · SharePoint · Power Apps",
         "~ 2,5 FTE Betriebs- & Integrations-Aufwand"),
        ("Manuelle Daten-Aggregation",
         "Excel-Konsolidierung für Reportings, Steckbriefe",
         "~ 2 Wochen/Monat Reporting-Aufwand"),
        ("Uneinheitliche Datenmodelle",
         "Jede Insel eigener Struktur",
         "~ 1 FTE Datenqualitäts-Arbeit"),
        ("Doppelte Datenpflege",
         "Projekte in PIT + Excel + SharePoint erfasst",
         "~ 0,8 FTE Doppelerfassung"),
        ("Fehlende Traceability",
         "Idee ↔ Portfolio ↔ Projekt ↔ Kosten nicht durchgängig",
         "Governance-Aufwand hoch, Audits schwierig"),
    ]
    for i, (h, sub, imp) in enumerate(hoy):
        y = CY_TOP + 5.5 + i * 11.5
        add_rect(s, 5.5, y, 71, 10.5, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=RED_HL, line_w=0.6)
        add_rect(s, 5.5, y, 0.6, 10.5, RED_HL)
        add_text(s, 7, y + 0.4, 68, 3.0, [
            (h, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 7, y + 3.2, 68, 3.4, [
            (sub, dict(size=9, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 7, y + 6.8, 68, 3.4, [
            (imp, dict(size=9.5, bold=True, color=RED_HL))],
            anchor=MSO_ANCHOR.MIDDLE)
    # Right: MIT SEPM
    add_rect(s, 82, CY_TOP, 74, 64, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, 82, CY_TOP, 74, 4, GREEN, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 83, CY_TOP, 72, 4, [
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
         "Vorhaben-Steckbrief einheitlich",
         "− 0,7 FTE Datenqualitäts-Arbeit"),
        ("Erfassung einmal – Nutzung überall",
         "Demand → Portfolio → Projekt in einem System",
         "− 0,6 FTE Doppelerfassung"),
        ("Durchgängige Traceability",
         "OKR → Portfolio → Projekt → Kosten → Nutzen",
         "Audit-fähig, Compliance-sicher"),
    ]
    for i, (h, sub, imp) in enumerate(mit):
        y = CY_TOP + 5.5 + i * 11.5
        add_rect(s, 83.5, y, 71, 10.5, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=GREEN, line_w=0.6)
        add_rect(s, 83.5, y, 0.6, 10.5, GREEN)
        add_text(s, 85, y + 0.4, 68, 3.0, [
            (h, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 85, y + 3.2, 68, 3.4, [
            (sub, dict(size=9, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 85, y + 6.8, 68, 3.4, [
            (imp, dict(size=9.5, bold=True, color=GREEN))],
            anchor=MSO_ANCHOR.MIDDLE)

slide_komplex_reduzierung()

# ============================================================
# SLIDE 6: Business Value
# ============================================================
def slide_business_value():
    s = add_slide_titled(
        "Business Value Analyse",
        "Nutzen-Kategorien und quantifizierter Steady-State p.a.")
    add_rect(s, 4, CY_TOP, 152, 8, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 6, CY_TOP, 60, 8, [
        ("Gesamt-Nutzen SEPM (Steady State p.a.)",
         dict(size=12.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 66, CY_TOP, 60, 8, [
        (f"~ {TOTAL_VALUE:,} k€ / Jahr".replace(",", "."),
         dict(size=22, bold=True, color=ORANGE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)

    add_text(s, 4, CY_TOP + 10, 152, 2.4, [
        ("Detail – Nutzen-Kategorien",
         dict(size=12.5, bold=True, color=ORANGE))])
    max_val = max(v[2] for v in VALUE_CATEGORIES)
    row_top = CY_TOP + 13
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
        bar_x = 58
        bar_w_max = 78
        bar_w_val = bar_w_max * (val / max_val)
        add_rect(s, bar_x, y + 3.0, bar_w_max, 4, LIGHT_GREY,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, bar_x, y + 3.0, bar_w_val, 4, ORANGE,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, 138, y, 18, row_h, [
            (f"{val:,} k€".replace(",", "."),
             dict(size=13, bold=True, color=NAVY, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)

slide_business_value()

# ============================================================
# SLIDE 7: TCO Detail
# ============================================================
def slide_tco():
    s = add_slide_titled(
        "TCO – Total Cost of Ownership (5 Jahre)",
        "Gesamt-Kosten SEPM: einmalig + laufend, aufgeschlüsselt")
    total_5y = IMPL_ONE + CHANGE_ONE + 5 * (LICENSE_Y_STEADY +
                                             SUPPORT_INT_FTE_STEADY * FTE_RATE + SUPPORT_EXT)
    add_rect(s, 4, CY_TOP, 152, 8, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 6, CY_TOP, 70, 8, [
        ("5-Jahres-TCO SEPM (Base Case)",
         dict(size=13, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 76, CY_TOP, 60, 8, [
        (f"~ {total_5y/1000:.1f} Mio. €",
         dict(size=22, bold=True, color=ORANGE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 136, CY_TOP, 20, 8, [
        (f"(~ {total_5y/5/1000:.2f} M€ p.a. Ø)",
         dict(size=10, italic=True, color=LIGHT_GREY, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)

    left_x = 4; right_x = 82; w = 74
    top = CY_TOP + 10
    add_rect(s, left_x, top, w, 52, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, left_x, top, w, 4, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, left_x + 1, top, w - 2, 4, [
        ("Einmalige Kosten (Y1-Y2)",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    einmalig = [
        ("Implementierungs-Partner (SI)", IMPL_ONE, "SNOW-Zertifiziert"),
        ("Interne FTE (Phase 1+2)", INT_FTE_IMPL * FTE_RATE * 2, f"{INT_FTE_IMPL} FTE × 2 J."),
        ("Change / Schulung", CHANGE_ONE, "Anwender + Multiplikatoren"),
        ("Contingency (10 %)", int(0.1 * (IMPL_ONE + INT_FTE_IMPL * FTE_RATE * 2 + CHANGE_ONE)), "Puffer"),
    ]
    for i, (name, val, note) in enumerate(einmalig):
        y = top + 5.5 + i * 9.5
        add_rect(s, left_x + 1.5, y, w - 3, 8.5, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=NAVY, line_w=0.5)
        add_text(s, left_x + 2.5, y + 0.3, 45, 3.2, [
            (name, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, left_x + 2.5, y + 3.5, 45, 4, [
            (note, dict(size=8.5, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.TOP)
        add_text(s, left_x + w - 20, y, 18, 8.5, [
            (f"{val:,} k€".replace(",", "."),
             dict(size=12.5, bold=True, color=NAVY, align=PP_ALIGN.RIGHT))],
            anchor=MSO_ANCHOR.MIDDLE)
    total_einmal = sum(v[1] for v in einmalig)
    add_text(s, left_x + 1, top + 46, w - 2, 4, [
        (f"Summe einmalig: {total_einmal:,} k€".replace(",", "."),
         dict(size=12, bold=True, color=NAVY, align=PP_ALIGN.RIGHT))],
        anchor=MSO_ANCHOR.MIDDLE)

    add_rect(s, right_x, top, w, 52, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, right_x, top, w, 4, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, right_x + 1, top, w - 2, 4, [
        ("Laufende Kosten p.a. (Steady State)",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    laufend = [
        ("Lizenzen SNOW SPM Pro", LICENSE_Y_STEADY, "~ 350-450 Named User"),
        ("Interne Betriebs-FTE", int(SUPPORT_INT_FTE_STEADY * FTE_RATE),
         f"~ {SUPPORT_INT_FTE_STEADY} FTE (IT + Business)"),
        ("Externer Support / Wartung", SUPPORT_EXT, "Retainer + Ticket-Volumen"),
        ("Weiterentwicklung (Backlog)", 220, "kontinuierlich"),
    ]
    for i, (name, val, note) in enumerate(laufend):
        y = top + 5.5 + i * 9.5
        add_rect(s, right_x + 1.5, y, w - 3, 8.5, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=ORANGE, line_w=0.5)
        add_text(s, right_x + 2.5, y + 0.3, 45, 3.2, [
            (name, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, right_x + 2.5, y + 3.5, 45, 4, [
            (note, dict(size=8.5, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.TOP)
        add_text(s, right_x + w - 20, y, 18, 8.5, [
            (f"{val:,} k€".replace(",", "."),
             dict(size=12.5, bold=True, color=ORANGE, align=PP_ALIGN.RIGHT))],
            anchor=MSO_ANCHOR.MIDDLE)
    total_laufend = sum(v[1] for v in laufend)
    add_text(s, right_x + 1, top + 46, w - 2, 4, [
        (f"Summe p.a.: {total_laufend:,} k€".replace(",", "."),
         dict(size=12, bold=True, color=ORANGE, align=PP_ALIGN.RIGHT))],
        anchor=MSO_ANCHOR.MIDDLE)

slide_tco()

# ============================================================
# SLIDE 8: Support-Kosten Detail
# ============================================================
def slide_support():
    s = add_slide_titled(
        "Support-Kosten – Detailbetrachtung",
        "Interner + externer Support, laufender Betrieb, Weiterentwicklung")
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
    top = CY_TOP
    w = (152 - 4) / 3
    gap = 2
    for i, (title, color, items) in enumerate(cols):
        x = 4 + i * (w + gap)
        add_rect(s, x, top, w, 55, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, x, top, w, 5, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x + 1, top + 0.4, w - 2, 4.4, [
            (title, dict(size=13, bold=True, color=WHITE))],
            anchor=MSO_ANCHOR.MIDDLE)
        total = 0
        for j, (name, note, val) in enumerate(items):
            y = top + 6 + j * 8
            add_rect(s, x + 1, y, w - 2, 7.5, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                     line=color, line_w=0.5)
            add_text(s, x + 2, y + 0.3, w - 20, 3.4, [
                (name, dict(size=10.5, bold=True, color=DARK))],
                anchor=MSO_ANCHOR.MIDDLE)
            add_text(s, x + 2, y + 3.4, w - 20, 4, [
                (note, dict(size=9, italic=True, color=GREY_TXT))],
                anchor=MSO_ANCHOR.MIDDLE)
            add_text(s, x + w - 18, y + 0.3, 16, 7, [
                (f"{val} k€",
                 dict(size=12, bold=True, color=color, align=PP_ALIGN.RIGHT))],
                anchor=MSO_ANCHOR.MIDDLE)
            total += val
        add_rect(s, x + 1, top + 49, w - 2, 4, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x + 1, top + 49, w - 2, 4, [
            (f"Summe: {total} k€ / Jahr",
             dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    tb_y = top + 57
    add_rect(s, 4, tb_y, 152, 6, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 6, tb_y, 90, 6, [
        ("Gesamt Support & Weiterentwicklung p.a.",
         dict(size=12.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 96, tb_y, 60, 6, [
        ("~ 860 k€",
         dict(size=20, bold=True, color=ORANGE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)

slide_support()

# ============================================================
# SLIDE 9: Konsolidierung Tools
# ============================================================
def slide_konsolidierung_tools():
    s = add_slide_titled(
        "Konsolidierungseinsparungen – Tool-Ablöse",
        "Wegfall bisheriger Tools durch SNOW SPM Konsolidierung")
    add_rect(s, 4, CY_TOP, 152, 8, GREEN, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 6, CY_TOP, 90, 8, [
        ("Gesamt Tool-Ablöse-Ersparnis (Steady State p.a.)",
         dict(size=12.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 96, CY_TOP, 60, 8, [
        (f"~ {TOTAL_TOOL_SAVINGS} k€ / Jahr",
         dict(size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 4, CY_TOP + 10, 152, 2.4, [
        ("Detail – Tool-für-Tool-Ablöse",
         dict(size=12.5, bold=True, color=ORANGE))])
    max_val = max(x[2] for x in CONSOL_TOOLS)
    row_top = CY_TOP + 13
    row_h = 8.5
    for i, (tool, cat, saving) in enumerate(CONSOL_TOOLS):
        y = row_top + i * (row_h + 0.4)
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        add_rect(s, 4, y, 152, row_h, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, 4, y, 0.7, row_h, GREEN)
        add_text(s, 6, y + 0.3, 40, row_h - 0.6, [
            (tool, dict(size=11.5, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 6, y + row_h - 3.0, 40, 2.5, [
            (cat, dict(size=8.5, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.TOP)
        bar_x = 48; bar_w_max = 88
        bar_w_val = bar_w_max * (saving / max_val)
        add_rect(s, bar_x, y + 3, bar_w_max, 3, LIGHT_GREY,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, bar_x, y + 3, bar_w_val, 3, GREEN,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, 138, y, 18, row_h, [
            (f"− {saving} k€",
             dict(size=13, bold=True, color=GREEN, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)

slide_konsolidierung_tools()

# ============================================================
# SLIDE 10: Konsolidierung Schnittstellen
# ============================================================
def slide_konsolidierung_interfaces():
    s = add_slide_titled(
        "Konsolidierungseinsparungen – Schnittstellen",
        "Wegfall komplexer Integrationen durch eine Plattform")
    add_rect(s, 4, CY_TOP, 152, 8, GREEN, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 6, CY_TOP, 90, 8, [
        ("Ersparnis Schnittstellen & Integration p.a.",
         dict(size=12.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 96, CY_TOP, 60, 8, [
        (f"~ {INT_SAVINGS} k€ / Jahr",
         dict(size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, 4, CY_TOP + 10, 74, 54, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, 4, CY_TOP + 10, 74, 4, RED_HL, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 5, CY_TOP + 10, 72, 4, [
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
        y = CY_TOP + 15.5 + i * 6.5
        add_rect(s, 6, y, 70, 5.4, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=RED_HL, line_w=0.5)
        add_rect(s, 6, y, 0.5, 5.4, RED_HL)
        add_text(s, 8, y, 68, 5.4, [
            (txt, dict(size=10, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, 82, CY_TOP + 10, 74, 54, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, 82, CY_TOP + 10, 74, 4, GREEN, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 83, CY_TOP + 10, 72, 4, [
        ("MIT SEPM – Native SNOW-Integration",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    mit_ints = [
        ("SAP FI Konnektor", "Standard, Integration Hub"),
        ("SAP HCM Konnektor", "Standard, Integration Hub"),
        ("Interne Datenmodelle", "keine Excel-Export-Ketten mehr"),
        ("OKR-Modul nativ", "kein Workpath-Bridge nötig"),
        ("App Engine (Low-Code)", "Custom-Data auf gleicher Plattform"),
        ("EV-Prozess nativ", "kein Lucom-Bridge nötig"),
        ("Live-Reporting", "Excel-Konsolidierung entfällt"),
    ]
    for i, (h, sub) in enumerate(mit_ints):
        y = CY_TOP + 15.5 + i * 6.5
        add_rect(s, 84, y, 70, 5.4, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=GREEN, line_w=0.5)
        add_rect(s, 84, y, 0.5, 5.4, GREEN)
        add_text(s, 86, y + 0.1, 30, 5.2, [
            (h, dict(size=10, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 116, y + 0.1, 38, 5.2, [
            (sub, dict(size=9, italic=True, color=GREEN))], anchor=MSO_ANCHOR.MIDDLE)

slide_konsolidierung_interfaces()

# ============================================================
# SLIDE 11: Flankierende Prozesse
# ============================================================
def slide_flanking():
    s = add_slide_titled(
        "Flankierende Prozess-Potentiale",
        "Zusätzliche Einsparungen durch SPM-nahe Prozess-Digitalisierung")
    add_rect(s, 4, CY_TOP, 152, 8, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 6, CY_TOP, 100, 8, [
        ("Zusätzliches Einspar-Potenzial (SPM-flankierend, p.a.)",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 106, CY_TOP, 50, 8, [
        (f"~ {TOTAL_FLANKING:,} k€ / Jahr".replace(",", "."),
         dict(size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    max_val = max(x[2] for x in FLANKING)
    row_top = CY_TOP + 10
    row_h = 10
    for i, (name, desc, val) in enumerate(FLANKING):
        y = row_top + i * (row_h + 0.6)
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        add_rect(s, 4, y, 152, row_h, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, 4, y, 0.7, row_h, ORANGE)
        add_oval(s, 6, y + 3, 4, 4, ORANGE)
        add_text(s, 6, y + 3, 4, 4, [
            (str(i + 1), dict(size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 12, y + 0.4, 100, 3.8, [
            (name, dict(size=11, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 12, y + 4.3, 100, 5, [
            (desc, dict(size=9.5, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.TOP)
        bar_x = 118; bar_w_max = 20
        bar_w_val = bar_w_max * (val / max_val)
        add_rect(s, bar_x, y + 4, bar_w_max, 3, LIGHT_GREY,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, bar_x, y + 4, bar_w_val, 3, ORANGE,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, 140, y, 15, row_h, [
            (f"{val} k€",
             dict(size=13, bold=True, color=NAVY, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)

slide_flanking()

# ============================================================
# SLIDE 12: Szenarien-Übersicht
# ============================================================
def slide_szenarien_uebersicht():
    s = add_slide_titled(
        "Einführungs-Szenarien im Vergleich",
        "Drei Szenarien zur Umsetzung SEPM – Kosten, Nutzen, Risiko")
    top = CY_TOP
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
          ("Ramp Up", "MVP live in 12 Mo"),
          ("Payback", "~ 3 Jahre auf Teil-Nutzen"),
          ("Business Value", "Teilnutzen früh"),
          ("Risiko", "Mittel — Nachreichung Module")],
         False),
    ]
    for i, (name, months, color, badge, items, recommend) in enumerate(scenarios):
        x = 4 + i * (w + gap)
        add_rect(s, x, top, w, 63, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=color, line_w=2 if recommend else 1)
        add_rect(s, x, top, w, 8, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        if recommend:
            add_rect(s, x + w - 20, top, 20, 8, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            add_text(s, x + w - 20, top, 20, 8, [
                ("★ EMPFOHLEN",
                 dict(size=9, bold=True, color=ORANGE, align=PP_ALIGN.CENTER))],
                anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 1, top + 0.5, w - 22 if recommend else w - 2, 7, [
            (name, dict(size=13, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, x + 2, top + 9, w - 4, 3.6, LIGHT_BG,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x + 2, top + 9, w - 4, 3.6, [
            (badge, dict(size=10, bold=True, color=color, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 2, top + 13.5, w - 4, 2.4, [
            (f"Dauer: {months} Monate",
             dict(size=9.5, italic=True, color=GREY_TXT, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        max_m = 36
        bar_w = (w - 4) * (months / max_m)
        add_rect(s, x + 2, top + 16, w - 4, 2, LIGHT_GREY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, x + 2, top + 16, bar_w, 2, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        for j, (k, v) in enumerate(items):
            y = top + 20 + j * 8
            add_rect(s, x + 2, y, w - 4, 7, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            add_text(s, x + 3, y + 0.2, w - 6, 2.8, [
                (k, dict(size=9, bold=True, color=GREY_TXT))], anchor=MSO_ANCHOR.MIDDLE)
            add_text(s, x + 3, y + 3, w - 6, 3.8, [
                (v, dict(size=10, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)

slide_szenarien_uebersicht()

# ============================================================
# SLIDE 13: Szenario A Detail
# ============================================================
def slide_szenario_A():
    s = add_slide_titled(
        "Szenario A – Full Scope 24 Monate (Empfohlen)",
        "Timeline · Investitionsprofil · Nutzen-Kurve · Break-even")
    add_text(s, 4, CY_TOP, 152, 2.4, [
        ("Umsetzungs-Timeline",
         dict(size=12.5, bold=True, color=ORANGE))])
    axis_y = CY_TOP + 3
    n_m = 24
    left_x = 30
    right_x = 156
    m_w = (right_x - left_x) / n_m
    def mx(m): return left_x + m * m_w
    add_rect(s, left_x, axis_y, 12 * m_w, 3.4, ORANGE_LT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, left_x, axis_y, 12 * m_w, 3.4, [
        ("Jahr 1", dict(size=11, bold=True, color=DARK, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, left_x + 12 * m_w, axis_y, 12 * m_w, 3.4, ORANGE2,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, left_x + 12 * m_w, axis_y, 12 * m_w, 3.4, [
        ("Jahr 2", dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    phases = [
        ("Konzeption",            NAVY,     0,  3),
        ("Impl. Phase 1 (Demand, Portfolio)", ORANGE, 3, 12),
        ("Go-Live Phase 1",       GREEN,    12, 13),
        ("Impl. Phase 2 (PM, Res., Fin.)", ORANGE2, 12, 23),
        ("Go-Live Phase 2",       GREEN,    23, 24),
        ("Change / Rollout",      BLUE_HD,  3,  24),
    ]
    lane_top = axis_y + 4
    lane_h = 4.5
    lane_gap = 0.4
    for i, (name, color, ms, me) in enumerate(phases):
        y = lane_top + i * (lane_h + lane_gap)
        add_text(s, 4, y, 25, lane_h, [
            (name, dict(size=9.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, left_x, y, n_m * m_w, lane_h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        bar_x = mx(ms)
        bar_w = (me - ms) * m_w - 0.2
        add_rect(s, bar_x + 0.1, y + 0.3, bar_w, lane_h - 0.6, color,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    curve_top = lane_top + len(phases) * (lane_h + lane_gap) + 3
    add_text(s, 4, curve_top, 152, 2.4, [
        ("Investitions- und Nutzenprofil (schematisch)",
         dict(size=12, bold=True, color=ORANGE))])
    ct = curve_top + 3
    ch = 20
    add_rect(s, 4, ct, 152, ch, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    year_w = 152 / 5
    for yi in range(5):
        x = 4 + yi * year_w
        col = WHITE if yi % 2 == 0 else GREY_BG
        add_rect(s, x, ct, year_w - 0.05, ch, col)
        add_text(s, x, ct + ch - 2.5, year_w, 2.4, [
            (f"Jahr {yi + 1}",
             dict(size=9, color=GREY_TXT, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    inv_vals = [2.0, 2.2, 1.2, 1.0, 1.0]
    ben_vals = [0.3, 1.5, 3.0, 4.0, 4.3]
    max_v = 4.5
    for yi in range(5):
        cx = 4 + yi * year_w + year_w / 2
        h_inv = (inv_vals[yi] / max_v) * (ch - 4.5)
        add_rect(s, cx - 3.5, ct + ch - 3 - h_inv, 3, h_inv, RED_HL,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, cx - 6, ct + ch - 2 - h_inv - 2, 8, 1.8, [
            (f"{inv_vals[yi]:.1f}", dict(size=8, bold=True, color=RED_HL, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        h_ben = (ben_vals[yi] / max_v) * (ch - 4.5)
        add_rect(s, cx + 0.5, ct + ch - 3 - h_ben, 3, h_ben, GREEN,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, cx - 2, ct + ch - 2 - h_ben - 2, 8, 1.8, [
            (f"{ben_vals[yi]:.1f}", dict(size=8, bold=True, color=GREEN, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 4 + 3.5 * year_w, ct + 0.6, year_w, 2.4, [
        ("★ Break-even", dict(size=10, bold=True, color=ORANGE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)

slide_szenario_A()

# ============================================================
# SLIDE 14: Szenario B Detail (Vergleich A vs B)
# ============================================================
def slide_szenario_B():
    s = add_slide_titled(
        "Szenario B – Extended Ramp-Up 36 Monate",
        "Vergleich zu Szenario A: konservativer, weniger Change-Druck")
    left_x = 4; right_x = 82; w = 74
    top = CY_TOP
    for k, (title, color, months, txt) in enumerate([
        ("Szenario A – 24 Monate", ORANGE, 24,
         ["Standard-Tempo, 2 Jahre",
          "Höherer Change-Druck",
          "Nutzen ab Jahr 3 vollständig",
          "Payback ~ 3,5 Jahre",
          "Empfohlen bei ausreichenden Ressourcen"]),
        ("Szenario B – 36 Monate", BLUE_HD, 36,
         ["Verteiltes Tempo, 3 Jahre",
          "Weniger Change-Druck",
          "Nutzen ab Jahr 4 vollständig",
          "Payback ~ 4,5 Jahre",
          "Empfohlen bei knappen Ressourcen"]),
    ]):
        x = left_x if k == 0 else right_x
        add_rect(s, x, top, w, 32, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=color, line_w=1.5)
        add_rect(s, x, top, w, 5, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x, top, w, 5, [
            (title, dict(size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        max_m = 36
        add_text(s, x + 2, top + 6, w - 4, 2.4, [
            (f"Dauer: {months} Monate",
             dict(size=10, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, x + 2, top + 9, w - 4, 2, LIGHT_GREY,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, x + 2, top + 9, (w - 4) * months / max_m, 2, color,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        for j, tx in enumerate(txt):
            y = top + 13 + j * 3.8
            add_oval(s, x + 2, y + 1.2, 1.2, 1.2, color)
            add_text(s, x + 4.3, y, w - 6, 3.6, [
                (tx, dict(size=10, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    tt = top + 35
    add_text(s, 4, tt, 152, 2.4, [
        ("Kosten- & Nutzen-Delta",
         dict(size=12.5, bold=True, color=ORANGE))])
    rows = [
        ("Einmalige Kosten",     "3,0 M€",  "2,7 M€",   "−0,3 M€"),
        ("Laufend p.a. steady",  "0,85 M€", "0,80 M€",  "−0,05 M€"),
        ("5-Jahres-TCO",         "7,0 M€",  "6,7 M€",   "−0,3 M€"),
        ("Nutzen Y3 (kumul.)",   "3,5 M€",  "1,8 M€",   "−1,7 M€"),
        ("Nutzen Y5 (kumul.)",   "16 M€",   "12 M€",    "−4 M€"),
        ("Payback",              "~ 3,5 J.","~ 4,5 J.", "+1 J."),
    ]
    hdr_y = tt + 3
    add_rect(s, 4, hdr_y, 152, 3.4, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    for i, h in enumerate(["Kennzahl", "Szenario A", "Szenario B", "Δ B vs A"]):
        add_text(s, [4, 60, 92, 124][i], hdr_y, [55, 32, 32, 32][i], 3.4, [
            (h, dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    for j, (k, va, vb, d) in enumerate(rows):
        y = hdr_y + 3.6 + j * 3.4
        bg = LIGHT_BG if j % 2 == 0 else ROW_ALT
        add_rect(s, 4, y, 152, 3.2, bg)
        add_text(s, 6, y, 52, 3.2, [
            (k, dict(size=10.5, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 60, y, 32, 3.2, [
            (va, dict(size=10.5, color=ORANGE, bold=True, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 92, y, 32, 3.2, [
            (vb, dict(size=10.5, color=BLUE_HD, bold=True, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 124, y, 32, 3.2, [
            (d, dict(size=10.5, color=RED_HL if "−" in d and "M€" in d else DARK,
                     bold=True, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)

slide_szenario_B()

# ============================================================
# SLIDE 15: Budgetplanung (Punkt 6)
# ============================================================
def slide_budget():
    s = add_slide_titled(
        "Budgetplanung SEPM",
        "Punkt 6 Folie 4: Investitionsprofil für Jahresplanung 2027 + 2028")
    years = [2027, 2028, 2029, 2030, 2031]
    invs = [(2200, "Phase 1 Impl. + Lizenz Y1"),
            (2200, "Phase 2 Impl. + Lizenz Y2"),
            (1200, "Steady State + Weiterentw."),
            (1000, "Steady State"),
            (1000, "Steady State")]
    total = sum(v[0] for v in invs)
    add_rect(s, 4, CY_TOP, 152, 8, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 6, CY_TOP, 90, 8, [
        ("Gesamt-Investition SEPM 2027–2031 (Base Case)",
         dict(size=12.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 96, CY_TOP, 60, 8, [
        (f"~ {total/1000:.1f} Mio. €",
         dict(size=22, bold=True, color=ORANGE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    chart_top = CY_TOP + 12
    chart_h = 38
    chart_x = 8; chart_w = 148
    n = len(years)
    slot_w = chart_w / n
    max_v = max(v[0] for v in invs)
    for i, (yr, (val, label)) in enumerate(zip(years, invs)):
        cx = chart_x + i * slot_w + slot_w / 2
        h = (val / max_v) * (chart_h - 8)
        color = ORANGE if i < 2 else (ORANGE2 if i < 3 else ORANGE_LT)
        add_rect(s, cx - 8, chart_top + chart_h - h - 4, 16, h, color,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, cx - 10, chart_top + chart_h - h - 6.5, 20, 2.4, [
            (f"{val/1000:.1f} M€",
             dict(size=10.5, bold=True, color=DARK, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, cx - 10, chart_top + chart_h - 3.5, 20, 3, [
            (str(yr), dict(size=11, bold=True, color=DARK, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, cx - 20, chart_top + chart_h - 0.5, 40, 3, [
            (label, dict(size=8, italic=True, color=GREY_TXT, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, chart_x, chart_top + chart_h - 4, chart_w, 0.15, DARK)
    add_text(s, 4, CY_TOP + 54, 152, 2.4, [
        ("Milestones Jahresplanung",
         dict(size=12, bold=True, color=ORANGE))])
    ms = [
        ("Q3 2026: Jahresplanung", "Investitions-Beschluss 2027 & Rahmen 2028"),
        ("Q3 2027: Jahresplanung", "Freigabe Phase 2"),
        ("Q3 2028: Review", "Übergang in Steady State"),
    ]
    for i, (t, sub) in enumerate(ms):
        x = 4 + i * 50
        add_rect(s, x, CY_TOP + 57, 48, 10, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=NAVY, line_w=0.8)
        add_diamond(s, x + 1, CY_TOP + 59.5, 3, 3, RED_HL)
        add_text(s, x + 5, CY_TOP + 57.4, 42, 3.8, [
            (t, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 5, CY_TOP + 60.8, 42, 6, [
            (sub, dict(size=9, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.MIDDLE)

slide_budget()

# ============================================================
# SLIDE 16: Risiken
# ============================================================
def slide_risk():
    s = add_slide_titled(
        "Risiken der ServiceNow-Plattform-Abhängigkeit",
        "Systematische Bewertung + Mitigationen")
    add_text(s, 4, CY_TOP, 152, 2.4, [
        ("Risiko-Matrix",
         dict(size=12.5, bold=True, color=ORANGE))])
    hdr_y = CY_TOP + 3
    add_rect(s, 4, hdr_y, 152, 3.4, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    hdrs = [("Risiko", 40), ("Schwere", 18), ("Mitigation", 94)]
    xh = 4
    for hd, w in hdrs:
        add_text(s, xh + 2, hdr_y, w - 2, 3.4, [
            (hd, dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.LEFT))],
            anchor=MSO_ANCHOR.MIDDLE)
        xh += w
    for i, (risk, sev, mit) in enumerate(RISKS):
        y = hdr_y + 3.6 + i * 7
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        add_rect(s, 4, y, 152, 6.6, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, 6, y + 0.3, 38, 6, [
            (risk, dict(size=10.5, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        for k in range(5):
            filled = k < sev
            color = (RED_HL if sev >= 4 else (YELLOW if sev == 3 else GREEN)) if filled else LIGHT_GREY
            add_oval(s, 46 + k * 3, y + 2.7, 2, 2, color)
        add_text(s, 64, y + 0.3, 90, 6, [
            (mit, dict(size=10, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    sy = hdr_y + 3.6 + len(RISKS) * 7 + 1
    add_rect(s, 4, sy, 152, 10, GREEN_LT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, 4, sy, 0.7, 10, GREEN)
    add_text(s, 6, sy, 148, 3.4, [
        ("Fazit Risiko-Bewertung",
         dict(size=12, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 6, sy + 3, 148, 6, [
        ("Kein Risiko ist ein Show-Stopper. Alle Risiken sind mit Standard-Verträgen "
         "(Preisgarantie, DSGVO-ADV, SLA) und Standard-First-Governance beherrschbar. "
         "Größter Hebel gegen Vendor Lock-in: konsequente Standard-Nutzung ohne tiefe Customizings.",
         dict(size=10, color=DARK))], anchor=MSO_ANCHOR.TOP)

slide_risk()

# ============================================================
# SLIDE 17: SNOW Vertragshistorie
# ============================================================
def slide_snow_history():
    s = add_slide_titled(
        "ServiceNow bei STIHL – Vertragshistorie 2014–2026",
        "Bewährter Partner: 12 Jahre gemeinsame Historie, kontrolliertes Wachstum")
    add_text(s, 4, CY_TOP, 152, 2.4, [
        ("Vertragsstufen und jährliche Lizenzausgaben (indikativ)",
         dict(size=12.5, bold=True, color=ORANGE))])
    ct = CY_TOP + 3
    ch = 35
    add_rect(s, 4, ct, 152, ch, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    def xp(yr): return 8 + (yr - 2014) / (2026 - 2014) * 144
    add_rect(s, 8, ct + ch - 5, 148, 0.15, DARK)
    max_val = max(x[3] for x in SNOW_HISTORY)
    for yr, event, note, val in SNOW_HISTORY:
        x = xp(yr)
        h = (val / max_val) * (ch - 10)
        y = ct + ch - 5 - h
        add_rect(s, x - 1.5, y, 3, h, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_oval(s, x - 1.2, y - 1.2, 2.4, 2.4, NAVY)
        add_text(s, x - 10, y - 4.5, 20, 2.4, [
            (f"{val} k€",
             dict(size=9, bold=True, color=NAVY, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x - 10, ct + ch - 4, 20, 3, [
            (str(yr), dict(size=10, bold=True, color=DARK, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 4, ct + ch + 1, 152, 2.4, [
        ("Meilensteine der Zusammenarbeit",
         dict(size=12, bold=True, color=ORANGE))])
    ev_top = ct + ch + 4
    for i, (yr, event, note, val) in enumerate(SNOW_HISTORY):
        col = i % 3
        row = i // 3
        x = 4 + col * 51
        y = ev_top + row * 7
        add_rect(s, x, y, 49, 6, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=ORANGE, line_w=0.5)
        add_rect(s, x, y, 0.5, 6, ORANGE)
        add_text(s, x + 1.5, y + 0.2, 46, 2.8, [
            (f"{yr} · {event}",
             dict(size=10, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 1.5, y + 3, 46, 3, [
            (note, dict(size=9, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)

slide_snow_history()

# ============================================================
# SLIDE 18: ROI & Empfehlung
# ============================================================
def slide_roi_recommendation():
    s = add_slide_titled(
        "ROI-Kalkulation & Empfehlung",
        "Gesamt-Wirtschaftlichkeit und Beschlussvorschlag")
    kpis = [
        ("NPV 5 Jahre",   "~ +8,5 M€", "Base Case, WACC 6 %", GREEN),
        ("IRR",           "~ 32 %",    "über Investitionshürde", GREEN),
        ("Payback",       "~ 3,5 J.",  "nach Kickoff", ORANGE),
        ("TCO 5 Jahre",   "~ 7,0 M€",  "einmalig + laufend",  NAVY),
    ]
    top = CY_TOP
    w = (152 - 3 * 1.5) / 4
    for i, (label, big, sub, col) in enumerate(kpis):
        x = 4 + i * (w + 1.5)
        add_rect(s, x, top, w, 15, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=col, line_w=1.5)
        add_rect(s, x, top, w, 3.4, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x, top, w, 3.4, [
            (label, dict(size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x, top + 4, w, 6, [
            (big, dict(size=22, bold=True, color=col, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x, top + 11, w, 3.5, [
            (sub, dict(size=9, italic=True, color=GREY_TXT, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    tt_y = top + 18
    add_text(s, 4, tt_y, 152, 2.4, [
        ("Kosten-Nutzen-Übersicht (5 Jahre kumuliert, Base Case)",
         dict(size=12.5, bold=True, color=ORANGE))])
    rows = [
        ("Kosten (TCO)",      "− 7,0 M€",   RED_HL),
        ("Konsolidierung Tools (5 J.)", "+ 4,7 M€", GREEN),
        ("Konsolidierung Schnittstellen", "+ 1,1 M€", GREEN),
        ("Business Value",    "+ 11,6 M€",  GREEN),
        ("Flankierende Prozesse", "+ 5,3 M€", GREEN),
        ("Netto-Effekt 5 Jahre", "+ 15,7 M€", DARK),
    ]
    tt = tt_y + 3
    for i, (name, val, color) in enumerate(rows):
        y = tt + i * 4.6
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        is_total = (i == len(rows) - 1)
        add_rect(s, 4, y, 152, 4.2, bg if not is_total else DARK,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, 6, y + 0.2, 90, 3.8, [
            (name, dict(size=11, bold=is_total,
                        color=DARK if not is_total else WHITE))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 100, y + 0.2, 52, 3.8, [
            (val, dict(size=13 if is_total else 12, bold=True,
                       color=color if not is_total else ORANGE, align=PP_ALIGN.RIGHT))],
            anchor=MSO_ANCHOR.MIDDLE)
    rec_y = tt + len(rows) * 4.6 + 2
    add_rect(s, 4, rec_y, 152, 10, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, 6, rec_y, 40, 10, [
        ("★ EMPFEHLUNG", dict(size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 46, rec_y + 0.5, 108, 9, [
        ("Beschluss zur Umsetzung von Szenario A (Full Scope 24 Monate) und "
         "Freigabe Detail-Planung / RFP-Prozess bis Q3 2026 empfohlen. "
         "NPV +8,5 M€, IRR 32 %, Payback ~3,5 J. — deutlich über Investitions-Hürde.",
         dict(size=11, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)

slide_roi_recommendation()

# ============================================================
# SLIDE 19: Vielen Dank
# ============================================================
def slide_thanks():
    s = prs.slides.add_slide(LAYOUT_CONCL)
    for ph in s.placeholders:
        tf = ph.text_frame
        tf.text = "Vielen Dank"
        break

slide_thanks()

prs.save(DST_PPTX)
print(f"Saved PPTX: {DST_PPTX}  ({len(prs.slides)} slides)")
print(f"Using STIHL template with {len(prs.slide_masters)} master(s), "
      f"{len(prs.slide_masters[0].slide_layouts)} layouts.")
