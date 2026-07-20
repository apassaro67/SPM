"""Business Case SEPM v3 — strukturiert nach Vorstandsprotokoll (01.07.2026).

Neue Struktur:
- Cover
- Inhaltsverzeichnis (NEU)
- Vorstandsprotokoll-Uebersicht (NEU): 6 Punkte + Farbstatus
- Vorgehensweise & Zeitplan
- Kapitel A: Klaerung Antraege (Punkt 2 out-of-scope, Punkt 3 Prozesse)
- Kapitel B: Business Case + Komplexitaets-Reduzierung (Punkt 4)
- Kapitel C: Alternativen & Umsetzung (Szenarien, Budget)
- Kapitel D: Risiken Plattform-Abhaengigkeit (Punkt 6)
- Kapitel E: Empfehlung
- Vielen Dank
"""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

SRC_TEMPLATE = "/home/user/SPM/2026_07_01_Business_Case_SEPM_STIHL_v2.pptx"
DST_PPTX = "/tmp/2026_07_01_Business_Case_SEPM_v3.pptx"

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
RED_LT   = RGBColor(0xFB, 0xE7, 0xE7)
YELLOW   = RGBColor(0xF5, 0xC2, 0x14)
YELLOW_LT= RGBColor(0xFF, 0xF6, 0xD5)
LIGHT_GREY = RGBColor(0xD1, 0xD5, 0xDB)
GREY_BG  = RGBColor(0xE9, 0xED, 0xF1)

INCH = 914400
CONTENT_TOP  = 1.05
CONTENT_BOT  = 7.05
CONTENT_LEFT = 0.33
CONTENT_RIGHT = 13.00
CONTENT_W = CONTENT_RIGHT - CONTENT_LEFT
CONTENT_H = CONTENT_BOT - CONTENT_TOP

prs = Presentation(SRC_TEMPLATE)
sldIdLst = prs.slides._sldIdLst
for sld in list(sldIdLst):
    rId = sld.get(qn("r:id"))
    sldIdLst.remove(sld)
    try:
        prs.part.drop_rel(rId)
    except Exception:
        pass

LAYOUT_COVER = prs.slide_masters[0].slide_layouts[2]
LAYOUT_TITLE = prs.slide_masters[0].slide_layouts[4]
LAYOUT_CONCL = prs.slide_masters[0].slide_layouts[18]

def add_rect(slide, left_in, top_in, w_in, h_in, fill, line=None,
             shape=MSO_SHAPE.RECTANGLE, line_w=1.0):
    s = slide.shapes.add_shape(shape, int(left_in*INCH), int(top_in*INCH),
                               int(w_in*INCH), int(h_in*INCH))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line; s.line.width = Pt(line_w)
    s.shadow.inherit = False
    return s

def add_oval(slide, l, t, w, h, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, int(l*INCH), int(t*INCH),
                               int(w*INCH), int(h*INCH))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    s.line.fill.background()
    return s

def add_diamond(slide, l, t, w, h, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.DIAMOND, int(l*INCH), int(t*INCH),
                               int(w*INCH), int(h*INCH))
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
        run.font.name = "Calibri"
        run.font.size = Pt(st.get("size", default_size))
        run.font.bold = st.get("bold", False)
        run.font.italic = st.get("italic", False)
        run.font.color.rgb = st.get("color", default_color)

def add_text(slide, l, t, w, h, paragraphs, *, anchor=MSO_ANCHOR.MIDDLE, **kw):
    tb = slide.shapes.add_textbox(int(l*INCH), int(t*INCH), int(w*INCH), int(h*INCH))
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
    s = prs.slides.add_slide(LAYOUT_TITLE)
    set_placeholder_text(s, 0, title)
    return s

# Assumptions (unchanged from v2)
LICENSE_Y1=720; LICENSE_Y_STEADY=850; IMPL_ONE=1300
INT_FTE_IMPL=4; CHANGE_ONE=250; SUPPORT_INT_FTE_STEADY=3.0
SUPPORT_EXT=180; FTE_RATE=130
CONSOL_TOOLS = [
    ("PIT / PLANTA","Projekt-/Ressourcen-Mgmt",380),
    ("MS Project Online","Klass. Projekt-Mgmt",120),
    ("Workpath","Strategie / OKR",150),
    ("Power Apps","Custom-Formulare",90),
    ("SharePoint-Listen","Demand-Listen",70),
    ("Lucom (EV, teilweise)","EV-Prozess",140),
]
TOTAL_TOOL_SAVINGS = sum(x[2] for x in CONSOL_TOOLS)
INT_SAVINGS = 220
VALUE_CATEGORIES = [
    ("Effizienz PM/PPM","Weniger manuelle Excels/Berichte",820),
    ("Time-to-Market","Schnellere Idee->Projekt-Start",480),
    ("Governance / Rework","Weniger Doppelarbeit",360),
    ("Ressourcen-Optimierung","Skill-Match, Auslastungs-Steuerung",520),
    ("Compliance / Audit","Nachvollziehbarkeit",140),
]
TOTAL_VALUE = sum(x[2] for x in VALUE_CATEGORIES)
FLANKING = [
    ("Massnahmen-Management Zielkostenmgmt Produktentwicklung","Zentraler Massnahmen-Backlog",380),
    ("Innovations- / Vorhaben-Trichter (LH/CRD -> Steckbrief)","Durchgaengige Pipeline",220),
    ("Application Portfolio Mgmt (APM)","Lifecycle-Management",160),
    ("Benefits Realization Tracking","Ist-/Soll-Nutzenabgleich",120),
    ("Compliance / EV-Prozess-Digitalisierung","EV-Workflow in SNOW",180),
]
TOTAL_FLANKING = sum(x[2] for x in FLANKING)
SNOW_HISTORY = [
    (2014,"Erstvertrag ITSM","Standard Package",180),
    (2016,"Erweiterung Global","Ausrollung inkl. China",260),
    (2018,"Enterprise-Vertrag","Volumen-Diskont, HR-SD add-on",340),
    (2020,"ITOM + CMDB","Discovery, Event Mgmt",410),
    (2022,"Contract Renewal 3J","Preisanpassung +5% p.a.",490),
    (2024,"Now Assist Preview","GenAI eingefuehrt",560),
    (2026,"SPM Pro (geplant)","Neuer Modul-Baustein",820),
]
RISKS = [
    ("Vendor Lock-in",4,"Standard-Konfiguration, Migrations-Klauseln, Marktscreening"),
    ("Preis-Eskalation",3,"Mehrjahresvertrag mit Preisgarantie, Volumen-Rabatte"),
    ("Ausfall/Betrieb",2,"SLA 99,9% + DC Frankfurt/Duesseldorf, DSGVO"),
    ("Roadmap-Divergenz",3,"Standard first, halbjaehrl. Review"),
    ("Datenhoheit / DSGVO",2,"DE-Hosting, ADV-Vertrag, verschluesselt"),
    ("Team-Kompetenzaufbau",3,"Ausbau ITSM-Team +1,5-2 FTE, Ausbildung"),
]

# ============================================================
# COVER
# ============================================================
def slide_cover():
    s = prs.slides.add_slide(LAYOUT_COVER)
    set_placeholder_text(s, 0, "Business Case SEPM")
    set_placeholder_text(s, 13, "Vorgehensweise nach Vorstandsprotokoll 01.07.2026")
    set_placeholder_text(s, 14, "Alex Passaro · Torsten Zahn")
    set_placeholder_text(s, 15, "Stand: 15.09.2026")
slide_cover()

# ============================================================
# 2 · INHALTSVERZEICHNIS
# ============================================================
def slide_toc():
    s = add_slide_stihl("Inhaltsverzeichnis")
    chapters = [
        ("A", "Klaerung Vorstands-Antraege",
         [("2", "Massnahmen-Mgmt Hype: Begruendung Out-of-Scope", "5", RED_HL),
          ("3", "SNOW-Prozessabdeckung + App-Engine-Strategie",   "6", GREEN)]),
        ("B", "Business Case & Komplexitaets-Reduzierung  (Punkt 4)",
         [("7", "Business Case Methodik",         "7", NAVY),
          ("8", "Business Value Analyse",         "8", NAVY),
          ("9", "TCO 5 Jahre",                    "9", NAVY),
          ("10","Support-Kosten Detail",          "10", NAVY),
          ("11","Konsolidierung Tools",           "11", NAVY),
          ("12","Konsolidierung Schnittstellen",  "12", NAVY),
          ("13","Flankierende Prozesse",          "13", NAVY),
          ("14","Komplexitaets-Reduzierung PC",   "14", NAVY)]),
        ("C", "Alternativen & Umsetzung",
         [("15","Alternative PM-Plattform (Hedge)","15", NAVY),
          ("16","Einfuehrungs-Szenarien A / B / C","16", NAVY),
          ("17","Szenario A Detail",               "17", NAVY),
          ("18","Budgetplanung Jahresplanung 2027","18", NAVY)]),
        ("D", "Risiken Plattform-Abhaengigkeit  (Punkt 6)",
         [("19","Risiko-Matrix ServiceNow",             "19", GREEN),
          ("20","Vertragshistorie SNOW + Referenzkunden","20", GREEN),
          ("21","Lockin, EAM Plattformstrategie & AI",   "21", GREEN)]),
        ("E", "Empfehlung",
         [("22","ROI & Empfehlung",  "22", ORANGE)]),
    ]
    # Two columns of chapters
    top = CONTENT_TOP
    left_w = (CONTENT_W - 0.30) / 2
    right_x = CONTENT_LEFT + left_w + 0.30
    y_col = [top, top]
    col_x = [CONTENT_LEFT, right_x]
    for i, (letter, title, items) in enumerate(chapters):
        col = 0 if i < 3 else 1
        x = col_x[col]
        y = y_col[col]
        # Chapter card
        chapter_h = 0.55 + len(items) * 0.35 + 0.15
        add_rect(s, x, y, left_w, chapter_h, LIGHT_BG,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, x, y, left_w, 0.45, ORANGE,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        # Chapter letter
        add_rect(s, x + 0.10, y + 0.06, 0.35, 0.34, NAVY,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x + 0.10, y + 0.06, 0.35, 0.34, [
            (letter, dict(size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.55, y + 0.06, left_w - 0.65, 0.34, [
            (title, dict(size=11.5, bold=True, color=WHITE))],
            anchor=MSO_ANCHOR.MIDDLE)
        # Items
        for j, (num, item_title, page, color) in enumerate(items):
            iy = y + 0.55 + j * 0.35
            add_rect(s, x + 0.10, iy, 0.35, 0.28, color,
                     shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            add_text(s, x + 0.10, iy, 0.35, 0.28, [
                (page, dict(size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
                anchor=MSO_ANCHOR.MIDDLE)
            add_text(s, x + 0.55, iy, left_w - 0.65, 0.28, [
                (item_title, dict(size=10, color=DARK))],
                anchor=MSO_ANCHOR.MIDDLE)
        y_col[col] += chapter_h + 0.15
    # Footer note
    add_text(s, CONTENT_LEFT, 6.75, CONTENT_W, 0.25, [
        ("Struktur folgt dem Vorstandsprotokoll SEPM (01.07.2026). Farben: gruen = in Arbeit, "
         "rot = out of scope, navy = Bestandteil Business Case.",
         dict(size=9, italic=True, color=GREY_TXT))],
        anchor=MSO_ANCHOR.MIDDLE)
slide_toc()

# ============================================================
# 3 · VORSTANDSPROTOKOLL-UEBERSICHT
# ============================================================
def slide_protokoll():
    s = add_slide_stihl("Vorstandsprotokoll SEPM — Ausgangsbasis der Vorgehensweise")
    # Header row
    hdr_y = CONTENT_TOP
    add_rect(s, CONTENT_LEFT, hdr_y, CONTENT_W, 0.32, DARK,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    cols = [("#", 0.5), ("Typ", 0.6), ("Thema", 6.5), ("Vorstands-Anmerkung / Status", 3.5), ("Verantw.", 0.9), ("Frist", 0.9)]
    xc = CONTENT_LEFT
    for name, w in cols:
        add_text(s, xc, hdr_y, w, 0.32, [
            (name, dict(size=10.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        xc += w
    # Rows
    rows = [
        ("1", "E", "Dem Antrag wird zugestimmt · Tiefergehende Analyse SNOW · Entscheidungsvorlage + Implementierungsplan · Jahresplanung 2027 VS",
         "Beschlossen — dieser Business Case ist die Umsetzung",  GREEN,     "beschlossen", GREEN,     "D8/ODS · D1/SEH", "Ende 08/2026"),
        ("2", "A", "Prueft, ob 'Hype' (Massnahmenmgmt) durch SNOW abgeloest werden kann",
         "Hype stark customized → Vorschlag Out-of-Scope (nur Begruendung)", RED_HL, "OUT OF SCOPE", RED_HL, "D8/ODS · D1/SEH", "Ende 08/2026"),
        ("3", "A", "Prueft SNOW-Abdeckung Schluesselprozesse / App Engine (keine Standardisierung um jeden Preis)",
         "In Progress — siehe Vorhaben-Plan",                        GREEN,      "IN PROGRESS", GREEN,     "D8/ODS · D1/SEH", "Ende 08/2026"),
        ("4", "A", "Governance-Strukturen, Prozess-Aufwand (FTE) im Folgeprojekt bewerten (Komplexitaet vermeiden)",
         "In Progress — Business Case + Komplexitaets-Reduzierung",   GREEN,      "IN PROGRESS", GREEN,     "D8/ODS · D1/SEH", "Ende 08/2026"),
        ("5", "I", "Support fuer PIT limitiert — Ablose durch SNOW im Folgeprojekt pruefen",
         "Information — als Kontext im Business Case beruecksichtigt", NAVY,       "INFORMATION", NAVY,      "—", "—"),
        ("6", "A", "SNOW Plattform-Abhaengigkeitsrisiko pruefen (Vergleich mit SAP)",
         "Risikomatrix, Referenzkunden, Lockin, SNOW-STIHL-Historie, AI Vibe Coding, EAM Plattformstrategie",
                                                                     GREEN,      "IN PROGRESS", GREEN,     "D8/ODS · D1/SEH", "Ende 08/2026"),
    ]
    for i, r in enumerate(rows):
        num, typ, thema, anm, anm_color, status, status_color, resp, frist = r
        y = hdr_y + 0.40 + i * 0.85
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        add_rect(s, CONTENT_LEFT, y, CONTENT_W, 0.80, bg,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, CONTENT_LEFT, y, 0.08, 0.80, anm_color)
        xc = CONTENT_LEFT
        # #
        add_text(s, xc + 0.10, y, 0.5, 0.80, [
            (num, dict(size=12, bold=True, color=DARK, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        xc += 0.5
        # Typ
        add_rect(s, xc + 0.10, y + 0.24, 0.40, 0.30, NAVY,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, xc + 0.10, y + 0.24, 0.40, 0.30, [
            (typ, dict(size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        xc += 0.6
        # Thema
        add_text(s, xc + 0.10, y, 6.30, 0.80, [
            (thema, dict(size=9.5, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        xc += 6.5
        # Anmerkung + Status badge
        add_rect(s, xc + 0.10, y + 0.05, 3.30, 0.22, status_color,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, xc + 0.10, y + 0.05, 3.30, 0.22, [
            (status, dict(size=8.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, xc + 0.10, y + 0.30, 3.30, 0.48, [
            (anm, dict(size=9, italic=True, color=anm_color))],
            anchor=MSO_ANCHOR.MIDDLE)
        xc += 3.5
        # Verantw.
        add_text(s, xc, y, 0.9, 0.80, [
            (resp, dict(size=8.5, color=DARK, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        xc += 0.9
        # Frist
        add_text(s, xc, y, 0.9, 0.80, [
            (frist, dict(size=9, bold=True, color=DARK, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    # Legend
    leg_y = 6.85
    add_text(s, CONTENT_LEFT, leg_y - 0.05, 3.0, 0.25, [
        ("Legende:", dict(size=9, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    items = [(GREEN, "In Progress / beschlossen"),
             (RED_HL, "Out of Scope"),
             (NAVY,  "Information (kein Auftrag)")]
    lx = CONTENT_LEFT + 0.9
    for color, lbl in items:
        add_rect(s, lx, leg_y + 0.02, 0.20, 0.16, color,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, lx + 0.28, leg_y - 0.05, 2.8, 0.25, [
            (lbl, dict(size=9, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        lx += 3.2
slide_protokoll()

# ============================================================
# 4 · VORGEHENSWEISE / ZEITPLAN
# ============================================================
def slide_vorgehensweise():
    s = add_slide_stihl("Vorgehensweise & Zeitplan bis Vorstandssitzung Q3 2026")
    add_text(s, CONTENT_LEFT, CONTENT_TOP, CONTENT_W, 0.30, [
        ("Umsetzungspfad: alle Antraege des Vorstandsprotokolls werden bis Ende 08/2026 abgearbeitet",
         dict(size=11, italic=True, color=GREY_TXT))])
    # Timeline: Juli - Aug - Sept 2026
    tl_top = CONTENT_TOP + 0.50
    tl_h = 4.5
    # Month axis
    axis_y = tl_top
    m_labels = ["Juli 26", "Aug 26", "Sept 26"]
    lbl_w = 3.0
    tl_x = CONTENT_LEFT + lbl_w
    tl_w = CONTENT_W - lbl_w
    m_w = tl_w / 3
    for i, lbl in enumerate(m_labels):
        color = ORANGE_LT if i % 2 == 0 else ORANGE2
        text_col = DARK if i % 2 == 0 else WHITE
        add_rect(s, tl_x + i * m_w, axis_y, m_w - 0.05, 0.35, color,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, tl_x + i * m_w, axis_y, m_w - 0.05, 0.35, [
            (lbl, dict(size=11, bold=True, color=text_col, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    # Milestone: VS Ende Sept
    line_y = axis_y + 0.42
    add_rect(s, tl_x, line_y, tl_w, 0.03, LIGHT_GREY)
    ms_x = tl_x + 2 * m_w + m_w * 0.85
    add_diamond(s, ms_x - 0.15, line_y - 0.13, 0.30, 0.30, RED_HL)
    add_text(s, ms_x - 1.20, line_y + 0.25, 2.40, 0.24, [
        ("VS Q3 2026",
         dict(size=10, bold=True, color=RED_HL, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    # Swim lanes with Vorgänge from protokoll
    lanes_top = line_y + 0.70
    lanes = [
        ("Punkt 3 — Prozess- & Toolanalyse",  "In Progress",  GREEN,    0.00, 0.55),
        ("Punkt 4 — Business Case",            "In Progress",  GREEN,    0.05, 0.68),
        ("Punkt 4 — Komplexitaets-Reduzierung","In Progress",  GREEN,    0.10, 0.60),
        ("Punkt 6 — Alternative PM-Plattform", "In Progress",  GREEN,    0.05, 0.70),
        ("Punkt 6 — Risiko / Lockin / EAM / AI","In Progress", GREEN,    0.15, 0.85),
        ("Alle Punkte — Stakeholder-Management","laufend",     BLUE_HD,  0.15, 1.00),
    ]
    lane_h = 0.42
    lane_gap = 0.08
    for i, (name, status, color, s_frac, e_frac) in enumerate(lanes):
        y = lanes_top + i * (lane_h + lane_gap)
        # Label
        add_rect(s, CONTENT_LEFT, y, lbl_w - 0.10, lane_h, LIGHT_BG,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, CONTENT_LEFT, y, 0.06, lane_h, color)
        add_text(s, CONTENT_LEFT + 0.12, y, lbl_w - 0.30, lane_h, [
            (name, dict(size=9.5, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        # Bar
        add_rect(s, tl_x, y, tl_w, lane_h, LIGHT_BG,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        bar_x = tl_x + s_frac * tl_w
        bar_w = (e_frac - s_frac) * tl_w
        add_rect(s, bar_x, y + 0.05, bar_w, lane_h - 0.10, color,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, bar_x + 0.10, y, bar_w - 0.20, lane_h, [
            (status, dict(size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
slide_vorgehensweise()

# ============================================================
# 5 · Kap A — Punkt 2 · Hype = Out of Scope
# ============================================================
def slide_hype_oos():
    s = add_slide_stihl("Punkt 2 — Massnahmen-Management 'Hype': Begruendung Out-of-Scope")
    # Big status banner
    add_rect(s, CONTENT_LEFT, CONTENT_TOP, CONTENT_W, 0.75, RED_LT,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, CONTENT_LEFT, CONTENT_TOP, 0.10, 0.75, RED_HL)
    add_text(s, CONTENT_LEFT + 0.20, CONTENT_TOP, 6.5, 0.75, [
        ("STATUS: OUT OF SCOPE",
         dict(size=13, bold=True, color=RED_HL))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, CONTENT_LEFT + 6.8, CONTENT_TOP, 5.9, 0.75, [
        ("Vorstands-Vorschlag 01.07.2026: Hype-Ablose nicht Teil dieser Untersuchung",
         dict(size=10.5, italic=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    # 4 Begründungs-Karten
    top = CONTENT_TOP + 0.95
    h = 2.30
    gap = 0.15
    w = (CONTENT_W - gap) / 2
    reasons = [
        ("1", "Stark customized ueber Jahre",
         "Hype wurde ueber mehrere Release-Zyklen fuer STIHL-spezifische Massnahmen-Logik erweitert. "
         "Eine Ablose muesste alle Customizings inhaltlich nachziehen — enormer Migrationsaufwand.", RED_HL),
        ("2", "Zusatzfunktionen mit hoher Fachtiefe",
         "Hype-Funktionen fuer Zielkosten-verfolgung / Massnahmen-Backlog sind stark im Prozess der "
         "Produktentwicklung verwurzelt. Nicht 1:1 auf SNOW SPM abbildbar ohne massive Anpassung.", RED_HL),
        ("3", "Keine direkte PPM-Kernfunktion",
         "Hype gehoert nicht in den Kern des Portfolio-Managements. Der Business Case konzentriert sich "
         "auf Strategy, Portfolio, Demand, Projekt- und Ressourcen-Management — nicht auf Massnahmen-Details.", ORANGE2),
        ("4", "Vermeidung von Scope-Creep",
         "Erweiterung um Hype wuerde die Aufwaende und Risiken des Business Case erheblich vergroessern "
         "und den Zeitplan fuer die Jahresplanung 2027 gefaehrden.", ORANGE2),
    ]
    for i, (num, head, body, color) in enumerate(reasons):
        row = i // 2
        col = i % 2
        x = CONTENT_LEFT + col * (w + gap)
        y = top + row * (h + gap)
        add_rect(s, x, y, w, h, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=color, line_w=1.2)
        add_rect(s, x, y, w, 0.42, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_oval(s, x + 0.10, y + 0.06, 0.30, 0.30, WHITE)
        add_text(s, x + 0.10, y + 0.06, 0.30, 0.30, [
            (num, dict(size=13, bold=True, color=color, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.45, y + 0.06, w - 0.55, 0.30, [
            (head, dict(size=12, bold=True, color=WHITE))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.15, y + 0.55, w - 0.30, h - 0.65, [
            (body, dict(size=10, color=DARK))], anchor=MSO_ANCHOR.TOP)
    # Alternative-Hinweis
    alt_y = top + 2 * h + 2 * gap + 0.05
    add_rect(s, CONTENT_LEFT, alt_y, CONTENT_W, 0.55, YELLOW_LT,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, CONTENT_LEFT, alt_y, 0.08, 0.55, YELLOW)
    add_text(s, CONTENT_LEFT + 0.15, alt_y, CONTENT_W - 0.25, 0.55, [
        ("Zukunftsoption: Prüfung einer losen Integration Hype ↔ SNOW (z.B. Massnahmen-Referenz "
         "auf Portfolio-Ebene) im spaeteren Projekt-Follow-up — nicht Teil dieses Business Case.",
         dict(size=10, italic=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
slide_hype_oos()

# ============================================================
# 6 · Kap A — Punkt 3 · Prozessabdeckung + App Engine
# ============================================================
def slide_prozessabdeckung():
    s = add_slide_stihl("Punkt 3 — SNOW-Prozessabdeckung + App-Engine-Strategie")
    add_rect(s, CONTENT_LEFT, CONTENT_TOP, CONTENT_W, 0.65, GREEN_LT,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, CONTENT_LEFT, CONTENT_TOP, 0.10, 0.65, GREEN)
    add_text(s, CONTENT_LEFT + 0.20, CONTENT_TOP, 5.5, 0.65, [
        ("STATUS: IN PROGRESS",
         dict(size=12, bold=True, color=GREEN))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, CONTENT_LEFT + 5.8, CONTENT_TOP, 6.9, 0.65, [
        ("Leitprinzip: Standard first — aber keine Standardisierung um jeden Preis",
         dict(size=10.5, italic=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    # Three-column classification
    top = CONTENT_TOP + 0.85
    h = CONTENT_BOT - top
    w = (CONTENT_W - 0.30) / 3
    gap = 0.15
    buckets = [
        ("Standard SNOW-Modul",
         GREEN, "hohe Abdeckung, kein Anpassungsbedarf",
         [("Demand Management",       "SNOW SPM Standard"),
          ("Portfolio Management",    "SNOW SPM Standard"),
          ("Project Management (Gantt/Agile)","SNOW SPM Standard"),
          ("Resource Management",     "SNOW SPM Standard"),
          ("Strategic Planning / OKR","SNOW SPM Pro"),
          ("Financial Management",    "SNOW SPM Pro"),
          ("Now Assist (KI)",         "SNOW SPM Pro")]),
        ("App Engine (Low-Code)",
         BLUE_HD, "Standardnahe Prozesse, ohne 3rd-Party-Tool",
         [("EV-Prozess",             "App-Engine-Workflow"),
          ("Vorhaben-Steckbrief",    "App-Engine-Formular"),
          ("Stage-Gate PEP",         "App-Engine + SPM-Verknuepfung"),
          ("Steckbrief-Kategorisierung","App-Engine + Portfolio-Link"),
          ("Freigabe-Workflows",     "App-Engine + Approvals")]),
        ("3rd-Party / Bestand",
         ORANGE, "sinnvoll wenn Kernprozess ausserhalb SPM",
         [("SAP FI/CO",              "Quelle Kosten (bleibt)"),
          ("SAP HCM",                "Quelle Personal-Stamm (bleibt)"),
          ("PLM / CAD",              "Engineering-Kern (bleibt)"),
          ("Hype",                   "Massnahmen-Mgmt (bleibt)"),
          ("Power BI (optional)",    "Cross-Reporting")]),
    ]
    for i, (title, color, sub, items) in enumerate(buckets):
        x = CONTENT_LEFT + i * (w + gap)
        add_rect(s, x, top, w, h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, x, top, w, 0.65, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x + 0.10, top + 0.04, w - 0.20, 0.30, [
            (title, dict(size=12, bold=True, color=WHITE))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.10, top + 0.36, w - 0.20, 0.25, [
            (sub, dict(size=9, italic=True, color=WHITE))],
            anchor=MSO_ANCHOR.MIDDLE)
        for j, (name, note) in enumerate(items):
            iy = top + 0.80 + j * 0.55
            add_rect(s, x + 0.12, iy, w - 0.24, 0.50, WHITE,
                     shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=color, line_w=0.4)
            add_rect(s, x + 0.12, iy, 0.06, 0.50, color)
            add_text(s, x + 0.25, iy + 0.02, w - 0.40, 0.24, [
                (name, dict(size=10, bold=True, color=DARK))],
                anchor=MSO_ANCHOR.MIDDLE)
            add_text(s, x + 0.25, iy + 0.26, w - 0.40, 0.22, [
                (note, dict(size=8.5, italic=True, color=color))],
                anchor=MSO_ANCHOR.MIDDLE)
slide_prozessabdeckung()

# ============================================================
# KAPITEL B — Business Case
# ============================================================

# 7 · Business Case Methodik
def slide_methodik():
    s = add_slide_stihl("Business Case Methodik")
    top = CONTENT_TOP
    h = 2.85
    gap = 0.15
    w = (CONTENT_W - gap) / 2
    quads = [
        ("TCO (5 Jahre)", "Lizenzen · Implementierung · Support · Change",
         ["Lizenzen SNOW SPM Pro","Implementierung","Change & Schulung","Support / Wartung"],
         ORANGE, "€"),
        ("Nutzen / Business Value", "Effizienz · Time-to-Market · Governance",
         ["Effizienz PPM","Time-to-Market","Ressourcen-Steuerung","Governance"], GREEN, "*"),
        ("Konsolidierung", "Tool-Abloese · Schnittstellen · flankierende Prozesse",
         ["PIT / MSPO / Workpath","Schnittstellen","Massnahmen-Mgmt","APM"], NAVY, ">>"),
        ("Risiko / Sensitivitaet", "Vendor · Preis · Betrieb · DSGVO",
         ["Vendor Lock-in","Preis-Eskalation","Ausfall / Betrieb","Sensitivity"], RED_HL, "!"),
    ]
    for i, (head, sub, items, color, icon) in enumerate(quads):
        row = i // 2; col = i % 2
        x = CONTENT_LEFT + col * (w + gap)
        y = top + row * (h + gap)
        add_rect(s, x, y, w, h, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=color, line_w=1.5)
        add_rect(s, x, y, w, 0.42, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_oval(s, x + 0.08, y + 0.06, 0.30, 0.30, WHITE)
        add_text(s, x + 0.08, y + 0.06, 0.30, 0.30, [
            (icon, dict(size=13, bold=True, color=color, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.45, y + 0.02, w - 0.55, 0.38, [
            (head, dict(size=13, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.15, y + 0.52, w - 0.30, 0.30, [
            (sub, dict(size=10, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.MIDDLE)
        for j, it in enumerate(items):
            iy = y + 0.90 + j * 0.42
            add_oval(s, x + 0.20, iy + 0.13, 0.08, 0.08, color)
            add_text(s, x + 0.35, iy, w - 0.50, 0.36, [
                (it, dict(size=10.5, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
slide_methodik()

# 8 · Business Value
def slide_business_value():
    s = add_slide_stihl("Business Value Analyse — Steady State p.a.")
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
        add_rect(s, CONTENT_LEFT, y, CONTENT_W, row_h, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, CONTENT_LEFT, y, 0.07, row_h, ORANGE)
        add_text(s, CONTENT_LEFT + 0.15, y + 0.04, 4.5, row_h - 0.10, [
            (name, dict(size=11, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, CONTENT_LEFT + 0.15, y + row_h - 0.34, 4.5, 0.28, [
            (desc, dict(size=9, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.TOP)
        bar_x = CONTENT_LEFT + 4.80; bar_w_max = 6.30
        bar_w_val = bar_w_max * (val / max_val)
        add_rect(s, bar_x, y + 0.30, bar_w_max, 0.34, LIGHT_GREY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, bar_x, y + 0.30, bar_w_val, 0.34, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, CONTENT_LEFT + 11.20, y, 1.40, row_h, [
            (f"{val:,} k€".replace(",", "."),
             dict(size=13, bold=True, color=NAVY, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
slide_business_value()

# 9 · TCO
def slide_tco():
    s = add_slide_stihl("TCO — Total Cost of Ownership (5 Jahre)")
    total_5y = IMPL_ONE + CHANGE_ONE + 5 * (LICENSE_Y_STEADY +
                SUPPORT_INT_FTE_STEADY * FTE_RATE + SUPPORT_EXT)
    add_rect(s, CONTENT_LEFT, CONTENT_TOP, CONTENT_W, 0.70, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.15, CONTENT_TOP, 6, 0.70, [
        ("5-Jahres-TCO SEPM (Base Case)",
         dict(size=13, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, CONTENT_LEFT + 6.2, CONTENT_TOP, 5, 0.70, [
        (f"~ {total_5y/1000:.1f} Mio. €",
         dict(size=22, bold=True, color=ORANGE, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, CONTENT_LEFT + 11.20, CONTENT_TOP, 1.60, 0.70, [
        (f"(~ {total_5y/5/1000:.2f} M€ p.a. Ø)",
         dict(size=10, italic=True, color=LIGHT_GREY, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)
    left_w = (CONTENT_W - 0.15) / 2
    right_x = CONTENT_LEFT + left_w + 0.15
    tt = CONTENT_TOP + 0.85
    h = CONTENT_BOT - tt
    add_rect(s, CONTENT_LEFT, tt, left_w, h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, CONTENT_LEFT, tt, left_w, 0.35, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.10, tt, left_w - 0.20, 0.35, [
        ("Einmalige Kosten (Y1-Y2)",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    einmalig = [("Implementierungs-Partner (SI)", IMPL_ONE),
                ("Interne FTE (Phase 1+2)", INT_FTE_IMPL * FTE_RATE * 2),
                ("Change / Schulung", CHANGE_ONE),
                ("Contingency (10%)", int(0.1 * (IMPL_ONE + INT_FTE_IMPL * FTE_RATE * 2 + CHANGE_ONE)))]
    for i, (n, v) in enumerate(einmalig):
        y = tt + 0.50 + i * 0.75
        add_rect(s, CONTENT_LEFT + 0.15, y, left_w - 0.30, 0.65, WHITE,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=NAVY, line_w=0.5)
        add_text(s, CONTENT_LEFT + 0.30, y, 3.8, 0.65, [
            (n, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, CONTENT_LEFT + left_w - 1.70, y, 1.50, 0.65, [
            (f"{v:,} k€".replace(",", "."),
             dict(size=12.5, bold=True, color=NAVY, align=PP_ALIGN.RIGHT))], anchor=MSO_ANCHOR.MIDDLE)
    total_e = sum(x[1] for x in einmalig)
    add_text(s, CONTENT_LEFT + 0.10, tt + 3.75, left_w - 0.20, 0.40, [
        (f"Summe einmalig: {total_e:,} k€".replace(",", "."),
         dict(size=12, bold=True, color=NAVY, align=PP_ALIGN.RIGHT))], anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, right_x, tt, left_w, h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, right_x, tt, left_w, 0.35, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, right_x + 0.10, tt, left_w - 0.20, 0.35, [
        ("Laufende Kosten p.a. (Steady State)",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    laufend = [("Lizenzen SNOW SPM Pro", LICENSE_Y_STEADY),
               ("Interne Betriebs-FTE", int(SUPPORT_INT_FTE_STEADY * FTE_RATE)),
               ("Externer Support / Wartung", SUPPORT_EXT),
               ("Weiterentwicklung", 220)]
    for i, (n, v) in enumerate(laufend):
        y = tt + 0.50 + i * 0.75
        add_rect(s, right_x + 0.15, y, left_w - 0.30, 0.65, WHITE,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=ORANGE, line_w=0.5)
        add_text(s, right_x + 0.30, y, 3.8, 0.65, [
            (n, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, right_x + left_w - 1.70, y, 1.50, 0.65, [
            (f"{v:,} k€".replace(",", "."),
             dict(size=12.5, bold=True, color=ORANGE, align=PP_ALIGN.RIGHT))], anchor=MSO_ANCHOR.MIDDLE)
    total_l = sum(x[1] for x in laufend)
    add_text(s, right_x + 0.10, tt + 3.75, left_w - 0.20, 0.40, [
        (f"Summe p.a.: {total_l:,} k€".replace(",", "."),
         dict(size=12, bold=True, color=ORANGE, align=PP_ALIGN.RIGHT))], anchor=MSO_ANCHOR.MIDDLE)
slide_tco()

# 10 · Support Kosten
def slide_support():
    s = add_slide_stihl("Support-Kosten — Detailbetrachtung")
    cols = [
        ("Interner Betrieb", ORANGE,
         [("SNOW-Plattform-Admin","1,0 FTE",130),("SPM Business Analyst","1,0 FTE",130),
          ("Integration / Schnittstellen","0,5 FTE",65),("Data Governance","0,3 FTE",39),
          ("User Support 2nd Level","0,2 FTE",26)]),
        ("Externer Support", NAVY,
         [("Wartungsvertrag SNOW","Standard 22%",70),("Managed-Service Partner","Retainer",90),
          ("Externes Consulting","20-30 Tage/y",60),("Zertifikate & Schulungen","5-8 Personen",30)]),
        ("Weiterentwicklung", GREEN,
         [("Feature-Backlog","~10 PT/Mo",100),("Release-Upgrades","2× / Jahr",50),
          ("Custom-Reports","4-6 PT/Mo",40),("Innovation Sprints","1× Q",30)]),
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
            (title, dict(size=13, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
        total = 0
        for j, (n, note, val) in enumerate(items):
            y = top + 0.50 + j * 0.65
            add_rect(s, x + 0.10, y, w - 0.20, 0.60, WHITE,
                     shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=color, line_w=0.5)
            add_text(s, x + 0.20, y + 0.02, w - 1.60, 0.28, [
                (n, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
            add_text(s, x + 0.20, y + 0.30, w - 1.60, 0.28, [
                (note, dict(size=9, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.MIDDLE)
            add_text(s, x + w - 1.50, y + 0.02, 1.35, 0.56, [
                (f"{val} k€", dict(size=12, bold=True, color=color, align=PP_ALIGN.RIGHT))],
                anchor=MSO_ANCHOR.MIDDLE)
            total += val
        add_rect(s, x + 0.10, top + col_h - 0.55, w - 0.20, 0.35, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x + 0.10, top + col_h - 0.55, w - 0.20, 0.35, [
            (f"Summe: {total} k€ / Jahr",
             dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)
slide_support()

# 11 · Konsolidierung Tools
def slide_konsolidierung_tools():
    s = add_slide_stihl("Konsolidierungseinsparungen — Tool-Abloese  (inkl. PIT-Ablose · Punkt 5)")
    add_rect(s, CONTENT_LEFT, CONTENT_TOP, CONTENT_W, 0.70, GREEN, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.15, CONTENT_TOP, 8, 0.70, [
        ("Gesamt Tool-Abloese-Ersparnis (Steady State p.a.)",
         dict(size=12.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, CONTENT_LEFT + 8.2, CONTENT_TOP, 4.4, 0.70, [
        (f"~ {TOTAL_TOOL_SAVINGS} k€ / Jahr",
         dict(size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)
    max_val = max(x[2] for x in CONSOL_TOOLS)
    row_top = CONTENT_TOP + 0.90
    row_h = 0.72
    for i, (tool, cat, saving) in enumerate(CONSOL_TOOLS):
        y = row_top + i * (row_h + 0.05)
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        add_rect(s, CONTENT_LEFT, y, CONTENT_W, row_h, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, CONTENT_LEFT, y, 0.07, row_h, GREEN)
        add_text(s, CONTENT_LEFT + 0.15, y + 0.03, 3.5, row_h - 0.06, [
            (tool, dict(size=11.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, CONTENT_LEFT + 0.15, y + row_h - 0.28, 3.5, 0.22, [
            (cat, dict(size=8.5, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.TOP)
        bar_x = CONTENT_LEFT + 3.90; bar_w_max = 7.30
        bar_w_val = bar_w_max * (saving / max_val)
        add_rect(s, bar_x, y + 0.28, bar_w_max, 0.24, LIGHT_GREY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, bar_x, y + 0.28, bar_w_val, 0.24, GREEN, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, CONTENT_LEFT + 11.30, y, 1.30, row_h, [
            (f"- {saving} k€",
             dict(size=13, bold=True, color=GREEN, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)
slide_konsolidierung_tools()

# 12 · Konsolidierung Schnittstellen
def slide_konsolidierung_interfaces():
    s = add_slide_stihl("Konsolidierungseinsparungen — Schnittstellen")
    add_rect(s, CONTENT_LEFT, CONTENT_TOP, CONTENT_W, 0.70, GREEN, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.15, CONTENT_TOP, 8, 0.70, [
        ("Ersparnis Schnittstellen & Integration p.a.",
         dict(size=12.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, CONTENT_LEFT + 8.2, CONTENT_TOP, 4.4, 0.70, [
        (f"~ {INT_SAVINGS} k€ / Jahr",
         dict(size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)
    left_w = (CONTENT_W - 0.15) / 2
    right_x = CONTENT_LEFT + left_w + 0.15
    tt = CONTENT_TOP + 0.85
    h = CONTENT_BOT - tt
    add_rect(s, CONTENT_LEFT, tt, left_w, h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, CONTENT_LEFT, tt, left_w, 0.35, RED_HL, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.10, tt, left_w - 0.20, 0.35, [
        ("HEUTE - Fragmentierte Integration",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    heute_ints = ["PIT -> SAP FI (Kostentransfer)", "MSPO -> SAP HCM (Personalstamm)",
        "SharePoint -> Excel (Datenexport)", "Workpath -> Excel (OKR-Reporting)",
        "PowerApps -> Dataverse", "Lucom EV -> SAP", "Excel-Konsolidierungen (div.)"]
    for i, txt in enumerate(heute_ints):
        y = tt + 0.50 + i * 0.55
        add_rect(s, CONTENT_LEFT + 0.15, y, left_w - 0.30, 0.50, WHITE,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=RED_HL, line_w=0.5)
        add_rect(s, CONTENT_LEFT + 0.15, y, 0.06, 0.50, RED_HL)
        add_text(s, CONTENT_LEFT + 0.30, y, left_w - 0.50, 0.50, [
            (txt, dict(size=10, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, right_x, tt, left_w, h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, right_x, tt, left_w, 0.35, GREEN, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, right_x + 0.10, tt, left_w - 0.20, 0.35, [
        ("MIT SEPM - Native SNOW-Integration",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    mit_ints = [("SAP FI Konnektor","Standard, Integration Hub"),
        ("SAP HCM Konnektor","Standard, Integration Hub"),
        ("Interne Datenmodelle","keine Excel-Export-Ketten"),
        ("OKR-Modul nativ","kein Workpath-Bridge"),
        ("App Engine (Low-Code)","Custom-Data auf gleicher Plattform"),
        ("EV-Prozess nativ","kein Lucom-Bridge"),
        ("Live-Reporting","Excel-Konsolidierung entfaellt")]
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

# 13 · Flankierende Prozesse
def slide_flanking():
    s = add_slide_stihl("Flankierende Prozess-Potentiale")
    add_rect(s, CONTENT_LEFT, CONTENT_TOP, CONTENT_W, 0.70, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.15, CONTENT_TOP, 8.5, 0.70, [
        ("Zusaetzliches Einspar-Potenzial (SPM-flankierend, p.a.)",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, CONTENT_LEFT + 8.7, CONTENT_TOP, 3.9, 0.70, [
        (f"~ {TOTAL_FLANKING:,} k€ / Jahr".replace(",", "."),
         dict(size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)
    max_val = max(x[2] for x in FLANKING)
    row_top = CONTENT_TOP + 0.90
    row_h = 0.86
    for i, (name, desc, val) in enumerate(FLANKING):
        y = row_top + i * (row_h + 0.06)
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        add_rect(s, CONTENT_LEFT, y, CONTENT_W, row_h, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
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
        add_rect(s, bar_x, y + 0.30, bar_w_max, 0.26, LIGHT_GREY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, bar_x, y + 0.30, bar_w_val, 0.26, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, CONTENT_LEFT + 11.30, y, 1.30, row_h, [
            (f"{val} k€",
             dict(size=13, bold=True, color=NAVY, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)
slide_flanking()

# 14 · Komplexitäts-Reduzierung
def slide_komplex():
    s = add_slide_stihl("Komplexitaets-Reduzierung ProjektControlling  (Punkt 4 Folie 4)")
    left_w = (CONTENT_W - 0.15) / 2
    right_x = CONTENT_LEFT + left_w + 0.15
    top = CONTENT_TOP
    h = CONTENT_BOT - top
    add_rect(s, CONTENT_LEFT, top, left_w, h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, CONTENT_LEFT, top, left_w, 0.35, RED_HL, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.10, top, left_w - 0.20, 0.35, [
        ("HEUTE - Komplexitaetstreiber",
         dict(size=11.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    hoy = [
        ("Multi-Tool-Landschaft PPM","PIT · MSPO · Excel · SharePoint · Power Apps","~ 2,5 FTE Betriebs-/Integrations-Aufwand"),
        ("Manuelle Daten-Aggregation","Excel-Konsolidierung fuer Reportings","~ 2 Wochen/Monat Reporting"),
        ("Uneinheitliche Datenmodelle","Jede Insel eigene Struktur","~ 1 FTE Datenqualitaet"),
        ("Doppelte Datenpflege","Projekte in PIT + Excel + SharePoint","~ 0,8 FTE Doppelerfassung"),
        ("Fehlende Traceability","Idee-Portfolio-Projekt-Kosten nicht durchgaengig","Governance-Aufwand hoch"),
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
    add_rect(s, right_x, top, left_w, h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, right_x, top, left_w, 0.35, GREEN, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, right_x + 0.10, top, left_w - 0.20, 0.35, [
        ("MIT SEPM - Konsolidierungs-Hebel",
         dict(size=11.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    mit = [
        ("Eine Plattform, ein Datenmodell","SNOW SPM als Single Source of Truth","- 6 Tools · - 0,5 FTE Betrieb"),
        ("Automatische Reports & Dashboards","Live-Dashboards ersetzen Excel","- 1,5 Wochen/Monat Reporting"),
        ("Standard-Datenmodell","Vorhaben-Steckbrief einheitlich","- 0,7 FTE Datenqualitaet"),
        ("Erfassung einmal - Nutzung ueberall","Demand-Portfolio-Projekt in einem System","- 0,6 FTE Doppelerfassung"),
        ("Durchgaengige Traceability","OKR-Portfolio-Projekt-Kosten","Audit-faehig"),
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
slide_komplex()

# ============================================================
# KAPITEL C — Alternativen & Umsetzung
# ============================================================

# 15 · Alternative PM-Plattform
def slide_alternative_pm():
    s = add_slide_stihl("Alternative PM-Plattform — Bewertung (Hedge)")
    left_w = (CONTENT_W - 0.15) / 2
    right_x = CONTENT_LEFT + left_w + 0.15
    top1 = CONTENT_TOP
    top1_h = 2.90
    add_rect(s, CONTENT_LEFT, top1, left_w, top1_h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, CONTENT_LEFT, top1, left_w, 0.35, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.10, top1, left_w - 0.20, 0.35, [
        ("Warum Alternative bewerten?",
         dict(size=11.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    rat = [("Risiko-Absicherung","Falls SNOW fuer VEW/VPM nicht tief genug"),
           ("Verhandlungsposition","Best-of-Breed-Angebot senkt SNOW-Preis"),
           ("Engineering-Spezifika","PLM-nahe Prozesse ggf. besser in Fach-Tool"),
           ("EWW-Bedenken","Explizite Alternative senkt Widerstand")]
    for i, (h, b) in enumerate(rat):
        y = top1 + 0.55 + i * 0.56
        add_oval(s, CONTENT_LEFT + 0.15, y + 0.12, 0.22, 0.22, NAVY)
        add_text(s, CONTENT_LEFT + 0.15, y + 0.12, 0.22, 0.22, [
            (str(i + 1), dict(size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, CONTENT_LEFT + 0.45, y, left_w - 0.55, 0.25, [
            (h, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, CONTENT_LEFT + 0.45, y + 0.26, left_w - 0.55, 0.24, [
            (b, dict(size=9, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, right_x, top1, left_w, top1_h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, right_x, top1, left_w, 0.35, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, right_x + 0.10, top1, left_w - 0.20, 0.35, [
        ("Shortlist Alternative PM-Plattformen",
         dict(size=11.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    alts = [("Planisware Enterprise","Marktfuehrer Engineering-PPM"),
            ("SAP EPPM","SAP-Oekosystem-Integration"),
            ("Sciforma","Mittelstands-Standard, hybrid")]
    for i, (n, sub) in enumerate(alts):
        y = top1 + 0.50 + i * 0.75
        add_rect(s, right_x + 0.15, y, left_w - 0.30, 0.65, WHITE,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=NAVY, line_w=0.5)
        add_rect(s, right_x + 0.15, y, 0.07, 0.65, NAVY)
        add_text(s, right_x + 0.30, y + 0.04, left_w - 1.90, 0.28, [
            (n, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, right_x + 0.30, y + 0.32, left_w - 1.90, 0.30, [
            (sub, dict(size=9, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.TOP)
        add_rect(s, right_x + left_w - 1.60, y + 0.15, 1.40, 0.35, ORANGE,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, right_x + left_w - 1.60, y + 0.15, 1.40, 0.35, [
            ("Lead Jonas", dict(size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    top2 = top1 + top1_h + 0.15
    top2_h = CONTENT_BOT - top2
    add_rect(s, CONTENT_LEFT, top2, CONTENT_W, top2_h, WHITE,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=LIGHT_GREY, line_w=0.7)
    add_rect(s, CONTENT_LEFT, top2, CONTENT_W, 0.35, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.10, top2, CONTENT_W - 0.20, 0.35, [
        ("Bewertungskriterien und Zwischenergebnis",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    criteria = [
        ("Funktions-Deckung","70%","80%","60%","65%"),
        ("Integration STIHL-Landschaft","80%","50%","70%","40%"),
        ("TCO 5 Jahre (rel.)","100%","115%","125%","95%"),
        ("Change-Aufwand","Mittel","Hoch","Mittel","Niedrig"),
        ("Zukunftssicherheit AI","Hoch","Mittel","Hoch","Niedrig"),
    ]
    col_x = [CONTENT_LEFT+0.10, CONTENT_LEFT+5.20, CONTENT_LEFT+7.10, CONTENT_LEFT+9.00, CONTENT_LEFT+10.90]
    col_w = [5.00, 1.85, 1.85, 1.85, 1.65]
    headers = ["Kriterium","SNOW SPM","Planisware","SAP EPPM","Sciforma"]
    hy = top2 + 0.42
    for i, hd in enumerate(headers):
        col = ORANGE if i == 1 else NAVY
        if i == 0:
            add_text(s, col_x[i], hy, col_w[i], 0.30, [
                (hd, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        else:
            add_rect(s, col_x[i], hy, col_w[i], 0.30, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
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
                (v, dict(size=10, color=DARK, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)
slide_alternative_pm()

# 16 · Szenarien-Uebersicht
def slide_szenarien_ov():
    s = add_slide_stihl("Einfuehrungs-Szenarien im Vergleich")
    top = CONTENT_TOP
    h = CONTENT_BOT - top
    w = (CONTENT_W - 0.30) / 3
    gap = 0.15
    scenarios = [
        ("A - Full Scope\n24 Monate", 24, ORANGE, "Empfohlen",
         [("Scope","Alle Module Phase 1 + 2"),("Ramp Up","Standard-Timeline"),
          ("Payback","~ 3,5 Jahre"),("Business Value","Voll ab Jahr 3"),
          ("Risiko","Mittel")], True),
        ("B - Extended Ramp\n36 Monate", 36, BLUE_HD, "Konservativ",
         [("Scope","Alle Module, verteilter"),("Ramp Up","Reduziertes Tempo"),
          ("Payback","~ 4,5 Jahre"),("Business Value","Voll ab Jahr 4"),
          ("Risiko","Niedrig")], False),
        ("C - MVP-First\n12 Mo + Ausbau", 30, GREEN, "Agil",
         [("Scope","Strategy + Portfolio zuerst"),("Ramp Up","MVP live in 12 Mo"),
          ("Payback","~ 3 Jahre auf Teil-Nutzen"),("Business Value","Teilnutzen frueh"),
          ("Risiko","Mittel")], False),
    ]
    for i, (name, months, color, badge, items, recommend) in enumerate(scenarios):
        x = CONTENT_LEFT + i * (w + gap)
        add_rect(s, x, top, w, h, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                 line=color, line_w=2 if recommend else 1)
        add_rect(s, x, top, w, 0.66, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        if recommend:
            add_rect(s, x + w - 1.60, top, 1.55, 0.66, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            add_text(s, x + w - 1.60, top, 1.55, 0.66, [
                ("★ EMPFOHLEN", dict(size=9, bold=True, color=ORANGE, align=PP_ALIGN.CENTER))],
                anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.08, top + 0.04, w - 1.75 if recommend else w - 0.15, 0.58, [
            (name, dict(size=13, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, x + 0.15, top + 0.75, w - 0.30, 0.30, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x + 0.15, top + 0.75, w - 0.30, 0.30, [
            (badge, dict(size=10, bold=True, color=color, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.15, top + 1.15, w - 0.30, 0.24, [
            (f"Dauer: {months} Monate",
             dict(size=9.5, italic=True, color=GREY_TXT, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        max_m = 36
        bar_w = (w - 0.30) * (months / max_m)
        add_rect(s, x + 0.15, top + 1.40, w - 0.30, 0.18, LIGHT_GREY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, x + 0.15, top + 1.40, bar_w, 0.18, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        for j, (k, v) in enumerate(items):
            y = top + 1.75 + j * 0.72
            add_rect(s, x + 0.15, y, w - 0.30, 0.66, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            add_text(s, x + 0.25, y + 0.02, w - 0.45, 0.26, [
                (k, dict(size=9, bold=True, color=GREY_TXT))], anchor=MSO_ANCHOR.MIDDLE)
            add_text(s, x + 0.25, y + 0.30, w - 0.45, 0.34, [
                (v, dict(size=10, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
slide_szenarien_ov()

# 17 · Szenario A Detail
def slide_szenario_A():
    s = add_slide_stihl("Szenario A - Full Scope 24 Monate (Empfohlen)")
    add_text(s, CONTENT_LEFT, CONTENT_TOP, CONTENT_W, 0.25, [
        ("Umsetzungs-Timeline",
         dict(size=12.5, bold=True, color=ORANGE))])
    axis_y = CONTENT_TOP + 0.30
    n_m = 24; lbl_w = 2.5
    left_x = CONTENT_LEFT + lbl_w
    right_x = CONTENT_LEFT + CONTENT_W
    m_w = (right_x - left_x) / n_m
    def mx(m): return left_x + m * m_w
    add_rect(s, left_x, axis_y, 12 * m_w, 0.28, ORANGE_LT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, left_x, axis_y, 12 * m_w, 0.28, [
        ("Jahr 1", dict(size=10.5, bold=True, color=DARK, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, left_x + 12 * m_w, axis_y, 12 * m_w, 0.28, ORANGE2, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, left_x + 12 * m_w, axis_y, 12 * m_w, 0.28, [
        ("Jahr 2", dict(size=10.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    phases = [("Konzeption", NAVY, 0, 3),("Impl. Phase 1", ORANGE, 3, 12),
              ("Go-Live Phase 1", GREEN, 12, 13),("Impl. Phase 2", ORANGE2, 12, 23),
              ("Go-Live Phase 2", GREEN, 23, 24),("Change / Rollout", BLUE_HD, 3, 24)]
    lane_top = axis_y + 0.35
    lane_h = 0.36; lane_gap = 0.05
    for i, (name, color, ms, me) in enumerate(phases):
        y = lane_top + i * (lane_h + lane_gap)
        add_text(s, CONTENT_LEFT, y, lbl_w - 0.10, lane_h, [
            (name, dict(size=9.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, left_x, y, n_m * m_w, lane_h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        bar_x = mx(ms); bar_w = (me - ms) * m_w - 0.02
        add_rect(s, bar_x + 0.01, y + 0.02, bar_w, lane_h - 0.04, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    curve_top = lane_top + len(phases) * (lane_h + lane_gap) + 0.25
    add_text(s, CONTENT_LEFT, curve_top, CONTENT_W, 0.25, [
        ("Investitions- und Nutzenprofil (schematisch)",
         dict(size=12, bold=True, color=ORANGE))])
    ct = curve_top + 0.30
    ch = CONTENT_BOT - ct
    add_rect(s, CONTENT_LEFT, ct, CONTENT_W, ch, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    year_w = CONTENT_W / 5
    for yi in range(5):
        x = CONTENT_LEFT + yi * year_w
        col = WHITE if yi % 2 == 0 else GREY_BG
        add_rect(s, x, ct, year_w - 0.01, ch, col)
        add_text(s, x, ct + ch - 0.28, year_w, 0.24, [
            (f"Jahr {yi + 1}",
             dict(size=9, color=GREY_TXT, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)
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

# 18 · Budgetplanung
def slide_budget():
    s = add_slide_stihl("Budgetplanung SEPM  (Jahresplanung 2027 + Rahmen 2028)")
    years = [2027, 2028, 2029, 2030, 2031]
    invs = [(2200, "Phase 1 Impl. + Lizenz"),(2200, "Phase 2 Impl. + Lizenz"),
            (1200, "Steady + Weiterentw."),(1000, "Steady State"),(1000, "Steady State")]
    total = sum(v[0] for v in invs)
    add_rect(s, CONTENT_LEFT, CONTENT_TOP, CONTENT_W, 0.70, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.15, CONTENT_TOP, 8, 0.70, [
        ("Gesamt-Investition SEPM 2027-2031 (Base Case)",
         dict(size=12.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, CONTENT_LEFT + 8.2, CONTENT_TOP, 4.4, 0.70, [
        (f"~ {total/1000:.1f} Mio. €",
         dict(size=22, bold=True, color=ORANGE, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)
    chart_top = CONTENT_TOP + 1.0
    chart_h = 3.6
    chart_x = CONTENT_LEFT + 0.30; chart_w = CONTENT_W - 0.60
    n = len(years); slot_w = chart_w / n
    max_v = max(v[0] for v in invs)
    for i, (yr, (val, label)) in enumerate(zip(years, invs)):
        cx = chart_x + i * slot_w + slot_w / 2
        h = (val / max_v) * (chart_h - 0.6)
        color = ORANGE if i < 2 else (ORANGE2 if i < 3 else ORANGE_LT)
        add_rect(s, cx - 0.75, chart_top + chart_h - h - 0.35, 1.50, h, color,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, cx - 1.0, chart_top + chart_h - h - 0.55, 2.0, 0.20, [
            (f"{val/1000:.1f} M€",
             dict(size=10.5, bold=True, color=DARK, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)
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
    ms = [("Q3 2026: Jahresplanung","Beschluss 2027 & Rahmen 2028"),
          ("Q3 2027: Jahresplanung","Freigabe Phase 2"),
          ("Q3 2028: Review","Uebergang Steady State")]
    ms_y = ms_top + 0.28
    m_w = (CONTENT_W - 0.30) / 3
    for i, (t, sub) in enumerate(ms):
        x = CONTENT_LEFT + i * (m_w + 0.15)
        add_rect(s, x, ms_y, m_w, 0.75, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=NAVY, line_w=0.8)
        add_diamond(s, x + 0.10, ms_y + 0.22, 0.28, 0.28, RED_HL)
        add_text(s, x + 0.50, ms_y + 0.04, m_w - 0.60, 0.30, [
            (t, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.50, ms_y + 0.36, m_w - 0.60, 0.35, [
            (sub, dict(size=9, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.MIDDLE)
slide_budget()

# ============================================================
# KAPITEL D — Risiken
# ============================================================

# 19 · Risikomatrix
def slide_risk():
    s = add_slide_stihl("Risiko-Matrix ServiceNow-Plattform-Abhaengigkeit")
    add_text(s, CONTENT_LEFT, CONTENT_TOP, CONTENT_W, 0.25, [
        ("Risiko-Matrix", dict(size=12.5, bold=True, color=ORANGE))])
    hdr_y = CONTENT_TOP + 0.30
    add_rect(s, CONTENT_LEFT, hdr_y, CONTENT_W, 0.28, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.15, hdr_y, 3.3, 0.28, [
        ("Risiko", dict(size=11, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, CONTENT_LEFT + 3.6, hdr_y, 1.5, 0.28, [
        ("Schwere", dict(size=11, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, CONTENT_LEFT + 5.3, hdr_y, 7.5, 0.28, [
        ("Mitigation", dict(size=11, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    for i, (risk, sev, mit) in enumerate(RISKS):
        y = hdr_y + 0.32 + i * 0.60
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        add_rect(s, CONTENT_LEFT, y, CONTENT_W, 0.56, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, CONTENT_LEFT + 0.15, y, 3.3, 0.56, [
            (risk, dict(size=10.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        for k in range(5):
            filled = k < sev
            color = (RED_HL if sev >= 4 else (YELLOW if sev == 3 else GREEN)) if filled else LIGHT_GREY
            add_oval(s, CONTENT_LEFT + 3.60 + k * 0.28, y + 0.22, 0.18, 0.18, color)
        add_text(s, CONTENT_LEFT + 5.30, y, 7.5, 0.56, [
            (mit, dict(size=10, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    sy = hdr_y + 0.32 + len(RISKS) * 0.60 + 0.10
    sh_h = CONTENT_BOT - sy
    add_rect(s, CONTENT_LEFT, sy, CONTENT_W, sh_h, GREEN_LT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, CONTENT_LEFT, sy, 0.07, sh_h, GREEN)
    add_text(s, CONTENT_LEFT + 0.15, sy + 0.04, CONTENT_W - 0.30, 0.30, [
        ("Fazit Risiko-Bewertung",
         dict(size=12, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, CONTENT_LEFT + 0.15, sy + 0.35, CONTENT_W - 0.30, sh_h - 0.40, [
        ("Kein Risiko ist ein Show-Stopper. Alle Risiken sind mit Standard-Vertraegen "
         "(Preisgarantie, DSGVO-ADV, SLA) und Standard-First-Governance beherrschbar. "
         "Vergleich mit SAP: aehnliches Vendor-Lock-in-Profil, kein Alleinstellungsmerkmal von SNOW.",
         dict(size=10, color=DARK))], anchor=MSO_ANCHOR.TOP)
slide_risk()

# 20 · Vertragshistorie + Referenzkunden
def slide_history():
    s = add_slide_stihl("Vertragshistorie SNOW 2014-2026 + Referenzkunden")
    add_text(s, CONTENT_LEFT, CONTENT_TOP, CONTENT_W, 0.25, [
        ("Vertragsstufen bei STIHL und Referenzkunden im Markt",
         dict(size=12, bold=True, color=ORANGE))])
    # Bar chart
    ct = CONTENT_TOP + 0.30
    ch = 2.60
    add_rect(s, CONTENT_LEFT, ct, CONTENT_W, ch, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
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
             dict(size=9, bold=True, color=NAVY, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x - 0.85, ct + ch - 0.32, 1.70, 0.24, [
            (str(yr), dict(size=10, bold=True, color=DARK, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    # Reference customers below
    ref_top = ct + ch + 0.25
    add_text(s, CONTENT_LEFT, ref_top, CONTENT_W, 0.24, [
        ("Deutsche Referenzkunden ServiceNow (Grosskonzerne)",
         dict(size=11.5, bold=True, color=ORANGE))])
    customers = [
        ("SIEMENS", "Industrie / Tech", RGBColor(0x00, 0x9D, 0x9D), WHITE),
        ("SAP", "Software", RGBColor(0x00, 0x3A, 0x5D), WHITE),
        ("Allianz", "Versicherung", RGBColor(0x00, 0x37, 0x81), WHITE),
        ("Deutsche Bank", "Banking", RGBColor(0x00, 0x18, 0xA8), WHITE),
        ("T", "Deutsche Telekom", RGBColor(0xE2, 0x00, 0x74), WHITE),
        ("BOSCH", "Industrie / Auto", RGBColor(0xC8, 0x10, 0x2E), WHITE),
        ("Mercedes-Benz", "Automotive", RGBColor(0x00, 0x00, 0x00), WHITE),
        ("Volkswagen", "Automotive", RGBColor(0x00, 0x1E, 0x50), WHITE),
    ]
    box_top = ref_top + 0.32
    box_h = 0.85
    n = len(customers)
    gap = 0.10
    box_w = (CONTENT_W - (n - 1) * gap) / n
    for i, (name, branche, bg, fg) in enumerate(customers):
        x = CONTENT_LEFT + i * (box_w + gap)
        add_rect(s, x, box_top, box_w, box_h, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        name_size = 14 if len(name) <= 4 else (12 if len(name) <= 8 else 10)
        add_text(s, x, box_top, box_w, box_h, [
            (name, dict(size=name_size, bold=True, color=fg, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x, box_top + box_h + 0.02, box_w, 0.24, [
            (branche, dict(size=8, italic=True, color=GREY_TXT, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.TOP)
    # Fazit note
    fz_y = box_top + box_h + 0.35
    add_rect(s, CONTENT_LEFT, fz_y, CONTENT_W, 0.55, GREEN_LT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, CONTENT_LEFT, fz_y, 0.07, 0.55, GREEN)
    add_text(s, CONTENT_LEFT + 0.15, fz_y, CONTENT_W - 0.25, 0.55, [
        ("Fazit: 12 Jahre bewaehrter STIHL-Partner + breite deutsche Referenzbasis "
         "senken Risiko einer strategischen Fehlentscheidung.",
         dict(size=10, italic=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
slide_history()

# 21 · Lockin + EAM + AI Vibe Coding
def slide_lockin_eam_ai():
    s = add_slide_stihl("Lockin-Effekt · EAM Plattformstrategie · AI Vibe Coding")
    top = CONTENT_TOP
    h = CONTENT_BOT - top
    w = (CONTENT_W - 0.30) / 3
    gap = 0.15
    cols = [
        ("Lockin-Effekt (Vergleich SAP)", RED_HL,
         "Wie stark bindet uns SNOW?",
         [("Vergleich mit SAP",
           "SAP-Lockin bei STIHL seit Jahrzehnten akzeptiert — SNOW-Lockin ist strukturell aehnlich, aber mit deutlich juengerer Historie."),
          ("Datenexport-Faehigkeit",
           "SNOW bietet Standard-Exports (JSON/CSV/APIs). Daten koennen jederzeit gezogen werden."),
          ("Migrations-Pfade",
           "Third-Party-Migrationstools verfuegbar (SPM -> Planisware, PPM Standard-Formate)."),
          ("Vertrags-Klauseln",
           "Ausstiegs-/Kuendigungsklauseln + Preisgarantien pro Renewal-Zyklus vertraglich absichern."),
          ("Fazit",
           "Lockin vorhanden aber im Marktvergleich moderat — kein K.O.-Kriterium.")]),
        ("EAM Plattform-Strategie", NAVY,
         "Wie fuegt sich SNOW in unsere IT-Architektur?",
         [("Konsolidierungs-Ziel",
           "STIHL EAM-Strategie: weniger, aber strategische Plattformen (SAP, SNOW, M365)."),
          ("Plattform-Rolle SNOW",
           "SNOW = Workflow-Plattform fuer bereichs- uebergreifende Prozesse (ITSM heute, SPM/HR/CSM morgen)."),
          ("Interoperabilitaet",
           "Native SAP-Integrationen; klare Rollenverteilung Daten-Systeme (SAP) vs Prozess-Systeme (SNOW)."),
          ("Governance",
           "Architektur-Board-Genehmigung fuer jede neue SNOW-Erweiterung — verhindert Wildwuchs."),
          ("Fazit",
           "SNOW passt strategisch in unsere EAM-Roadmap und reduziert Tool-Vielfalt.")]),
        ("AI Vibe Coding", GREEN,
         "Vendor-Neutralitaet bei GenAI?",
         [("Ist-Zustand SNOW AI",
           "Now Assist als eingebauter AI-Assistent (basiert auf Multi-Model — u.a. Azure OpenAI, eigene LLMs)."),
          ("Vendor-Wahl",
           "SNOW erlaubt Wahl des LLM-Providers (Azure, Google, Anthropic) — kein Zwang zu einem KI-Anbieter."),
          ("Datenhoheit",
           "AI-Verarbeitung optional in EU-Rechenzentren; STIHL-Daten bleiben unter Kontrolle."),
          ("Ergaenzung Multi-AI",
           "STIHL kann parallel eigene KI-Tools (z.B. Copilot, Claude, on-prem) einsetzen — SNOW blockiert das nicht."),
          ("Fazit",
           "AI-Neutralitaet gegeben — kein Lock-in bei KI-Providern.")]),
    ]
    for i, (title, color, sub, items) in enumerate(cols):
        x = CONTENT_LEFT + i * (w + gap)
        add_rect(s, x, top, w, h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, x, top, w, 0.65, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x + 0.10, top + 0.03, w - 0.20, 0.32, [
            (title, dict(size=11.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.10, top + 0.36, w - 0.20, 0.25, [
            (sub, dict(size=9.5, italic=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
        for j, (k, body) in enumerate(items):
            iy = top + 0.75 + j * 1.10
            is_fazit = (k == "Fazit")
            add_rect(s, x + 0.12, iy, w - 0.24, 1.00,
                     GREEN_LT if is_fazit else WHITE,
                     shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                     line=(GREEN if is_fazit else color), line_w=0.5)
            add_rect(s, x + 0.12, iy, 0.06, 1.00, GREEN if is_fazit else color)
            add_text(s, x + 0.25, iy + 0.02, w - 0.40, 0.28, [
                (k, dict(size=10, bold=True,
                         color=GREEN if is_fazit else color))],
                anchor=MSO_ANCHOR.MIDDLE)
            add_text(s, x + 0.25, iy + 0.28, w - 0.40, 0.70, [
                (body, dict(size=9, color=DARK))], anchor=MSO_ANCHOR.TOP)
slide_lockin_eam_ai()

# ============================================================
# KAPITEL E — Empfehlung
# ============================================================

# 22 · ROI & Empfehlung
def slide_roi_recommendation():
    s = add_slide_stihl("ROI-Kalkulation & Empfehlung")
    kpis = [("NPV 5 Jahre","~ +8,5 M€","Base Case, WACC 6 %", GREEN),
            ("IRR","~ 32 %","ueber Investitionshuerde", GREEN),
            ("Payback","~ 3,5 J.","nach Kickoff", ORANGE),
            ("TCO 5 Jahre","~ 7,0 M€","einmalig + laufend", NAVY)]
    top = CONTENT_TOP
    kpi_h = 1.20
    gap = 0.12
    w = (CONTENT_W - 3 * gap) / 4
    for i, (label, big, sub, col) in enumerate(kpis):
        x = CONTENT_LEFT + i * (w + gap)
        add_rect(s, x, top, w, kpi_h, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=col, line_w=1.5)
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
    rows = [("Kosten (TCO)","- 7,0 M€", RED_HL),
            ("Konsolidierung Tools (5 J.)","+ 4,7 M€", GREEN),
            ("Konsolidierung Schnittstellen","+ 1,1 M€", GREEN),
            ("Business Value","+ 11,6 M€", GREEN),
            ("Flankierende Prozesse","+ 5,3 M€", GREEN),
            ("Netto-Effekt 5 Jahre","+ 15,7 M€", DARK)]
    tt = tt_y + 0.30
    for i, (name, val, color) in enumerate(rows):
        y = tt + i * 0.38
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        is_total = (i == len(rows) - 1)
        add_rect(s, CONTENT_LEFT, y, CONTENT_W, 0.34,
                 bg if not is_total else DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, CONTENT_LEFT + 0.15, y + 0.02, 8.0, 0.30, [
            (name, dict(size=11, bold=is_total,
                        color=DARK if not is_total else WHITE))], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, CONTENT_LEFT + 8.5, y + 0.02, CONTENT_W - 8.65, 0.30, [
            (val, dict(size=13 if is_total else 12, bold=True,
                       color=color if not is_total else ORANGE, align=PP_ALIGN.RIGHT))],
            anchor=MSO_ANCHOR.MIDDLE)
    rec_y = tt + len(rows) * 0.38 + 0.15
    rec_h = CONTENT_BOT - rec_y
    add_rect(s, CONTENT_LEFT, rec_y, CONTENT_W, rec_h, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, CONTENT_LEFT + 0.20, rec_y, 3.5, rec_h, [
        ("★ EMPFEHLUNG", dict(size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, CONTENT_LEFT + 3.8, rec_y + 0.10, CONTENT_W - 4.0, rec_h - 0.20, [
        ("Beschluss zur Umsetzung von Szenario A (Full Scope 24 Monate) und "
         "Freigabe Detail-Planung / RFP-Prozess. NPV +8,5 M€, IRR 32 %, Payback ~3,5 J. — "
         "deutlich ueber Investitions-Huerde. Alle Vorstands-Antraege (2, 3, 4, 6) beantwortet.",
         dict(size=11, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
slide_roi_recommendation()

# 23 · Vielen Dank
def slide_thanks():
    s = prs.slides.add_slide(LAYOUT_CONCL)
    for ph in s.placeholders:
        ph.text_frame.text = "Vielen Dank"
        break
slide_thanks()

prs.save(DST_PPTX)
print(f"Saved: {DST_PPTX}  ({len(prs.slides)} slides)")
