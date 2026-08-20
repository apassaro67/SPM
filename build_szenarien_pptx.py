"""Build Szenarien-PowerPoint SNOW SPM homogen vs SNOW+Planisware Kombinationen."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from copy import deepcopy
from lxml import etree

SRC = "/home/user/SPM/2026_05_18_VS_Zielbild_SNOW_SPM_v23.pptx"
DST = "/tmp/2026_10_20_SNOW_Planisware_Szenarien.pptx"

ORANGE   = RGBColor(0xF0, 0x7F, 0x12)
ORANGE_D = RGBColor(0xC0, 0x5F, 0x00)
DARK     = RGBColor(0x1F, 0x29, 0x37)
GREY_TXT = RGBColor(0x4B, 0x55, 0x63)
NAVY     = RGBColor(0x2C, 0x3E, 0x50)
NAVY_LT  = RGBColor(0x3B, 0x52, 0x6B)
GREEN    = RGBColor(0x2E, 0x7D, 0x32)
GREEN_LT = RGBColor(0xDD, 0xF1, 0xDE)
RED      = RGBColor(0xB9, 0x1C, 0x1C)
RED_LT   = RGBColor(0xFB, 0xE0, 0xE0)
YELLOW   = RGBColor(0xE0, 0xA8, 0x0B)
YELLOW_LT= RGBColor(0xFF, 0xF6, 0xD5)
GREY_LT  = RGBColor(0xEF, 0xF1, 0xF4)
GREY_BOX = RGBColor(0xE0, 0xE4, 0xEA)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
BLUE_SN  = RGBColor(0x28, 0x5A, 0xB8)
GREEN_PW = RGBColor(0x0E, 0x8A, 0x63)

prs = Presentation(SRC)
# Remove all existing slides
sldIdLst = prs.slides._sldIdLst
for sld in list(sldIdLst):
    rId = sld.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    sldIdLst.remove(sld)
    try: prs.part.drop_rel(rId)
    except Exception: pass

LAY_TITLE   = prs.slide_layouts[2]   # Title slide empty
LAY_ONLY    = prs.slide_layouts[4]   # Nur Titel
LAY_CONCL   = prs.slide_layouts[18]  # Conclusion
LAY_CHAPTER = prs.slide_layouts[19]  # Chapter separator without picture

def set_text(tf, text, size=18, bold=False, color=DARK, align=PP_ALIGN.LEFT, italic=False, name="Calibri"):
    tf.text = ""
    tf.word_wrap = True
    lines = text.split("\n") if isinstance(text, str) else text
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = ln
        r.font.name = name
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color

def add_box(slide, l, t, w, h, fill=WHITE, line=None, line_w=0.75):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(line_w)
    s.shadow.inherit = False
    return s

def add_round_box(slide, l, t, w, h, fill=WHITE, line=None, line_w=0.75, radius=0.15):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    s.adjustments[0] = radius
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(line_w)
    s.shadow.inherit = False
    return s

def add_text(slide, l, t, w, h, text, size=14, bold=False, color=DARK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.margin_left = Inches(0.05); tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.03); tf.margin_bottom = Inches(0.03)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    set_text(tf, text, size=size, bold=bold, color=color, align=align, italic=italic)
    return tb

def add_bullets(slide, l, t, w, h, items, size=12, color=DARK, bullet="•", line_spacing=1.15):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.margin_left = Inches(0.05); tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.03); tf.margin_bottom = Inches(0.03)
    tf.word_wrap = True
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = line_spacing
        r = p.add_run()
        r.text = f"{bullet}  {it}"
        r.font.name = "Calibri"
        r.font.size = Pt(size)
        r.font.color.rgb = color
    return tb

def add_slide_title(slide, title, subtitle=None):
    add_text(slide, 0.33, 0.30, 12.7, 0.55, title, size=24, bold=True, color=DARK)
    # orange underline
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.33), Inches(0.85), Inches(0.75), Inches(0.06))
    s.fill.solid(); s.fill.fore_color.rgb = ORANGE
    s.line.fill.background(); s.shadow.inherit = False
    if subtitle:
        add_text(slide, 0.33, 0.92, 12.7, 0.36, subtitle, size=12, color=GREY_TXT, italic=True)

# ==========================================================
# SLIDE 1 — Titel
# ==========================================================
s1 = prs.slides.add_slide(LAY_TITLE)
add_text(s1, 0.75, 2.3, 11.8, 1.1,
         "Evaluierung Plattform-Szenarien für PPM",
         size=36, bold=True, color=DARK)
add_text(s1, 0.75, 3.4, 11.8, 0.7,
         "ServiceNow SPM homogen vs. Kombination mit Planisware",
         size=22, color=ORANGE)
add_text(s1, 0.75, 4.25, 11.8, 0.5,
         "Entscheidungsvorlage für Team-Evaluierung und Vorstands-Empfehlung",
         size=14, color=GREY_TXT, italic=True)
add_text(s1, 0.75, 6.55, 11.8, 0.4,
         "Alex Passaro · Torsten Zahn  ·  20.10.2026",
         size=11, color=GREY_TXT)

# ==========================================================
# SLIDE 2 — Ausgangslage / Rahmen
# ==========================================================
s2 = prs.slides.add_slide(LAY_ONLY)
add_slide_title(s2, "Ausgangslage & Rahmen der Evaluierung",
                "Was ist gesetzt, was steht zur Entscheidung")

# 2 columns: gesetzt / zu entscheiden
add_round_box(s2, 0.33, 1.35, 6.30, 5.55, fill=GREEN_LT, line=GREEN, line_w=1.0, radius=0.08)
add_text(s2, 0.55, 1.50, 6.0, 0.45, "✅  Gesetzt (nicht Teil dieser Evaluierung)",
         size=14, bold=True, color=GREEN)
add_bullets(s2, 0.55, 1.95, 6.0, 4.90, [
    "ServiceNow bleibt strategische Plattform (seit 2014)",
    "ServiceNow SPM als Modul für Strategy / OKR",
    "ServiceNow SPM als Modul für Demand Management",
    "ServiceNow SPM als Modul für Projekt- und Produktportfolio (Roadmap)",
    "SNOW Demand-Management ist bereits produktiv aktiv",
    "Ablöse des Altsystems (PIT/MSPO) ist beschlossen",
    "Funktionale Eignung SNOW für PM (inkl. PEP, Ressourcen,\n    Restwert, SAP-Obligo) wurde bereits nachgewiesen —\n    Massenänderungen per Low-Code/App lösbar",
], size=12)

add_round_box(s2, 6.70, 1.35, 6.30, 5.55, fill=YELLOW_LT, line=YELLOW, line_w=1.0, radius=0.08)
add_text(s2, 6.92, 1.50, 6.0, 0.45, "❓  Zu entscheiden in dieser Evaluierung",
         size=14, bold=True, color=RGBColor(0xB8, 0x8A, 0x00))
add_bullets(s2, 6.92, 1.95, 6.0, 4.90, [
    "Welche Plattform trägt das Projekt-Management?",
    "  → SNOW SPM homogen für alles?",
    "  → SNOW + Planisware nur für Entwicklung/Produktion?",
    "  → SNOW + Planisware weltweit für alle Fachbereiche?",
    "  → Alternativer Schnitt (z.B. nur ausgewählter PEP-Subset)?",
    "Zielbild für Betriebsmodell (Team, Support, Verträge)",
    "Wirtschaftliche Rahmung (Lizenz, Implementierung, TCO 5J)",
], size=12)

add_text(s2, 0.33, 7.05, 12.7, 0.30,
         "Fokus dieser Präsentation: Beschreibung der Szenarien, Vor-/Nachteile, gemeinsame Scorecard für die Bewertung.",
         size=10, italic=True, color=GREY_TXT, align=PP_ALIGN.LEFT)

# ==========================================================
# SLIDE 3 — Szenario-Übersicht (visuelle Karte)
# ==========================================================
s3 = prs.slides.add_slide(LAY_ONLY)
add_slide_title(s3, "Vier bewertete Szenarien im Überblick",
                "Gemeinsamer Kern: SNOW für Strategy · Portfolio · Demand · Roadmap")

# 4 columns, one per scenario
scenarios_hdr = [
    ("SZ 1",  "SNOW homogen",           BLUE_SN,   "Ein System für alles"),
    ("SZ 2a", "SNOW + Planisware\n(PEP-Subset)", ORANGE,  "Zusatzoption: nur ausgewählte NPD-Projekte"),
    ("SZ 2",  "SNOW + Planisware\n(Entwicklung/Produktion)",   NAVY_LT,   "PW für gesamten E&P-Bereich"),
    ("SZ 3",  "SNOW + Planisware\n(weltweit alle FB)", GREEN_PW,   "PW als Konzern-PM-Tool"),
]
box_w = 3.05; box_gap = 0.11; base_l = 0.33
for i, (tag, name, col, sub) in enumerate(scenarios_hdr):
    l = base_l + i * (box_w + box_gap)
    add_round_box(s3, l, 1.35, box_w, 0.85, fill=col, line=col)
    add_text(s3, l, 1.42, box_w, 0.35, tag, size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s3, l, 1.68, box_w, 0.50, name, size=11.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Content rows: SNOW-Anteil / Planisware-Anteil / typischer User
rows = [
    ("Strategy / OKR",              ["SNOW", "SNOW", "SNOW", "SNOW"]),
    ("Portfolio & Roadmap",         ["SNOW", "SNOW", "SNOW", "SNOW"]),
    ("Demand Management",           ["SNOW", "SNOW", "SNOW", "SNOW"]),
    ("PM IT / Business / HR / Compliance", ["SNOW", "SNOW", "SNOW", "PW"]),
    ("PM Entwicklung & Produktion", ["SNOW", "SNOW*", "PW", "PW"]),
    ("Zeiterfassung / Ressourcen",  ["SNOW", "gemischt", "gemischt", "PW"]),
    ("Anzahl PPM-Systeme im Konzern", ["1", "1 + Add-on", "2", "2"]),
]
r_top = 2.35; r_h = 0.42; label_w = 3.20
for ri, (lbl, cells) in enumerate(rows):
    y = r_top + ri * r_h
    fill = GREY_LT if ri % 2 == 0 else WHITE
    add_box(s3, base_l, y, label_w, r_h, fill=fill, line=RGBColor(0xD1,0xD5,0xDB), line_w=0.5)
    add_text(s3, base_l+0.10, y, label_w-0.12, r_h, lbl, size=11, bold=True, color=DARK,
             anchor=MSO_ANCHOR.MIDDLE)
    for ci, v in enumerate(cells):
        cx = base_l + label_w + 0.10 + ci * ((4*box_w + 3*box_gap - label_w - 0.10)/4)
        cw = (4*box_w + 3*box_gap - label_w - 0.10)/4 - 0.05
        # Actually align with header columns
        cx = base_l + (ci+1) * (box_w + box_gap) - (box_w * 0.03)
        cx = base_l + label_w + 0.02 + ci * ((4*(box_w+box_gap) - label_w - 0.02)/4)
        cw = (4*(box_w+box_gap) - label_w - 0.02)/4 - 0.05
        cell_fill = fill
        text_col = DARK
        if v == "SNOW":
            cell_fill = RGBColor(0xE3, 0xEF, 0xFB); text_col = BLUE_SN
        elif v == "PW":
            cell_fill = RGBColor(0xE0, 0xF3, 0xEB); text_col = GREEN_PW
        elif v == "SNOW*":
            cell_fill = RGBColor(0xFD, 0xE6, 0xCC); text_col = ORANGE_D
        elif v == "gemischt":
            cell_fill = RGBColor(0xFF, 0xF3, 0xC4); text_col = RGBColor(0x8A, 0x63, 0x00)
        add_box(s3, cx, y, cw, r_h, fill=cell_fill, line=RGBColor(0xD1,0xD5,0xDB), line_w=0.5)
        add_text(s3, cx, y, cw, r_h, v, size=11, bold=True, color=text_col,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_text(s3, 0.33, 6.95, 12.7, 0.28,
         "* SZ 2a nur ausgewählte Neuprodukt-Entwicklungen (Filter: Größe, Komplexität) in Planisware — alle anderen E&P-Projekte in SNOW.",
         size=9.5, italic=True, color=GREY_TXT)
add_text(s3, 0.33, 7.20, 12.7, 0.28,
         "SNOW = ServiceNow SPM Pro   ·   PW = Planisware Orchestra   ·   In allen Szenarien: SNOW trägt Strategy/Portfolio/Demand.",
         size=9.5, italic=True, color=GREY_TXT)

# ==========================================================
# SLIDES 4-7 — Detail je Szenario (Beschreibung + Pro + Contra)
# ==========================================================
def scenario_detail(tag, title, desc_bullets, pros, cons, kpis, color):
    s = prs.slides.add_slide(LAY_ONLY)
    add_slide_title(s, f"{tag} — {title}",
                    "Beschreibung · Vorteile · Nachteile · Kern-KPIs")

    # Left: Beschreibung
    add_round_box(s, 0.33, 1.35, 4.15, 5.60, fill=WHITE, line=color, line_w=1.5, radius=0.05)
    add_box(s, 0.33, 1.35, 4.15, 0.42, fill=color, line=color)
    add_text(s, 0.33, 1.35, 4.15, 0.42, "Beschreibung",
             size=12.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_bullets(s, 0.50, 1.90, 3.85, 5.00, desc_bullets, size=11)

    # Middle: Pro
    add_round_box(s, 4.58, 1.35, 4.15, 5.60, fill=WHITE, line=GREEN, line_w=1.5, radius=0.05)
    add_box(s, 4.58, 1.35, 4.15, 0.42, fill=GREEN, line=GREEN)
    add_text(s, 4.58, 1.35, 4.15, 0.42, "✔  Vorteile",
             size=12.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_bullets(s, 4.75, 1.90, 3.85, 5.00, pros, size=11)

    # Right: Contra
    add_round_box(s, 8.83, 1.35, 4.15, 5.60, fill=WHITE, line=RED, line_w=1.5, radius=0.05)
    add_box(s, 8.83, 1.35, 4.15, 0.42, fill=RED, line=RED)
    add_text(s, 8.83, 1.35, 4.15, 0.42, "✘  Nachteile",
             size=12.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_bullets(s, 9.00, 1.90, 3.85, 5.00, cons, size=11)

    # KPI strip bottom
    add_round_box(s, 0.33, 7.05, 12.65, 0.32, fill=RGBColor(0xF7,0xF7,0xF7), line=None)
    add_text(s, 0.45, 7.05, 12.5, 0.32, "  ·  ".join(kpis),
             size=10, bold=True, color=DARK, anchor=MSO_ANCHOR.MIDDLE)

scenario_detail(
    "SZ 1", "ServiceNow SPM homogen für alles",
    desc_bullets=[
        "Ein System — SNOW SPM Pro — für Strategy, Portfolio, Demand, Roadmap UND das gesamte Projekt-Management",
        "Alle Fachbereiche (IT, Business, HR, Compliance, Entwicklung, Produktion) arbeiten in der gleichen Plattform",
        "PEP im E&P-Bereich wird per Standard-SPM + App Engine (Low-Code) abgebildet",
        "Massenänderungen über App/Low-Code-Ergänzung",
        "Ein Betriebsteam, ein Lizenzvertrag, eine Roadmap",
    ],
    pros=[
        "Datenkontinuität End-to-End (Strategie → Ergebnis)",
        "Niedrigste TCO (kein zweiter Vendor)",
        "Time-to-Value am kürzesten",
        "Ein UI, ein Support-Team, ein Training",
        "Now Assist (KI) über die gesamte Prozesskette",
        "Klare Governance und Reporting-Basis",
    ],
    cons=[
        "PEP-Spezifika im E&P über App Engine bleiben Custom-Aufwand (Wartung, Roadmap-Abhängigkeit)",
        "Bei sehr komplexen R&D-Simulationen weniger reif als spezialisiertes Tool",
        "Massenänderungen benötigen Low-Code-Ergänzung",
        "Vendor-Konzentration auf ServiceNow (Klumpenrisiko)",
    ],
    kpis=["Systeme: 1", "Vendors: 1", "Betriebs-FTE (est.): 4–5", "TCO 5J (Richtwert): ≈ Basis"],
    color=BLUE_SN,
)

scenario_detail(
    "SZ 2a", "SNOW + Planisware nur für definierten PEP-Subset (Vorschlag)",
    desc_bullets=[
        "SNOW SPM Pro trägt Strategy, Portfolio, Demand, Roadmap sowie PM in IT/Business/HR/Compliance UND den Großteil E&P",
        "Planisware Orchestra kommt nur für einen klar definierten Subset zum Einsatz: z.B. NPD-Projekte > X Mio. Budget oder > Y parallele Ressourcen",
        "Filter-/Übergabeprozess von SNOW zu Planisware definiert",
        "Portfolio-Sicht bleibt in SNOW — Planisware liefert Portfolio-Rohdaten zurück",
    ],
    pros=[
        "Best-of-Breed nur dort, wo tatsächlich benötigt",
        "Sehr begrenztes zweites Lizenz-/Betriebsmodell",
        "Behält SNOW-Vorteile für 80–90% der Projekte",
        "Reduziert Klumpenrisiko moderat",
        "Klare Filterkriterien möglich (Governance)",
    ],
    cons=[
        "Zweiter Vendor trotz kleinem Scope",
        "Schnittstelle SNOW ↔ Planisware muss gebaut & betrieben werden",
        "Zwei UIs für E&P-Anwender bei Grenzprojekten",
        'Filter-Diskussion („was ist Subset?") politisch anspruchsvoll',
        "Zusätzliche Trainings-/Support-Kosten für kleinen Nutzerkreis",
    ],
    kpis=["Systeme: 2", "Vendors: 2", "Betriebs-FTE (est.): 5–6", "TCO 5J (Richtwert): +0,3–0,6 M€"],
    color=ORANGE,
)

scenario_detail(
    "SZ 2", "SNOW + Planisware für Entwicklung & Produktion",
    desc_bullets=[
        "SNOW SPM Pro trägt Strategy, Portfolio, Demand, Roadmap sowie PM in IT/Business/HR/Compliance",
        "Planisware Orchestra übernimmt komplettes PM für die Bereiche Entwicklung und Produktion",
        "SNOW → Planisware: Übergabe Demand+Portfolio-Rahmen, Planisware liefert Projekt-Ergebnisse zurück",
        "Zwei parallele Betriebs- und Support-Modelle",
    ],
    pros=[
        "Best-in-Class PM-Tool im wichtigsten Wertschöpfungsbereich",
        "Planisware-Reife für Stage-Gate / PEP direkt nutzbar",
        "SNOW-Nutzen (Integration, Now Assist) bleibt für Rest bestehen",
        "Reduziert Klumpenrisiko",
        "Klare fachliche Trennung (E&P vs. Rest)",
    ],
    cons=[
        "Zweiter Vollvendor mit signifikantem Lizenz-/Betriebsvolumen",
        "Umfangreiche Schnittstelle SNOW ↔ PW notwendig",
        "Zwei UIs, zwei Support-Prozesse, zwei Roadmaps",
        "Datenqualität an der Schnittstelle als Dauerthema",
        "Bereits nachgewiesene SNOW-Eignung für PEP wird nicht genutzt → höhere Kosten ohne klaren Fach-Mehrwert",
        "Change-Aufwand E&P: neuer Vendor, neue UI, neuer Support",
    ],
    kpis=["Systeme: 2", "Vendors: 2", "Betriebs-FTE (est.): 7–9", "TCO 5J (Richtwert): +1,5–2,5 M€"],
    color=NAVY_LT,
)

scenario_detail(
    "SZ 3", "SNOW + Planisware weltweit für alle Fachbereiche",
    desc_bullets=[
        "SNOW SPM Pro trägt Strategy, Portfolio, Demand, Roadmap",
        "Planisware Orchestra übernimmt komplettes PM für alle Fachbereiche weltweit (IT, Business, HR, Compliance, E&P, ...)",
        "Alle projektbezogenen Aktivitäten in Planisware; SNOW-SPM nur oberhalb der Projektebene",
        "Zwei Vendors, zwei Betriebsteams, zwei Trainingsprogramme weltweit",
    ],
    pros=[
        "Eine einheitliche PM-Welt für alle Anwender",
        "Best-in-Class PM-Reife über den ganzen Konzern",
        "Fachliche Trennung SPM-oben / PM-unten sauber",
        "Planisware-Reporting konsistent für alle Bereiche",
    ],
    cons=[
        "Höchste TCO (Lizenzen + Implementierung + Betrieb + Change weltweit)",
        "Ignoriert bereits nachgewiesene SNOW-Eignung für PM",
        "SNOW SPM-Investition wird um wichtigen Nutzen beschnitten",
        "Umfangreichster Change-Aufwand aller Szenarien",
        "Umfangreichste globale Rollout-Komplexität",
        "Massives Klumpenrisiko am zweiten Vendor",
        "Doppelte KI-Investition (Now Assist + Oscar)",
    ],
    kpis=["Systeme: 2", "Vendors: 2", "Betriebs-FTE (est.): 10–14", "TCO 5J (Richtwert): +4–7 M€"],
    color=GREEN_PW,
)

# ==========================================================
# SLIDE — Alternative Schnitt-Vorschlag (Erklärung SZ 2a)
# ==========================================================
sX = prs.slides.add_slide(LAY_ONLY)
add_slide_title(sX, "Vorschlag: Alternativer Schnitt (Szenario 2a)",
                "Ergänzung zu den drei Basis-Szenarien — für die Team-Diskussion")

add_round_box(sX, 0.33, 1.35, 12.65, 0.55, fill=ORANGE_LT_HEX if False else RGBColor(0xFD,0xE6,0xCC), line=ORANGE, line_w=1.0, radius=0.20)
add_text(sX, 0.55, 1.35, 12.35, 0.55, "Grundgedanke: Trenn-Kriterium ist nicht der Fachbereich, sondern das einzelne Projekt.",
         size=13, bold=True, color=ORANGE_D, anchor=MSO_ANCHOR.MIDDLE)

add_text(sX, 0.33, 2.05, 12.65, 0.45,
         "SNOW ist Standard. Planisware wird ausschließlich für Projekte aktiviert, die harte, objektive Filterkriterien erfüllen:",
         size=12.5, color=DARK)

# Filter criteria left
add_round_box(sX, 0.33, 2.60, 6.30, 3.20, fill=WHITE, line=ORANGE, line_w=1.2, radius=0.05)
add_box(sX, 0.33, 2.60, 6.30, 0.40, fill=ORANGE, line=ORANGE)
add_text(sX, 0.33, 2.60, 6.30, 0.40, "Beispiel-Filter (in Detail-Phase zu schärfen)",
         size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_bullets(sX, 0.50, 3.10, 6.00, 2.65, [
    "Budget > 5 Mio. € über gesamte Laufzeit",
    "Laufzeit > 24 Monate",
    "> 30 parallele Ressourcen aus mind. 3 Bereichen",
    "Mind. 5 Stage-Gates (Standard-PEP)",
    "Regulatorische R&D-Dokumentationspflicht",
    "Physisches Produkt mit BOM/PLM-Anbindung",
], size=11.5)

# Right: Nutzen
add_round_box(sX, 6.70, 2.60, 6.30, 3.20, fill=WHITE, line=GREEN, line_w=1.2, radius=0.05)
add_box(sX, 6.70, 2.60, 6.30, 0.40, fill=GREEN, line=GREEN)
add_text(sX, 6.70, 2.60, 6.30, 0.40, "Erwarteter Nutzen dieses Schnitts",
         size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_bullets(sX, 6.87, 3.10, 6.00, 2.65, [
    "Planisware nur für 10–20 wirklich komplexe Projekte",
    "80–90% aller Projekte bleiben in SNOW",
    "Deutlich niedrigere PW-Lizenzkosten",
    "Reduziertes Change-/Trainings-Volumen",
    'Klare, messbare Governance („warum ist es in PW?")',
    "Behält SNOW-Investition und Plattform-Nutzen",
], size=11.5)

# Bottom warnings
add_round_box(sX, 0.33, 5.95, 12.65, 1.05, fill=YELLOW_LT, line=YELLOW, line_w=0.8, radius=0.10)
add_text(sX, 0.55, 6.00, 12.20, 0.30, "⚠  Zu klären in der Detail-Phase",
         size=12, bold=True, color=RGBColor(0x8A,0x63,0x00))
add_bullets(sX, 0.55, 6.30, 12.20, 0.65, [
    'Wer entscheidet: Projekt „gehört" in PW oder SNOW? — Governance-Prozess nötig.',
    "Übergangs-Fälle (Projekt startet klein, wächst über Schwellwert): Migration in PW oder Verbleib?  ·  Schnittstellen-Anforderungen konkret aus Beispiel-Projekten ableiten.",
], size=10.5)

# ==========================================================
# SLIDE — Bewertungsmethodik + Kriteriendimensionen (aus Excel)
# ==========================================================
sM = prs.slides.add_slide(LAY_ONLY)
add_slide_title(sM, "Bewertungsmethodik der Scorecard",
                "6 Dimensionen · gewichtete Kriterien · Score 1 – 5 · gewichtete Gesamtsumme")

dims = [
    ("Strategisch",        "20%", "Fit zur Konzern-IT-Strategie, Plattform-Ansatz, Vendor-Landschaft"),
    ("Funktional",         "20%", "PM-Tiefe, PEP/Stage-Gate, Ressourcen-, Finanz-Management"),
    ("Technisch",          "15%", "Integration, Schnittstellen, Extensibility, KI/GenAI"),
    ("Wirtschaftlich",     "20%", "Lizenzen, Implementierung, Betrieb, TCO 5J"),
    ("Organisatorisch",    "15%", "Change, Training, Betrieb, Vendor-Management"),
    ("Risiko/Zukunft",     "10%", "Klumpenrisiko, Roadmap, Skalierung, Migrationsrisiko"),
]

col_w = (12.65) / 3
box_h = 1.55
for i, (name, wt, desc) in enumerate(dims):
    row = i // 3; col = i % 3
    l = 0.33 + col * col_w
    t = 1.55 + row * (box_h + 0.30)
    add_round_box(sM, l+0.05, t, col_w-0.10, box_h, fill=WHITE, line=NAVY, line_w=1.0, radius=0.06)
    add_box(sM, l+0.05, t, col_w-0.10, 0.40, fill=NAVY, line=NAVY)
    add_text(sM, l+0.15, t, col_w-0.30, 0.40, name, size=13, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_text(sM, l+col_w-0.60, t, 0.45, 0.40, wt, size=12, bold=True, color=WHITE,
             align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    add_text(sM, l+0.15, t+0.48, col_w-0.30, box_h-0.55, desc, size=11, color=GREY_TXT)

add_round_box(sM, 0.33, 6.55, 12.65, 0.55, fill=GREY_LT, line=None, radius=0.15)
add_text(sM, 0.45, 6.55, 12.50, 0.55,
         "Score-Skala je Kriterium:  1 = sehr schlecht  ·  2 = schlecht  ·  3 = ok  ·  4 = gut  ·  5 = sehr gut     |     Gewichtungen in Excel anpassbar.",
         size=11, bold=False, color=DARK, anchor=MSO_ANCHOR.MIDDLE)

# ==========================================================
# SLIDE — Zusammenfassungs-Karte / Empfehlung
# ==========================================================
sZ = prs.slides.add_slide(LAY_ONLY)
add_slide_title(sZ, "Vorläufige Einschätzung & Empfehlung für die Team-Diskussion",
                "Ausgangs-Scoring — die finale Bewertung erfolgt im Team über die Excel-Scorecard")

# Ranking rows
rank = [
    ("1", "SZ 1  — SNOW homogen für alles",                     GREEN,   "Empfehlung",        "Bestes Kosten-Nutzen-Verhältnis, klarster Plattform-Case; PEP-Restrisiko durch App Engine adressiert"),
    ("2", "SZ 2a — SNOW + PW nur für definierten PEP-Subset",   YELLOW,  "Bedingt tragfähig", "Sinnvolle Absicherung, wenn App-Engine-PEP nach Phase 1 als unzureichend bewertet wird"),
    ("3", "SZ 2  — SNOW + PW für gesamte E&P",                  ORANGE,  "Nicht empfohlen",   "Wirtschaftlich nur tragbar, wenn PEP-Fach-Nutzen den Zusatzaufwand rechtfertigt — aktuell nicht belegt"),
    ("4", "SZ 3  — SNOW + PW weltweit alle FB",                 RED,     "Nicht empfohlen",   "TCO und Change-Aufwand nicht durch Fach-Nutzen begründet; ignoriert nachgewiesene SNOW-Eignung"),
]

r_top = 1.55; r_h = 1.20
for i, (rk, name, col, verdict, note) in enumerate(rank):
    y = r_top + i * (r_h + 0.10)
    add_round_box(sZ, 0.33, y, 12.65, r_h, fill=WHITE, line=col, line_w=1.5, radius=0.05)
    add_box(sZ, 0.33, y, 0.75, r_h, fill=col, line=col)
    add_text(sZ, 0.33, y, 0.75, r_h, rk, size=28, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(sZ, 1.20, y+0.10, 8.10, 0.45, name, size=14, bold=True, color=DARK)
    add_text(sZ, 1.20, y+0.55, 8.10, r_h-0.60, note, size=11, color=GREY_TXT)
    add_round_box(sZ, 9.50, y+0.30, 3.30, 0.60, fill=col, line=col, radius=0.20)
    add_text(sZ, 9.50, y+0.30, 3.30, 0.60, verdict, size=13, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_text(sZ, 0.33, 7.05, 12.65, 0.32,
         "Diese Einschätzung ist Startpunkt — die verbindliche Reihenfolge ergibt sich aus der Team-Scorecard (Excel, beigefügt).",
         size=10, italic=True, color=GREY_TXT, align=PP_ALIGN.CENTER)

# ==========================================================
# SLIDE — Nächste Schritte / Team-Arbeitsweise
# ==========================================================
sN = prs.slides.add_slide(LAY_ONLY)
add_slide_title(sN, "Vorgehen zur Team-Bewertung",
                "Wie die Scorecard ausgefüllt und ausgewertet wird")

steps = [
    ("1", "Individuell scoren", "Jeder Teilnehmer bewertet alle 4 Szenarien in allen Kriterien (1–5) — ohne vorherige Absprache."),
    ("2", "Gewichte diskutieren", "Team-Diskussion: Sind die vorgeschlagenen Gewichte je Dimension passend? Anpassung im Excel möglich."),
    ("3", "Deltas identifizieren", "Wo weichen die Einzelbewertungen stark ab? Diese Kriterien im Team klären (Faktencheck, ggf. Nachbewertung)."),
    ("4", "Konsens-Score bilden", "Team einigt sich auf einen konsolidierten Score je Zelle. Ergebnis ist die Team-Empfehlung."),
    ("5", "Sensitivität testen", "Was ändert sich, wenn Gewichte ±10% variieren? Robuste Empfehlung überlebt Sensitivität."),
    ("6", "Vorstands-Vorlage", "Ergebnis der Scorecard als objektive Basis für die Entscheidungsvorlage an den Vorstand."),
]

col_w2 = 4.20
for i, (n, title, desc) in enumerate(steps):
    row = i // 3; col = i % 3
    l = 0.33 + col * (col_w2 + 0.10)
    t = 1.55 + row * 2.55
    add_round_box(sN, l, t, col_w2, 2.40, fill=WHITE, line=NAVY, line_w=1.0, radius=0.04)
    # circle number
    c = sN.shapes.add_shape(MSO_SHAPE.OVAL, Inches(l+0.20), Inches(t+0.20), Inches(0.70), Inches(0.70))
    c.fill.solid(); c.fill.fore_color.rgb = ORANGE; c.line.fill.background(); c.shadow.inherit=False
    tf = c.text_frame; tf.margin_left = Inches(0.0); tf.margin_right = Inches(0.0)
    set_text(tf, n, size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    add_text(sN, l+1.00, t+0.25, col_w2-1.10, 0.60, title, size=13, bold=True, color=DARK)
    add_text(sN, l+0.20, t+1.00, col_w2-0.35, 1.30, desc, size=11, color=GREY_TXT)

add_text(sN, 0.33, 7.10, 12.65, 0.30,
         "→ Beilage: Excel-Scorecard   ·   Zeitrahmen Vorschlag: 2 Wochen individuelle Bewertung, 1 Konsens-Workshop (½ Tag), 1 Woche Vorstands-Aufbereitung.",
         size=10.5, italic=True, color=GREY_TXT)

# ==========================================================
# CLOSING SLIDE
# ==========================================================
sC = prs.slides.add_slide(LAY_CONCL)
add_text(sC, 0.75, 2.5, 11.8, 1.2, "Vielen Dank",
         size=52, bold=True, color=DARK)
add_text(sC, 0.75, 3.7, 11.8, 0.6,
         "Diskussion, Ergänzung der Filterkriterien und Start Individual-Scoring",
         size=18, color=ORANGE)
add_text(sC, 0.75, 6.55, 11.8, 0.4,
         "Alex Passaro · Torsten Zahn  ·  20.10.2026",
         size=11, color=GREY_TXT)

prs.save(DST)
print(f"Saved: {DST} ({len(prs.slides)} slides)")
