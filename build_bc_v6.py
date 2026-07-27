"""Business Case SEPM v4 — Struktur nach 4 Vorstands-Aufgaben (Protokoll)."""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

SRC = "/home/user/SPM/2026_07_01_Business_Case_SEPM_STIHL_v2.pptx"
DST = "/tmp/2026_07_08_Business_Case_SEPM_v6.pptx"

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
PURPLE   = RGBColor(0x86, 0x5D, 0xE4)
GREEN_MID= RGBColor(0x66, 0xA6, 0x5A)
YELLOW_GRN= RGBColor(0xB4, 0xC7, 0x39)
GREY_NA  = RGBColor(0x9C, 0xA3, 0xAF)
POT_COLOR = {"Sehr Hoch": GREEN, "Hoch": GREEN_MID, "Mittel-Hoch": YELLOW_GRN,
             "Mittel": ORANGE, "Gering": RED_HL, "na": GREY_NA}

INCH = 914400
CT, CB, CL, CR = 1.05, 7.05, 0.33, 13.00
CW = CR - CL

prs = Presentation(SRC)
sldIdLst = prs.slides._sldIdLst
for sld in list(sldIdLst):
    rId = sld.get(qn("r:id"))
    sldIdLst.remove(sld)
    try: prs.part.drop_rel(rId)
    except Exception: pass

LC = prs.slide_masters[0].slide_layouts[2]
LT = prs.slide_masters[0].slide_layouts[4]
LConcl = prs.slide_masters[0].slide_layouts[18]

def R(s, l, t, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE, lw=1.0):
    x = s.shapes.add_shape(shape, int(l*INCH), int(t*INCH), int(w*INCH), int(h*INCH))
    x.fill.solid(); x.fill.fore_color.rgb = fill
    if line is None: x.line.fill.background()
    else: x.line.color.rgb = line; x.line.width = Pt(lw)
    x.shadow.inherit = False
    return x

def O(s, l, t, w, h, fill):
    x = s.shapes.add_shape(MSO_SHAPE.OVAL, int(l*INCH), int(t*INCH), int(w*INCH), int(h*INCH))
    x.fill.solid(); x.fill.fore_color.rgb = fill
    x.line.fill.background()
    return x

def D(s, l, t, w, h, fill):
    x = s.shapes.add_shape(MSO_SHAPE.DIAMOND, int(l*INCH), int(t*INCH), int(w*INCH), int(h*INCH))
    x.fill.solid(); x.fill.fore_color.rgb = fill
    x.line.fill.background()
    return x

def SP(sh, paras, ds=10, dc=DARK, da=PP_ALIGN.LEFT):
    tf = sh.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(30000); tf.margin_right = Emu(30000)
    tf.margin_top = Emu(10000); tf.margin_bottom = Emu(10000)
    tf.clear()
    for i, it in enumerate(paras):
        if isinstance(it, str): text, st = it, {}
        else: text, st = it
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = st.get("align", da)
        r = p.add_run(); r.text = text
        r.font.name = "Calibri"
        r.font.size = Pt(st.get("size", ds))
        r.font.bold = st.get("bold", False)
        r.font.italic = st.get("italic", False)
        r.font.color.rgb = st.get("color", dc)

def T(s, l, t, w, h, paras, anchor=MSO_ANCHOR.MIDDLE, **kw):
    tb = s.shapes.add_textbox(int(l*INCH), int(t*INCH), int(w*INCH), int(h*INCH))
    tb.fill.background(); tb.line.fill.background()
    tb.text_frame.vertical_anchor = anchor
    SP(tb, paras, **kw)
    return tb

def placeholder(s, idx, text):
    for ph in s.placeholders:
        if ph.placeholder_format.idx == idx:
            ph.text_frame.text = text
            return

def add_stihl(title):
    s = prs.slides.add_slide(LT)
    placeholder(s, 0, title)
    return s

# Bewertung mapping
BEW_COLOR = {"niedrig": GREEN, "mittel": ORANGE, "hoch": RED_HL,
             "gering": GREEN, "hoch": RED_HL}

# ================================
# 1 · COVER
# ================================
def s_cover():
    s = prs.slides.add_slide(LC)
    placeholder(s, 0, "Business Case SEPM")
    placeholder(s, 13, "Bewertung der Plattform-Strategie mit ServiceNow · nach 4 Vorstands-Aufgaben")
    placeholder(s, 14, "Alex Passaro · Torsten Zahn")
    placeholder(s, 15, "Stand: 08.07.2026")
s_cover()

# ================================
# 2 · INHALTSVERZEICHNIS
# ================================
def s_toc():
    s = add_stihl("Inhaltsverzeichnis")
    chapters = [
        ("A1", "AUFGABE 1 — Massnahmen-Mgmt 'Hype'",
         [(4, "Hype = OUT OF SCOPE — Begruendung", RED_HL)]),
        ("A2", "AUFGABE 2 — SNOW-Prozessabdeckung  (20 Schluesselprozesse)",
         [(5, "Schluesselprozesse @ STIHL — Uebersicht",  NAVY),
          (6, "Bewertungsergebnis + Harmonisierungspotenzial", NAVY),
          (7, "Top-Konsolidierungs-Chancen ('Sehr Hoch' + 'Hoch')", NAVY),
          (8, "Alternativen (mit Nachteilen)",              NAVY)]),
        ("A3", "AUFGABE 3 — Governance-Strukturen",
         [(9,  "Governance-Modelle A / B / C",             NAVY),
          (10, "Was aendert sich pro Schluesselprozess",   NAVY),
          (11, "Benefits (Zeit · Qualitaet · Kosten)",     NAVY),
          (12, "Einmalaufwaende + Betriebsthemen",         NAVY)]),
        ("A4", "AUFGABE 4 — Plattform-Abhaengigkeit",
         [(13, "Risikomatrix neu (BoB vs. Plattform)",     NAVY),
          (14, "Lockin Pro/Con + Anpassungsaufwaende",     NAVY),
          (15, "SAP-SNOW Vergleich + Vertragshistorie",    NAVY),
          (16, "Referenzkunden",                            NAVY)]),
        ("G",  "GESAMTDARSTELLUNG",
         [(17, "Qualitative Analyse (5 Dimensionen)",      ORANGE),
          (18, "Strategische Ausrichtung / neue Geschaeftsfelder", ORANGE),
          (19, "TCO und Betrieb",                          ORANGE),
          (20, "Umsetzungsvorschlag Szenario 1 + 2",       ORANGE),
          (21, "Empfehlung + Entscheidungsvorlage",        ORANGE)]),
    ]
    top = CT
    lw = (CW - 0.30) / 2
    rx = CL + lw + 0.30
    yc = [top, top]
    cx = [CL, rx]
    for i, (letter, title, items) in enumerate(chapters):
        col = 0 if i < 3 else 1
        x = cx[col]
        y = yc[col]
        h = 0.50 + len(items) * 0.32 + 0.12
        R(s, x, y, lw, h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        R(s, x, y, lw, 0.40, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        R(s, x + 0.08, y + 0.06, 0.36, 0.28, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, x + 0.08, y + 0.06, 0.36, 0.28, [
            (letter, dict(size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER))])
        T(s, x + 0.52, y + 0.04, lw - 0.60, 0.32, [
            (title, dict(size=11.5, bold=True, color=WHITE))])
        for j, (num, it_title, color) in enumerate(items):
            iy = y + 0.50 + j * 0.32
            R(s, x + 0.10, iy, 0.32, 0.25, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            T(s, x + 0.10, iy, 0.32, 0.25, [
                (str(num), dict(size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER))])
            T(s, x + 0.50, iy, lw - 0.60, 0.25, [
                (it_title, dict(size=10, color=DARK))])
        yc[col] += h + 0.15
    T(s, CL, 6.75, CW, 0.24, [
        ("Farben: rot = out of scope, navy = Kapitel-Detail, orange = Gesamtdarstellung / Empfehlung",
         dict(size=9, italic=True, color=GREY_TXT))])
s_toc()

# ================================
# 3 · RAHMEN: 4 AUFGABEN
# ================================
def s_rahmen():
    s = add_stihl("Rahmen — 4 Aufgaben aus dem Vorstandsprotokoll SEPM (01.07.2026)")
    hdr_y = CT
    R(s, CL, hdr_y, CW, 0.32, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    hcols = [("Aufgabe", 1.5), ("Fragestellung", 6.0), ("Ansatz", 3.7), ("Status / Ergebnis", 1.5)]
    xc = CL
    for name, w in hcols:
        T(s, xc + 0.10, hdr_y, w, 0.32, [
            (name, dict(size=10.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
        xc += w
    T(s, CL + 12.7, hdr_y, 0.63, 0.32, [
        ("Folie", dict(size=10.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    aufgaben = [
        ("A1", "Ablose Software 'Hype' (Massnahmenmgmt) durch SNOW?",
         "Hype stark angepasst → OUT-OF-SCOPE-Begruendung",
         "OUT OF SCOPE", RED_HL, "4"),
        ("A2", "Kann SNOW die STIHL-Schluesselprozesse abbilden (oder App Engine) — ohne Standardisierung um jeden Preis?",
         "Schluesselprozesse listen · Reife SNOW bewerten · Alternativen mit Nachteilen",
         "IN PROGRESS", GREEN, "5-8"),
        ("A3", "Governance-Strukturen bewerten — Komplexitaet in der Organisation vermeiden",
         "Governance-Modelle A/B/C · Prozess-Governance · Benefits · Aufwaende · Betrieb",
         "IN PROGRESS", GREEN, "9-12"),
        ("A4", "Abhaengigkeitsrisiko SNOW (Vergleich SAP) im Folgeprojekt bewerten",
         "Risikomatrix neu · Lockin/Vibe Coding · SAP-Vergleich · Vertragshistorie · Referenzen",
         "IN PROGRESS", GREEN, "13-16"),
        ("G",  "Gesamtdarstellung fuer Vorstand + Entscheidungsvorlage (VS Q3/2026)",
         "Qualitative Analyse · Strategische Ausrichtung · TCO · Szenarien 1+2 · Empfehlung",
         "IN PROGRESS", ORANGE, "17-21"),
    ]
    for i, (nr, frage, ansatz, status, s_col, folien) in enumerate(aufgaben):
        y = hdr_y + 0.40 + i * 1.10
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        R(s, CL, y, CW, 1.05, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        R(s, CL, y, 0.10, 1.05, s_col)
        xc = CL
        # A#
        T(s, xc + 0.20, y, 1.5, 1.05, [
            (nr, dict(size=18, bold=True, color=s_col, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        xc += 1.5
        # Frage
        T(s, xc + 0.10, y + 0.08, 5.9, 0.90, [
            (frage, dict(size=10, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        xc += 6.0
        # Ansatz
        T(s, xc + 0.10, y + 0.08, 3.6, 0.90, [
            (ansatz, dict(size=9.5, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)
        xc += 3.7
        # Status Badge
        R(s, xc + 0.10, y + 0.35, 1.30, 0.36, s_col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, xc + 0.10, y + 0.35, 1.30, 0.36, [
            (status, dict(size=8.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        xc += 1.5
        # Folie
        T(s, xc, y, 0.7, 1.05, [
            (folien, dict(size=13, bold=True, color=DARK, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
s_rahmen()

# ================================
# 4 · AUFGABE 1 — HYPE OUT OF SCOPE
# ================================
def s_a1():
    s = add_stihl("AUFGABE 1 — Massnahmen-Management 'Hype': OUT OF SCOPE")
    # Big status banner
    R(s, CL, CT, CW, 0.75, RED_LT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    R(s, CL, CT, 0.10, 0.75, RED_HL)
    T(s, CL + 0.20, CT, 6.5, 0.75, [
        ("STATUS: OUT OF SCOPE",
         dict(size=13, bold=True, color=RED_HL))], anchor=MSO_ANCHOR.MIDDLE)
    T(s, CL + 6.8, CT, 5.9, 0.75, [
        ("Vorstands-Vorschlag: Hype-Ablose nicht Teil dieser Untersuchung",
         dict(size=10.5, italic=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    # 4 reason cards
    top = CT + 0.95
    h = 2.30
    gap = 0.15
    w = (CW - gap) / 2
    reasons = [
        ("1", "Stark customized ueber Jahre",
         "Hype wurde ueber mehrere Release-Zyklen fuer STIHL-spezifische Massnahmen-Logik erweitert. "
         "Eine Ablose muesste alle Customizings inhaltlich nachziehen — enormer Migrationsaufwand.", RED_HL),
        ("2", "Zusatzfunktionen mit hoher Fachtiefe",
         "Hype-Funktionen fuer Zielkosten-Verfolgung / Massnahmen-Backlog sind stark in der "
         "Produktentwicklung verwurzelt. Nicht 1:1 auf SNOW SPM abbildbar ohne massive Anpassung.", RED_HL),
        ("3", "Keine PPM-Kernfunktion",
         "Hype gehoert nicht in den Kern des Portfolio-Managements. Der Business Case konzentriert sich auf "
         "Strategy, Portfolio, Demand, Projekt- und Ressourcen-Management — nicht auf Massnahmen-Details.", ORANGE2),
        ("4", "Vermeidung von Scope-Creep",
         "Erweiterung um Hype wuerde die Aufwaende und Risiken erheblich vergroessern und den Zeitplan fuer "
         "die Jahresplanung 2027 gefaehrden.", ORANGE2),
    ]
    for i, (num, head, body, color) in enumerate(reasons):
        row = i // 2; col = i % 2
        x = CL + col * (w + gap)
        y = top + row * (h + gap)
        R(s, x, y, w, h, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=color, lw=1.2)
        R(s, x, y, w, 0.42, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        O(s, x + 0.10, y + 0.06, 0.30, 0.30, WHITE)
        T(s, x + 0.10, y + 0.06, 0.30, 0.30, [
            (num, dict(size=13, bold=True, color=color, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, x + 0.45, y + 0.06, w - 0.55, 0.30, [
            (head, dict(size=12, bold=True, color=WHITE))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, x + 0.15, y + 0.55, w - 0.30, h - 0.65, [
            (body, dict(size=10, color=DARK))], anchor=MSO_ANCHOR.TOP)
    # Alternative hint
    alt_y = top + 2 * h + 2 * gap + 0.05
    R(s, CL, alt_y, CW, 0.55, YELLOW_LT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    R(s, CL, alt_y, 0.08, 0.55, YELLOW)
    T(s, CL + 0.15, alt_y, CW - 0.25, 0.55, [
        ("Zukunftsoption: lose Integration Hype ↔ SNOW (z.B. Massnahmen-Referenz auf Portfolio-Ebene) "
         "im spaeteren Projekt-Follow-up — nicht Teil dieses Business Case.",
         dict(size=10, italic=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
s_a1()

# ================================
# 5-8 · AUFGABE 2
# ================================
# ============================================================
# 20 Schluesselprozesse aus Excel-Bewertung (Harmonisierungs-Analyse)
# ============================================================
PROC = [
    (1,  "Strategy",  "Strategische Planung",  "5/6", "Corp&Strat · MPM · VEW · VPM · IT",
     "Mittel", "Stark fragmentiert (workpath, Excel/SPP-XLS, PIT, PowerPoint)",
     "Mittel-Hoch", "Gleiches Ziel · MPM produktspezifisch · SPM als gemeinsamer Rahmen"),
    (2,  "Strategy",  "OKR / Goals & Targets", "2/6", "Corp&Strat · VPM",
     "Hoch", "Identisch — beide nutzen workpath",
     "na", "Nur 2 Bereiche · bereits harmonisiert · Tool ersetzbar"),
    (3,  "Strategy",  "Portfolio Steuerung",   "5/6", "Corp&Strat · MPM · VEW · VPM · IT",
     "Mittel", "Stark fragmentiert (workpath, Excel/PowerBI, PIT, PowerPoint)",
     "na", "Unterschiedliche Auspraegungen · SPM schafft uebergreifende Steuerung"),
    (4,  "Strategy",  "Strategy Reporting",    "5/6", "Corp&Strat · MPM · VEW · VPM · IT",
     "Hoch", "PowerBI verbreitet, Quellen fragmentiert (PIT, Excel, DAP)",
     "Sehr Hoch", "Alle mit Management-Berichten · einheitliche Datenbasis · hoher Nutzen"),
    (5,  "Alignment", "Ideenmanagement",       "5/6", "Corp&Strat · MPM · VEW · VPM · IT",
     "Mittel", "Fragmentiert (ServiceNow, PIT, Excel/Planner, MEMPE/SAP-CRD)",
     "Hoch", "Gemeinsamer Intake-Kanal · Ideen-Sichtbarkeit ueber Silos"),
    (6,  "Alignment", "Demand Management",     "5/6", "Corp&Strat · MPM · VEW · VPM · IT",
     "Mittel-Hoch", "3 Parallelsysteme: ServiceNow (IT), PIT/Excel (VEW), Excel/MEMPE (MPM)",
     "Sehr Hoch", "Ablose 3 Systeme durch einheitlichen SPM-Demand-Intake"),
    (7,  "Alignment", "Portfolio Planung / Roadmaps", "5/6", "Corp&Strat · MPM · VEW · VPM · IT",
     "Hoch", "Fragmentiert (workpath, Excel/PowerBI, PIT, MSPO abgekuendigt)",
     "na", "SPM deckt Roadmap universell · MPM-spezifika beachten"),
    (8,  "Alignment", "Portfolio Kapazitaetsplanung", "4/6", "Corp&Strat · MPM · VEW · IT",
     "Mittel", "Sehr fragmentiert (Alevo, Lucom, PIT, SAP, ServiceNow, PowerApp)",
     "Mittel-Hoch", "Reife-Unterschiede · Potenzial bereichsuebergreifende Kapazitaet"),
    (9,  "Alignment", "Portfolio Finanzplanung", "5/6", "MPM · VEW · VPM · IT · Controlling",
     "Hoch", "Fragmentiert (PIT, SAP/SAC, Alevo, Lucom, ServiceNow sfinx)",
     "na", "SPM Portfolioplanung kann Plattform bieten"),
    (10, "Alignment", "Approval Prozesse",     "6/6", "Alle 6 Bereiche",
     "Hoch", "4 parallele Systeme (PowerApp, Lucom, ServiceNow, E-Mail)",
     "Sehr Hoch", "Universelles Prozessmuster · einheitliche SNOW-Approval-Workflows"),
    (11, "Alignment", "Portfolio Reporting",   "6/6", "Alle 6 Bereiche",
     "Hoch", "PowerBI verbreitet, Daten aus 4+ Quellen",
     "Hoch", "Alle brauchen Portfolio-Gesamtbild · einheitl. Datenbasis"),
    (12, "Alignment", "Portfolio Szenarien",   "3/6", "Corp&Strat · VEW · VPM",
     "Gering", "Kaum vorhanden (VPM: PIT kaum nutzbar, VEW/C&S: kein Tool)",
     "na", "Kein etablierter Prozess · SPM eher Roadmap-Thema"),
    (13, "Lieferung", "Projektplanung",        "4/6", "Corp&Strat · VEW · VPM · IT",
     "Hoch", "VEW+VPM fast identisch (PIT-Templates), IT muss migrieren (MSPO)",
     "Hoch", "PIT-Standardisierung als Basis · Deep Dive 22.06."),
    (14, "Lieferung", "Finanzplanung und Steuerung", "5/6", "Corp&Strat · VEW · VPM · IT · Controlling",
     "Hoch", "VEW+VPM fast identisch (PIT+SAP+SAC), IT mit Lucom-Zusatz",
     "Hoch", "Nahezu gleicher Prozess · hoechste Konvergenz Lieferung"),
    (15, "Lieferung", "Ressourcenplanung",     "3/6", "Corp&Strat · VEW · VPM · IT",
     "Hoch", "VEW+VPM identisch (PIT), IT mit zusaetzl. Systemen",
     "Hoch", "VEW/VPM identisch — einfachste Harmonisierung"),
    (16, "Lieferung", "Programm Management",   "3/6", "VEW · VPM · IT",
     "Mittel", "PIT bei allen unzureichend · IT: MSPO abgekuendigt",
     "Mittel-Hoch", "Gleiche Anforderung · heute schwach etabliert · SPM schliesst Luecke"),
    (17, "Lieferung", "Zeitaufschreibung",     "3/6", "VEW · VPM · IT",
     "Sehr Hoch", "PIT bei allen 3 + SAP-Integration",
     "na", "Fast identisch · MSPO-Abloese erzwingt Migration"),
    (18, "Lieferung", "Aufgabenmanagement",    "3/6", "VEW · VPM · IT",
     "Mittel", "Stark fragmentiert (PIT, Jira, Planner, Excel)",
     "Mittel", "Grundbedarf gleich · Umsetzung sehr verschieden · CWM als Basis"),
    (19, "Lieferung", "Risikomanagement",      "4/6", "Corp&Strat · VEW · VPM · IT",
     "Hoch", "Excel ueberall vorhanden · VEW mit Governance",
     "Hoch", "Aehnliche Struktur · SPM-Risikomodul + Portfolio-Aggregation"),
    (20, "Lieferung", "Projekt Reporting",     "4/6", "VEW · VPM · IT · Controlling",
     "Hoch", "VEW+VPM fast identisch (PIT+PowerBI via DAP), IT fragmentiert",
     "Hoch", "PEP-Struktur · IT eigenes PPM Reporting"),
]



# T-Shirt Size Aufwand-Mapping (Anpassungen Planhorizon)
AUFWAND = {
    # Prozess-Nr -> ("Kategorie", detail_text)
    1:  ("Standard", "Standard"),
    2:  ("Standard", "Standard"),
    3:  ("Standard", "Standard"),
    4:  ("Standard", "Standard"),
    5:  ("M-L",      "MPM Produkt Demand (Identification, Roadmap Review)"),
    6:  ("L",        "Approval Prozesse: Abloese EV Tool + Controlling"),
    7:  ("M-L",      "MPM Roadmap Review + Realisierungsbewertung"),
    8:  ("L",        "MPM Produkt Mengenplanung + Datenmodell + Dashboarding"),
    9:  ("XL",       "SAP Schnittstelle · Restwerte · Sachkosten · Personal"),
    10: ("L",        "Approval Prozesse (EV-Ablose, Controlling)"),
    11: ("Standard", "Standard"),
    12: ("Standard", "Standard"),
    13: ("M",        "VEW/VPM Master Templates (S) + PEP Struktur (M) + PEP Rollen (M)"),
    14: ("XL",       "SAP Schnittstelle (XL) + Restwerte/Sachkosten/Personal (S)"),
    15: ("S",        "Planung Personal nach Phasen"),
    16: ("Standard", "Standard"),
    17: ("XL",       "SAP Schnittstelle (XL) + Anonymisierung Zeiterfassung (S/M)"),
    18: ("Standard", "Standard"),
    19: ("Standard", "Standard"),
    20: ("M-XL",     "Status/Ampel Logik (M) + Massenaenderungen (M-XL)"),
}
# Farbmapping T-Shirt-Sizes (kompakt)
AUFWAND_COLOR = {
    "Standard": GREEN,
    "S":        RGBColor(0x60, 0xA5, 0xFA),   # hellblau
    "M":        RGBColor(0x3B, 0x82, 0xF6),   # mittelblau
    "M-L":      RGBColor(0x25, 0x63, 0xEB),   # dunkler blau
    "L":        RGBColor(0x1E, 0x40, 0xAF),   # navy-blau
    "XL":       RGBColor(0x6D, 0x28, 0xD9),   # violett
    "M-XL":     RGBColor(0x5B, 0x21, 0xB6),   # dunkleres violett
}
AUFWAND_LEGEND = [
    ("Standard", "(kein Aufwand)"),
    ("S",        "1-3 PT"),
    ("M",        "4-10 PT"),
    ("M-L",      "Uebergang"),
    ("L",        "12-20 PT"),
    ("XL",       "25 PT+"),
    ("M-XL",     "variabel"),
]

def s_a2_uebersicht():
    s = add_stihl("AUFGABE 2 — 20 Schluesselprozesse @ STIHL: Uebersicht + Harmonisierungspotenzial + Umsetzung")
    # Header mit zwei Legenden
    R(s, CL, CT, CW, 0.50, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    T(s, CL + 0.15, CT, 2.5, 0.50, [("Bewertungsergebnis", dict(size=10.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    # Legende Harmonisierung
    lx = CL + 2.70
    T(s, lx, CT, 1.2, 0.50, [("Harmonisierung:", dict(size=8, italic=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    lx += 1.10
    for lbl, col in [("Sehr Hoch", GREEN), ("Hoch", GREEN_MID), ("Mittel-Hoch", YELLOW_GRN), ("Mittel", ORANGE), ("na", GREY_NA)]:
        R(s, lx, CT + 0.16, 0.18, 0.18, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, lx + 0.22, CT, 0.85, 0.50, [(lbl, dict(size=7.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
        lx += 1.02
    # Legende Aufwand
    T(s, lx, CT, 0.95, 0.50, [("Umsetzung:", dict(size=8, italic=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    lx += 0.95
    for lbl, note in [("Standard", ""), ("S", "1-3"), ("M", "4-10"), ("L", "12-20"), ("XL", "25+")]:
        col = AUFWAND_COLOR[lbl]
        R(s, lx, CT + 0.16, 0.18, 0.18, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, lx + 0.22, CT, 0.60, 0.50, [(lbl, dict(size=7.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
        lx += 0.80

    # 3 Gruppen-Spalten
    gruppen = ["Strategy", "Alignment", "Lieferung"]
    gcol = {"Strategy": ORANGE, "Alignment": BLUE_HD, "Lieferung": PURPLE}
    top = CT + 0.60
    h = CB - top
    w = (CW - 0.30) / 3
    gap = 0.15
    for gi, gname in enumerate(gruppen):
        x = CL + gi * (w + gap)
        color = gcol[gname]
        R(s, x, top, w, h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        R(s, x, top, w, 0.36, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        items = [p for p in PROC if p[1] == gname]
        T(s, x + 0.15, top + 0.02, w - 0.30, 0.32, [(f"{gname.upper()}  ({len(items)} Prozesse)", dict(size=11, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
        row_h = min((h - 0.50) / max(len(items), 1), 0.75)
        for j, p in enumerate(items):
            nr, _, prz, anz, _, _, _, pot, _ = p
            auf_kat, auf_detail = AUFWAND.get(nr, ("Standard", "Standard"))
            y = top + 0.42 + j * row_h
            R(s, x + 0.10, y, w - 0.20, row_h - 0.05, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=color, lw=0.5)
            R(s, x + 0.10, y, 0.06, row_h - 0.05, color)
            # Nummer-Kreis
            O(s, x + 0.20, y + 0.06, 0.24, 0.24, color)
            T(s, x + 0.20, y + 0.06, 0.24, 0.24, [(str(nr), dict(size=8.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)
            # Prozess-Name (obere Zeile)
            T(s, x + 0.50, y + 0.03, w - 0.60, 0.32, [(prz, dict(size=9.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
            # Untere Zeile: 3 Badges (Bereiche · Potenzial · Umsetzung)
            b_y = y + row_h - 0.28
            # Anz. Bereiche
            R(s, x + 0.20, b_y, 0.45, 0.22, GREY_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            T(s, x + 0.20, b_y, 0.45, 0.22, [(anz, dict(size=7.5, bold=True, color=DARK, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)
            # Potenzial
            pcol = POT_COLOR[pot]
            R(s, x + 0.70, b_y, 1.20, 0.22, pcol, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            T(s, x + 0.70, b_y, 1.20, 0.22, [(pot, dict(size=7, bold=True, color=WHITE, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)
            # Umsetzung-Badge
            acol = AUFWAND_COLOR[auf_kat]
            aw_w = 0.85 if len(auf_kat) <= 3 else 1.05
            aw_x = x + w - aw_w - 0.15
            R(s, aw_x, b_y, aw_w, 0.22, acol, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            T(s, aw_x, b_y, aw_w, 0.22, [(auf_kat, dict(size=7.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)

    # Zusammenfassung
    counts_pot = {}
    counts_auf = {}
    for p in PROC:
        counts_pot[p[7]] = counts_pot.get(p[7], 0) + 1
        auf = AUFWAND.get(p[0], ("Standard","Standard"))[0]
        counts_auf[auf] = counts_auf.get(auf, 0) + 1
    T(s, CL, CB - 0.05, CW, 0.20, [
        (f"Harmon.:  {counts_pot.get('Sehr Hoch',0)} Sehr Hoch · {counts_pot.get('Hoch',0)} Hoch · {counts_pot.get('Mittel-Hoch',0)} Mittel-Hoch · {counts_pot.get('na',0)} na          |          Umsetzung:  {counts_auf.get('Standard',0)} Standard · {counts_auf.get('S',0)} S · {counts_auf.get('M',0)} M · {counts_auf.get('M-L',0)} M-L · {counts_auf.get('L',0)} L · {counts_auf.get('XL',0)} XL · {counts_auf.get('M-XL',0)} M-XL",
         dict(size=8, italic=True, color=GREY_TXT, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)


def s_a2_bewertung():
    s = add_stihl("AUFGABE 2 — Bewertungsergebnis: Prozess-Aehnlichkeit x Tool-Situation -> Harmonisierungspotenzial")
    hdr_y = CT
    R(s, CL, hdr_y, CW, 0.32, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    cols = [("#", 0.30), ("Prozess", 3.0), ("Ber.", 0.50), ("Aehnlichkeit", 1.15), ("Tool-Situation heute", 5.35), ("Potenzial", 1.15), ("Key Insight", 2.22)]
    xc = CL
    for name, w in cols:
        T(s, xc + 0.08, hdr_y, w - 0.08, 0.32, [(name, dict(size=9.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
        xc += w
    focus = [p for p in PROC if p[7] != "na"]
    order = {"Sehr Hoch": 0, "Hoch": 1, "Mittel-Hoch": 2, "Mittel": 3, "Gering": 4}
    focus.sort(key=lambda p: order.get(p[7], 99))
    for i, p in enumerate(focus):
        nr, gr, prz, anz, ber, aehn, tool, pot, insight = p
        y = hdr_y + 0.36 + i * 0.42
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        R(s, CL, y, CW, 0.40, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, CL + 0.05, y, 0.28, 0.40, [(str(nr), dict(size=9, bold=True, color=DARK, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)
        T(s, CL + 0.34, y, 2.94, 0.40, [(prz, dict(size=9.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        T(s, CL + 3.30, y, 0.44, 0.40, [(anz, dict(size=9, color=DARK, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)
        aehn_col = {"Sehr Hoch": GREEN, "Hoch": GREEN_MID, "Mittel-Hoch": YELLOW_GRN, "Mittel": ORANGE, "Gering": RED_HL}.get(aehn, GREY_NA)
        R(s, CL + 3.85, y + 0.08, 1.02, 0.24, aehn_col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, CL + 3.85, y + 0.08, 1.02, 0.24, [(aehn, dict(size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)
        T(s, CL + 4.95, y, 5.30, 0.40, [(tool, dict(size=8.5, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        pcol = POT_COLOR[pot]
        R(s, CL + 10.30, y + 0.08, 1.05, 0.24, pcol, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, CL + 10.30, y + 0.08, 1.05, 0.24, [(pot, dict(size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)
        T(s, CL + 11.45, y, 2.22, 0.40, [(insight, dict(size=8, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.MIDDLE)
    fy = hdr_y + 0.36 + len(focus) * 0.42 + 0.10
    if fy > CB - 0.45: fy = CB - 0.45
    R(s, CL, fy, CW, 0.40, GREEN_LT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    R(s, CL, fy, 0.08, 0.40, GREEN)
    T(s, CL + 0.15, fy, CW - 0.25, 0.40, [("Fazit:  17 von 20 Prozessen mit Harmonisierungspotenzial · 10 davon 'Sehr Hoch' oder 'Hoch' — SEPM hebt hier den groessten Nutzen", dict(size=10, bold=True, italic=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)


def s_a2_top_potenziale():
    s = add_stihl("AUFGABE 2 — Top-Konsolidierungs-Chancen  (10 Prozesse: 'Sehr Hoch' + 'Hoch')")
    sehr_hoch = [p for p in PROC if p[7] == "Sehr Hoch"]
    hoch = [p for p in PROC if p[7] == "Hoch"]
    top1 = CT
    top1_h = 2.15
    top1_w = (CW - 0.30) / 3
    for i, p in enumerate(sehr_hoch):
        nr, gr, prz, anz, ber, aehn, tool, pot, insight = p
        x = CL + i * (top1_w + 0.15)
        R(s, x, top1, top1_w, top1_h, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=GREEN, lw=2.0)
        R(s, x, top1, top1_w, 0.48, GREEN, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        R(s, x + top1_w - 1.25, top1 + 0.08, 1.15, 0.32, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, x + top1_w - 1.25, top1 + 0.08, 1.15, 0.32, [("★ SEHR HOCH", dict(size=8.5, bold=True, color=GREEN, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)
        O(s, x + 0.15, top1 + 0.08, 0.32, 0.32, WHITE)
        T(s, x + 0.15, top1 + 0.08, 0.32, 0.32, [(str(nr), dict(size=11, bold=True, color=GREEN, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)
        T(s, x + 0.55, top1 + 0.04, top1_w - 1.85, 0.40, [(prz, dict(size=12.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
        T(s, x + 0.15, top1 + 0.55, top1_w - 0.30, 0.25, [(f"{anz} Bereiche  ·  Aehnl.: {aehn}", dict(size=9, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.MIDDLE)
        T(s, x + 0.15, top1 + 0.82, top1_w - 0.30, 0.42, [("Heute:  " + tool, dict(size=8.5, color=DARK))], anchor=MSO_ANCHOR.TOP)
        R(s, x + 0.15, top1 + 1.32, top1_w - 0.30, 0.75, GREEN_LT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        R(s, x + 0.15, top1 + 1.32, 0.05, 0.75, GREEN)
        T(s, x + 0.30, top1 + 1.35, top1_w - 0.50, 0.68, [(insight, dict(size=8.5, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    top2 = top1 + top1_h + 0.22
    T(s, CL, top2 - 0.28, CW, 0.24, [("Weitere 7 Top-Kandidaten (Bewertung 'Hoch'):", dict(size=11, bold=True, color=GREEN_MID))])
    top2_h = (CB - top2 - 0.20) / 2 - 0.05
    top2_w = (CW - 3 * 0.12) / 4
    for i, p in enumerate(hoch):
        nr, gr, prz, anz, ber, aehn, tool, pot, insight = p
        row = i // 4; col = i % 4
        x = CL + col * (top2_w + 0.12)
        y = top2 + row * (top2_h + 0.10)
        R(s, x, y, top2_w, top2_h, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=GREEN_MID, lw=1.2)
        R(s, x, y, top2_w, 0.35, GREEN_MID, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        O(s, x + 0.10, y + 0.05, 0.25, 0.25, WHITE)
        T(s, x + 0.10, y + 0.05, 0.25, 0.25, [(str(nr), dict(size=9, bold=True, color=GREEN_MID, align=PP_ALIGN.CENTER))], anchor=MSO_ANCHOR.MIDDLE)
        T(s, x + 0.42, y + 0.03, top2_w - 0.55, 0.30, [(prz, dict(size=10.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
        T(s, x + 0.10, y + 0.42, top2_w - 0.20, 0.22, [(f"{anz} · Aehnl. {aehn}", dict(size=8, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.MIDDLE)
        T(s, x + 0.10, y + 0.66, top2_w - 0.20, top2_h - 0.72, [(insight, dict(size=8, color=DARK))], anchor=MSO_ANCHOR.TOP)


def s_a2_alternativen():
    s = add_stihl("AUFGABE 2 — Alternativen mit aehnlicher Reife: Nachteile im Vergleich")
    hdr_y = CT
    R(s, CL, hdr_y, CW, 0.32, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    hcols = [("Schluesselprozess", 3.5), ("Alternative (Best-of-Breed)", 3.0),
             ("Reife Alternative", 1.5), ("Wesentliche Nachteile", 4.67)]
    xc = CL
    for name, w in hcols:
        T(s, xc + 0.15, hdr_y, w - 0.15, 0.32, [
            (name, dict(size=10.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
        xc += w
    alt = [
        ("Strategisches Portfolio-Mgmt",  "Planisware, SAP EPPM",   4,
         "Separates Tool, keine ITSM-Integration, hoehere TCO"),
        ("Demand-Management",             "Workpath / SharePoint",  3,
         "Custom-Bau, kein Standardprodukt, Wartung teuer"),
        ("Projekt- & Ressourcenplanung",  "Planisware Enterprise",  5,
         "Sehr maechtig aber personell-hungrig, Insellandschaft"),
        ("Reporting & Dashboards",        "Power BI + SharePoint",  4,
         "Nur Dashboarding, keine PPM-Prozess-Integration"),
        ("Entwicklungsprojekte (PEP)",    "Planisware / Siemens Teamcenter", 4,
         "Enge Kopplung an PLM, hohe Umstellungskosten"),
        ("EV-Prozess",                    "Lucom + SharePoint",     3,
         "Bestand, aber Wartungsvertrag limitiert, Insel"),
        ("Application Portfolio Mgmt",    "LeanIX / eigenes Excel", 4,
         "Zusatzliches Tool, kein Anschluss an SPM/Demand"),
    ]
    for i, (prozess, alter, reife, nachteile) in enumerate(alt):
        y = hdr_y + 0.36 + i * 0.68
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        R(s, CL, y, CW, 0.64, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, CL + 0.15, y, 3.35, 0.64, [
            (prozess, dict(size=10.5, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, CL + 3.60, y, 2.85, 0.64, [
            (alter, dict(size=10, color=NAVY))],
            anchor=MSO_ANCHOR.MIDDLE)
        # Reife stars
        star_start = CL + 6.60
        for k in range(5):
            filled = k < reife
            color = GREEN if reife >= 4 else (ORANGE if reife >= 3 else YELLOW)
            O(s, star_start + k * 0.26, y + 0.19, 0.22, 0.22,
              color if filled else LIGHT_GREY)
        T(s, CL + 8.10, y, 2.10, 0.64, [
            ("(Best-of-Breed)", dict(size=8.5, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, CL + 8.35, y, 4.60, 0.64, [
            (nachteile, dict(size=9.5, color=RED_HL))],
            anchor=MSO_ANCHOR.MIDDLE)
    # Kernaussage
    ky = hdr_y + 0.36 + len(alt) * 0.68 + 0.15
    R(s, CL, ky, CW, 0.55, ORANGE_LT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    R(s, CL, ky, 0.08, 0.55, ORANGE)
    T(s, CL + 0.15, ky, CW - 0.25, 0.55, [
        ("Kernaussage:  Alternativen haben aehnliche Reife, aber alle bringen Insel-Betrieb, hoehere TCO "
         "und fehlende Prozess-Integration mit sich — SNOW ist strategisch ueberlegen.",
         dict(size=10, italic=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
s_a2_uebersicht()
s_a2_bewertung()
s_a2_top_potenziale()
s_a2_alternativen()

# ================================
# 8 · AUFGABE 3 — GOVERNANCE-MODELLE A/B/C
# ================================
def s_a3_governance():
    s = add_stihl("AUFGABE 3 — Governance-Modelle:  A · Vollharmonisiert  ·  B · Foederiert  ·  C · Status quo")
    R(s, CL, CT, CW, 0.55, GREEN_LT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    R(s, CL, CT, 0.08, 0.55, GREEN)
    T(s, CL + 0.15, CT, CW - 0.25, 0.55, [
        ("Bewertung dreier Governance-Modelle — B (Foederiert) wird empfohlen",
         dict(size=10.5, italic=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    hdr_y = CT + 0.75
    R(s, CL, hdr_y, CW, 0.42, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    cols = [("Szenario", 3.0), ("Governance-Modell", 4.0),
            ("FTE (laufend)", 1.9), ("Komplexitaet", 1.9), ("Change (einmalig)", 1.87)]
    xc = CL
    for name, w in cols:
        T(s, xc + 0.15, hdr_y, w - 0.15, 0.42, [
            (name, dict(size=11, bold=True, color=WHITE))],
            anchor=MSO_ANCHOR.MIDDLE)
        xc += w
    rows = [
        ("A · Vollharmonisiert",          "1 Standard, comply-or-explain",
         "niedrig", "niedrig", "hoch",     False),
        ("B · Foederiert  (Empfehlung)",  "Standard + wenige def. Varianten",
         "mittel",  "mittel",  "mittel",   True),
        ("C · Status quo digitalisiert",  "je Fachbereich eigenes Modell",
         "hoch",    "hoch",    "niedrig",  False),
    ]
    for i, (szen, modell, fte, komp, chg, emp) in enumerate(rows):
        y = hdr_y + 0.46 + i * 1.20
        bg = GREEN_LT if emp else (LIGHT_BG if i % 2 == 0 else ROW_ALT)
        R(s, CL, y, CW, 1.15, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        # Left accent
        R(s, CL, y, 0.12, 1.15, GREEN if emp else NAVY)
        xc = CL
        # Szenario
        T(s, xc + 0.25, y, 2.85, 1.15, [
            (szen, dict(size=12.5, bold=True, color=DARK if not emp else GREEN))],
            anchor=MSO_ANCHOR.MIDDLE)
        xc += 3.0
        # Modell
        T(s, xc + 0.15, y, 3.85, 1.15, [
            (modell, dict(size=11, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        xc += 4.0
        # FTE
        color_fte = BEW_COLOR[fte]
        R(s, xc + 0.30, y + 0.30, 1.30, 0.55, color_fte, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, xc + 0.30, y + 0.30, 1.30, 0.55, [
            (fte, dict(size=11.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        xc += 1.9
        # Komplexität
        color_k = BEW_COLOR[komp]
        R(s, xc + 0.30, y + 0.30, 1.30, 0.55, color_k, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, xc + 0.30, y + 0.30, 1.30, 0.55, [
            (komp, dict(size=11.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        xc += 1.9
        # Change
        color_c = BEW_COLOR[chg]
        R(s, xc + 0.30, y + 0.30, 1.30, 0.55, color_c, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, xc + 0.30, y + 0.30, 1.30, 0.55, [
            (chg, dict(size=11.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        # Empfehlungs-Star
        if emp:
            R(s, CL + CW - 0.85, y + 0.10, 0.75, 0.30, GREEN, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            T(s, CL + CW - 0.85, y + 0.10, 0.75, 0.30, [
                ("★ EMPF", dict(size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
                anchor=MSO_ANCHOR.MIDDLE)
    # Footnote
    fn_y = hdr_y + 0.46 + 3 * 1.20 + 0.10
    T(s, CL, fn_y, CW, 0.28, [
        ("FTE-Zahlen werden im Folgeprojekt bottom-up je Governance-Rolle × Fachbereich × Standardisierungsgrad ermittelt.",
         dict(size=9, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.MIDDLE)
    # Legende
    leg_y = fn_y + 0.35
    T(s, CL, leg_y, 1.5, 0.24, [
        ("Bewertung:", dict(size=9, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    lx = CL + 1.5
    for lbl, col in [("niedrig", GREEN), ("mittel", ORANGE), ("hoch", RED_HL)]:
        R(s, lx, leg_y + 0.02, 0.24, 0.18, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, lx + 0.32, leg_y, 1.5, 0.24, [
            (lbl, dict(size=9, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        lx += 1.7
s_a3_governance()

# ================================
# 9 · AUFGABE 3 — WAS AENDERT SICH PRO SCHLUESSELPROZESS
# ================================
def s_a3_aenderung():
    s = add_stihl("AUFGABE 3 — Was aendert sich pro Schluesselprozess  (Prozess-Governance)")
    hdr_y = CT
    R(s, CL, hdr_y, CW, 0.32, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    hcols = [("Schluesselprozess", 3.5), ("Heute (Rollen · Owner)", 3.8),
             ("Mit SEPM (Rollen · Owner)", 3.8), ("Delta / Governance-Anspruch", 1.57)]
    xc = CL
    for name, w in hcols:
        T(s, xc + 0.15, hdr_y, w - 0.15, 0.32, [
            (name, dict(size=10.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
        xc += w
    T(s, CL + 12.6, hdr_y, 0.73, 0.32, [
        ("Aufwand", dict(size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    rows = [
        ("Strat. Portfolio-Mgmt",  "Excel · dezentral, VS-Sekretariat",
         "SNOW SPM · Portfolio-Board (SCD/SEH)",   "gering",  GREEN),
        ("Demand-Management",      "SharePoint · manuell je Ressort",
         "SNOW Demand · Fachbereich + IT-Demand",  "mittel",  ORANGE),
        ("Projekt-/Ressourcen",    "PIT · Ressort-Leitung",
         "SNOW PM/RM · zentral + Ressort-Owner",   "mittel",  ORANGE),
        ("Reporting/Dashboards",   "Excel · manuell Zusammenstellen",
         "SNOW Live-Dashboards · Self-Service",    "gering",  GREEN),
        ("Entwicklungs-Projekte",  "PIT + Excel · EWW Portfolio",
         "SPM + App Engine · EWW-Board",           "hoch",    RED_HL),
        ("EV-Prozess",             "Lucom + manuelle Freigabe",
         "SNOW App-Engine-Workflow",               "mittel",  ORANGE),
        ("App Portfolio Mgmt",     "kein zentraler Owner",
         "SNOW APM · IT-Architektur-Board",        "mittel",  ORANGE),
    ]
    for i, (proz, heute, mit_sepm, aufw, color) in enumerate(rows):
        y = hdr_y + 0.36 + i * 0.72
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        R(s, CL, y, CW, 0.68, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, CL + 0.15, y, 3.35, 0.68, [
            (proz, dict(size=10.5, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, CL + 3.60, y, 3.65, 0.68, [
            (heute, dict(size=9.5, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, CL + 7.40, y, 3.65, 0.68, [
            (mit_sepm, dict(size=9.5, color=NAVY, bold=True))],
            anchor=MSO_ANCHOR.MIDDLE)
        R(s, CL + 11.20, y + 0.15, 1.20, 0.38, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, CL + 11.20, y + 0.15, 1.20, 0.38, [
            (aufw, dict(size=9.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, CL + 12.45, y, 0.85, 0.68, [
            ("★" if aufw == "hoch" else ("●" if aufw == "mittel" else "○"),
             dict(size=14, bold=True, color=color, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    # Fazit
    fy = hdr_y + 0.36 + len(rows) * 0.72 + 0.10
    R(s, CL, fy, CW, 0.55, GREEN_LT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    R(s, CL, fy, 0.08, 0.55, GREEN)
    T(s, CL + 0.15, fy, CW - 0.25, 0.55, [
        ("Prinzip:  Klare Owner-Zuordnung je Prozess-Rolle in einer Plattform — Komplexitaet in der Organisation sinkt.",
         dict(size=10, italic=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
s_a3_aenderung()

# ================================
# 10 · AUFGABE 3 — BENEFITS (Zeit / Qualitaet / Kosten)
# ================================
def s_a3_benefits():
    s = add_stihl("AUFGABE 3 — Benefits durch neue Governance:  Zeit · Qualitaet · Kosten")
    top = CT
    h = CB - top
    w = (CW - 0.30) / 3
    gap = 0.15
    kats = [
        ("Zeit",     ORANGE, "⏱",
         [("Idee → Freigabe", "6-8 Wochen", "1-2 Wochen", "− 75 %"),
          ("Reporting-Zyklus", "monatlich manuell", "live", "sofort"),
          ("Governance-Entscheidungen", "3-4 Meetings", "1 Meeting + Dashboard", "− 60 %"),
          ("Onboarding neuer PM", "3-4 Wochen", "1 Woche (Standardrollen)", "− 70 %")]),
        ("Qualitaet", GREEN, "✓",
         [("Datenkonsistenz", "Insel-Excels", "1 Datenmodell", "100 %"),
          ("Audit-Faehigkeit", "manuell", "Standard-Reports", "hoch"),
          ("Fehlerquote Berichte", "5-10 %", "< 1 %", "− 90 %"),
          ("Nachvollziehbarkeit OKR→Projekt", "keine", "durchgaengig", "neu")]),
        ("Kosten",   NAVY,  "€",
         [("Reporting-Aufwand p.a.", "~ 3 FTE", "~ 1 FTE", "− 260 k€"),
          ("Governance-FTE Ressorts", "~ 4-6 FTE", "~ 2-3 FTE", "− 350 k€"),
          ("Tool-Betrieb (Konsol.)", "~ 6 Tools", "1 Plattform", "− 950 k€"),
          ("Change-Aufwand Ressorts (einm.)", "—", "~ 200 k€", "+ 200 k€ einm.")]),
    ]
    for i, (title, color, icon, items) in enumerate(kats):
        x = CL + i * (w + gap)
        R(s, x, top, w, h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        R(s, x, top, w, 0.65, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        O(s, x + 0.15, top + 0.12, 0.42, 0.42, WHITE)
        T(s, x + 0.15, top + 0.12, 0.42, 0.42, [
            (icon, dict(size=14, bold=True, color=color, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, x + 0.65, top + 0.15, w - 0.75, 0.40, [
            (title, dict(size=15, bold=True, color=WHITE))],
            anchor=MSO_ANCHOR.MIDDLE)
        # Header row for sub-columns
        R(s, x + 0.10, top + 0.72, w - 0.20, 0.30, GREY_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, x + 0.15, top + 0.72, 1.90, 0.30, [
            ("Kennzahl", dict(size=8.5, bold=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, x + 2.05, top + 0.72, 1.10, 0.30, [
            ("Heute", dict(size=8.5, bold=True, color=GREY_TXT, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, x + 3.15, top + 0.72, 1.30, 0.30, [
            ("Mit SEPM", dict(size=8.5, bold=True, color=GREY_TXT, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, x + 4.45, top + 0.72, 0.30, 0.30, [
            ("Δ", dict(size=8.5, bold=True, color=GREY_TXT, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        # Items rows
        for j, (kenn, heute, mit, delta) in enumerate(items):
            yy = top + 1.10 + j * 1.15
            R(s, x + 0.10, yy, w - 0.20, 1.05, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
              line=color, lw=0.4)
            R(s, x + 0.10, yy, 0.06, 1.05, color)
            T(s, x + 0.25, yy + 0.04, 1.80, 1.00, [
                (kenn, dict(size=9.5, bold=True, color=DARK))],
                anchor=MSO_ANCHOR.MIDDLE)
            T(s, x + 2.05, yy + 0.05, 1.10, 0.45, [
                (heute, dict(size=9, color=RED_HL, align=PP_ALIGN.CENTER))],
                anchor=MSO_ANCHOR.MIDDLE)
            T(s, x + 3.15, yy + 0.05, 1.30, 0.45, [
                (mit, dict(size=9, color=GREEN, bold=True, align=PP_ALIGN.CENTER))],
                anchor=MSO_ANCHOR.MIDDLE)
            T(s, x + 0.25, yy + 0.55, w - 0.55, 0.48, [
                (delta, dict(size=11, bold=True, color=color, align=PP_ALIGN.CENTER))],
                anchor=MSO_ANCHOR.MIDDLE)
s_a3_benefits()

# ================================
# 11 · AUFGABE 3 — EINMALAUFWAENDE + BETRIEBSTHEMEN
# ================================
def s_a3_aufwaende():
    s = add_stihl("AUFGABE 3 — Einmalaufwaende (Training/Enablement) + Betriebsthemen (Fachbereich/IT)")
    lw = (CW - 0.20) / 2
    rx = CL + lw + 0.20
    top = CT
    h = CB - top
    # LEFT: Einmalaufwände
    R(s, CL, top, lw, h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    R(s, CL, top, lw, 0.45, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    T(s, CL + 0.15, top, lw - 0.25, 0.45, [
        ("EINMALIGE AUFWAENDE — Training & Enablement",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    einmalig = [
        ("Basis-Training Anwender",       "~ 400 Personen · 4 h",      "~ 80 k€",  "e-Learning + Praesenz"),
        ("Power-User-Training",           "~ 40 Personen · 2 Tage",    "~ 60 k€",  "SNOW-Standard + STIHL-Prozesse"),
        ("Governance-Board-Onboarding",   "je Ressort · 1 Tag",         "~ 30 k€",  "SCD/SEH + Fachbereichsleiter"),
        ("Change-Kommunikation",          "Kick-off · Roadshow",        "~ 50 k€",  "internes Marketing"),
        ("Prozess-Doku + Playbooks",      "~ 30 Dokumente",             "~ 40 k€",  "Standard first — Anpassung dokumentiert"),
        ("Datenmigration + Bereinigung",  "PIT → SNOW",                 "~ 120 k€", "einmalig, mit Data-Owner"),
    ]
    for i, (name, umf, kost, note) in enumerate(einmalig):
        y = top + 0.60 + i * 1.02
        R(s, CL + 0.15, y, lw - 0.30, 0.95, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
          line=ORANGE, lw=0.5)
        R(s, CL + 0.15, y, 0.06, 0.95, ORANGE)
        T(s, CL + 0.28, y + 0.05, lw - 2.00, 0.28, [
            (name, dict(size=10.5, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, CL + 0.28, y + 0.30, lw - 2.00, 0.28, [
            (umf, dict(size=9, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, CL + 0.28, y + 0.60, lw - 2.00, 0.30, [
            (note, dict(size=9, color=DARK))],
            anchor=MSO_ANCHOR.TOP)
        R(s, CL + lw - 1.70, y + 0.28, 1.50, 0.40, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, CL + lw - 1.70, y + 0.28, 1.50, 0.40, [
            (kost, dict(size=10.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    total_e = 380
    T(s, CL + 0.15, top + h - 0.35, lw - 0.30, 0.30, [
        (f"Summe einmalig: ~ {total_e} k€",
         dict(size=11, bold=True, color=ORANGE, align=PP_ALIGN.RIGHT))],
        anchor=MSO_ANCHOR.MIDDLE)

    # RIGHT: Betriebsthemen (FB / IT)
    R(s, rx, top, lw, h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    R(s, rx, top, lw, 0.45, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    T(s, rx + 0.15, top, lw - 0.25, 0.45, [
        ("BETRIEBSTHEMEN — Fachbereich & IT",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    betrieb = [
        ("Fachbereich",  "Prozess-Owner je Schluesselprozess", "0,5-1,0 FTE / Ressort",
         "Owner haelt Standard, Fachbereichs-Anpassungen im Rahmen der Foederierung"),
        ("Fachbereich",  "Governance-Boards (SCD/SEH + FB)",   "quartalsweise 0,5 Tag",
         "Portfolio-Priorisierung, OKR-Alignment, Approvals"),
        ("Fachbereich",  "Anwender-Support 1st Level",         "0,2 FTE",
         "Fragen Standard-Bedienung, im Ressort"),
        ("IT",           "SNOW-Plattform-Admin",               "1,0 FTE",
         "im bestehenden ITSM-Team, Erweiterung"),
        ("IT",           "SPM Business Analyst / SI-Partner",  "1,0 FTE + Externer",
         "Erweiterungen, App Engine, Prozess-Anpassungen"),
        ("IT",           "Integration / API-Betreuung",        "0,5 FTE",
         "SAP FI/HCM, PIT-Ausrollung, Monitoring"),
    ]
    fb_col = {"Fachbereich": GREEN, "IT": BLUE_HD}
    for i, (typ, name, fte, note) in enumerate(betrieb):
        y = top + 0.60 + i * 1.02
        color = fb_col[typ]
        R(s, rx + 0.15, y, lw - 0.30, 0.95, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
          line=color, lw=0.5)
        R(s, rx + 0.15, y, 0.06, 0.95, color)
        # Typ badge
        R(s, rx + 0.28, y + 0.06, 1.15, 0.32, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, rx + 0.28, y + 0.06, 1.15, 0.32, [
            (typ, dict(size=8.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, rx + 1.55, y + 0.04, lw - 3.30, 0.34, [
            (name, dict(size=10.5, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, rx + 1.55, y + 0.36, lw - 1.70, 0.24, [
            (fte, dict(size=9, italic=True, color=color, bold=True))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, rx + 0.28, y + 0.60, lw - 0.55, 0.30, [
            (note, dict(size=9, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)
    T(s, rx + 0.15, top + h - 0.35, lw - 0.30, 0.30, [
        ("Gesamt Betriebs-FTE:  ~ 2,5 FTE IT · ~ 0,7 FTE / Ressort Fachbereich",
         dict(size=10.5, bold=True, color=NAVY, align=PP_ALIGN.RIGHT))],
        anchor=MSO_ANCHOR.MIDDLE)
s_a3_aufwaende()

# ================================
# 12 · AUFGABE 4 — RISIKOMATRIX (BoB vs Plattform)
# ================================
def s_a4_risiko():
    s = add_stihl("AUFGABE 4 — Risikomatrix neu:  Best-of-Breed vs. SNOW als Plattform")
    R(s, CL, CT, CW, 0.55, GREEN_LT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    R(s, CL, CT, 0.08, 0.55, GREEN)
    T(s, CL + 0.15, CT, CW - 0.25, 0.55, [
        ("Umdenken: heute BoB-Risiken (viele Tools) — kuenftig Plattform-Risiken (eine Plattform)",
         dict(size=10.5, italic=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    lw = (CW - 0.20) / 2
    rx = CL + lw + 0.20
    top = CT + 0.75
    h = CB - top
    # LEFT: BoB Risiken
    R(s, CL, top, lw, h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    R(s, CL, top, lw, 0.45, RED_HL, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    T(s, CL + 0.15, top, lw - 0.25, 0.45, [
        ("HEUTE — Best-of-Breed-Risiken (~ 6 Tools)",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    bob_risks = [
        ("Multi-Vendor-Abhaengigkeit",   "je Tool eigenes Update-/Support-Risiko"),
        ("Schnittstellen-Bruchgefahr",   "APIs zwischen Tools brechen bei Upgrades"),
        ("Datenkonsistenz nicht sicher", "Insel-Datenmodelle, Aggregation manuell"),
        ("Support-EOL (PIT/MSPO)",       "Bestandsprodukte werden abgekuendigt"),
        ("Hohe Betriebs-FTE",            "~ 4-6 FTE fuer Multi-Tool-Betrieb"),
        ("Governance-Fragmentierung",    "je Tool eigene Rollen, keine Uebersicht"),
    ]
    for i, (r, sub) in enumerate(bob_risks):
        y = top + 0.55 + i * 0.85
        R(s, CL + 0.15, y, lw - 0.30, 0.78, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
          line=RED_HL, lw=0.5)
        R(s, CL + 0.15, y, 0.06, 0.78, RED_HL)
        O(s, CL + 0.30, y + 0.28, 0.24, 0.24, RED_HL)
        T(s, CL + 0.30, y + 0.28, 0.24, 0.24, [
            ("!", dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, CL + 0.65, y + 0.06, lw - 0.85, 0.32, [
            (r, dict(size=10.5, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, CL + 0.65, y + 0.38, lw - 0.85, 0.35, [
            (sub, dict(size=9, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)
    # RIGHT: Plattform Risiken
    R(s, rx, top, lw, h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    R(s, rx, top, lw, 0.45, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    T(s, rx + 0.15, top, lw - 0.25, 0.45, [
        ("MIT SEPM — Plattform-Risiken (1 Plattform)",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    plat_risks = [
        ("Vendor Lock-in SNOW",          "Standard-Config, Migrations-Klauseln, Marktscreening"),
        ("Preis-Eskalation Renewal",     "Vertragsklausel 7% nach 4 Jahren (siehe F14)"),
        ("Plattform-Ausfall",            "SLA 99,9%, DC Frankfurt/Duesseldorf, DSGVO"),
        ("Roadmap-Divergenz SNOW",       "Standard first · halbjaehrliche Roadmap-Reviews"),
        ("Kompetenz-Aufbau",             "Erweiterung ITSM-Team +1,5-2 FTE, Ausbildung"),
        ("Governance-Zentralismus",      "Foederiertes Modell (Empfehlung B) mildert das ab"),
    ]
    for i, (r, sub) in enumerate(plat_risks):
        y = top + 0.55 + i * 0.85
        R(s, rx + 0.15, y, lw - 0.30, 0.78, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
          line=ORANGE, lw=0.5)
        R(s, rx + 0.15, y, 0.06, 0.78, ORANGE)
        O(s, rx + 0.30, y + 0.28, 0.24, 0.24, ORANGE)
        T(s, rx + 0.30, y + 0.28, 0.24, 0.24, [
            (str(i + 1), dict(size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, rx + 0.65, y + 0.06, lw - 0.85, 0.32, [
            (r, dict(size=10.5, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, rx + 0.65, y + 0.38, lw - 0.85, 0.35, [
            (sub, dict(size=9, italic=True, color=GREEN))],
            anchor=MSO_ANCHOR.MIDDLE)
s_a4_risiko()

# ================================
# 13 · AUFGABE 4 — LOCKIN PRO/CON + ANPASSUNGSAUFWAENDE
# ================================
def s_a4_lockin():
    s = add_stihl("AUFGABE 4 — Lockin Pro/Con  +  Anpassungsaufwaende (Vibe Coding · Beratungsumfeld)")
    lw = (CW - 0.20) / 2
    rx = CL + lw + 0.20
    top = CT
    h = CB - top
    # LEFT: Lockin Pro/Con
    R(s, CL, top, lw, h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    R(s, CL, top, lw, 0.45, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    T(s, CL + 0.15, top, lw - 0.25, 0.45, [
        ("LOCKIN-EFFEKTE — Pro / Con",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    # Pro sub-header
    R(s, CL + 0.15, top + 0.60, lw - 0.30, 0.30, GREEN, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    T(s, CL + 0.15, top + 0.60, lw - 0.30, 0.30, [
        ("PRO — Vorteile der Plattform-Bindung",
         dict(size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    pros = [
        ("Ein Datenmodell", "durchgaengige Traceability, keine Aggregations-Insel"),
        ("Einheitliche Rollen", "reduzierter Governance-Overhead"),
        ("Skaleneffekte Lizenz", "Volumen-Rabatte bei mehreren SNOW-Modulen"),
    ]
    for i, (h_, sub) in enumerate(pros):
        y = top + 1.00 + i * 0.60
        R(s, CL + 0.15, y, lw - 0.30, 0.55, GREEN_LT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        R(s, CL + 0.15, y, 0.06, 0.55, GREEN)
        T(s, CL + 0.28, y + 0.05, 2.30, 0.20, [
            (h_, dict(size=9.5, bold=True, color=GREEN))], anchor=MSO_ANCHOR.MIDDLE)
        T(s, CL + 0.28, y + 0.25, lw - 0.55, 0.28, [
            (sub, dict(size=8.5, italic=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    # Con sub-header
    con_y = top + 1.00 + 3 * 0.60 + 0.15
    R(s, CL + 0.15, con_y, lw - 0.30, 0.30, RED_HL, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    T(s, CL + 0.15, con_y, lw - 0.30, 0.30, [
        ("CON — Nachteile der Plattform-Bindung",
         dict(size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    cons = [
        ("Vendor-Abhaengigkeit",    "SNOW-Preis-/Roadmap-Diktat moeglich"),
        ("Exit-Kosten",             "Migrationsaufwand bei Vendor-Wechsel"),
        ("Skalier-Kosten",          "mit mehr Named Users steigen Lizenzkosten stark"),
    ]
    for i, (h_, sub) in enumerate(cons):
        y = con_y + 0.40 + i * 0.60
        R(s, CL + 0.15, y, lw - 0.30, 0.55, RED_LT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        R(s, CL + 0.15, y, 0.06, 0.55, RED_HL)
        T(s, CL + 0.28, y + 0.05, 2.30, 0.20, [
            (h_, dict(size=9.5, bold=True, color=RED_HL))], anchor=MSO_ANCHOR.MIDDLE)
        T(s, CL + 0.28, y + 0.25, lw - 0.55, 0.28, [
            (sub, dict(size=8.5, italic=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    # RIGHT: Anpassungsaufwaende
    R(s, rx, top, lw, h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    R(s, rx, top, lw, 0.45, PURPLE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    T(s, rx + 0.15, top, lw - 0.25, 0.45, [
        ("ANPASSUNGSAUFWAENDE — Vibe Coding · Beratungsumfeld",
         dict(size=12, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
    anpassungen = [
        ("SNOW Standard",           "Konfiguration via Now Studio",
         "0-5 PT / Prozess",   "kein Coding — nur Klicks", GREEN),
        ("Vibe Coding / App Engine","Low-Code Custom-Apps",
         "5-15 PT / Prozess",  "eigene Entwickler + SNOW-Zertifizierung", ORANGE),
        ("Custom-Skripte (Server)", "JavaScript/GlideScript",
         "10-30 PT / Modul",   "wird moeglichst vermieden — Standard first!", RED_HL),
        ("Beratungsumfeld DE",      "SNOW-Partner (T-Systems, Devoteam, ...)",
         "800-1500 €/PT",      "reifer Markt in Deutschland, Wettbewerb", GREEN),
        ("STIHL internes Know-how", "ITSM-Team seit 2014 + Ausbau",
         "wachsend",           "kein Blackbox-Wissen, gute Basis", GREEN),
    ]
    for i, (kat, sub, aufw, note, color) in enumerate(anpassungen):
        y = top + 0.60 + i * 1.08
        R(s, rx + 0.15, y, lw - 0.30, 1.00, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
          line=color, lw=0.5)
        R(s, rx + 0.15, y, 0.06, 1.00, color)
        T(s, rx + 0.28, y + 0.04, lw - 2.20, 0.28, [
            (kat, dict(size=10.5, bold=True, color=DARK))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, rx + 0.28, y + 0.30, lw - 2.20, 0.28, [
            (sub, dict(size=9, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, rx + 0.28, y + 0.60, lw - 0.55, 0.36, [
            (note, dict(size=9, color=color))], anchor=MSO_ANCHOR.MIDDLE)
        R(s, rx + lw - 1.90, y + 0.30, 1.75, 0.40, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, rx + lw - 1.90, y + 0.30, 1.75, 0.40, [
            (aufw, dict(size=9.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
s_a4_lockin()

# ================================
# 14 · AUFGABE 4 — SAP vs SNOW + VERTRAGSHISTORIE (7% Klausel)
# ================================
def s_a4_sap_history():
    s = add_stihl("AUFGABE 4 — SAP vs. ServiceNow-Vergleich  +  Vertragshistorie mit Erhoehungsklausel")
    # Top: Comparison table
    T(s, CL, CT, CW, 0.28, [
        ("Vergleich Plattform-Abhaengigkeit",
         dict(size=12.5, bold=True, color=ORANGE))])
    hdr_y = CT + 0.35
    R(s, CL, hdr_y, CW, 0.32, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    cols = [("Dimension", 3.6), ("SAP (Bestand)", 4.5), ("ServiceNow (SEPM)", 4.5), ("Bewertung", 0.07)]
    xc = CL
    for name, w in cols[:-1]:
        T(s, xc + 0.15, hdr_y, w - 0.15, 0.32, [
            (name, dict(size=10.5, bold=True, color=WHITE))], anchor=MSO_ANCHOR.MIDDLE)
        xc += w
    T(s, CL + CW - 0.55, hdr_y, 0.50, 0.32, [
        ("Δ", dict(size=10.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    vgl = [
        ("Lockin-Historie STIHL",   "seit ~ 25 Jahren, tief verwurzelt",  "seit 2014, kontrolliert wachsend", "≈"),
        ("Preisentwicklung",        "regelmaessige Anpassungen",           "Renewal +7 % nach 4 Jahren",       "="),
        ("Migrations-Optionen",     "sehr eingeschraenkt",                  "Standard-Exports, APIs, Migrationstools", "+"),
        ("Marktreife DACH-Partner", "sehr hoch",                            "hoch — Devoteam, T-Systems, ...",  "≈"),
        ("Cloud-Betrieb",           "S/4HANA-Roadmap",                      "SaaS DE, DSGVO, DC FRA/DUS",       "+"),
        ("AI-Faehigkeiten",         "Joule (im Aufbau)",                    "Now Assist (produktiv seit 2024)", "+"),
    ]
    for i, (dim, sap, snow, delta) in enumerate(vgl):
        y = hdr_y + 0.36 + i * 0.45
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        R(s, CL, y, CW, 0.42, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, CL + 0.15, y, 3.45, 0.42, [
            (dim, dict(size=10, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        T(s, CL + 3.75, y, 4.35, 0.42, [
            (sap, dict(size=9.5, color=NAVY))], anchor=MSO_ANCHOR.MIDDLE)
        T(s, CL + 8.25, y, 4.35, 0.42, [
            (snow, dict(size=9.5, color=NAVY))], anchor=MSO_ANCHOR.MIDDLE)
        d_col = GREEN if delta == "+" else (ORANGE if delta in ("=", "≈") else RED_HL)
        R(s, CL + CW - 0.55, y + 0.06, 0.42, 0.30, d_col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, CL + CW - 0.55, y + 0.06, 0.42, 0.30, [
            (delta, dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    # Middle: Vertragshistorie bar chart with 7%-Klausel highlight
    ch_top = hdr_y + 0.36 + len(vgl) * 0.45 + 0.20
    ch_h = 1.60
    T(s, CL, ch_top - 0.30, CW, 0.28, [
        ("Vertragshistorie ServiceNow bei STIHL  —  Renewal 2028 mit +7 % Erhoehungsklausel",
         dict(size=11.5, bold=True, color=ORANGE))])
    R(s, CL, ch_top, CW, ch_h, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    hist = [(2014,180),(2016,260),(2018,340),(2020,410),(2022,490),(2024,560),(2026,820),(2028,878)]
    max_v = max(v for _, v in hist)
    def xp(yr): return CL + 0.35 + (yr - 2014) / (2028 - 2014) * (CW - 0.70)
    R(s, CL + 0.35, ch_top + ch_h - 0.30, CW - 0.70, 0.02, DARK)
    for yr, val in hist:
        x = xp(yr)
        hh = (val / max_v) * (ch_h - 0.65)
        y = ch_top + ch_h - 0.30 - hh
        color = RED_HL if yr == 2028 else ORANGE
        R(s, x - 0.16, y, 0.32, hh, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        O(s, x - 0.10, y - 0.10, 0.20, 0.20, NAVY)
        T(s, x - 0.85, y - 0.32, 1.70, 0.20, [
            (f"{val}", dict(size=9, bold=True, color=color, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, x - 0.85, ch_top + ch_h - 0.24, 1.70, 0.22, [
            (str(yr), dict(size=9, bold=True, color=DARK, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    # 7% Klausel annotation
    ann_x = xp(2028)
    R(s, ann_x - 1.30, ch_top + 0.06, 1.20, 0.32, RED_HL, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    T(s, ann_x - 1.30, ch_top + 0.06, 1.20, 0.32, [
        ("+7 % Klausel", dict(size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    # Fazit
    fy = ch_top + ch_h + 0.20
    R(s, CL, fy, CW, 0.45, GREEN_LT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    R(s, CL, fy, 0.08, 0.45, GREEN)
    T(s, CL + 0.15, fy, CW - 0.25, 0.45, [
        ("Fazit: SNOW-Abhaengigkeit strukturell aehnlich SAP — kontrollierbar durch Standard-first + Vertragsklauseln.",
         dict(size=10, italic=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
s_a4_sap_history()

# ================================
# 15 · AUFGABE 4 — REFERENZKUNDEN
# ================================
def s_a4_referenzen():
    s = add_stihl("AUFGABE 4 — Referenzkunden ServiceNow  (Grosskonzerne in Deutschland)")
    R(s, CL, CT, CW, 0.55, GREEN_LT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    R(s, CL, CT, 0.08, 0.55, GREEN)
    T(s, CL + 0.15, CT, CW - 0.25, 0.55, [
        (">85 % der Fortune 500 · nahezu alle DAX-Konzerne · fuehrende deutsche Industrieunternehmen",
         dict(size=11, italic=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    customers = [
        ("SIEMENS",        "Industrie / Tech",       RGBColor(0x00,0x9D,0x9D), WHITE),
        ("SAP",            "Software",               RGBColor(0x00,0x3A,0x5D), WHITE),
        ("Allianz",        "Versicherung",           RGBColor(0x00,0x37,0x81), WHITE),
        ("Deutsche Bank",  "Banking",                RGBColor(0x00,0x18,0xA8), WHITE),
        ("T",              "Deutsche Telekom",       RGBColor(0xE2,0x00,0x74), WHITE),
        ("BOSCH",          "Industrie / Auto",       RGBColor(0xC8,0x10,0x2E), WHITE),
        ("Mercedes-Benz",  "Automotive",             RGBColor(0x00,0x00,0x00), WHITE),
        ("Volkswagen",     "Automotive",             RGBColor(0x00,0x1E,0x50), WHITE),
    ]
    box_top = CT + 0.85
    box_h = 2.20
    n = len(customers)
    gap = 0.15
    box_w = (CW - (n - 1) * gap) / n
    for i, (name, branche, bg, fg) in enumerate(customers):
        x = CL + i * (box_w + gap)
        R(s, x, box_top, box_w, box_h, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        name_size = 20 if len(name) <= 4 else (16 if len(name) <= 8 else 13)
        T(s, x, box_top, box_w, box_h, [
            (name, dict(size=name_size, bold=True, color=fg, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, x, box_top + box_h + 0.05, box_w, 0.28, [
            (branche, dict(size=9, italic=True, color=GREY_TXT, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.TOP)
    # Below: Use-Case-Kategorien
    uc_top = box_top + box_h + 0.60
    T(s, CL, uc_top, CW, 0.28, [
        ("Referenz-Use-Cases im Markt (nicht nur IT):",
         dict(size=11.5, bold=True, color=ORANGE))])
    ucs = [
        ("ITSM", "Basis-Modul, seit 10+ Jahren produktiv", NAVY),
        ("SPM / Portfolio", "Wachstumsbereich, viele Bosch/Siemens-Referenzen", ORANGE),
        ("HR Service Delivery", "Employee Portal, Onboarding-Workflows", GREEN),
        ("Customer Service Mgmt", "Field Service, Kundenprozesse (v.a. Telekom)", BLUE_HD),
        ("GRC / Risk", "Compliance-Automation (v.a. Banking / Insurance)", PURPLE),
    ]
    uw = (CW - 4 * 0.15) / 5
    for i, (name, note, color) in enumerate(ucs):
        x = CL + i * (uw + 0.15)
        R(s, x, uc_top + 0.35, uw, 1.35, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
          line=color, lw=1.0)
        R(s, x, uc_top + 0.35, uw, 0.35, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, x, uc_top + 0.35, uw, 0.35, [
            (name, dict(size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, x + 0.15, uc_top + 0.80, uw - 0.30, 0.85, [
            (note, dict(size=9, italic=True, color=DARK, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
s_a4_referenzen()

# ================================
# 16 · GESAMT — QUALITATIVE ANALYSE (5 Dimensionen)
# ================================
def s_g_qualitative():
    s = add_stihl("GESAMTDARSTELLUNG — Qualitative Analyse der Plattform  (5 Dimensionen)")
    dims = [
        ("Datenintegritaet",  GREEN,   "ein Datenmodell, keine Insel-Aggregation",
         5, "OKR → Portfolio → Projekt → Kosten durchgaengig verfolgbar"),
        ("KI (Now Assist)",   PURPLE,  "GenAI produktiv seit 2024, Vibe Coding moeglich",
         4, "Risk-Frueherkennung, Auto-Reports, Chat-Assistent"),
        ("Schnelligkeit",     ORANGE,  "Reporting live, Freigaben in Stunden statt Wochen",
         5, "Live-Dashboards, Auto-Approvals nach Regelwerk"),
        ("Prozesse",          NAVY,    "Standard-Prozesse SPM Pro + App-Engine-Erweiterungen",
         4, "Standardisiert wo moeglich, foederiert wo noetig"),
        ("Transparenz",       BLUE_HD, "Vorstand sieht Portfolio-Status in Echtzeit",
         5, "Ampeln, Budget, OKR-Beitrag pro Vorhaben"),
    ]
    top = CT
    h = (CB - top - 0.20) / len(dims)
    for i, (name, color, sub, score, benefit) in enumerate(dims):
        y = top + i * h
        R(s, CL, y, CW, h - 0.10, LIGHT_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        R(s, CL, y, 0.15, h - 0.10, color)
        # Icon-Zone
        O(s, CL + 0.30, y + 0.20, 0.60, 0.60, color)
        T(s, CL + 0.30, y + 0.20, 0.60, 0.60, [
            (str(i + 1), dict(size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        # Title + sub
        T(s, CL + 1.10, y + 0.10, 4.20, 0.42, [
            (name, dict(size=14, bold=True, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
        T(s, CL + 1.10, y + 0.50, 4.20, 0.32, [
            (sub, dict(size=9.5, italic=True, color=GREY_TXT))], anchor=MSO_ANCHOR.MIDDLE)
        # Score visualization (5 dots)
        star_x = CL + 5.50
        for k in range(5):
            filled = k < score
            O(s, star_x + k * 0.36, y + 0.30, 0.30, 0.30,
              color if filled else LIGHT_GREY)
        T(s, star_x + 5 * 0.36 + 0.15, y + 0.20, 0.6, 0.50, [
            (f"{score}/5", dict(size=11, bold=True, color=color, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        # Benefit text
        T(s, CL + 8.20, y + 0.15, CW - 8.35, h - 0.35, [
            (benefit, dict(size=10.5, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    T(s, CL, CB - 0.15, CW, 0.20, [
        ("Gesamt-Bewertung:  4,6 / 5,0  —  reife Plattform, mit klarem Vorteil in allen 5 Dimensionen",
         dict(size=10, italic=True, color=GREEN, bold=True, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
s_g_qualitative()

# ================================
# 17 · GESAMT — STRATEGISCHE AUSRICHTUNG / NEUE GESCHAEFTSFELDER
# ================================
def s_g_strategisch():
    s = add_stihl("GESAMTDARSTELLUNG — Strategische Ausrichtung + neue Geschaeftsfelder")
    T(s, CL, CT, CW, 0.28, [
        ("Von PPM-Konsolidierung zu einer echten Plattform-Strategie fuer STIHL",
         dict(size=11.5, italic=True, color=GREY_TXT))])
    # Horizontal roadmap
    tl_top = CT + 0.50
    tl_h = 0.50
    R(s, CL, tl_top, CW, tl_h, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    stages = ["Heute 2026 — ITSM", "2027 — SEPM Phase 1",
              "2028 — SEPM Phase 2", "2029+ — HR/CSM/GRC", "Zukunft — Vibe Coding"]
    seg_w = CW / len(stages)
    for i, name in enumerate(stages):
        x = CL + i * seg_w
        T(s, x, tl_top, seg_w, tl_h, [
            (name, dict(size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        if i < len(stages) - 1:
            R(s, x + seg_w - 0.03, tl_top + 0.10, 0.06, tl_h - 0.20, WHITE)
    # Field cards
    top = tl_top + tl_h + 0.35
    fields = [
        ("SEPM heute",            "Portfolio · Projekt · Ressourcen · Reporting",
         "Ableitung aus laufendem Business Case",                              ORANGE, "In Umsetzung"),
        ("HR Service Delivery",   "Onboarding, Employee Self-Service, HR-Ticketing",
         "Nutzung SNOW HR-Modul, gemeinsame Plattform mit ITSM",               GREEN,  "12-18 Mo Vorlauf"),
        ("Customer Service Mgmt", "Field Service Techniker, Ersatzteil-Case-Mgmt",
         "SNOW CSM/FSM — Bindeglied zu Aftersales & Handel",                   BLUE_HD,"Potenzial gross"),
        ("GRC / Risk-Mgmt",       "Compliance-Automation, DSGVO, Audit-Trails",
         "SNOW GRC — Ergaenzung zu SAP-Kernprozessen",                          PURPLE, "Mittelfristig"),
        ("Custom Solutions (App Engine)","Vibe Coding fuer STIHL-spezifische Workflows",
         "Interne Loesungen ohne 3rd-Party-Tool, schneller als Custom Coding", NAVY,   "Ab Jahr 2"),
        ("KI / Now Assist",       "Generative KI ueber alle Bereiche hinweg",
         "STIHL-eigene GPT-Anwendungen auf SNOW-Datenbasis",                    RED_HL, "Zukunft")
    ]
    w = (CW - 0.30) / 3
    h = 2.10
    gap = 0.15
    for i, (name, sub, note, color, badge) in enumerate(fields):
        row = i // 3; col = i % 3
        x = CL + col * (w + gap)
        y = top + row * (h + gap)
        R(s, x, y, w, h, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=color, lw=1.2)
        R(s, x, y, w, 0.42, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, x + 0.15, y + 0.04, w - 0.30, 0.34, [
            (name, dict(size=12, bold=True, color=WHITE))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, x + 0.15, y + 0.55, w - 0.30, 0.36, [
            (sub, dict(size=9.5, italic=True, color=GREY_TXT))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, x + 0.15, y + 1.00, w - 0.30, 0.70, [
            (note, dict(size=9.5, color=DARK))], anchor=MSO_ANCHOR.TOP)
        R(s, x + w - 1.80, y + h - 0.45, 1.65, 0.34, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, x + w - 1.80, y + h - 0.45, 1.65, 0.34, [
            (badge, dict(size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
s_g_strategisch()

# ================================
# 18 · GESAMT — TCO UND BETRIEB
# ================================
def s_g_tco():
    s = add_stihl("GESAMTDARSTELLUNG — TCO und Betrieb  (kompakte Zusammenfassung)")
    # KPI-Kacheln oben
    kpis = [
        ("5-Jahres-TCO",   "~ 7,0 M€",  "einmalig + laufend",       NAVY),
        ("Nutzen (5 J.)",  "~ 22,7 M€", "Konsol.+Value+Flanking",   GREEN),
        ("Netto-Effekt",   "~ +15,7 M€","Cash Flow ueber 5 Jahre",  ORANGE),
        ("Betriebs-FTE",   "~ 2,5 FTE", "im Endzustand, Erweiterung ITSM", BLUE_HD),
    ]
    top = CT
    kpi_h = 1.35
    gap = 0.15
    w = (CW - 3 * gap) / 4
    for i, (label, big, sub, color) in enumerate(kpis):
        x = CL + i * (w + gap)
        R(s, x, top, w, kpi_h, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=color, lw=1.5)
        R(s, x, top, w, 0.32, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, x, top, w, 0.32, [
            (label, dict(size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, x, top + 0.38, w, 0.55, [
            (big, dict(size=22, bold=True, color=color, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, x, top + 1.00, w, 0.30, [
            (sub, dict(size=9, italic=True, color=GREY_TXT, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
    # Cost + Value breakdown
    tt = top + kpi_h + 0.30
    T(s, CL, tt, CW, 0.28, [
        ("Zusammensetzung TCO und Nutzen (5 Jahre kumuliert, Base Case)",
         dict(size=12, bold=True, color=ORANGE))])
    rows = [
        ("KOSTEN — Implementierung + Change (einmalig)",       "- 3,0 M€",  RED_HL),
        ("KOSTEN — Lizenz + Betrieb + Weiterentwicklung (5 J.)", "- 4,0 M€",  RED_HL),
        ("NUTZEN — Konsolidierung Tools & Schnittstellen",     "+ 5,8 M€", GREEN),
        ("NUTZEN — Business Value (Effizienz, TTM, Governance)","+ 11,6 M€",GREEN),
        ("NUTZEN — Flankierende Prozesse",                     "+ 5,3 M€", GREEN),
        ("NETTO 5 Jahre",                                       "+ 15,7 M€",DARK),
    ]
    tt2 = tt + 0.35
    for i, (name, val, color) in enumerate(rows):
        y = tt2 + i * 0.42
        bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
        is_total = (i == len(rows) - 1)
        R(s, CL, y, CW, 0.38, bg if not is_total else DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, CL + 0.15, y + 0.02, 9.0, 0.34, [
            (name, dict(size=11, bold=is_total, color=DARK if not is_total else WHITE))],
            anchor=MSO_ANCHOR.MIDDLE)
        T(s, CL + 9.5, y + 0.02, CW - 9.65, 0.34, [
            (val, dict(size=14 if is_total else 12, bold=True,
                       color=color if not is_total else ORANGE, align=PP_ALIGN.RIGHT))],
            anchor=MSO_ANCHOR.MIDDLE)
s_g_tco()

# ================================
# 19 · GESAMT — UMSETZUNGSVORSCHLAG SZENARIO 1 + 2
# ================================
def s_g_szenarien():
    s = add_stihl("GESAMTDARSTELLUNG — Umsetzungsvorschlag Szenario 1 & 2")
    lw = (CW - 0.20) / 2
    rx = CL + lw + 0.20
    top = CT
    h = CB - top - 0.55
    szenarien = [
        (CL, "SZENARIO 1 — Full Scope 24 Monate", ORANGE, "★ EMPFEHLUNG",
         [("Scope",       "Alle Module Phase 1 (Demand · Strategy · Portfolio · Reporting) + Phase 2 (PM · Ress.)"),
          ("Zeit",        "24 Monate ab Kickoff (Nov 2026)  ·  Go-Live Phase 1 Q1/2028"),
          ("Kosten",      "~ 3,0 M€ einmalig  ·  ~ 0,8 M€ p.a. Steady State"),
          ("Milestones",  "Q4/26 Kickoff · Q3/27 Phase 1 UAT · Q1/28 Go-Live · Q3/28 Phase 2 GL"),
          ("Quickwins",   "Demand-Backlog (Q2/27) · Live-Portfolio-Dashboard (Q3/27) · OKR-Cascade (Q4/27)"),
          ("EV",          "Investitionsbeschluss zur Jahresplanung 2027 (Q3/2026)"),
          ("Risiko",      "Mittel — Change-intensiv, ausreichende Ressourcen noetig")]),
        (rx, "SZENARIO 2 — Extended Ramp 36 Monate", BLUE_HD, "Konservativ",
         [("Scope",       "Alle Module verteilt auf 3 Jahre — Phase 1 in 18 Mo, Phase 2 danach"),
          ("Zeit",        "36 Monate ab Kickoff  ·  Go-Live Phase 1 Q3/2028"),
          ("Kosten",      "~ 2,7 M€ einmalig  ·  ~ 0,75 M€ p.a. Steady State (spaeter)"),
          ("Milestones",  "Q4/26 Kickoff · Q1/28 Phase 1 UAT · Q3/28 Go-Live · Q4/29 Phase 2 GL"),
          ("Quickwins",   "Demand-Backlog erst Q3/27 · Dashboards Q1/28 · Full Portfolio Q3/28"),
          ("EV",          "Zweistufig: Rahmen 2027, Vollbudget 2028"),
          ("Risiko",      "Niedrig — mehr Puffer, aber laengerer Payback")]),
    ]
    for xoff, title, color, badge, items in szenarien:
        R(s, xoff, top, lw, h, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=color, lw=1.5)
        R(s, xoff, top, lw, 0.55, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, xoff + 0.15, top + 0.05, lw - 2.10, 0.45, [
            (title, dict(size=12.5, bold=True, color=WHITE))],
            anchor=MSO_ANCHOR.MIDDLE)
        R(s, xoff + lw - 1.90, top + 0.10, 1.75, 0.35, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, xoff + lw - 1.90, top + 0.10, 1.75, 0.35, [
            (badge, dict(size=9, bold=True, color=color, align=PP_ALIGN.CENTER))],
            anchor=MSO_ANCHOR.MIDDLE)
        # Items rows
        for j, (k, v) in enumerate(items):
            y = top + 0.70 + j * 0.65
            bg = LIGHT_BG if j % 2 == 0 else ROW_ALT
            R(s, xoff + 0.10, y, lw - 0.20, 0.60, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            R(s, xoff + 0.10, y, 0.06, 0.60, color)
            T(s, xoff + 0.25, y + 0.04, 1.30, 0.52, [
                (k, dict(size=9.5, bold=True, color=color))], anchor=MSO_ANCHOR.MIDDLE)
            T(s, xoff + 1.65, y + 0.04, lw - 1.80, 0.52, [
                (v, dict(size=9.5, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
    # Recommendation bar
    ry = top + h + 0.15
    R(s, CL, ry, CW, 0.40, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    T(s, CL + 0.20, ry, CW - 0.40, 0.40, [
        ("Empfehlung:  Szenario 1 — Full Scope 24 Monate. Deutlich schnellerer Payback + hoehere Nutzen-Kumulation.",
         dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
s_g_szenarien()

# ================================
# 20 · GESAMT — EMPFEHLUNG + EV
# ================================
def s_g_empfehlung():
    s = add_stihl("GESAMTDARSTELLUNG — Empfehlung + Entscheidungsvorlage (EV)")
    # Kernaussage banner
    R(s, CL, CT, CW, 0.75, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    T(s, CL + 0.20, CT, CW - 0.40, 0.75, [
        ("★ BESCHLUSSVORSCHLAG:  Umsetzung Szenario 1 (Full Scope 24 Mo) · Governance-Modell B (Foederiert)",
         dict(size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER))],
        anchor=MSO_ANCHOR.MIDDLE)
    # 4 EV-Bausteine
    top = CT + 0.95
    h = 2.30
    gap = 0.15
    w = (CW - gap) / 2
    ev_blocks = [
        ("1 — Was wird beschlossen?", GREEN,
         ["Investitionsentscheidung SEPM (Szenario 1) fuer Jahresplanung 2027",
          "Governance-Modell B (Foederiert) — Standard first, definierte Varianten je Ressort",
          "SI-Partner-Auswahl per RFP im Q4/2026",
          "Kickoff Detail-Konzept Q4/2026, Umsetzungs-Start Q1/2027"]),
        ("2 — Was ist zu erwarten?", NAVY,
         ["TCO ~ 7,0 M€ ueber 5 Jahre  ·  Netto-Effekt ~ + 15,7 M€",
          "Payback ~ 3,5 Jahre  ·  IRR ~ 32 %  ·  NPV ~ + 8,5 M€",
          "Konsolidierung von 6+ Tools · Ablose PIT im Rahmen Phase 1/2",
          "Foederiertes Governance-Modell — Komplexitaet mittel gehalten"]),
        ("3 — Welche Ressourcen sind noetig?", BLUE_HD,
         ["4 FTE intern (Phase 1+2) · ~ 2,5 FTE Steady State",
          "SI-Partner (SNOW-zertifiziert) · Retainer + T&M",
          "Change-Programm + Trainings (~ 380 k€ einmalig)",
          "Governance-Board etablieren (Q4/2026)"]),
        ("4 — Wann und wie geht es weiter?", ORANGE,
         ["Q3/2026: EV zur Jahresplanung (Vorstand)",
          "Q4/2026: SI-Auswahl + Detail-Konzept",
          "Q1/2027: Kickoff Umsetzung Phase 1",
          "Q1/2028: Go-Live Phase 1 (Demand, Portfolio, Strategy)"]),
    ]
    for i, (title, color, items) in enumerate(ev_blocks):
        row = i // 2; col = i % 2
        x = CL + col * (w + gap)
        y = top + row * (h + gap)
        R(s, x, y, w, h, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=color, lw=1.5)
        R(s, x, y, w, 0.42, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        T(s, x + 0.15, y + 0.04, w - 0.30, 0.34, [
            (title, dict(size=12.5, bold=True, color=WHITE))],
            anchor=MSO_ANCHOR.MIDDLE)
        for j, it in enumerate(items):
            iy = y + 0.55 + j * 0.42
            O(s, x + 0.20, iy + 0.12, 0.16, 0.16, color)
            T(s, x + 0.45, iy, w - 0.55, 0.40, [
                (it, dict(size=10, color=DARK))], anchor=MSO_ANCHOR.MIDDLE)
s_g_empfehlung()

# ================================
# 21 · VIELEN DANK
# ================================
def s_thanks():
    s = prs.slides.add_slide(LConcl)
    for ph in s.placeholders:
        ph.text_frame.text = "Vielen Dank"
        break
s_thanks()

prs.save(DST)
print(f"Saved: {DST}  ({len(prs.slides)} slides)")
