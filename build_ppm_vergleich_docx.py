"""Build Word-Bericht: ServiceNow SPM Pro vs. Planisware."""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DST = "/tmp/2026_10_15_PPM_Vergleich_SNOW_Planisware.docx"

# STIHL colors
ORANGE   = RGBColor(0xF0, 0x7F, 0x12)
DARK     = RGBColor(0x1F, 0x29, 0x37)
GREY_TXT = RGBColor(0x4B, 0x55, 0x63)
NAVY     = RGBColor(0x2C, 0x3E, 0x50)
GREEN    = RGBColor(0x2E, 0x7D, 0x32)
RED_HL   = RGBColor(0xB9, 0x1C, 0x1C)
LIGHT_BG_HEX = "F7F7F7"
ORANGE_HEX   = "F07F12"
NAVY_HEX     = "2C3E50"
GREEN_HEX    = "2E7D32"
ORANGE_LT_HEX= "FDE6CC"
GREY_BG_HEX  = "E9EDF1"

doc = Document()

for section in doc.sections:
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

styles = doc.styles
normal = styles["Normal"]
normal.font.name = "Calibri"
normal.font.size = Pt(11)
normal.font.color.rgb = DARK

def add_heading(text, level=1, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18 if level == 1 else 14)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.bold = True
    r.font.size = Pt({1: 20, 2: 15, 3: 12}.get(level, 12))
    r.font.color.rgb = color or (ORANGE if level == 1 else NAVY)
    return p

def add_para(text, italic=False, bold=False, color=None, size=11, space_after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.italic = italic
    r.bold = bold
    r.font.size = Pt(size)
    if color is not None:
        r.font.color.rgb = color
    return p

def add_bullets(items):
    for it in items:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(it)
        r.font.name = "Calibri"
        r.font.size = Pt(11)

def cell_shade(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)

def cell_set_text(cell, text, *, bold=False, color=None, size=10.5, align=None,
                  hex_fill=None, italic=False):
    cell.text = ""
    p = cell.paragraphs[0]
    if align is not None:
        p.alignment = align
    r = p.add_run(str(text))
    r.font.name = "Calibri"
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    if color is not None:
        r.font.color.rgb = color
    if hex_fill:
        cell_shade(cell, hex_fill)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

def set_cell_border(cell, color_hex="D1D5DB", size="4"):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        b = OxmlElement(f"w:{edge}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), size)
        b.set(qn("w:color"), color_hex)
        tcBorders.append(b)
    tcPr.append(tcBorders)

def build_table(headers, rows, col_widths_cm=None, header_fill=NAVY_HEX,
                first_col_fill=None, first_col_bold=True, alt_row_fill=True,
                center_cols=None):
    tbl = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    if col_widths_cm:
        for i, w in enumerate(col_widths_cm):
            for c in range(len(rows) + 1):
                tbl.cell(c, i).width = Cm(w)
    for i, h in enumerate(headers):
        c = tbl.cell(0, i)
        cell_set_text(c, h, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF),
                      size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, hex_fill=header_fill)
        set_cell_border(c)
    for r, row in enumerate(rows, start=1):
        for i, v in enumerate(row):
            c = tbl.cell(r, i)
            fill = None
            if alt_row_fill and r % 2 == 0:
                fill = "EEF1F4"
            elif alt_row_fill:
                fill = LIGHT_BG_HEX
            if i == 0 and first_col_fill:
                fill = first_col_fill
            align = None
            if center_cols and i in center_cols:
                align = WD_ALIGN_PARAGRAPH.CENTER
            cell_set_text(c, v, bold=(i == 0 and first_col_bold),
                          size=10.5, align=align, hex_fill=fill)
            set_cell_border(c)
    return tbl

def add_callout(text, color_hex, title=None):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tbl.columns[0].width = Cm(17)
    c = tbl.cell(0, 0)
    c.text = ""
    if title:
        p = c.paragraphs[0]
        r = p.add_run(title + "\n")
        r.bold = True
        r.font.name = "Calibri"
        r.font.size = Pt(12)
    p2 = c.add_paragraph()
    r2 = p2.add_run(text)
    r2.font.name = "Calibri"
    r2.font.size = Pt(10.5)
    cell_shade(c, color_hex)
    set_cell_border(c, color_hex="808080", size="2")
    doc.add_paragraph()

# ============================================================
# TITLE + META
# ============================================================
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
r = p.add_run("ServiceNow SPM Pro vs. Planisware")
r.bold = True
r.font.name = "Calibri"
r.font.size = Pt(24)
r.font.color.rgb = ORANGE

p = doc.add_paragraph()
r = p.add_run("Vergleich und Empfehlung im Kontext einer bestehenden ServiceNow-Basis")
r.font.name = "Calibri"
r.font.size = Pt(13)
r.font.color.rgb = GREY_TXT
r.italic = True

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(12)
meta = p.add_run(
    "Referenz: Gartner Magic Quadrant for Adaptive Project Management and Reporting (APMR), "
    "3. August 2026, ID G00842775  ·  "
    "Fokus: Projekt-Management-Fähigkeiten im STIHL-Kontext"
)
meta.font.size = Pt(10)
meta.font.color.rgb = GREY_TXT
meta.italic = True

p = doc.add_paragraph()
pPr = p._p.get_or_add_pPr()
pBdr = OxmlElement("w:pBdr")
bottom = OxmlElement("w:bottom")
bottom.set(qn("w:val"), "single")
bottom.set(qn("w:sz"), "8")
bottom.set(qn("w:space"), "1")
bottom.set(qn("w:color"), ORANGE_HEX)
pBdr.append(bottom)
pPr.append(pBdr)

# ============================================================
# EXECUTIVE SUMMARY
# ============================================================
add_heading("Executive Summary", level=1)

add_para(
    "Planisware wird im aktuellen Gartner APMR Magic Quadrant (3. August 2026) als Leader "
    "geführt — mit Planisware Orchestra als Kernprodukt und dem AI-Agent Oscar für Portfolio- "
    'und Projektmanagement-Aufgaben. Gartner attestiert Planisware ein starkes „Market '
    'Understanding" und eine belastbare Post-IPO-Stabilität.',
    space_after=8
)
add_para(
    "ServiceNow SPM Pro wird im APMR Magic Quadrant nicht bewertet — nicht wegen fehlender "
    "Qualität, sondern per Gartner-Ausschlussregel: Plattform-Vendors, die APMR als Modul "
    "ihrer Basis-Plattform anbieten, sind vom APMR-MQ ausgeschlossen. Für STIHL ist genau "
    "dieser Plattform-Charakter der entscheidende Vorteil.",
    space_after=8
)
add_para(
    "Für ein Unternehmen mit bestehender ServiceNow-Basis und aktivem SPM Demand-Management "
    "ist die Empfehlung: ServiceNow SPM Pro als strategische Ziel-Plattform, ergänzt durch "
    "App Engine für spezifische Prozesse (z.B. Stage-Gate-PEP im Bereich Entwicklung "
    "Waldwirtschaft, EWW). Planisware bleibt als Hedge-Option für den engen R&D-Bereich "
    "denkbar, ist aber wirtschaftlich meist unterlegen.",
    space_after=12
)

add_callout(
    "ServiceNow SPM Pro empfohlen — Datenkontinuität, Betriebskontinuität und Time-to-Value "
    "überwiegen deutlich die Funktionsvorteile des spezialisierten Best-of-Breed-Tools.",
    color_hex="DDF1DE",
    title="🎯 Kernempfehlung"
)

# ============================================================
# METHODISCHER HINWEIS ZUM GARTNER-MQ
# ============================================================
add_heading("Einordnung: Warum ServiceNow im APMR MQ fehlt", level=1)

add_para(
    "Der Gartner APMR Magic Quadrant bewertet 11 Vendors: Asana, monday.com, Planforge, "
    "Planisware, Planview, Prism PPM, ProSymmetry, Smartsheet, Triskell, Uppwise und Wrike. "
    "ServiceNow ist bewusst nicht enthalten.",
    space_after=8
)

add_para("Wörtliches Exclusion Criterion des Gartner-Berichts (Seite 17):", space_after=4)
add_callout(
    '„ERP, ITSM, SFA and other similar platform vendors that offer APMR extensions or '
    'modules from their base platforms are not included."'
    "\n\nÜbersetzt: ServiceNow ist als ITSM-Plattform mit APMR-Modul (SPM Pro) per Definition "
    "kein reiner APMR-Player und wird deshalb im MQ nicht gewertet.",
    color_hex="EEF1F4",
    title="📖 Zitat Gartner APMR MQ"
)

add_para(
    "Für den STIHL-Kontext ist das kein Nachteil — im Gegenteil: wir suchen genau den "
    "Plattform-Effekt (Integration mit ITSM, HR-SD, GRC, App Engine). "
    "Reine APMR-Best-of-Breed-Tools wie Planisware bieten diese Integration definitionsgemäß nicht.",
    italic=True, color=GREY_TXT, space_after=12
)

# ============================================================
# STECKBRIEFE
# ============================================================
add_heading("Steckbriefe der beiden Player", level=1)

# Planisware
add_heading("🎯 Planisware — Leader im Gartner APMR MQ 2026", level=2)
add_bullets([
    "Kernprodukt: Planisware Orchestra (SaaS, mehrsprachig)",
    "AI-Agent Oscar: Demand Management, Szenario-Planung, Natural Language Queries, automatisierte Statusberichte",
    "Gartner-Stärken laut MQ 2026: Market Understanding (Reifegrad-Skalierung von Projekt- zu Portfolio-Ebene), Marketing Strategy (out-of-the-box Dashboards), Overall Viability (Post-IPO-Stabilität, neue Präsenz in den Amerikas)",
    "Gartner-Cautions laut MQ 2026: (1) Steile Lernkurve durch breites Feature-Set, (2) Industry-Starter-Packages decken Anforderungen nicht immer ab — Zusatzkonfiguration nötig, (3) AI-Fokus deckt evtl. nicht alle funktionalen APMR-Anforderungen",
    "Betriebsmodell: Insellösung ohne nativen Anschluss an ITSM/HR/CSM/GRC",
])

# ServiceNow
add_heading("🔷 ServiceNow SPM Pro — Plattform-Modul, nicht im APMR MQ bewertet", level=2)
add_bullets([
    "Teil der Now Platform, native Integration mit ITSM, ITOM, HR, CSM, GRC",
    "Kernstärken: Plattform-Integration, Demand-to-Delivery, App Engine, Now Assist (KI)",
    "Zielgruppe: Enterprises mit vorhandener SNOW-Basis, Fokus auf End-to-End-Digitalisierung",
    "Preis-Level: mittel bis hoch, aber deutlicher Skaleneffekt bei bestehendem SNOW-Vertrag",
    "Wichtiger Kontext: Gartner-APMR-MQ schließt Plattform-Vendors wie SNOW definitorisch aus — dies ist kein Qualitätsurteil, sondern eine Marktabgrenzung",
])

doc.add_page_break()

# ============================================================
# VERGLEICHSMATRIX
# ============================================================
add_heading("Head-to-Head: Projekt-Management-Fähigkeiten", level=1)
add_para(
    "Die folgende Matrix vergleicht beide Player anhand von 16 Kriterien im Kontext "
    "Projekt-Management (5 Sterne = Best-in-Class). Die Planisware-Bewertungen "
    "orientieren sich am Gartner APMR MQ 2026, die SNOW-Bewertungen an "
    "STIHL-Erfahrungswerten und Marktbeobachtung.",
    italic=True, color=GREY_TXT, space_after=10
)

compare_rows = [
    ("Klassische Projektplanung (Gantt, WBS, Meilensteine)",  "★★★★★", "★★★★"),
    ("Multi-Projekt-Abhängigkeiten",                          "★★★★★", "★★★★"),
    ("Ressourcen-Management (Skills, Capacity, Optimierung)", "★★★★★", "★★★★"),
    ("Was-wäre-wenn / Szenario-Analysen (Oscar-AI)",          "★★★★★", "★★★★ (Pro)"),
    ("Finanz-Management (Kosten, Forecasts, Investment)",     "★★★★★", "★★★★"),
    ("Programm-Management",                                   "★★★★★", "★★★★"),
    ("Stage-Gate / PEP-Prozesse",                             "★★★★★", "★★★ / ★★★★ mit App Engine"),
    ("Agile / Hybrid-PM (SAFe, Scrum)",                       "★★★★",  "★★★★ (EAP Pro)"),
    ("Reporting & Live-Dashboards",                           "★★★★",  "★★★★★"),
    ("KI / GenAI (Vorhersagen, Auto-Reports)",                "★★★★ (Oscar)", "★★★★★ (Now Assist)"),
    ("Zeiterfassung + Approval",                              "★★★★★", "★★★★"),
    ("Integration mit ITSM / bestehende SNOW-Landschaft",     "★★",    "★★★★★ (nativ)"),
    ("Extensibility / Low-Code Custom",                       "★★★",   "★★★★★ (App Engine)"),
    ("User Experience / Familiarität bei SNOW-Kunden",        "★★",    "★★★★"),
    ("Time-to-Value bei bestehender SNOW-Basis",              "★",     "★★★★★"),
    ("TCO (5J) bei bestehender SNOW-Basis",                   "★",     "★★★★"),
]
build_table(
    headers=["Fähigkeit", "Planisware", "ServiceNow SPM Pro"],
    rows=compare_rows,
    col_widths_cm=[9.0, 3.5, 4.5],
    header_fill=NAVY_HEX,
    center_cols=[1, 2],
)
doc.add_paragraph()

# ============================================================
# FACH-BEWERTUNG
# ============================================================
add_heading("Fach-Bewertung nach Projekt-Management-Aspekten", level=1)

add_heading("Wo Planisware technisch überlegen ist (laut Gartner APMR MQ 2026)", level=3, color=RED_HL)
add_bullets([
    "Reife Portfolio-Ebene mit strukturierter Governance und adaptiven Controls",
    "AI-Agent Oscar mit Fokus auf Demand, Szenario-Planung, Natural Language Queries",
    "Out-of-the-box Dashboards und einheitliche Datenschicht als Fundament für AI",
    "Post-IPO-Stabilität und globale Präsenz (jüngst neue Amerikas-Präsenz)",
    "ABER laut Gartner: Steile Lernkurve, Industry-Packages oft nicht ausreichend, Insel-Betrieb ohne Anschluss an ITSM/HR/CSM",
])

add_heading("Wo ServiceNow SPM Pro überlegen ist", level=3, color=GREEN)
add_bullets([
    "Plattform-Integration — Demand aus ITSM fließt direkt in Projekte (nativ, kein Konnektor)",
    "Time-to-Value — Team, Betrieb und UI im Haus seit 2014 bekannt",
    "KI durchgängig — Now Assist für Reporting, Risikofrüherkennung, Story Generation über alle SNOW-Bereiche",
    "Erweiterbarkeit über App Engine — spezifische PEP-Steps ohne 3rd-Party-Tool nachbaubar",
    "Ganzheitliche Sicht über IT-Projekte, Business-Projekte, HR-Projekte, Compliance",
    "ABER: in tiefstem Engineering-PM strukturell weniger reif als das spezialisierte Best-of-Breed-Tool",
])

doc.add_page_break()

# ============================================================
# EMPFEHLUNG STIHL-KONTEXT
# ============================================================
add_heading("Empfehlung für den STIHL-Kontext", level=1)
add_para(
    "Setup: ServiceNow seit 2014 im Einsatz  ·  SPM Demand bereits aktiv  ·  "
    "Migration von altem PM-Tool (PIT/MSPO)",
    italic=True, color=GREY_TXT, space_after=10
)

add_callout(
    "Klare Empfehlung: ServiceNow SPM Pro als strategische PPM-Zielplattform. "
    "Add-on zur bestehenden SNOW-Basis statt Neu-Vendor mit Integrationsaufwand.",
    color_hex="FDE6CC",
    title="👉 Empfehlung"
)

add_heading("Argumente in Vorstands-Reihenfolge", level=2)
argu_rows = [
    ("1", "Datenkontinuität — Demand fließt bereits in SNOW → Projekt-Umsetzung ohne Bruch/Bridge",
         "Sofortiger Prozess-Fluss Demand → Projekt"),
    ("2", "Betriebskontinuität — ITSM-Team seit 12 Jahren im Haus, keine neue Vendor-Beziehung nötig",
         "~ 2 FTE eingespart vs. neuer Vendor"),
    ("3", "UI-Konsistenz — Anwender kennen SNOW-Look, deutlich geringerer Change-Aufwand",
         "~ 200 k€ eingesparte Trainings"),
    ("4", "TCO — Add-on zur bestehenden Plattform statt Zweitvendor + Integration",
         "~ 1,5–2,5 M€/5J. günstiger als Planisware"),
    ("5", "KI-Roadmap — Now Assist skaliert über alle Bereiche, nicht nur PM",
         "Zukunftsinvestition schützt"),
    ("6", "App Engine für PEP — spezifische EWW-Anforderungen ohne 3rd-Party abbildbar",
         "Kein Doppel-System"),
]
build_table(
    headers=["#", "Argument", "Impact"],
    rows=argu_rows,
    col_widths_cm=[1.0, 10.5, 5.5],
    header_fill=NAVY_HEX,
)
doc.add_paragraph()

add_heading("Ehrlicher Vorbehalt (Vorstands-tauglich)", level=2)

add_heading("Engineering-PM-Tiefe (EWW / R&D)", level=3, color=NAVY)
add_para(
    "Für die absolute Spitzentiefe im Stage-Gate-PEP (Bereich Entwicklung Waldwirtschaft, EWW) "
    "ist Planisware technisch weiter. Wenn STIHL in EWW industrialisierte Produktentwicklung "
    "mit > 100 parallelen R&D-Projekten hat und dort maximale Reife will → hybrider Ansatz denkbar:"
)
add_bullets([
    "ServiceNow SPM Pro: Konzern-Portfolio, Demand, Strategy, IT/Operations-PM, Reporting",
    "Planisware Orchestra (nur EWW): Deep-R&D-PEP-Steuerung mit Übergabe an SNOW für Portfolio-Sicht",
])
add_para(
    "Aber: dieser Split verdoppelt Betriebskosten. Nur sinnvoll, wenn App-Engine-Erweiterung "
    "nachweislich nicht ausreicht.",
    color=RED_HL, italic=True
)

doc.add_page_break()

# ============================================================
# MIGRATIONS-EMPFEHLUNG
# ============================================================
add_heading("Migrations-Empfehlung von altem PM-Tool", level=1)
add_para(
    'Statt „Big Bang" auf Planisware — schrittweise Erweiterung ServiceNow SPM Pro:',
    space_after=8
)

phase_rows = [
    ("Phase 0 (heute)", "läuft",       "Demand-Management bereits aktiv"),
    ("Phase 1",         "12 Monate",   "Portfolio + Strategy + Standard-PM (Gantt, Ressourcen, Finanzen) → Ablöse PIT/MSPO für Standard-Projekte"),
    ("Phase 2",         "12 Monate",   "EAP/Agile + tiefes Ressourcen-Management + PEP über App Engine → Ablöse EWW-Spezifika (evtl. mit Planisware-Hedge-Bewertung)"),
    ("Phase 3+",        "rollierend",  "HR-SD, CSM, GRC — SNOW-Plattform breiter nutzen"),
]
build_table(
    headers=["Phase", "Dauer", "Scope"],
    rows=phase_rows,
    col_widths_cm=[3.5, 2.5, 11.0],
    header_fill=ORANGE_HEX,
    center_cols=[1],
)
doc.add_paragraph()

add_callout(
    "Nach Phase 1, wenn sich zeigt, dass App Engine PEP nicht in vernünftiger Tiefe abbildbar "
    "ist — dann Planisware Orchestra als spezialisiertes Engineering-PM in Betracht ziehen. "
    "Bis dahin: Standard-First mit klarer Migration-Klausel.",
    color_hex="FFF6D5",
    title='⚠ Break-Punkt für „Doch Planisware"-Diskussion'
)

# ============================================================
# FAZIT
# ============================================================
add_heading("Fazit", level=1)
add_para(
    "Für Unternehmen mit bestehender ServiceNow-Basis wie STIHL überwiegen die Vorteile "
    "einer SNOW SPM Pro-Migration klar die Funktionsdefizite gegenüber dem spezialisierten "
    "Engineering-PM-Tool Planisware. Die entscheidenden Argumente sind nicht in der "
    "PM-Funktionstiefe zu finden, sondern in Plattform-Integration, Betriebskosten, "
    "Zeit-bis-Nutzen und der Konsistenz der Anwenderführung.",
    space_after=8
)
add_para(
    "Der Gartner APMR MQ 2026 bestätigt Planisware als Leader im reinen APMR-Markt. "
    "ServiceNow ist in diesem MQ per Ausschlussregel nicht enthalten, weil es als "
    "Plattform-Vendor mit APMR-Modul kategorisiert wird — genau die Eigenschaft, die im "
    "STIHL-Kontext den ausschlaggebenden Mehrwert liefert.",
    space_after=8
)
add_para(
    "Planisware bleibt technisch überlegen in den tiefsten R&D-Stage-Gate-Prozessen und "
    "sollte als Hedge-Option in der Detail-Auftragsklärung (Vorgang 3 im Vorstandsprotokoll) "
    "bewertet werden — jedoch bewusst als Ergänzung für den engen EWW-Bereich, nicht als "
    "Konzern-Ziel-Plattform.",
    space_after=12
)

# Footer
p = doc.add_paragraph()
pPr = p._p.get_or_add_pPr()
pBdr = OxmlElement("w:pBdr")
top = OxmlElement("w:top")
top.set(qn("w:val"), "single")
top.set(qn("w:sz"), "4")
top.set(qn("w:space"), "1")
top.set(qn("w:color"), "D1D5DB")
pBdr.append(top)
pPr.append(pBdr)

p = doc.add_paragraph()
r = p.add_run(
    "Quellen: Gartner Magic Quadrant for Adaptive Project Management and Reporting (APMR), "
    "3. August 2026, ID G00842775 (Clegg, Jackson, Ali, Stang, Choi); "
    "Vendor-Aussagen zu ServiceNow SPM Pro basierend auf öffentlich verfügbaren "
    "Produktinformationen und STIHL-interner Marktbeobachtung. "
    "Bindende Bewertung nur nach Detail-Prüfung im Rahmen des RFP-Prozesses "
    "(Vorgang 6, Q4/2026)."
)
r.font.name = "Calibri"
r.font.size = Pt(8.5)
r.font.color.rgb = GREY_TXT
r.italic = True

p = doc.add_paragraph()
r = p.add_run("Autoren: Alex Passaro · Torsten Zahn  ·  Stand: 15.10.2026")
r.font.name = "Calibri"
r.font.size = Pt(8.5)
r.font.color.rgb = GREY_TXT
r.italic = True

doc.save(DST)
print(f"Saved: {DST}")
