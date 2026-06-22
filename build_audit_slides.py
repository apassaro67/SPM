# -*- coding: utf-8 -*-
"""
Baut fuer das DQS Audit (Teilaudit 17 IT1, 01.07.2026) den Foliensatz von
Alexander Passaro: 1 Uebersichtsfolie + Inhaltsfolien zu den 3 Themen
Digital Collaboration, IT Service Management und Support.

Inhalte basieren auf der Audit-Vorbereitungsrecherche (Copilot-Zusammenfassung).
Design folgt der STIHL / IPS Vorlage (DQS_Audit_Bereich_IPS_2026_07_01.pptx).
"""
import copy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

SRC = "/root/.claude/uploads/03b2ce4b-0ccb-5007-a1e5-57f9bf8c6baa/2f087795-DQS_Audit_Bereich_IPS_2026_07_01.pptx"
OUT = "/home/user/SPM/DQS_Audit_IPS_2026_07_01_Passaro.pptx"
IMG = "/home/user/SPM/kpi_img/flat"  # aufbereitete KPI-Screenshots aus der E-Mail WG: KPIs

# --- STIHL Palette ---
ORANGE = RGBColor(0xF3, 0x7A, 0x1F)   # Digital Collaboration
BLUE   = RGBColor(0x24, 0x9A, 0xBE)   # IT Service Management
GREEN  = RGBColor(0x74, 0x9F, 0x4A)   # Support
DARK   = RGBColor(0x33, 0x33, 0x33)
GREY   = RGBColor(0x87, 0x87, 0x87)
LIGHT  = RGBColor(0xF2, 0xF2, 0xF2)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
FONT   = "STIHL Contraface Text"
TITLEFONT = "STIHL Contraface Display Title"

INSERT_AT = 18  # 0-basiert -> neue Folien werden ab Position 19 eingefuegt
LAYOUT_TITLE_ONLY = 4  # "Nur Titel"


def add_slide_at(prs, layout_idx, index):
    layout = prs.slide_layouts[layout_idx]
    slide = prs.slides.add_slide(layout)
    # angehaengte Folie an die gewuenschte Position verschieben
    sldIdLst = prs.slides._sldIdLst
    new_id = sldIdLst[-1]
    sldIdLst.remove(new_id)
    sldIdLst.insert(index, new_id)
    return slide


def clear_placeholders(slide, keep_title=True):
    for ph in list(slide.placeholders):
        if keep_title and ph.placeholder_format.idx == 0:
            continue
        ph._element.getparent().remove(ph._element)


def set_title(slide, text):
    title = slide.placeholders[0]
    title.text = text
    for p in title.text_frame.paragraphs:
        for r in p.runs:
            r.font.size = Pt(22)
            r.font.bold = True
            r.font.name = FONT
            r.font.color.rgb = DARK
    return title


def accent_bar(slide, color, top=Inches(0.55), left=Inches(0.0), w=Inches(0.18), h=Inches(0.35)):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, w, h)
    bar.fill.solid(); bar.fill.fore_color.rgb = color
    bar.line.fill.background()
    bar.shadow.inherit = False
    return bar


def hline(slide, top, left=Inches(0.52), right=Inches(12.81), color=GREY, weight=Pt(0.75)):
    ln = slide.shapes.add_connector(2, left, top, right, top)
    ln.line.color.rgb = color
    ln.line.width = weight
    return ln


def textbox(slide, l, t, w, h, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(0.05); tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02); tf.margin_bottom = Inches(0.02)
    return tb, tf


def add_para(tf, text, size=11, bold=False, color=DARK, first=False,
             space_after=3, bullet=False, level=0, italic=False):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.space_after = Pt(space_after)
    p.space_before = Pt(0)
    p.level = level
    r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    r.font.name = FONT; r.font.color.rgb = color
    if bullet:
        _bullet(p, color)
    return p, r


def _bullet(p, color):
    pPr = p._pPr if p._pPr is not None else p.get_or_add_pPr()
    # kleiner orangefarbener Punkt-Bullet
    buFont = pPr.makeelement(qn('a:buFont'), {'typeface': 'Arial'})
    buChar = pPr.makeelement(qn('a:buChar'), {'char': '•'})
    buClr = pPr.makeelement(qn('a:buClr'), {})
    srgb = pPr.makeelement(qn('a:srgbClr'), {'val': '%02X%02X%02X' % (color[0], color[1], color[2])})
    buClr.append(srgb)
    pPr.set('indent', '-137160')
    pPr.set('marL', '137160')
    pPr.append(buClr); pPr.append(buFont); pPr.append(buChar)


def footer(slide):
    tb, tf = textbox(slide, Inches(0.52), Inches(7.18), Inches(9.0), Inches(0.2))
    add_para(tf, "DQS Audit 2026 | Teilaudit 17 IT1 – IT | Digital Collaboration, Service Management & Support | A. Passaro | 01.07.2026",
             size=8, color=GREY, first=True)


# =====================================================================
prs = Presentation(SRC)

# ---------------------------------------------------------------------
# FOLIE 19 - UEBERSICHT / AGENDA
# ---------------------------------------------------------------------
s = add_slide_at(prs, LAYOUT_TITLE_ONLY, INSERT_AT)
clear_placeholders(s)
accent_bar(s, ORANGE)
set_title(s, "Digital Collaboration, Service Management & Support")
tb, tf = textbox(s, Inches(0.52), Inches(0.98), Inches(12.3), Inches(0.4))
add_para(tf, "Wer wir sind – unsere Prozesse – unsere KPI & Nachweise", size=13, bold=True, color=GREY, first=True)
hline(s, Inches(1.5))

cards = [
    (ORANGE, "1", "Digital Collaboration",
     "Bereitstellung & Betrieb der digitalen Arbeitsplatz- und Kollaborationsservices (M365).",
     ["Business Services: Teamwork, Messaging,", "Knowledge & Document Collaboration, Voice",
      "Betriebsregelwerk: M365 Operational Manual", "Support L1–L3 & Service Owner je Service"]),
    (BLUE, "2", "IT Service Management",
     "ITSM als zentrales Steuerungsinstrument der Group IT (IPS) – Services & Value.",
     ["Service Catalogue & Service Owner", "201 Services / 890 Offerings (TBM, 11/25)",
      "Monthly Practice Reviews + CSI Register", "Service-based Working in ServiceNow"]),
    (GREEN, "3", "Support",
     "Service Desk und durchgaengige Incident- & Service-Level-Prozesse.",
     ["First Contact: lokaler Service Desk", "Incident Mgmt (5 Phasen, P1–P4)",
      "Service Level Management (3 Klassen)", "KPI: SLA, FCR, CSAT, Backlog, AHT"]),
]
cw, gap, x0, y0, ch = Inches(3.92), Inches(0.2), Inches(0.52), Inches(1.75), Inches(4.05)
for i, (col, num, title, lead, bullets) in enumerate(cards):
    x = x0 + i * (cw + gap)
    card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y0, cw, ch)
    card.fill.solid(); card.fill.fore_color.rgb = WHITE
    card.line.color.rgb = col; card.line.width = Pt(1.25)
    card.shadow.inherit = False
    # Kopfband
    head = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y0, cw, Inches(0.95))
    head.fill.solid(); head.fill.fore_color.rgb = col
    head.line.fill.background(); head.shadow.inherit = False
    htf = head.text_frame; htf.word_wrap = True
    htf.margin_left = Inches(0.15); htf.margin_top = Inches(0.08)
    p = htf.paragraphs[0]
    r = p.add_run(); r.text = num + "  "; r.font.size = Pt(20); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = TITLEFONT
    r2 = p.add_run(); r2.text = title; r2.font.size = Pt(15); r2.font.bold = True; r2.font.color.rgb = WHITE; r2.font.name = FONT
    # Body
    btb, btf = textbox(s, x + Inches(0.18), y0 + Inches(1.08), cw - Inches(0.36), ch - Inches(1.2))
    add_para(btf, lead, size=11, bold=True, color=DARK, first=True, space_after=8)
    for b in bullets:
        add_para(btf, b, size=10.5, color=DARK, bullet=True, space_after=5)

tb, tf = textbox(s, Inches(0.52), Inches(6.05), Inches(12.3), Inches(0.5))
add_para(tf, "Roter Faden je Thema:  Wer wir sind  →  Prozess & Regelwerk (wo geregelt)  →  KPI & Nachweise (Tracking in ServiceNow / Dashboards)",
         size=11, bold=True, color=ORANGE, first=True)
footer(s)

# ---------------------------------------------------------------------
# FOLIE 20 - WER WIR SIND & SCOPE
# ---------------------------------------------------------------------
s = add_slide_at(prs, LAYOUT_TITLE_ONLY, INSERT_AT + 1)
clear_placeholders(s)
accent_bar(s, ORANGE)
set_title(s, "Wer wir sind – Digital Collaboration Service Management & Support")
hline(s, Inches(1.05))

# Mission Box
mb = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.52), Inches(1.2), Inches(12.3), Inches(0.95))
mb.fill.solid(); mb.fill.fore_color.rgb = LIGHT; mb.line.color.rgb = ORANGE; mb.line.width = Pt(1.0); mb.shadow.inherit = False
mtf = mb.text_frame; mtf.word_wrap = True; mtf.margin_left = Inches(0.2); mtf.margin_top = Inches(0.1)
add_para(mtf, "Mission", size=11, bold=True, color=ORANGE, first=True, space_after=2)
add_para(mtf, "„We provide IT solutions and services to support and enable collaboration, communication and content creation.“",
         size=13, bold=True, italic=True, color=DARK, space_after=0)

# linke Spalte: Organisation + Business Services
tb, tf = textbox(s, Inches(0.52), Inches(2.4), Inches(6.0), Inches(4.4))
add_para(tf, "Organisation", size=13, bold=True, color=ORANGE, first=True, space_after=4)
add_para(tf, "Abteilung D8/ODS – Department „Digital Collaboration Service Management & Support“", size=11, color=DARK, bullet=True, space_after=4)
add_para(tf, "Verantwortung: Bereitstellung & Weiterentwicklung digitaler Arbeitsplatz- und Kollaborationsservices", size=11, color=DARK, bullet=True, space_after=10)
add_para(tf, "Business Services (Portfolio)", size=13, bold=True, color=ORANGE, space_after=4)
for bs in ["Teamwork", "Messaging", "Knowledge Management", "Document Collaboration",
           "Workflow & Automation", "Voice", "Media Technology"]:
    add_para(tf, bs, size=11, color=DARK, bullet=True, space_after=3)
add_para(tf, "flankiert durch Enablement, Community, Support und Governance", size=10.5, italic=True, color=GREY, space_after=0)

# rechte Spalte: Einordnung in IPS
rb = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.82), Inches(2.4), Inches(6.0), Inches(4.4))
rb.fill.solid(); rb.fill.fore_color.rgb = WHITE; rb.line.color.rgb = GREY; rb.line.width = Pt(1.0); rb.shadow.inherit = False
tb, tf = textbox(s, Inches(7.0), Inches(2.55), Inches(5.65), Inches(4.1))
add_para(tf, "Einordnung in der Zielorganisation IPS", size=13, bold=True, color=DARK, first=True, space_after=6)
for h, d in [
    ("Verantworten", "Workplace-Plattformen inkl. Client-Management, Mobile Devices und Collaboration-Tools"),
    ("Steuern", "Service Desk, Request-Management und Incident-Prozesse"),
    ("Betreiben", "zentrale Plattformen wie IDM und ServiceNow"),
    ("Unterstuetzen", "lokale IT-Operations, Onsite-Services und die „letzte Support-Meile“"),
]:
    p = tf.add_paragraph(); p.space_after = Pt(8)
    r = p.add_run(); r.text = h + ":  "; r.font.bold = True; r.font.size = Pt(11.5); r.font.color.rgb = ORANGE; r.font.name = FONT
    r2 = p.add_run(); r2.text = d; r2.font.size = Pt(11); r2.font.color.rgb = DARK; r2.font.name = FONT
footer(s)

# ---------------------------------------------------------------------
# Generische Themen-Inhaltsfolie (Prozesse / Governance / KPI)
# ---------------------------------------------------------------------
def topic_slide(index, color, title, blocks, kpis, regel):
    s = add_slide_at(prs, LAYOUT_TITLE_ONLY, index)
    clear_placeholders(s)
    accent_bar(s, color)
    set_title(s, title)
    hline(s, Inches(1.05), color=color, weight=Pt(1.5))

    # linke Spalte: inhaltliche Bloecke
    tb, tf = textbox(s, Inches(0.52), Inches(1.25), Inches(7.3), Inches(5.6))
    first = True
    for head, items in blocks:
        add_para(tf, head, size=13, bold=True, color=color, first=first, space_after=4)
        first = False
        for it in items:
            add_para(tf, it, size=11, color=DARK, bullet=True, space_after=3)
        # kleiner Abstand
        sp = tf.add_paragraph(); sp.space_after = Pt(2)

    # rechte Spalte: KPI-Karte
    kx, kw = Inches(8.1), Inches(4.72)
    kcard = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, kx, Inches(1.25), kw, Inches(4.35))
    kcard.fill.solid(); kcard.fill.fore_color.rgb = LIGHT
    kcard.line.color.rgb = color; kcard.line.width = Pt(1.25); kcard.shadow.inherit = False
    ktb, ktf = textbox(s, kx + Inches(0.2), Inches(1.4), kw - Inches(0.4), Inches(4.1))
    add_para(ktf, "KPI & Nachweise (Tracking)", size=13, bold=True, color=color, first=True, space_after=6)
    for k in kpis:
        add_para(ktf, k, size=10.8, color=DARK, bullet=True, space_after=5)

    # Regelwerk-Leiste unten
    rbar = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.1), Inches(5.75), kw, Inches(1.05))
    rbar.fill.solid(); rbar.fill.fore_color.rgb = color; rbar.line.fill.background(); rbar.shadow.inherit = False
    rtf = rbar.text_frame; rtf.word_wrap = True; rtf.margin_left = Inches(0.18); rtf.margin_top = Inches(0.08)
    add_para(rtf, "Wo ist es geregelt?", size=11, bold=True, color=WHITE, first=True, space_after=2)
    add_para(rtf, regel, size=10.3, color=WHITE, space_after=0)
    footer(s)
    return s

# kleine Erweiterung von add_para um space_before_zero ignorieren (kompat.)
# (oben bereits ueber try/False geloest)

# ---------------------------------------------------------------------
# FOLIE 21 - DIGITAL COLLABORATION
# ---------------------------------------------------------------------
topic_slide(
    INSERT_AT + 2, ORANGE,
    "Digital Collaboration – Prozesse, Governance & KPI",
    blocks=[
        ("Betriebsregelwerk", [
            "M365 Operational Manual: alle betriebsnotwendigen Infos, Verantwortlichkeiten, Prozesse & Konfigurationen",
            "Business Criticality je Service eingestuft (z. B. Teams / Power Platform = 2, Copilot = 3)",
        ]),
        ("Prozesse", [
            "Monitoring · Ticketing · Alerts & Escalation",
            "Business Continuity / Backup & Restore",
        ]),
        ("Support-Organisation & Ownership", [
            "First Contact lokaler Service Desk → GRP Collaboration L1 / L2 / L3 → Microsoft",
            "Service Manager + Deputy je Service: SharePoint, OneDrive, Exchange, Teams, Planner, To Do",
            "L2/L3-Taetigkeiten dokumentiert: Incident Resolution, Problem Mgmt, Herstellereskalation",
        ]),
    ],
    kpis=[
        "Teams: 76,5 % active users (06/2026)",
        "Copilot Chat: 7.944 active users / 768.343 Prompts (90 T.)",
        "M365 Copilot: 1.431 active · 95,9 % (1.064 lizenziert)",
        "DC Community: 1.017 Mitglieder · 61 % aktiv",
        "DC-Seite: 3.963 Unique Visitors (30 Tage)",
    ],
    regel="M365_Operational_Manual (Word) · Live-Dashboards: M365 Admin Center, DC Community & Site Analytics (Stand 06/2026)",
)

# ---------------------------------------------------------------------
# FOLIE 22 - DIGITAL COLLABORATION: KPI-DASHBOARDS (NACHWEISE)
# ---------------------------------------------------------------------
def fit_image(slide, path, bx, by, bw, bh):
    """platziert ein Bild seitenverhaeltnis-erhaltend zentriert in eine Box."""
    from PIL import Image
    iw, ih = Image.open(path).size
    scale = min(bw / iw, bh / ih)
    w = int(iw * scale); h = int(ih * scale)
    x = bx + (bw - w) // 2
    y = by + (bh - h) // 2
    pic = slide.shapes.add_picture(path, x, y, w, h)
    pic.line.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
    pic.line.width = Pt(0.5)
    return pic

s = add_slide_at(prs, LAYOUT_TITLE_ONLY, INSERT_AT + 3)
clear_placeholders(s)
accent_bar(s, ORANGE)
set_title(s, "Digital Collaboration – KPI-Dashboards & Nachweise (Stand 06/2026)")
hline(s, Inches(1.05), color=ORANGE, weight=Pt(1.5))

tiles = [
    ("Teams – Active Users je Monat (in %)", "Outlook-4bwwffr1.png"),
    ("Copilot Chat – Active Users (90 Tage)", "Outlook-offjn1gv.png"),
    ("M365 Copilot – Usage Overview", "Outlook-vtm5zhsf.png"),
    ("DC Community – Mitglieder & Aktivitaet", "Outlook-qzua1fhd.png"),
    ("DC-Seite – Unique Visitors", "Outlook-5zzsa5tj.png"),
    ("Kommunikationskanaele – E-Mail vs. Teams", "Outlook-vbqlsl05.png"),
]
import os
col_w, gap_x = Emu(Inches(3.95)), Emu(Inches(0.22))
row_h, gap_y = Emu(Inches(2.55)), Emu(Inches(0.2))
x0, y0 = Emu(Inches(0.52)), Emu(Inches(1.25))
for i, (cap, fn) in enumerate(tiles):
    r, c = divmod(i, 3)
    x = x0 + c * (col_w + gap_x)
    y = y0 + r * (row_h + gap_y)
    # Karten-Hintergrund
    card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, col_w, row_h)
    card.fill.solid(); card.fill.fore_color.rgb = WHITE
    card.line.color.rgb = ORANGE; card.line.width = Pt(1.0); card.shadow.inherit = False
    # Caption
    ctb, ctf = textbox(s, Emu(x + Emu(Inches(0.1))), Emu(y + Emu(Inches(0.06))),
                       Emu(col_w - Emu(Inches(0.2))), Inches(0.3))
    add_para(ctf, cap, size=10, bold=True, color=ORANGE, first=True, space_after=0)
    # Bild
    fit_image(s, os.path.join(IMG, fn),
              x + Emu(Inches(0.12)), y + Emu(Inches(0.42)),
              col_w - Emu(Inches(0.24)), row_h - Emu(Inches(0.54)))

qtb, qtf = textbox(s, Inches(0.52), Inches(6.78), Inches(12.3), Inches(0.32))
add_para(qtf, "Quellen: Microsoft 365 Admin Center · DC Community Analytics · SharePoint Site Analytics · Copilot Usage Reports  |  weitere Nachweise: Teams-Tabelle, Copilot-Lizenzgruppe, M365 Active-Users-Dashboard",
         size=8, color=GREY, first=True)
footer(s)

# ---------------------------------------------------------------------
# FOLIE 23 - IT SERVICE MANAGEMENT
# ---------------------------------------------------------------------
topic_slide(
    INSERT_AT + 4, BLUE,
    "IT Service Management – Prozesse, Governance & KPI",
    blocks=[
        ("Steuerungsansatz", [
            "ITSM als zentrales Steuerungsinstrument der Group IT (IPS) – Shift zu Services & Value",
            "Strategische, community-getriebene Practice-Initiative",
        ]),
        ("Service Catalogue & Service Owner", [
            "201 Services / 890 Service Offerings nach TBM-Struktur (Stand 11/25)",
            "Single Source of consistent information zu allen operativen IT-Services",
            "Service Owner sichern Performance & Lifecycle-Governance",
        ]),
        ("Governance & Tooling", [
            "Monthly Practice Reviews: KPI Dashboard Review, Root Cause / Corrective Actions, CSI Register",
            "Service-based Working dokumentiert in ServiceNow",
            "Historie: 9 ITSM Practices (2023) → globaler Service-Katalog & SBW GoLive (2025)",
        ]),
    ],
    kpis=[
        "SCM 01 – Number of active services",
        "SCM 02 – Ø Service Offerings je Service",
        "SCM 03 – Number of used offerings in process",
        "SCM 04 – New/updated services per month",
        "Quelle: ServiceNow · Reporting: monatlich",
    ],
    regel="Practice „Service Catalogue Management“ (SharePoint) · Group ITSM Briefing · ServiceNow",
)

# ---------------------------------------------------------------------
# FOLIE 24 - SUPPORT
# ---------------------------------------------------------------------
topic_slide(
    INSERT_AT + 5, GREEN,
    "Support – Prozesse, Governance & KPI",
    blocks=[
        ("Incident Management", [
            "Klassifikation ueber Impacted Service + Service Offering (zweistufig, seit 07/2025)",
            "Priorisierung aus Impact × Urgency – Priority 1–4 fachlich definiert",
            "Workflow (5 Phasen): Registration → 1st Analysis → Routing → Tracking/Escalation → Review & Closing",
        ]),
        ("Service Level Management", [
            "3 Service Level Classes – designed & approved via Demand-to-Portfolio",
            "Features: response time, resolution time, qualified support time",
            "KPI dort, wo automatisch in ServiceNow generierbar",
        ]),
        ("Support-Steuerung", [
            "First Contact lokaler Service Desk → Eskalation L1 → L2 → L3 / Vendor",
            "Management- / Operations- / Vendor-KPI-Felder definiert",
        ]),
    ],
    kpis=[
        "INC 01 – Number of incidents",
        "INC 02 – Initial resolution rate (≥ 65 %)",
        "INC 03 – Incident resolution time (≥ 80 %)",
        "SLM – Response Time & Resolution Time",
        "Support-Felder: SLA, FCR, CSAT, AHT, Backlog, Escalation rate",
    ],
    regel="Incident-Management-Details & Service-Level-Management-Details (SharePoint) · ServiceNow",
)

# ---------------------------------------------------------------------
# FOLIE 25 - KPI-COCKPIT & AKTUELLE PROJEKTE / ABSCHLUSS
# ---------------------------------------------------------------------
s = add_slide_at(prs, LAYOUT_TITLE_ONLY, INSERT_AT + 6)
clear_placeholders(s)
accent_bar(s, ORANGE)
set_title(s, "KPI-Cockpit & aktuelle Projekte – auditfeste Nachweise")
hline(s, Inches(1.05))

# linke Seite: KPI-Tabelle
tb, tf = textbox(s, Inches(0.52), Inches(1.2), Inches(7.4), Inches(0.35))
add_para(tf, "KPI-Uebersicht – definiert, mit Datenquelle & Frequenz", size=13, bold=True, color=ORANGE, first=True)

rows = [
    ("KPI", "Schwelle", "Quelle", "Frequenz"),
    ("INC 02 Initial resolution rate", "≥ 65 %", "ServiceNow", "monatl."),
    ("INC 03 Incident resolution time", "≥ 80 %", "ServiceNow", "monatl."),
    ("INC 01 Number of incidents", "–", "ServiceNow", "monatl."),
    ("SLM Response / Resolution Time", "SLA-def.", "ServiceNow", "monatl."),
    ("SCM 01–04 Service Catalogue", "–", "ServiceNow", "monatl."),
    ("DC Teams active users", "76,5 %", "M365 Admin", "06/2026"),
    ("DC M365 Copilot active rate", "95,9 %", "Copilot Rep.", "90 Tage"),
    ("DC Community Mitglieder", "1.017 / 61%", "DC Analytics", "06/2026"),
    ("Support: SLA, FCR, CSAT, Backlog", "Felder def.", "Dashboard", "lfd."),
]
nrows, ncols = len(rows), 4
gt = s.shapes.add_table(nrows, ncols, Inches(0.52), Inches(1.6), Inches(7.4), Inches(3.9)).table
gt.columns[0].width = Inches(3.7); gt.columns[1].width = Inches(1.3)
gt.columns[2].width = Inches(1.35); gt.columns[3].width = Inches(1.05)
for ci in range(ncols):
    for ri in range(nrows):
        cell = gt.cell(ri, ci)
        cell.margin_left = Inches(0.06); cell.margin_right = Inches(0.04)
        cell.margin_top = Inches(0.02); cell.margin_bottom = Inches(0.02)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = cell.text_frame.paragraphs[0]
        run = para.add_run(); run.text = rows[ri][ci]
        run.font.name = FONT; run.font.size = Pt(10.5 if ri == 0 else 10)
        if ri == 0:
            run.font.bold = True; run.font.color.rgb = WHITE
            cell.fill.solid(); cell.fill.fore_color.rgb = ORANGE
        else:
            run.font.color.rgb = DARK
            cell.fill.solid(); cell.fill.fore_color.rgb = WHITE if ri % 2 else LIGHT

# rechte Seite: aktuelle Projekte
rb = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.15), Inches(1.2), Inches(4.67), Inches(4.3))
rb.fill.solid(); rb.fill.fore_color.rgb = WHITE; rb.line.color.rgb = BLUE; rb.line.width = Pt(1.25); rb.shadow.inherit = False
tb, tf = textbox(s, Inches(8.33), Inches(1.32), Inches(4.32), Inches(4.05))
add_para(tf, "Aktuelle Projekte 2026", size=13, bold=True, color=BLUE, first=True, space_after=6)
for proj in [
    "Teams Phone Rollout",
    "M365 Copilot – Pilot & Rollout",
    "Copilot Studio Agents (ONE STIHL Hypercare Buddy, IT Support Agent, Ask HR)",
    "Copilot Governance + Training",
    "DLP Project · Cloud Infrastructure Migration",
    "Supplier Collaboration Portal · Places",
]:
    add_para(tf, proj, size=11, color=DARK, bullet=True, space_after=6)

# Abschlussstatement
fb = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.52), Inches(5.75), Inches(12.3), Inches(1.0))
fb.fill.solid(); fb.fill.fore_color.rgb = ORANGE; fb.line.fill.background(); fb.shadow.inherit = False
ftf = fb.text_frame; ftf.word_wrap = True; ftf.vertical_anchor = MSO_ANCHOR.MIDDLE
ftf.margin_left = Inches(0.2)
add_para(ftf, "Fazit: Prozesse sind dokumentiert, Rollen & Service Owner geklaert, KPI definiert – und das Tracking ist in ServiceNow / den Dashboards vorhanden.",
         size=13, bold=True, color=WHITE, first=True, space_after=0)
footer(s)

prs.save(OUT)
print("Gespeichert:", OUT, "| Folien gesamt:", len(prs.slides.__iter__.__self__._sldIdLst))
