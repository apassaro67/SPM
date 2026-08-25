# -*- coding: utf-8 -*-
"""
Erzeugt aus dem sFinx-Export (FY26/FY27) eine aufbereitete Analyse-Arbeitsmappe:
  Dashboard (interaktiv)  |  Cost Center  |  Themen  |  Positionen  |  Rohdaten  |  Legende
Alle Auswertungen sind formelbasiert (SUMIFS/INDEX/MATCH) und rechnen bei
Aenderungen in 'Positionen' automatisch neu.
"""
import re
from collections import OrderedDict

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.chart import BarChart, DoughnutChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.data_source import NumFmt as NumberFormat
from openpyxl.chart.marker import DataPoint
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.shapes import GraphicalProperties

SRC = '/home/user/SPM/SfinxExport_25082026.xlsx'
OUT = '/home/user/SPM/2026_08_25_Budgetanalyse_FY26_FY27.xlsx'

FONT = 'Arial'
C_DARK   = '1F3864'   # Überschriften / dunkelblau
C_HEAD   = '2E5C8A'   # Tabellenkopf
C_FY26   = 'A6B8CC'   # Balken FY26
C_FY27   = '2E5C8A'   # Balken FY27
C_UP     = 'C0392B'   # Kostenanstieg
C_DOWN   = '1E8449'   # Kostenreduktion
C_BAND   = 'EEF3F9'   # Zeilenband
C_INPUT  = 'FFF2CC'   # Eingabe-/Filterzelle
C_GREY   = '595959'

EUR  = '#,##0;-#,##0;"–"'
PCT  = '0.0%;-0.0%;"–"'

HINWEIS_AUSSCHLUSS = ('Hinweis: Cost Center DE2060502 (IPS Management Office NCE) ist auf Anforderung '
                      'vollständig aus dieser Auswertung ausgeschlossen – mit beiden Jahreswerten.')

thin = Side(style='thin', color='BFBFBF')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

# --------------------------------------------------------------------------
# 1) Quelldaten einlesen
# --------------------------------------------------------------------------
wb_src = openpyxl.load_workbook(SRC, data_only=True)
ws_src = wb_src['sfinx']

raw_rows = []
for r in ws_src.iter_rows(min_row=3, values_only=True):
    if r[0] is None:
        continue
    cc, ccname, name, cur, accno, accname, fy26, fy27 = r[:8]
    raw_rows.append(dict(
        cc=str(cc).strip(), ccname=str(ccname).strip(), raw=str(name).strip(),
        cur=str(cur).strip(), accno=str(accno).strip(), accname=str(accname).strip(),
        fy26=float(fy26) if fy26 not in (None, '') else 0.0,
        fy27=float(fy27) if fy27 not in (None, '') else 0.0))


def split_code(raw):
    m = re.match(r'^(R\d+)\s*-\s*(.*)$', raw)
    return (m.group(1), m.group(2).strip()) if m else ('', raw)


def norm_key(n):
    """Normalisiert die Positionsbezeichnung für den Jahresvergleich."""
    s = n.lower()
    s = re.sub(r'\s*\(d-\d+\)\s*$', '', s)          # Demand-Nummer FY27
    s = re.sub(r'^itsm:\s*', '', s)                  # Präfix 'ITSM:' FY27
    s = re.sub(r'^servicenow modul\s+', 'servicenow ', s)
    s = re.sub(r'\s*\(dopplung sac 2025\)', '', s)
    s = re.sub(r'[\s ]+', ' ', s).strip()
    return s.replace('€', 'eur')


for d in raw_rows:
    d['rcode'], d['item'] = split_code(d['raw'])
    d['key'] = norm_key(d['item'])
    d['code'] = re.sub(r'^\d{4}\s*-\s*', '', d['cc']).strip()
    d['ccl'] = re.sub(r'^\d{4}\s*-\s*', '', d['ccname']).strip()

rows26 = [d for d in raw_rows if d['fy26']]
rows27 = [d for d in raw_rows if d['fy27']]
m26 = {d['key']: d for d in rows26}
m27 = {d['key']: d for d in rows27}

# Cost-Center-Bezeichnungen je Jahr
CC_NAME = {}
for d in rows26:
    CC_NAME.setdefault(d['code'], {})['FY26'] = d['ccl']
for d in rows27:
    CC_NAME.setdefault(d['code'], {})['FY27'] = d['ccl']

# Organisatorische Ueberfuehrung FY26 -> FY27 (Vorgabe des Fachbereichs):
#   DE2060202 (ITSM+ServiceNow)       wurde in DE2060204 integriert
#   DE2060203 (Digital Collaboration) wird unveraendert fortgefuehrt
#   DE2060204 (Serv. Desk+Support)    wird unveraendert fortgefuehrt
CC_NACHFOLGER = {'DE2060202': 'DE2060204',
                 'DE2060203': 'DE2060203',
                 'DE2060204': 'DE2060204'}
CC_UEBERFUEHRUNG_ART = {'DE2060202': 'in DE2060204 integriert',
                        'DE2060203': 'unverändert fortgeführt',
                        'DE2060204': 'unverändert fortgeführt'}

# Feste Lage der Ueberfuehrungstabelle auf dem Blatt 'Cost Center'. Die Spalte
# 'Bemerkung' im Blatt 'Positionen' prueft ueber diese Tabelle, ob ein
# Cost-Center-Wechsel der Vorgabe entspricht.
CC_VORGAENGER = {}
for _alt, _neu in CC_NACHFOLGER.items():
    CC_VORGAENGER.setdefault(_neu, []).append(_alt)

MAP_ROW_FIRST = 7
MAP_ROW_LAST = MAP_ROW_FIRST + len(CC_NACHFOLGER) - 1
MAP_26 = f"'Cost Center'!$A${MAP_ROW_FIRST}:$A${MAP_ROW_LAST}"
MAP_27 = f"'Cost Center'!$C${MAP_ROW_FIRST}:$C${MAP_ROW_LAST}"

# Auf Anforderung vollstaendig aus der Auswertung ausgeschlossene Cost Center.
# Betroffene Positionen entfallen mit BEIDEN Jahreswerten – auch dann, wenn der
# FY26-Wert urspruenglich auf einem anderen Cost Center gebucht war.
CC_AUSGESCHLOSSEN = {'DE2060502'}

# --------------------------------------------------------------------------
# 2) Themen-Zuordnung
# --------------------------------------------------------------------------
THEMEN_REGELN = [
    ('ITSM & ServiceNow', [r'servicenow', r'\bitsm\b', r'smo operations', r'sepm', r'\bspm\b', r'\bitom\b']),
    ('KI & Copilot', [r'copilot']),
    ('Microsoft M365 Lizenzen', [r'microsoft m365 ea', r'm365 governance', r'sharepoint storage',
                                 r'project online', r'powerbi', r'entra und intune', r'github']),
    ('Unified Communications & Telefonie', [r'webex', r'teams rooms', r'teams phone', r'telefon', r'\buc\b',
                                            r'unified communication', r'cisco cvi', r'\bfax\b', r'avodaq',
                                            r'heydt', r'videkonferenz', r'videokonferenz', r'phone \+ meeting',
                                            r'alarming']),
    ('E-Mail & Messaging Security', [r'mailjet', r'mail-signature', r'dmarc']),
    ('Collaboration- & Produktivitäts-Tools', [r'think-cell', r'efficient elements', r'mural', r'visio',
                                                r'boardwise', r'sharegate', r'treesize', r'novacapta',
                                                r'm365 beratung', r'mhs consulting']),
    ('Digital Signage', [r'digital signage']),
    ('Client & Endgeräte (L-EDV)', [r'l-edv', r'gwg', r'kleinteile']),
    ('Output Management & Druck', [r'\blrs\b', r'lexmark', r'plot management', r'drucker']),
    ('Service Desk & User Support', [r'bechtle', r'service desk', r'on-hour', r'off hour']),
    ('IT-Management & Governance', [r'sfinx', r'flexera', r'liz[ei]n[zm]management', r'pmo support',
                                    r'demand management', r'trainings', r'it-finanzmanagement']),
]


def thema_of(item):
    s = item.lower()
    for name, pats in THEMEN_REGELN:
        if any(re.search(p, s) for p in pats):
            return name
    return 'Sonstiges'


# --------------------------------------------------------------------------
# 3) Positionsliste (Union FY26/FY27) aufbauen
# --------------------------------------------------------------------------
positions = []
for key in sorted(set(m26) | set(m27)):
    a, b = m26.get(key), m27.get(key)
    src = b or a
    cc26 = a['code'] if a else ''
    cc27 = b['code'] if b else ''
    if cc27:
        cc_view = cc27
    else:
        cc_view = CC_NACHFOLGER.get(cc26, cc26)
    positions.append(dict(
        cc_view=cc_view,
        cc_view_name=CC_NAME.get(cc_view, {}).get('FY27') or CC_NAME.get(cc_view, {}).get('FY26', ''),
        cc26=cc26, cc27=cc27,
        thema=thema_of(src['item']),
        item=src['item'],
        accname=src['accname'], accno=src['accno'],
        r26=a['rcode'] if a else '', r27=b['rcode'] if b else '',
        fy26=a['fy26'] if a else 0.0, fy27=b['fy27'] if b else 0.0))

ausgeschlossen = [d for d in positions
                 if d['cc_view'] in CC_AUSGESCHLOSSEN
                 or d['cc26'] in CC_AUSGESCHLOSSEN or d['cc27'] in CC_AUSGESCHLOSSEN]
positions = [d for d in positions if d not in ausgeschlossen]
positions.sort(key=lambda d: (d['cc_view'], d['thema'], d['item'].lower()))

# Ist-Sicht: jeder Code, der in einem der beiden Jahre gebucht ist.
# Bereinigte Sicht: nur die FY27-Zielcodes, auf die Positionen zusammenlaufen.
CC_CODES_IST = sorted({c for p in positions for c in (p['cc26'], p['cc27']) if c})
CC_CODES = sorted({p['cc_view'] for p in positions})
THEMEN = sorted({p['thema'] for p in positions})
KOSTENARTEN = sorted({p['accname'] for p in positions})

N = len(positions)
FIRST, LAST = 3, 2 + N            # Datenzeilen im Blatt 'Positionen'
P = "Positionen!"
R_CCV = f"{P}$A${FIRST}:$A${LAST}"
R_CC26 = f"{P}$C${FIRST}:$C${LAST}"
R_CC27 = f"{P}$D${FIRST}:$D${LAST}"
R_THEMA = f"{P}$E${FIRST}:$E${LAST}"
R_ITEM = f"{P}$F${FIRST}:$F${LAST}"
R_KA = f"{P}$G${FIRST}:$G${LAST}"
R_F26 = f"{P}$K${FIRST}:$K${LAST}"
R_F27 = f"{P}$L${FIRST}:$L${LAST}"
R_DELTA = f"{P}$M${FIRST}:$M${LAST}"
R_STAT = f"{P}$O${FIRST}:$O${LAST}"

# --------------------------------------------------------------------------
# 4) Arbeitsmappe / Hilfsfunktionen
# --------------------------------------------------------------------------
wb = openpyxl.Workbook()
wb.remove(wb.active)


def base_font(ws, rows=200, cols=30, size=10):
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            ws.cell(row=r, column=c).font = Font(name=FONT, size=size)


def title(ws, text, sub=''):
    ws['A1'] = text
    ws['A1'].font = Font(name=FONT, size=16, bold=True, color=C_DARK)
    if sub:
        ws['A2'] = sub
        ws['A2'].font = Font(name=FONT, size=9, italic=True, color=C_GREY)
    ws.row_dimensions[1].height = 24


def header(ws, row, col, labels, widths=None):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=lab)
        c.font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor=C_HEAD)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = 30
    if widths:
        for i, w in enumerate(widths):
            ws.column_dimensions[get_column_letter(col + i)].width = w


def block_title(ws, row, text):
    ws.cell(row=row, column=1, value=text).font = Font(name=FONT, size=12, bold=True, color=C_DARK)


def note(ws, row, col, text):
    c = ws.cell(row=row, column=col, value=text)
    c.font = Font(name=FONT, size=8, italic=True, color=C_GREY)


def style_table(ws, r1, r2, c1, c2, band=True):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            if band and (r - r1) % 2 == 1:
                cell.fill = PatternFill('solid', fgColor=C_BAND)


def numfmt(axis, code):
    """Zahlenformat an einer Diagrammachse erzwingen (nicht von der Quelle ableiten)."""
    axis.numFmt = NumberFormat(formatCode=code, sourceLinked=False)


def dlbls(fmt=EUR, size=800):
    """Datenbeschriftung: nur der Wert – alle uebrigen Anzeigeflags explizit aus,
    sonst blenden Excel/LibreOffice zusaetzlich Kategorie- und Reihennamen ein."""
    d = DataLabelList()
    d.showVal = True
    d.showCatName = False
    d.showSerName = False
    d.showLegendKey = False
    d.showPercent = False
    d.showBubbleSize = False
    d.numFmt = fmt
    d.dLblPos = None
    return d


def gp(color):
    return GraphicalProperties(solidFill=color, ln=LineProperties(noFill=True))


def style_chart(ch, title_txt, height=8.5, width=17.5):
    ch.title = title_txt
    ch.height, ch.width = height, width
    ch.style = 2
    ch.gapWidth = 60
    if ch.y_axis is not None:
        numfmt(ch.y_axis, EUR)
        ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill='D9D9D9'))
    return ch



# ==========================================================================
# BLATT: Positionen  (Detailtabelle, sortiert nach Cost-Center-Code)
# ==========================================================================
wsP = wb.create_sheet('Positionen')
base_font(wsP, rows=N + 12, cols=17)
wsP['A1'] = 'Positionen FY26 / FY27 – sortiert nach Cost Center Code'
wsP['A1'].font = Font(name=FONT, size=14, bold=True, color=C_DARK)
note(wsP, 1, 6, HINWEIS_AUSSCHLUSS)

COLS = [('Cost Center (Sicht)', 15), ('Cost Center Bezeichnung (FY27)', 34), ('CC FY26', 12), ('CC FY27', 12),
        ('Thema', 30), ('Position', 52), ('Kostenart (Konto)', 34), ('Konto-Nr.', 11),
        ('R-Code FY26', 12), ('R-Code FY27', 12), ('Budget FY26 (EUR)', 15), ('Budget FY27 (EUR)', 15),
        ('Delta (EUR)', 14), ('Delta %', 10), ('Status', 13), ('Bemerkung', 30)]
header(wsP, 2, 1, [c[0] for c in COLS], [c[1] for c in COLS])

for i, p in enumerate(positions):
    r = FIRST + i
    wsP.cell(row=r, column=1, value=p['cc_view'])
    wsP.cell(row=r, column=2, value=p['cc_view_name'])
    wsP.cell(row=r, column=3, value=p['cc26'])
    wsP.cell(row=r, column=4, value=p['cc27'])
    wsP.cell(row=r, column=5, value=p['thema'])
    wsP.cell(row=r, column=6, value=p['item'])
    wsP.cell(row=r, column=7, value=p['accname'])
    wsP.cell(row=r, column=8, value=p['accno'])
    wsP.cell(row=r, column=9, value=p['r26'])
    wsP.cell(row=r, column=10, value=p['r27'])
    wsP.cell(row=r, column=11, value=p['fy26'])
    wsP.cell(row=r, column=12, value=p['fy27'])
    wsP.cell(row=r, column=13, value=f'=L{r}-K{r}')
    wsP.cell(row=r, column=14, value=f'=IF(K{r}=0,"",M{r}/K{r})')
    wsP.cell(row=r, column=15, value=(f'=IF(K{r}=0,"neu",IF(L{r}=0,"entfallen",'
                                      f'IF(M{r}>0,"erhöht",IF(M{r}<0,"reduziert","unverändert"))))'))
    wsP.cell(row=r, column=16, value=(
        f'=IF(C{r}="","erstmals in FY27",'
        f'IF(D{r}="","nur in FY26",'
        f'IF(C{r}=D{r},"",'
        f'IF(D{r}=IFERROR(INDEX({MAP_27},MATCH(C{r},{MAP_26},0)),""),'
        f'"Reorg-Überführung "&C{r}&" -> "&D{r},'
        f'"abweichende Verlagerung "&C{r}&" -> "&D{r}))))'))
    for c in (11, 12, 13):
        wsP.cell(row=r, column=c).number_format = EUR
    wsP.cell(row=r, column=14).number_format = PCT
    for c in (1, 3, 4, 8, 9, 10, 15):
        wsP.cell(row=r, column=c).alignment = Alignment(horizontal='center')
    wsP.cell(row=r, column=6).alignment = Alignment(wrap_text=False)

# Summenzeile
tot = LAST + 1
wsP.cell(row=tot, column=1, value='Summe')
for c in (11, 12, 13):
    L_ = get_column_letter(c)
    wsP.cell(row=tot, column=c, value=f'=SUM({L_}{FIRST}:{L_}{LAST})').number_format = EUR
wsP.cell(row=tot, column=14, value=f'=IF(K{tot}=0,"",M{tot}/K{tot})').number_format = PCT
for c in range(1, 17):
    cell = wsP.cell(row=tot, column=c)
    cell.font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
    cell.fill = PatternFill('solid', fgColor=C_DARK)
    cell.border = BORDER

style_table(wsP, FIRST, LAST, 1, 16)
wsP.auto_filter.ref = f'A2:P{LAST}'
wsP.freeze_panes = 'F3'
wsP.conditional_formatting.add(
    f'M{FIRST}:M{LAST}',
    CellIsRule(operator='greaterThan', formula=['0'], font=Font(name=FONT, size=10, color=C_UP, bold=True)))
wsP.conditional_formatting.add(
    f'M{FIRST}:M{LAST}',
    CellIsRule(operator='lessThan', formula=['0'], font=Font(name=FONT, size=10, color=C_DOWN, bold=True)))
wsP.conditional_formatting.add(
    f'N{FIRST}:N{LAST}',
    CellIsRule(operator='greaterThan', formula=['0'], font=Font(name=FONT, size=10, color=C_UP)))
wsP.conditional_formatting.add(
    f'N{FIRST}:N{LAST}',
    CellIsRule(operator='lessThan', formula=['0'], font=Font(name=FONT, size=10, color=C_DOWN)))
note(wsP, tot + 2, 1, 'Spalten K und L sind Quellwerte aus dem sFinx-Export (Blatt "Rohdaten sfinx"); '
                      'alle übrigen Auswertungsspalten sind Formeln. Zuordnung "Cost Center (Sicht)" siehe Blatt "Legende".')

# ==========================================================================
# BLATT: Berechnung (Hilfsdaten für die interaktiven Grafiken)
# ==========================================================================
wsB = wb.create_sheet('Berechnung')
base_font(wsB, rows=max(120, N + 20), cols=14, size=9)
wsB['A1'] = 'Hilfsdaten für das Dashboard – bitte nicht löschen'
wsB['A1'].font = Font(name=FONT, size=12, bold=True, color=C_DARK)

# Auswahllisten und Filterkriterien (Spalten J/L/N – bewusst ausserhalb der Datenbloecke A:G)
wsB['J3'] = 'Liste Cost Center'
wsB['L3'] = 'Liste Themen'
wsB['N3'] = 'Filterkriterien (SUMIFS)'
for c in ('J3', 'L3', 'N3'):
    wsB[c].font = Font(name=FONT, size=10, bold=True)
for col, w in (('J', 18), ('L', 38), ('N', 16), ('O', 38)):
    wsB.column_dimensions[col].width = w
wsB['J4'] = 'Alle'
for i, code in enumerate(CC_CODES):
    wsB.cell(row=5 + i, column=10, value=code)
CC_LIST_REF = f"Berechnung!$J$4:$J${4 + len(CC_CODES)}"
wsB['L4'] = 'Alle'
for i, t in enumerate(THEMEN):
    wsB.cell(row=5 + i, column=12, value=t)
TH_LIST_REF = f"Berechnung!$L$4:$L${4 + len(THEMEN)}"

SEL_CC, SEL_TH = 'Dashboard!$C$5', 'Dashboard!$C$7'
wsB['N4'] = 'Cost Center'
wsB['O4'] = f'=IF({SEL_CC}="Alle","*",{SEL_CC})'
wsB['N5'] = 'Thema'
wsB['O5'] = f'=IF({SEL_TH}="Alle","*",{SEL_TH})'
note(wsB, 6, 14, '"*" = Platzhalter, trifft jeden Eintrag (Auswahl "Alle").')
CRIT_CC, CRIT_TH = 'Berechnung!$O$4', 'Berechnung!$O$5'

# Block 1: Themen (nur Cost-Center-Filter)
b1 = 9
wsB.cell(row=b1, column=1, value='Block 1 – Themen (Filter: Cost Center)').font = Font(name=FONT, size=10, bold=True)
header(wsB, b1 + 1, 1, ['Thema', 'FY26', 'FY27', 'Delta'], [34, 13, 13, 13])
for i, t in enumerate(THEMEN):
    r = b1 + 2 + i
    wsB.cell(row=r, column=1, value=t)
    wsB.cell(row=r, column=2, value=f'=SUMIFS({R_F26},{R_THEMA},$A{r},{R_CCV},{CRIT_CC})').number_format = EUR
    wsB.cell(row=r, column=3, value=f'=SUMIFS({R_F27},{R_THEMA},$A{r},{R_CCV},{CRIT_CC})').number_format = EUR
    wsB.cell(row=r, column=4, value=f'=C{r}-B{r}').number_format = EUR
B1_F, B1_L = b1 + 2, b1 + 1 + len(THEMEN)

# Block 2: Cost Center (nur Themen-Filter)
b2 = B1_L + 3
wsB.cell(row=b2, column=1, value='Block 2 – Cost Center (Filter: Thema)').font = Font(name=FONT, size=10, bold=True)
header(wsB, b2 + 1, 1, ['Cost Center', 'FY26', 'FY27', 'Delta'], [34, 13, 13, 13])
for i, code in enumerate(CC_CODES):
    r = b2 + 2 + i
    wsB.cell(row=r, column=1, value=code)
    wsB.cell(row=r, column=2, value=f'=SUMIFS({R_F26},{R_CCV},$A{r},{R_THEMA},{CRIT_TH})').number_format = EUR
    wsB.cell(row=r, column=3, value=f'=SUMIFS({R_F27},{R_CCV},$A{r},{R_THEMA},{CRIT_TH})').number_format = EUR
    wsB.cell(row=r, column=4, value=f'=C{r}-B{r}').number_format = EUR
B2_F, B2_L = b2 + 2, b2 + 1 + len(CC_CODES)

# Block 3: Kostenarten (beide Filter)
b3 = B2_L + 3
wsB.cell(row=b3, column=1, value='Block 3 – Kostenarten (Filter: Cost Center + Thema)').font = Font(name=FONT, size=10, bold=True)
header(wsB, b3 + 1, 1, ['Kostenart', 'FY26', 'FY27', 'Delta'], [34, 13, 13, 13])
for i, ka in enumerate(KOSTENARTEN):
    r = b3 + 2 + i
    wsB.cell(row=r, column=1, value=ka)
    wsB.cell(row=r, column=2, value=f'=SUMIFS({R_F26},{R_KA},$A{r},{R_CCV},{CRIT_CC},{R_THEMA},{CRIT_TH})').number_format = EUR
    wsB.cell(row=r, column=3, value=f'=SUMIFS({R_F27},{R_KA},$A{r},{R_CCV},{CRIT_CC},{R_THEMA},{CRIT_TH})').number_format = EUR
    wsB.cell(row=r, column=4, value=f'=C{r}-B{r}').number_format = EUR
B3_F, B3_L = b3 + 2, b3 + 1 + len(KOSTENARTEN)

# Block 4: Rangfolge der Positionen nach |Delta| (beide Filter)
b4 = B3_L + 3
wsB.cell(row=b4, column=1, value='Block 4 – Rangwerte je Position (Filter: Cost Center + Thema)').font = Font(name=FONT, size=10, bold=True)
header(wsB, b4 + 1, 1, ['Position', 'Rangwert |Delta|'], [52, 16])
rank_first = b4 + 2
for i in range(N):
    r = rank_first + i
    pr = FIRST + i
    wsB.cell(row=r, column=1, value=f'={P}F{pr}')
    wsB.cell(row=r, column=2, value=(
        f'=IF(AND(OR({SEL_CC}="Alle",{P}A{pr}={SEL_CC}),OR({SEL_TH}="Alle",{P}E{pr}={SEL_TH})),'
        f'ABS({P}M{pr})+ROW()/1000000,"")'))
rank_last = rank_first + N - 1
RANK_RNG = f'Berechnung!$B${rank_first}:$B${rank_last}'

# Block 5: Top 12 Veränderungen (Ergebnis der Rangfolge)
TOPN = 12
b5 = rank_last + 3
wsB.cell(row=b5, column=1, value='Block 5 – Top-Veränderungen (Diagrammquelle)').font = Font(name=FONT, size=10, bold=True)
header(wsB, b5 + 1, 1, ['Position', 'Delta', 'FY26', 'FY27', 'Cost Center', 'Thema', 'Status'],
       [52, 14, 13, 13, 14, 30, 13])
TOP_F = b5 + 2
for k in range(1, TOPN + 1):
    r = TOP_F + k - 1
    idx = f'MATCH(LARGE({RANK_RNG},{k}),{RANK_RNG},0)'
    wsB.cell(row=r, column=1, value=f'=IFERROR(INDEX({R_ITEM},{idx}),"")')
    wsB.cell(row=r, column=2, value=f'=IFERROR(INDEX({R_DELTA},{idx}),"")').number_format = EUR
    wsB.cell(row=r, column=3, value=f'=IFERROR(INDEX({R_F26},{idx}),"")').number_format = EUR
    wsB.cell(row=r, column=4, value=f'=IFERROR(INDEX({R_F27},{idx}),"")').number_format = EUR
    wsB.cell(row=r, column=5, value=f'=IFERROR(INDEX({R_CCV},{idx}),"")')
    wsB.cell(row=r, column=6, value=f'=IFERROR(INDEX({R_THEMA},{idx}),"")')
    wsB.cell(row=r, column=7, value=f'=IFERROR(INDEX({R_STAT},{idx}),"")')
TOP_L = TOP_F + TOPN - 1

# Aggregate (nur für die Farbgebung der Diagramm-Datenpunkte)
agg_th = OrderedDict((t, [0.0, 0.0]) for t in THEMEN)
for p in positions:
    agg_th[p['thema']][0] += p['fy26']
    agg_th[p['thema']][1] += p['fy27']
agg_cc_view = OrderedDict((c, [0.0, 0.0]) for c in CC_CODES)
agg_cc_ist = OrderedDict((c, [0.0, 0.0]) for c in CC_CODES_IST)
for p in positions:
    agg_cc_view[p['cc_view']][0] += p['fy26']
    agg_cc_view[p['cc_view']][1] += p['fy27']
    if p['cc26']:
        agg_cc_ist[p['cc26']][0] += p['fy26']
    if p['cc27']:
        agg_cc_ist[p['cc27']][1] += p['fy27']


def delta_points(deltas):
    """Datenpunkt-Formatierung: Anstieg rot, Reduktion gruen."""
    return [DataPoint(idx=i, spPr=gp(C_UP if d >= 0 else C_DOWN)) for i, d in enumerate(deltas)]


# ==========================================================================
# BLATT: Cost Center
# ==========================================================================
wsC = wb.create_sheet('Cost Center')
base_font(wsC, rows=80, cols=20)
title(wsC, 'Budgetentwicklung je Cost Center: FY26 → FY27',
      'Quelle: sFinx-Export vom 25.08.2026 · Werte in EUR · alle Zellen sind Formeln auf Blatt "Positionen"')
note(wsC, 3, 1, HINWEIS_AUSSCHLUSS)

HDR_CC = ['Cost Center Code', 'Bezeichnung FY26', 'Bezeichnung FY27', 'Budget FY26', 'Budget FY27',
          'Delta (EUR)', 'Delta %', 'Anteil FY27']
WID_CC = [18, 48, 40, 15, 15, 15, 11, 11]

# --- 0) Organisatorische Überführung FY26 → FY27 -----------------------------
block_title(wsC, 4, 'A) Organisatorische Überführung FY26 → FY27')
note(wsC, 5, 1, 'Vorgabe des Fachbereichs. Diese Zuordnung steuert die bereinigte Sicht (Block C) '
                'und die Spalte "Bemerkung" auf dem Blatt "Positionen".')
hM = 6
header(wsC, hM, 1, ['Cost Center FY26', 'Bezeichnung FY26', 'Cost Center FY27', 'Bezeichnung FY27',
                    'Art der Überführung'], [18, 48, 18, 40, 30])
M_F = hM + 1
for i, code in enumerate(sorted(CC_NACHFOLGER)):
    r = M_F + i
    ziel = CC_NACHFOLGER[code]
    wsC.cell(row=r, column=1, value=code).alignment = Alignment(horizontal='center')
    wsC.cell(row=r, column=2, value=CC_NAME.get(code, {}).get('FY26', '–'))
    wsC.cell(row=r, column=3, value=ziel).alignment = Alignment(horizontal='center')
    wsC.cell(row=r, column=4, value=CC_NAME.get(ziel, {}).get('FY27', '–'))
    wsC.cell(row=r, column=5, value=CC_UEBERFUEHRUNG_ART.get(code, ''))
M_L = M_F + len(CC_NACHFOLGER) - 1
style_table(wsC, M_F, M_L, 1, 5)
assert (M_F, M_L) == (MAP_ROW_FIRST, MAP_ROW_LAST), 'Lage der Überführungstabelle verschoben'

# --- B) Ist-Sicht -----------------------------------------------------------
bA = M_L + 2
block_title(wsC, bA, 'B) Ist-Sicht – Budget wie im Export gebucht')
note(wsC, bA + 1, 1, 'FY26-Werte am FY26-Cost-Center, FY27-Werte am FY27-Cost-Center. Die Überführung aus Block A '
                     'schlägt hier voll durch: abgebende Cost Center weisen FY27 keinen Wert mehr aus.')
hA = bA + 2
header(wsC, hA, 1, HDR_CC, WID_CC)
for i, code in enumerate(CC_CODES_IST):
    r = hA + 1 + i
    wsC.cell(row=r, column=1, value=code).alignment = Alignment(horizontal='center')
    wsC.cell(row=r, column=2, value=CC_NAME.get(code, {}).get('FY26', '– (in FY26 nicht vorhanden)'))
    wsC.cell(row=r, column=3, value=CC_NAME.get(code, {}).get('FY27', '– (in FY27 entfallen)'))
    wsC.cell(row=r, column=4, value=f'=SUMIFS({R_F26},{R_CC26},$A{r})').number_format = EUR
    wsC.cell(row=r, column=5, value=f'=SUMIFS({R_F27},{R_CC27},$A{r})').number_format = EUR
    wsC.cell(row=r, column=6, value=f'=E{r}-D{r}').number_format = EUR
    wsC.cell(row=r, column=7, value=f'=IF(D{r}=0,"",F{r}/D{r})').number_format = PCT
    wsC.cell(row=r, column=8, value=f'=IF($E${hA + 1 + len(CC_CODES_IST)}=0,"",E{r}/$E${hA + 1 + len(CC_CODES_IST)})').number_format = PCT
A_F, A_L = hA + 1, hA + len(CC_CODES_IST)
A_T = A_L + 1
wsC.cell(row=A_T, column=1, value='Gesamt')
for c in (4, 5, 6):
    L_ = get_column_letter(c)
    wsC.cell(row=A_T, column=c, value=f'=SUM({L_}{A_F}:{L_}{A_L})').number_format = EUR
wsC.cell(row=A_T, column=7, value=f'=IF(D{A_T}=0,"",F{A_T}/D{A_T})').number_format = PCT
wsC.cell(row=A_T, column=8, value=f'=IF(E{A_T}=0,"",E{A_T}/E{A_T})').number_format = PCT
style_table(wsC, A_F, A_L, 1, 8)
for c in range(1, 9):
    cell = wsC.cell(row=A_T, column=c)
    cell.font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
    cell.fill = PatternFill('solid', fgColor=C_DARK)
    cell.border = BORDER

# --- B) organisatorisch bereinigte Sicht ------------------------------------
bB = A_T + 3
block_title(wsC, bB, 'C) Vergleichbare Sicht – organisatorisch bereinigt')
note(wsC, bB + 1, 1, 'FY26- und FY27-Werte einer Position werden gemeinsam dem FY27-Cost-Center aus Block A zugeordnet '
                     '(Spalte "Cost Center (Sicht)"). So wird die reine Budgetentwicklung ohne Reorganisationseffekt sichtbar.')
hB = bB + 2
header(wsC, hB, 1, ['Cost Center Code (FY27)', 'Vorgänger FY26', 'Bezeichnung FY27', 'Budget FY26',
                    'Budget FY27', 'Delta (EUR)', 'Delta %', 'Anteil FY27'], WID_CC)
for i, code in enumerate(CC_CODES):
    r = hB + 1 + i
    vorg = ' + '.join(f"{v} ({CC_NAME.get(v, {}).get('FY26', '').split(' - ')[-1]})"
                      for v in sorted(CC_VORGAENGER.get(code, [])))
    wsC.cell(row=r, column=1, value=code).alignment = Alignment(horizontal='center')
    wsC.cell(row=r, column=2, value=vorg or '– (kein FY26-Vorgänger)')
    wsC.cell(row=r, column=3, value=CC_NAME.get(code, {}).get('FY27', '– (in FY27 entfallen)'))
    wsC.cell(row=r, column=4, value=f'=SUMIFS({R_F26},{R_CCV},$A{r})').number_format = EUR
    wsC.cell(row=r, column=5, value=f'=SUMIFS({R_F27},{R_CCV},$A{r})').number_format = EUR
    wsC.cell(row=r, column=6, value=f'=E{r}-D{r}').number_format = EUR
    wsC.cell(row=r, column=7, value=f'=IF(D{r}=0,"",F{r}/D{r})').number_format = PCT
    wsC.cell(row=r, column=8, value=f'=IF($E${hB + 1 + len(CC_CODES)}=0,"",E{r}/$E${hB + 1 + len(CC_CODES)})').number_format = PCT
B_F, B_L = hB + 1, hB + len(CC_CODES)
B_T = B_L + 1
wsC.cell(row=B_T, column=1, value='Gesamt')
for c in (4, 5, 6):
    L_ = get_column_letter(c)
    wsC.cell(row=B_T, column=c, value=f'=SUM({L_}{B_F}:{L_}{B_L})').number_format = EUR
wsC.cell(row=B_T, column=7, value=f'=IF(D{B_T}=0,"",F{B_T}/D{B_T})').number_format = PCT
wsC.cell(row=B_T, column=8, value=f'=IF(E{B_T}=0,"",E{B_T}/E{B_T})').number_format = PCT
style_table(wsC, B_F, B_L, 1, 8)
for c in range(1, 9):
    cell = wsC.cell(row=B_T, column=c)
    cell.font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
    cell.fill = PatternFill('solid', fgColor=C_DARK)
    cell.border = BORDER

for rng in (f'F{A_F}:F{A_L}', f'F{B_F}:F{B_L}', f'G{A_F}:G{A_L}', f'G{B_F}:G{B_L}'):
    wsC.conditional_formatting.add(rng, CellIsRule(operator='greaterThan', formula=['0'],
                                                   font=Font(name=FONT, size=10, bold=True, color=C_UP)))
    wsC.conditional_formatting.add(rng, CellIsRule(operator='lessThan', formula=['0'],
                                                   font=Font(name=FONT, size=10, bold=True, color=C_DOWN)))

# --- Diagramme --------------------------------------------------------------
ch1 = BarChart()
ch1.type, ch1.grouping = 'col', 'clustered'
data = Reference(wsC, min_col=4, max_col=5, min_row=hA, max_row=A_L)
cats = Reference(wsC, min_col=1, max_col=1, min_row=A_F, max_row=A_L)
ch1.add_data(data, titles_from_data=True)
ch1.set_categories(cats)
ch1.series[0].graphicalProperties = gp(C_FY26)
ch1.series[1].graphicalProperties = gp(C_FY27)
style_chart(ch1, 'Ist-Sicht: Budget je Cost Center FY26 vs. FY27')
ch1.dLbls = dlbls()
wsC.add_chart(ch1, 'J4')

ch2 = BarChart()
ch2.type, ch2.grouping = 'col', 'clustered'
data = Reference(wsC, min_col=4, max_col=5, min_row=hB, max_row=B_L)
cats = Reference(wsC, min_col=1, max_col=1, min_row=B_F, max_row=B_L)
ch2.add_data(data, titles_from_data=True)
ch2.set_categories(cats)
ch2.series[0].graphicalProperties = gp(C_FY26)
ch2.series[1].graphicalProperties = gp(C_FY27)
style_chart(ch2, 'Bereinigte Sicht: Budget je Cost Center FY26 vs. FY27')
ch2.dLbls = dlbls()
wsC.add_chart(ch2, 'J22')

ch3 = BarChart()
ch3.type, ch3.grouping = 'col', 'clustered'
data = Reference(wsC, min_col=6, max_col=6, min_row=hB, max_row=B_L)
cats = Reference(wsC, min_col=1, max_col=1, min_row=B_F, max_row=B_L)
ch3.add_data(data, titles_from_data=True)
ch3.set_categories(cats)
ch3.series[0].data_points = delta_points([agg_cc_view[c][1] - agg_cc_view[c][0] for c in CC_CODES])
style_chart(ch3, 'Veränderung je Cost Center (bereinigt): rot = Anstieg, grün = Reduktion')
ch3.dLbls = dlbls()
ch3.legend = None
wsC.add_chart(ch3, 'J40')

# ==========================================================================
# BLATT: Themen
# ==========================================================================
wsT = wb.create_sheet('Themen')
base_font(wsT, rows=140, cols=22)
title(wsT, 'Budgetentwicklung je Thema: FY26 → FY27',
      'Themen sind fachliche Bündel der Einzelpositionen (Zuordnungsregeln siehe Blatt "Legende")')
note(wsT, 3, 1, HINWEIS_AUSSCHLUSS)

block_title(wsT, 4, 'A) Entwicklung je Thema (alle Cost Center)')
HDR_T = ['Thema', 'Budget FY26', 'Budget FY27', 'Delta (EUR)', 'Delta %', 'Anteil FY27',
         'Positionen', 'davon neu', 'davon entfallen']
header(wsT, 5, 1, HDR_T, [38, 15, 15, 15, 11, 11, 11, 11, 13])
T_F = 6
T_L = 5 + len(THEMEN)
T_T = T_L + 1
for i, t in enumerate(THEMEN):
    r = T_F + i
    wsT.cell(row=r, column=1, value=t)
    wsT.cell(row=r, column=2, value=f'=SUMIFS({R_F26},{R_THEMA},$A{r})').number_format = EUR
    wsT.cell(row=r, column=3, value=f'=SUMIFS({R_F27},{R_THEMA},$A{r})').number_format = EUR
    wsT.cell(row=r, column=4, value=f'=C{r}-B{r}').number_format = EUR
    wsT.cell(row=r, column=5, value=f'=IF(B{r}=0,"",D{r}/B{r})').number_format = PCT
    wsT.cell(row=r, column=6, value=f'=IF($C${T_T}=0,"",C{r}/$C${T_T})').number_format = PCT
    wsT.cell(row=r, column=7, value=f'=COUNTIFS({R_THEMA},$A{r})')
    wsT.cell(row=r, column=8, value=f'=COUNTIFS({R_THEMA},$A{r},{R_STAT},"neu")')
    wsT.cell(row=r, column=9, value=f'=COUNTIFS({R_THEMA},$A{r},{R_STAT},"entfallen")')
wsT.cell(row=T_T, column=1, value='Gesamt')
for c in (2, 3, 4, 7, 8, 9):
    L_ = get_column_letter(c)
    cell = wsT.cell(row=T_T, column=c, value=f'=SUM({L_}{T_F}:{L_}{T_L})')
    if c in (2, 3, 4):
        cell.number_format = EUR
wsT.cell(row=T_T, column=5, value=f'=IF(B{T_T}=0,"",D{T_T}/B{T_T})').number_format = PCT
wsT.cell(row=T_T, column=6, value=f'=IF(C{T_T}=0,"",C{T_T}/C{T_T})').number_format = PCT
style_table(wsT, T_F, T_L, 1, 9)
for c in range(1, 10):
    cell = wsT.cell(row=T_T, column=c)
    cell.font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
    cell.fill = PatternFill('solid', fgColor=C_DARK)
    cell.border = BORDER
for rng in (f'D{T_F}:D{T_L}', f'E{T_F}:E{T_L}'):
    wsT.conditional_formatting.add(rng, CellIsRule(operator='greaterThan', formula=['0'],
                                                   font=Font(name=FONT, size=10, bold=True, color=C_UP)))
    wsT.conditional_formatting.add(rng, CellIsRule(operator='lessThan', formula=['0'],
                                                   font=Font(name=FONT, size=10, bold=True, color=C_DOWN)))

chT1 = BarChart()
chT1.type, chT1.grouping = 'bar', 'clustered'
chT1.add_data(Reference(wsT, min_col=2, max_col=3, min_row=5, max_row=T_L), titles_from_data=True)
chT1.set_categories(Reference(wsT, min_col=1, max_col=1, min_row=T_F, max_row=T_L))
chT1.series[0].graphicalProperties = gp(C_FY26)
chT1.series[1].graphicalProperties = gp(C_FY27)
style_chart(chT1, 'Budget je Thema FY26 vs. FY27', height=11, width=19)
wsT.add_chart(chT1, 'K4')

chT2 = BarChart()
chT2.type, chT2.grouping = 'bar', 'clustered'
chT2.add_data(Reference(wsT, min_col=4, max_col=4, min_row=5, max_row=T_L), titles_from_data=True)
chT2.set_categories(Reference(wsT, min_col=1, max_col=1, min_row=T_F, max_row=T_L))
chT2.series[0].data_points = delta_points([agg_th[t][1] - agg_th[t][0] for t in THEMEN])
style_chart(chT2, 'Veränderung je Thema (rot = Anstieg, grün = Reduktion)', height=11, width=19)
chT2.legend = None
chT2.dLbls = dlbls()
wsT.add_chart(chT2, 'K27')

chT3 = DoughnutChart(holeSize=55)
chT3.add_data(Reference(wsT, min_col=3, max_col=3, min_row=5, max_row=T_L), titles_from_data=True)
chT3.set_categories(Reference(wsT, min_col=1, max_col=1, min_row=T_F, max_row=T_L))
chT3.title = 'Budgetstruktur FY27 nach Themen'
chT3.height, chT3.width = 11, 19
wsT.add_chart(chT3, 'K50')

# --- B) Matrix Thema x Cost Center: FY27 ------------------------------------
mB = T_T + 3
block_title(wsT, mB, 'B) Budget FY27 je Thema und Cost Center (bereinigte Sicht)')
header(wsT, mB + 1, 1, ['Thema'] + CC_CODES + ['Gesamt FY27'], [38] + [15] * len(CC_CODES) + [15])
MB_F = mB + 2
for i, t in enumerate(THEMEN):
    r = MB_F + i
    wsT.cell(row=r, column=1, value=t)
    for j, code in enumerate(CC_CODES):
        wsT.cell(row=r, column=2 + j,
                 value=f'=SUMIFS({R_F27},{R_THEMA},$A{r},{R_CCV},{get_column_letter(2 + j)}${mB + 1})'
                 ).number_format = EUR
    last = get_column_letter(1 + len(CC_CODES))
    wsT.cell(row=r, column=2 + len(CC_CODES), value=f'=SUM(B{r}:{last}{r})').number_format = EUR
MB_L = MB_F + len(THEMEN) - 1
MB_T = MB_L + 1
wsT.cell(row=MB_T, column=1, value='Gesamt')
for j in range(len(CC_CODES) + 1):
    L_ = get_column_letter(2 + j)
    cell = wsT.cell(row=MB_T, column=2 + j, value=f'=SUM({L_}{MB_F}:{L_}{MB_L})')
    cell.number_format = EUR
style_table(wsT, MB_F, MB_L, 1, 2 + len(CC_CODES))
for c in range(1, 3 + len(CC_CODES)):
    cell = wsT.cell(row=MB_T, column=c)
    cell.font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
    cell.fill = PatternFill('solid', fgColor=C_DARK)
    cell.border = BORDER

chT4 = BarChart()
chT4.type, chT4.grouping, chT4.overlap = 'col', 'stacked', 100
chT4.add_data(Reference(wsT, min_col=2, max_col=1 + len(CC_CODES), min_row=mB + 1, max_row=MB_L),
              titles_from_data=True)
chT4.set_categories(Reference(wsT, min_col=1, max_col=1, min_row=MB_F, max_row=MB_L))
style_chart(chT4, 'FY27 je Thema, gestapelt nach Cost Center', height=11, width=19)
wsT.add_chart(chT4, 'K73')

# --- C) Matrix Thema x Cost Center: Delta -----------------------------------
mC = MB_T + 3
block_title(wsT, mC, 'C) Veränderung FY26 → FY27 je Thema und Cost Center (bereinigte Sicht)')
header(wsT, mC + 1, 1, ['Thema'] + CC_CODES + ['Delta gesamt'], [38] + [15] * len(CC_CODES) + [15])
MC_F = mC + 2
for i, t in enumerate(THEMEN):
    r = MC_F + i
    wsT.cell(row=r, column=1, value=t)
    for j, code in enumerate(CC_CODES):
        cl = get_column_letter(2 + j)
        wsT.cell(row=r, column=2 + j,
                 value=(f'=SUMIFS({R_F27},{R_THEMA},$A{r},{R_CCV},{cl}${mC + 1})'
                        f'-SUMIFS({R_F26},{R_THEMA},$A{r},{R_CCV},{cl}${mC + 1})')).number_format = EUR
    last = get_column_letter(1 + len(CC_CODES))
    wsT.cell(row=r, column=2 + len(CC_CODES), value=f'=SUM(B{r}:{last}{r})').number_format = EUR
MC_L = MC_F + len(THEMEN) - 1
MC_T = MC_L + 1
wsT.cell(row=MC_T, column=1, value='Gesamt')
for j in range(len(CC_CODES) + 1):
    L_ = get_column_letter(2 + j)
    wsT.cell(row=MC_T, column=2 + j, value=f'=SUM({L_}{MC_F}:{L_}{MC_L})').number_format = EUR
style_table(wsT, MC_F, MC_L, 1, 2 + len(CC_CODES))
for c in range(1, 3 + len(CC_CODES)):
    cell = wsT.cell(row=MC_T, column=c)
    cell.font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
    cell.fill = PatternFill('solid', fgColor=C_DARK)
    cell.border = BORDER
rng = f'B{MC_F}:{get_column_letter(2 + len(CC_CODES))}{MC_L}'
wsT.conditional_formatting.add(rng, CellIsRule(operator='greaterThan', formula=['0'],
                                               font=Font(name=FONT, size=10, bold=True, color=C_UP),
                                               fill=PatternFill('solid', fgColor='FDEDEC')))
wsT.conditional_formatting.add(rng, CellIsRule(operator='lessThan', formula=['0'],
                                               font=Font(name=FONT, size=10, bold=True, color=C_DOWN),
                                               fill=PatternFill('solid', fgColor='EAF7EF')))
note(wsT, MC_T + 2, 1, 'Lesehilfe: rot = Budgetanstieg FY27 gegenüber FY26, grün = Budgetreduktion. '
                       'Die Spaltensumme entspricht der bereinigten Sicht auf Blatt "Cost Center" (Block B).')

# ==========================================================================
# BLATT: Dashboard (interaktiv)
# ==========================================================================
wsD = wb.create_sheet('Dashboard', 0)
base_font(wsD, rows=90, cols=22)
title(wsD, 'Budget-Dashboard FY26 → FY27',
      'Interaktive Auswertung: die beiden Auswahlfelder unten steuern Kennzahlen, Tabellen und Diagramme.')
note(wsD, 3, 1, HINWEIS_AUSSCHLUSS)
for col, w in zip('ABCDEFGHIJKLMNOPQRSTUV',
                  [24, 15, 30, 13, 13, 13, 13, 13, 4, 13, 13, 13, 13, 13, 9, 9, 9, 9, 9, 9, 9, 9]):
    wsD.column_dimensions[col].width = w

wsD['A4'] = 'Filter'
wsD['A4'].font = Font(name=FONT, size=12, bold=True, color=C_DARK)
for row, label, ref, default in ((5, 'Cost Center:', CC_LIST_REF, 'Alle'), (7, 'Thema:', TH_LIST_REF, 'Alle')):
    wsD.cell(row=row, column=1, value=label).font = Font(name=FONT, size=11, bold=True)
    c = wsD.cell(row=row, column=3, value=default)
    c.font = Font(name=FONT, size=11, bold=True, color='0000FF')
    c.fill = PatternFill('solid', fgColor=C_INPUT)
    c.border = BORDER
    c.alignment = Alignment(horizontal='center')
    dv = DataValidation(type='list', formula1=ref, allow_blank=False, showDropDown=False)
    dv.prompt, dv.promptTitle = 'Bitte Eintrag aus der Liste wählen', 'Auswahl'
    wsD.add_data_validation(dv)
    dv.add(c)
note(wsD, 8, 3, 'Gelbe Felder = Auswahl (Dropdown).')

# --- Kennzahlen -------------------------------------------------------------
KPI = [
    ('Budget FY26', f'=SUMIFS({R_F26},{R_CCV},{CRIT_CC},{R_THEMA},{CRIT_TH})', EUR),
    ('Budget FY27', f'=SUMIFS({R_F27},{R_CCV},{CRIT_CC},{R_THEMA},{CRIT_TH})', EUR),
    ('Delta (EUR)', '=C11-B11', EUR),
    ('Delta %', '=IF(B11=0,"",D11/B11)', PCT),
    ('Positionen', f'=COUNTIFS({R_CCV},{CRIT_CC},{R_THEMA},{CRIT_TH})', '#,##0'),
    ('davon neu', f'=COUNTIFS({R_CCV},{CRIT_CC},{R_THEMA},{CRIT_TH},{R_STAT},"neu")', '#,##0'),
    ('davon entfallen', f'=COUNTIFS({R_CCV},{CRIT_CC},{R_THEMA},{CRIT_TH},{R_STAT},"entfallen")', '#,##0'),
]
header(wsD, 10, 2, [k[0] for k in KPI])
for i, (lab, f, fmt) in enumerate(KPI):
    c = wsD.cell(row=11, column=2 + i, value=f)
    c.number_format = fmt
    c.font = Font(name=FONT, size=13, bold=True, color=C_DARK)
    c.alignment = Alignment(horizontal='center', vertical='center')
    c.border = BORDER
    c.fill = PatternFill('solid', fgColor=C_BAND)
wsD.row_dimensions[11].height = 26
wsD.conditional_formatting.add('D11:E11', CellIsRule(operator='greaterThan', formula=['0'],
                                                     font=Font(name=FONT, size=13, bold=True, color=C_UP)))
wsD.conditional_formatting.add('D11:E11', CellIsRule(operator='lessThan', formula=['0'],
                                                     font=Font(name=FONT, size=13, bold=True, color=C_DOWN)))
wsD.cell(row=10, column=1, value='Kennzahlen').font = Font(name=FONT, size=12, bold=True, color=C_DARK)

# --- Diagramme (Quelle: Blatt "Berechnung") ---------------------------------
def dash_bar(min_row_hdr, min_row, max_row, ttl, typ='bar', delta=False, w=17.5, h=9.5):
    ch = BarChart()
    ch.type, ch.grouping = typ, 'clustered'
    if delta:
        ch.add_data(Reference(wsB, min_col=4, max_col=4, min_row=min_row_hdr, max_row=max_row), titles_from_data=True)
    else:
        ch.add_data(Reference(wsB, min_col=2, max_col=3, min_row=min_row_hdr, max_row=max_row), titles_from_data=True)
    ch.set_categories(Reference(wsB, min_col=1, max_col=1, min_row=min_row, max_row=max_row))
    if not delta:
        ch.series[0].graphicalProperties = gp(C_FY26)
        ch.series[1].graphicalProperties = gp(C_FY27)
    style_chart(ch, ttl, height=h, width=w)
    ch.visible_cells_only = False
    return ch


c_th = dash_bar(B1_F - 1, B1_F, B1_L, 'Budget je Thema FY26 vs. FY27  (Filter: Cost Center)')
wsD.add_chart(c_th, 'A14')
c_cc = dash_bar(B2_F - 1, B2_F, B2_L, 'Budget je Cost Center FY26 vs. FY27  (Filter: Thema)', typ='col')
wsD.add_chart(c_cc, 'J14')
c_ka = dash_bar(B3_F - 1, B3_F, B3_L, 'Budget je Kostenart FY26 vs. FY27  (Filter: Cost Center + Thema)')
wsD.add_chart(c_ka, 'A35')

c_top = BarChart()
c_top.type, c_top.grouping = 'bar', 'clustered'
c_top.add_data(Reference(wsB, min_col=2, max_col=2, min_row=TOP_F - 1, max_row=TOP_L), titles_from_data=True)
c_top.set_categories(Reference(wsB, min_col=1, max_col=1, min_row=TOP_F, max_row=TOP_L))
c_top.series[0].graphicalProperties = gp(C_FY27)
style_chart(c_top, f'Top {TOPN} Veränderungen nach Betrag  (Filter: Cost Center + Thema)', height=9.5, width=17.5)
c_top.legend = None
c_top.visible_cells_only = False
wsD.add_chart(c_top, 'J35')

# --- Tabelle Top-Veränderungen --------------------------------------------
tt = 56
wsD.cell(row=tt, column=1, value=f'Top {TOPN} Veränderungen (gemäß Filter)').font = Font(
    name=FONT, size=12, bold=True, color=C_DARK)
header(wsD, tt + 1, 1, ['Position', 'Cost Center', 'Thema', 'FY26', 'FY27', 'Delta (EUR)', 'Status'],
       [52, 14, 30, 14, 14, 14, 13])
for k in range(TOPN):
    r, br = tt + 2 + k, TOP_F + k
    wsD.cell(row=r, column=1, value=f'=IF(Berechnung!A{br}="","",Berechnung!A{br})')
    wsD.cell(row=r, column=2, value=f'=IF(Berechnung!E{br}="","",Berechnung!E{br})').alignment = Alignment(horizontal='center')
    wsD.cell(row=r, column=3, value=f'=IF(Berechnung!F{br}="","",Berechnung!F{br})')
    wsD.cell(row=r, column=4, value=f'=IF(Berechnung!C{br}="","",Berechnung!C{br})').number_format = EUR
    wsD.cell(row=r, column=5, value=f'=IF(Berechnung!D{br}="","",Berechnung!D{br})').number_format = EUR
    wsD.cell(row=r, column=6, value=f'=IF(Berechnung!B{br}="","",Berechnung!B{br})').number_format = EUR
    wsD.cell(row=r, column=7, value=f'=IF(Berechnung!G{br}="","",Berechnung!G{br})').alignment = Alignment(horizontal='center')
style_table(wsD, tt + 2, tt + 1 + TOPN, 1, 7)
wsD.conditional_formatting.add(f'F{tt + 2}:F{tt + 1 + TOPN}',
                               CellIsRule(operator='greaterThan', formula=['0'],
                                          font=Font(name=FONT, size=10, bold=True, color=C_UP)))
wsD.conditional_formatting.add(f'F{tt + 2}:F{tt + 1 + TOPN}',
                               CellIsRule(operator='lessThan', formula=['0'],
                                          font=Font(name=FONT, size=10, bold=True, color=C_DOWN)))
note(wsD, tt + 3 + TOPN, 1,
     'Kreuzfilter-Logik: Das Themen-Diagramm folgt dem Cost-Center-Filter, das Cost-Center-Diagramm dem Themen-Filter; '
     'Kostenarten, Kennzahlen und Top-Liste folgen beiden Filtern. Grundlage ist stets die bereinigte Cost-Center-Sicht.')

# ==========================================================================
# BLATT: Rohdaten sfinx (Original-Export, unverändert)
# ==========================================================================
wsR = wb.create_sheet('Rohdaten sfinx')
base_font(wsR, rows=ws_src.max_row + 2, cols=8, size=9)
for r, row in enumerate(ws_src.iter_rows(min_row=1, max_row=ws_src.max_row, max_col=8, values_only=True), 1):
    for c, v in enumerate(row, 1):
        if v is None:
            continue
        cell = wsR.cell(row=r, column=c, value=v)
        if r <= 2:
            cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
            cell.fill = PatternFill('solid', fgColor=C_HEAD)
        elif c in (7, 8):
            cell.number_format = EUR
for col, w in zip('ABCDEFGH', [20, 40, 60, 12, 14, 40, 15, 15]):
    wsR.column_dimensions[col].width = w
note(wsR, ws_src.max_row + 2, 1,
     'Unveränderter Originalexport. Enthält auch den aus der Auswertung ausgeschlossenen Cost Center DE2060502 – '
     'die Summen dieses Blatts weichen deshalb von den Auswertungsblättern ab (siehe Blatt "Legende", Abschnitt 5).')
wsR.freeze_panes = 'A3'
wsR.auto_filter.ref = f'A2:H{ws_src.max_row}'

# ==========================================================================
# BLATT: Legende
# ==========================================================================
wsL = wb.create_sheet('Legende')
base_font(wsL, rows=90, cols=8)
title(wsL, 'Legende, Aufbau und Annahmen')
wsL.column_dimensions['A'].width = 34
wsL.column_dimensions['B'].width = 110

def sec(row, head):
    c = wsL.cell(row=row, column=1, value=head)
    c.font = Font(name=FONT, size=12, bold=True, color=C_DARK)

def kv(row, k, v):
    a = wsL.cell(row=row, column=1, value=k)
    a.font = Font(name=FONT, size=10, bold=True)
    a.alignment = Alignment(vertical='top')
    b = wsL.cell(row=row, column=2, value=v)
    b.alignment = Alignment(vertical='top', wrap_text=True)

r = 4
sec(r, '1) Quelle')
r += 1
kv(r, 'Datei', 'SfinxExport_25082026.xlsx, Blatt "sfinx" (unverändert übernommen in Blatt "Rohdaten sfinx")')
r += 1
kv(r, 'Umfang', f'{len(rows26)} Positionen FY26 und {len(rows27)} Positionen FY27, Währung EUR, '
                f'zusammengeführt zu {N} Vergleichszeilen im Blatt "Positionen"')
r += 1
kv(r, 'Quellwerte', 'Nur die Spalten "Budget FY26" und "Budget FY27" im Blatt "Positionen" sind Werte aus dem Export. '
                    'Alle anderen Zellen der Mappe sind Formeln und rechnen automatisch neu.')
r += 2
sec(r, '2) Blattaufbau')
for k, v in [
    ('Dashboard', 'Interaktive Sicht mit zwei Dropdowns (Cost Center / Thema), Kennzahlen, vier Diagrammen und der Top-Liste der Veränderungen.'),
    ('Cost Center', 'A) organisatorische Überführung FY26 → FY27, B) Ist-Sicht wie gebucht, '
                    'C) organisatorisch bereinigte Sicht – B und C jeweils mit Diagramm.'),
    ('Themen', 'Entwicklung je Thema, Matrix Thema x Cost Center (FY27 und Veränderung) sowie vier Diagramme.'),
    ('Positionen', 'Detailtabelle aller Positionen, sortiert nach Cost-Center-Code, Thema und Bezeichnung. Mit Autofilter.'),
    ('Berechnung', 'Hilfsblatt: Auswahllisten und formelbasierte Diagrammquellen des Dashboards. Bitte nicht löschen.'),
    ('Rohdaten sfinx', 'Originalexport zur Nachvollziehbarkeit.'),
]:
    r += 1
    kv(r, k, v)

r += 2
sec(r, '3) Zuordnung der Positionen über die Jahre')
r += 1
kv(r, 'Matching', 'FY26- und FY27-Positionen werden über die bereinigte Bezeichnung verknüpft (ohne R-Nummer, ohne '
                  'Demand-Nummer "(D-0xxxxx)", ohne Präfix "ITSM:"). Die R-Nummern beider Jahre bleiben in den Spalten I und J sichtbar.')
r += 1
kv(r, 'Status', '"neu" = nur FY27 vorhanden · "entfallen" = nur FY26 vorhanden · "erhöht"/"reduziert"/"unverändert" = in beiden Jahren vorhanden.')
r += 1
kv(r, 'Bemerkung', 'Weist auf einen Cost-Center-Wechsel einer Position zwischen FY26 und FY27 hin (Verlagerung).')

r += 2
sec(r, '4) Annahmen (bitte prüfen)')
r += 1
kv(r, 'Überführung (Vorgabe)', 'Vom Fachbereich vorgegeben und in Block A des Blatts "Cost Center" dokumentiert: '
                              'DE2060202 (ITSM+ServiceNow) wurde in DE2060204 integriert · DE2060203 (Digital Collaboration bzw. '
                              'Modern Workplace & Experience) wird unverändert fortgeführt · DE2060204 wird unverändert fortgeführt. '
                              'DE2060502 bleibt unberücksichtigt (siehe Abschnitt 5).')
r += 1
kv(r, 'Bereinigte Sicht', 'Für die vergleichbare Sicht wird eine Position mit beiden Jahreswerten dem FY27-Cost-Center zugeordnet '
                          '(z. B. ServiceNow Enterprise Contract: FY26 auf DE2060202, FY27 auf DE2060204 – beide Werte werden DE2060204 zugerechnet). '
                          'Es verbleiben damit die beiden fortgeführten Cost Center DE2060203 und DE2060204.')
r += 1
kv(r, 'FY26-Positionen ohne Nachfolger', 'Positionen, die es nur in FY26 gibt, folgen der Überführung aus Block A: '
                                         'ITSM CSI Fokus, ITSM Development und ServiceNow Modul Release Management (SAP) '
                                         '(alle DE2060202) werden in der bereinigten Sicht DE2060204 zugeordnet, '
                                         'Alarming - SMS und Lexmark RFID Support bleiben bei DE2060204.')
r += 1
kv(r, 'Prüfhinweis', 'Die Spalte "Bemerkung" im Blatt "Positionen" gleicht jeden Cost-Center-Wechsel gegen Block A ab. '
                     '"Reorg-Überführung" = entspricht der Vorgabe; "abweichende Verlagerung" = weicht davon ab und ist zu prüfen. '
                     'Aktuell tritt kein abweichender Fall auf.')
r += 1
kv(r, 'Themen', 'Die Themen sind eine fachliche Bündelung der Positionsbezeichnungen (Regelwerk im Skript). '
                'Ein Treffer entscheidet in dieser Reihenfolge: ITSM & ServiceNow, KI & Copilot, Microsoft M365, Unified Communications, '
                'E-Mail & Messaging, Collaboration-Tools, Digital Signage, Client & Endgeräte, Output Management, Service Desk, IT-Management. '
                'Beispiel: "Microsoft M365 EA - Copilot Studio ..." zählt zu "KI & Copilot", nicht zu den M365-Lizenzen.')
r += 1
kv(r, 'ITSM & ServiceNow', 'ITSM- und ServiceNow-Positionen bilden gemäß Vorgabe ein gemeinsames Thema. Es umfasst alle Positionen '
                           'mit "ServiceNow", "ITSM", "SPM", "SEPM", "ITOM" oder "SMO Operations" im Namen – unabhängig davon, ob sie '
                           'als Lizenz, Plattformbetrieb oder Projektleistung gebucht sind.')

r += 2
sec(r, '5) Ausgeschlossener Cost Center')
r += 1
kv(r, 'DE2060502', 'Der Cost Center DE2060502 (IPS Management Office NCE) ist auf Anforderung vollständig aus der '
                   'Auswertung ausgeschlossen. Betroffen sind 8 Positionen, die auf keinem Blatt der Auswertung '
                   'mehr enthalten sind – weder in den Tabellen noch in den Diagrammen oder Auswahllisten.')
r += 1
kv(r, 'Folge für FY26', 'Die Position "ITSM: SMO Operations" (FY26 300.000 EUR, gebucht auf DE2060202; FY27 300.000 EUR '
                        'auf DE2060502) entfällt mit beiden Jahreswerten. Der FY26-Ausweis von DE2060202 in der Ist-Sicht '
                        'reduziert sich dadurch um 300.000 EUR gegenüber dem Rohexport.')
r += 1
kv(r, 'Entfallenes Thema', 'Das Thema "IT-Management & Governance" bestand ausschließlich aus DE2060502-Positionen und '
                           'erscheint deshalb nicht mehr in der Auswertung.')
r += 1
kv(r, 'Rohdaten', 'Das Blatt "Rohdaten sfinx" bleibt bewusst der unveränderte Originalexport und enthält DE2060502 '
                  'weiterhin – als Nachweis der Quelle. Die Summen dort weichen daher von den Auswertungsblättern ab.')

r += 2
sec(r, '6) Farb- und Formatlogik')
for k, v in [('Rot', 'Budgetanstieg FY27 gegenüber FY26 (Kostensteigerung).'),
             ('Grün', 'Budgetreduktion FY27 gegenüber FY26.'),
             ('Gelbe Zelle', 'Eingabe-/Auswahlfeld (Dropdown) im Dashboard.'),
             ('Zahlenformat', 'Beträge in EUR ohne Nachkommastellen, Nullwerte als "–". Prozentwerte mit einer Nachkommastelle.')]:
    r += 1
    kv(r, k, v)

r += 2
kv(r, 'Erstellt am', '25.08.2026 · erzeugt mit dem Skript build_budget_analysis.py (reproduzierbar)')

# ==========================================================================
# Seiteneinrichtung (Querformat, auf Breite skaliert, Kopfzeilen wiederholen)
# ==========================================================================
# Druckbereiche schliessen die rechts/unten liegenden Diagramme mit ein
for ws, area, rep in ((wsD, f'A1:V{tt + 3 + TOPN}', None),
                      (wsC, f'A1:V{max(B_T, 60)}', None),
                      (wsT, f'A1:W{max(MC_T + 2, 103)}', None),
                      (wsP, f'A1:P{tot}', '2:2'),
                      (wsR, f'A1:H{ws_src.max_row}', '2:2'),
                      (wsL, f'A1:B{r}', None)):
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = area
    if rep:
        ws.print_title_rows = rep
    ws.oddFooter.right.text = '&P / &N'
    ws.oddFooter.left.text = 'Budgetanalyse FY26/FY27 – &A'

wb._sheets = [wb['Dashboard'], wb['Cost Center'], wb['Themen'], wb['Positionen'],
              wb['Rohdaten sfinx'], wb['Legende'], wb['Berechnung']]
for ws in wb.worksheets:
    ws.sheet_view.showGridLines = False
wb['Positionen'].sheet_view.showGridLines = True
wb['Rohdaten sfinx'].sheet_view.showGridLines = True
wb['Berechnung'].sheet_view.showGridLines = True
wb.active = 0
wb.save(OUT)
print('gespeichert:', OUT)
print('Positionen:', N, '| Themen:', len(THEMEN), '| Cost Center:', len(CC_CODES), '| Kostenarten:', len(KOSTENARTEN))
