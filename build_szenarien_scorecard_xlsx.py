"""Balanced Scorecard: 4 Szenarien SNOW SPM vs SNOW+Planisware Kombinationen."""
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
from openpyxl.worksheet.table import Table, TableStyleInfo

DST = "/tmp/2026_10_20_SNOW_Planisware_Scorecard.xlsx"

# STIHL colors (hex without #)
ORANGE = "F07F12"
DARK   = "1F2937"
NAVY   = "2C3E50"
GREEN  = "2E7D32"
GREY_L = "F2F4F7"
GREY_H = "E5E8EE"
WHITE  = "FFFFFF"
RED    = "B91C1C"
YELLOW = "E0A80B"
BLUE_SN= "285AB8"
GREEN_PW="0E8A63"

def font(color="000000", size=11, bold=False, italic=False):
    return Font(name="Calibri", color=color, size=size, bold=bold, italic=italic)

def fill(color): return PatternFill("solid", fgColor=color)
def align(h="left", v="center", wrap=True): return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

thin = Side(border_style="thin", color="D1D5DB")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()

# ============================================================
# SHEET 1 — Anleitung
# ============================================================
ws0 = wb.active
ws0.title = "Anleitung"
ws0.sheet_view.showGridLines = False
ws0["B2"] = "Bewertungs-Scorecard  ·  SNOW SPM vs. Planisware-Kombinationen"
ws0["B2"].font = font(ORANGE, 20, True)
ws0["B3"] = "Team-Evaluierung der vier PPM-Plattform-Szenarien"
ws0["B3"].font = font(DARK, 13, italic=True)

ws0["B5"] = "Vorgehen"
ws0["B5"].font = font(NAVY, 14, True)
steps = [
    "1.  Sheet »Scorecard« öffnen. Jeder Teilnehmer trägt die Scores 1–5 je Kriterium und Szenario ein.",
    "2.  Gewichte je Dimension links unten anpassen (Summe = 100%). Kriterien-Gewichte innerhalb der Dimension in Spalte D.",
    "3.  Die gewichteten Werte, Dimensions-Scores und der Gesamt-Score werden automatisch berechnet.",
    "4.  Sheet »Konsens« für die Team-Konsens-Bewertung nutzen; individuelle Scores können in Kopien des Scorecard-Sheets gehalten werden.",
    "5.  Sheet »Ergebnis« zeigt die Rangfolge und eine Sensitivitäts-Übersicht (Gewichte ±10%).",
]
for i, t in enumerate(steps):
    ws0.cell(row=6+i, column=2, value=t).font = font(DARK, 11)

ws0["B12"] = "Score-Skala"
ws0["B12"].font = font(NAVY, 14, True)
sk = [
    ("1", "sehr schlecht", "Anforderung deutlich verfehlt, kritische Lücke"),
    ("2", "schlecht",      "Substanzielle Lücke, nur mit erheblichem Aufwand behebbar"),
    ("3", "ok",            "Anforderung im Grundsatz erfüllt, mit vertretbarem Aufwand"),
    ("4", "gut",           "Anforderung sehr gut erfüllt, marginale Lücken"),
    ("5", "sehr gut",      "Best-in-Class, vollständig erfüllt"),
]
for i, (s, l, d) in enumerate(sk):
    r = 13+i
    ws0.cell(row=r, column=2, value=s).font = font(ORANGE, 14, True)
    ws0.cell(row=r, column=2).alignment = align("center")
    ws0.cell(row=r, column=3, value=l).font = font(DARK, 11, True)
    ws0.cell(row=r, column=4, value=d).font = font(DARK, 11)

ws0["B20"] = "Szenarien"
ws0["B20"].font = font(NAVY, 14, True)
sz = [
    ("SZ 1",  "SNOW homogen",                       BLUE_SN,  "ServiceNow SPM Pro für Strategy, Portfolio, Demand, Roadmap UND komplettes PM aller Fachbereiche"),
    ("SZ 2a", "SNOW + PW nur PEP-Subset",           ORANGE,   "SNOW als Standard; Planisware nur für definierten PEP-Subset (Filterkriterien: Budget, Größe, Komplexität)"),
    ("SZ 2",  "SNOW + PW Entwicklung/Produktion",   NAVY,     "SNOW für Strategy/Portfolio/Demand/Rest; Planisware für gesamtes PM in Entwicklung + Produktion"),
    ("SZ 3",  "SNOW + PW weltweit alle FB",         GREEN_PW, "SNOW für Strategy/Portfolio/Demand; Planisware für PM in allen Fachbereichen weltweit"),
]
for i, (tag, name, col, desc) in enumerate(sz):
    r = 21+i
    ws0.cell(row=r, column=2, value=tag).font = font(WHITE, 12, True)
    ws0.cell(row=r, column=2).fill = fill(col)
    ws0.cell(row=r, column=2).alignment = align("center")
    ws0.cell(row=r, column=3, value=name).font = font(DARK, 11, True)
    ws0.cell(row=r, column=4, value=desc).font = font(DARK, 11)

# Column widths
ws0.column_dimensions["A"].width = 2
ws0.column_dimensions["B"].width = 12
ws0.column_dimensions["C"].width = 32
ws0.column_dimensions["D"].width = 80
for r in range(2, 30):
    ws0.row_dimensions[r].height = 22

# ============================================================
# SHEET 2 — Scorecard (Kriterien × Szenarien)
# ============================================================
ws = wb.create_sheet("Scorecard")
ws.sheet_view.showGridLines = False

# Kriteriendefinition: (Dimension, Gewicht Dim %, Kriterium, Gewicht innerhalb Dim %, [initial Scores SZ1, SZ2a, SZ2, SZ3])
# Gewichte je Dimension: müssen zu 100% summieren
# Gewichte je Kriterium innerhalb Dim: müssen zu 100% summieren
DIMS = [
    ("Strategisch",      20, [
        ("Fit zur Konzern-IT-Strategie (Plattform statt Silo)",         25, [5, 4, 3, 2]),
        ("Konsistenz Vendor-Landschaft (Reduktion Klumpenrisiko)",      15, [2, 3, 4, 4]),
        ("Zukunftssicherheit / KI-Roadmap konsolidiert",                20, [5, 4, 3, 3]),
        ("Business-IT-Alignment über Prozesskette Demand→Ergebnis",     25, [5, 4, 3, 2]),
        ("Beschleunigung Digitalisierungs-Programm",                    15, [5, 4, 3, 2]),
    ]),
    ("Funktional",       20, [
        ("Klassische Projektplanung (Gantt, WBS, Meilensteine)",        15, [4, 4, 5, 5]),
        ("Ressourcen-Management (Skills, Capacity, Optimierung)",       15, [4, 4, 5, 5]),
        ("Finanz-Management (Budget, Forecast, SAP-Obligo)",            15, [4, 4, 4, 4]),
        ("PEP / Stage-Gate (App Engine bei SNOW)",                      20, [4, 4, 5, 5]),
        ("Massenänderungen in Projekten",                               10, [3, 4, 5, 5]),
        ("Reporting & Live-Dashboards",                                 10, [5, 4, 4, 4]),
        ("KI / GenAI im PM (Now Assist vs. Oscar)",                     15, [5, 4, 4, 4]),
    ]),
    ("Technisch",        15, [
        ("Native Integration Demand → PM (ohne Schnittstelle)",         30, [5, 4, 3, 3]),
        ("Anbindung an ITSM / HR-SD / GRC (Plattform-Nutzen)",          20, [5, 4, 3, 3]),
        ("Extensibility / Low-Code (App Engine)",                       20, [5, 4, 4, 4]),
        ("Schnittstellen-Aufwand SNOW ↔ PW",                            15, [5, 3, 2, 2]),
        ("Datenhaltung, Master-Data, Konsistenz",                       15, [5, 4, 3, 3]),
    ]),
    ("Wirtschaftlich",   20, [
        ("Lizenzkosten 5J (Delta zu Basis)",                            25, [5, 4, 2, 1]),
        ("Implementierungskosten (einmalig)",                           20, [5, 4, 3, 2]),
        ("Betriebskosten p.a. (Team, Support, Wartung)",                20, [5, 4, 3, 2]),
        ("TCO 5J (gesamt, gewichtet)",                                  25, [5, 4, 2, 1]),
        ("Cost of Change / Migrations-Aufwand",                         10, [4, 4, 3, 2]),
    ]),
    ("Organisatorisch",  15, [
        ("Nutzbarkeit bestehendes SNOW-Team (Skills im Haus)",          25, [5, 4, 3, 2]),
        ("Change-Aufwand für Anwender (neuer Vendor, neue UI)",         25, [5, 4, 3, 2]),
        ("Trainings-Aufwand (Anzahl Anwender × neue Tools)",            15, [5, 4, 3, 2]),
        ("Vendor-Management-Komplexität",                               15, [5, 4, 3, 2]),
        ("Governance-Aufwand (Regeln, Übergaben)",                      20, [5, 3, 3, 2]),
    ]),
    ("Risiko/Zukunft",   10, [
        ("Klumpenrisiko ein Vendor (SNOW)",                             25, [2, 3, 4, 4]),
        ("Klumpenrisiko zweiter Vendor (PW-Lock-in)",                   20, [5, 4, 3, 2]),
        ("Migrations-/Rollout-Risiko",                                  20, [4, 4, 3, 2]),
        ("Skalierung bei künftigen Anforderungen",                      15, [4, 4, 4, 4]),
        ("Reversibilität der Entscheidung",                             20, [4, 4, 3, 2]),
    ]),
]

# Header
ws["A1"] = "Balanced Scorecard  —  4 Szenarien"
ws["A1"].font = font(ORANGE, 18, True)
ws.merge_cells("A1:J1")

ws["A2"] = "Dim.-Gewichte sind in Zeile ganz rechts editierbar. Kriterien-Gewichte innerhalb der Dimension in Spalte D."
ws["A2"].font = font(GREY_H if False else "4B5563", 10, italic=True)
ws.merge_cells("A2:J2")

# Column widths
widths = [16, 46, 10, 10, 11, 12, 12, 12, 12, 14]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w

# Header row (row 4)
hdr = ["Dimension", "Kriterium", "Dim-Gewicht %", "Krit-Gewicht %",
       "SZ 1 SNOW", "SZ 2a SNOW+PW\n(PEP-Subset)", "SZ 2 SNOW+PW\n(E&P)",
       "SZ 3 SNOW+PW\n(weltweit)", "Notiz / Begründung", "Gewichteter\nBeitrag SZ1"]
r_hdr = 4
for ci, h in enumerate(hdr, start=1):
    c = ws.cell(row=r_hdr, column=ci, value=h)
    c.font = font(WHITE, 11, True)
    c.fill = fill(NAVY)
    c.alignment = align("center", "center", True)
    c.border = border
ws.row_dimensions[r_hdr].height = 34

# Data rows — grouped by dimension
current_row = r_hdr + 1
DIM_HEAD_ROWS = []  # for merges

for dim_name, dim_wt, criteria in DIMS:
    dim_start = current_row
    for ki, (k_name, k_wt, initials) in enumerate(criteria):
        # A: Dimension (only first row)
        cA = ws.cell(row=current_row, column=1, value=dim_name if ki == 0 else "")
        cA.font = font(NAVY, 12, True)
        cA.alignment = align("center", "center", True)
        cA.fill = fill(GREY_H)
        cA.border = border
        # B: Kriterium
        cB = ws.cell(row=current_row, column=2, value=k_name)
        cB.font = font(DARK, 10.5)
        cB.alignment = align("left", "center", True)
        cB.fill = fill(GREY_L if ki % 2 == 0 else WHITE)
        cB.border = border
        # C: Dim gewicht (only first row)
        if ki == 0:
            cC = ws.cell(row=current_row, column=3, value=dim_wt/100)
            cC.number_format = "0.0%"
            cC.font = font(NAVY, 11, True)
            cC.alignment = align("center", "center")
            cC.fill = fill(GREY_H)
        else:
            cC = ws.cell(row=current_row, column=3, value="")
            cC.fill = fill(GREY_H)
        cC.border = border
        # D: Krit-Gewicht
        cD = ws.cell(row=current_row, column=4, value=k_wt/100)
        cD.number_format = "0.0%"
        cD.font = font(DARK, 11)
        cD.alignment = align("center", "center")
        cD.border = border
        # E-H: Scores 1..5 for 4 scenarios
        for si, sc in enumerate(initials):
            cell = ws.cell(row=current_row, column=5+si, value=sc)
            cell.font = font(DARK, 11, True)
            cell.alignment = align("center", "center")
            cell.border = border
        # I: Notiz
        ws.cell(row=current_row, column=9, value="").border = border
        # J: Gewichteter Beitrag SZ1 (Beispiel-Anzeige) — Krit × Score1
        j = ws.cell(row=current_row, column=10,
                    value=f"=D{current_row}*E{current_row}")
        j.number_format = "0.00"
        j.font = font(GREY_H if False else "6B7280", 10)
        j.alignment = align("center", "center")
        j.border = border
        current_row += 1
    # Merge column A across dimension block
    if len(criteria) > 1:
        ws.merge_cells(start_row=dim_start, start_column=1, end_row=current_row-1, end_column=1)
        ws.merge_cells(start_row=dim_start, start_column=3, end_row=current_row-1, end_column=3)
    DIM_HEAD_ROWS.append((dim_name, dim_start, current_row-1))
    current_row += 0  # no blank row

# Conditional formatting for scores (E-H)
score_range = f"E{r_hdr+1}:H{current_row-1}"
rule_bad  = CellIsRule(operator="lessThan",      formula=["3"], fill=fill("F8D7DA"), font=font(RED, 11, True))
rule_mid  = CellIsRule(operator="between",       formula=["3","3"], fill=fill("FFF3C4"), font=font("8A6300", 11, True))
rule_good = CellIsRule(operator="greaterThan",   formula=["3"], fill=fill("D4EDDA"), font=font(GREEN, 11, True))
ws.conditional_formatting.add(score_range, rule_bad)
ws.conditional_formatting.add(score_range, rule_mid)
ws.conditional_formatting.add(score_range, rule_good)

# ============================================================
# Summary block (below scorecard): Dimensionen-Scores + Total
# ============================================================
sum_start = current_row + 2
ws.cell(row=sum_start, column=1, value="Zusammenfassung je Dimension").font = font(NAVY, 14, True)
ws.merge_cells(start_row=sum_start, start_column=1, end_row=sum_start, end_column=10)

hdr2_row = sum_start + 1
hdr2 = ["Dimension", "Dim-Gewicht %", "SZ 1", "SZ 2a", "SZ 2", "SZ 3"]
for ci, h in enumerate(hdr2, start=1):
    c = ws.cell(row=hdr2_row, column=ci, value=h)
    c.font = font(WHITE, 11, True); c.fill = fill(NAVY); c.alignment = align("center","center"); c.border=border

row = hdr2_row + 1
dim_score_rows_by_scenario = {i: [] for i in range(4)}  # column index 5..8 → SZ index 0..3
dim_weight_col = 2

for dim_name, dstart, dend in DIM_HEAD_ROWS:
    ws.cell(row=row, column=1, value=dim_name).font = font(NAVY, 11, True)
    ws.cell(row=row, column=1).fill = fill(GREY_L)
    ws.cell(row=row, column=1).border = border
    # Dim gewicht = the C cell at dstart
    ws.cell(row=row, column=2, value=f"=C{dstart}")
    ws.cell(row=row, column=2).number_format = "0.0%"
    ws.cell(row=row, column=2).alignment = align("center","center")
    ws.cell(row=row, column=2).border = border
    # For each scenario: SUMPRODUCT(krit_gewichte * score) / SUM(krit_gewichte)
    for si in range(4):
        col_letter_score = get_column_letter(5+si)
        formula = (f"=SUMPRODUCT(D{dstart}:D{dend},{col_letter_score}{dstart}:{col_letter_score}{dend})"
                   f"/SUM(D{dstart}:D{dend})")
        c = ws.cell(row=row, column=3+si, value=formula)
        c.number_format = "0.00"
        c.font = font(DARK, 11, True)
        c.alignment = align("center","center")
        c.border = border
        dim_score_rows_by_scenario[si].append(row)
    row += 1

# Total row
total_row = row
ws.cell(row=total_row, column=1, value="GESAMT (gewichtet)").font = font(WHITE, 12, True)
ws.cell(row=total_row, column=1).fill = fill(ORANGE)
ws.cell(row=total_row, column=1).border = border
ws.cell(row=total_row, column=1).alignment = align("center", "center")
ws.cell(row=total_row, column=2, value=f"=SUM(B{hdr2_row+1}:B{total_row-1})")
ws.cell(row=total_row, column=2).number_format = "0.0%"
ws.cell(row=total_row, column=2).font = font(WHITE, 12, True)
ws.cell(row=total_row, column=2).fill = fill(ORANGE)
ws.cell(row=total_row, column=2).alignment = align("center","center")
ws.cell(row=total_row, column=2).border = border
for si in range(4):
    col_letter_score = get_column_letter(3+si)
    formula = (f"=SUMPRODUCT(B{hdr2_row+1}:B{total_row-1},{col_letter_score}{hdr2_row+1}:{col_letter_score}{total_row-1})"
               f"/SUM(B{hdr2_row+1}:B{total_row-1})")
    c = ws.cell(row=total_row, column=3+si, value=formula)
    c.number_format = "0.00"
    c.font = font(WHITE, 12, True)
    c.fill = fill(ORANGE)
    c.alignment = align("center","center")
    c.border = border

# Color scale for the total row scenarios
ws.conditional_formatting.add(
    f"C{total_row}:F{total_row}",
    ColorScaleRule(start_type="min", start_color="F8D7DA",
                   mid_type="percentile", mid_value=50, mid_color="FFF3C4",
                   end_type="max", end_color="D4EDDA")
)

# Rank row
rank_row = total_row + 1
ws.cell(row=rank_row, column=1, value="Rang").font = font(NAVY, 11, True)
ws.cell(row=rank_row, column=1).fill = fill(GREY_H)
ws.cell(row=rank_row, column=1).alignment = align("center","center")
ws.cell(row=rank_row, column=1).border = border
ws.cell(row=rank_row, column=2, value="").border = border
for si in range(4):
    col_letter_score = get_column_letter(3+si)
    formula = f"=RANK({col_letter_score}{total_row},$C${total_row}:$F${total_row})"
    c = ws.cell(row=rank_row, column=3+si, value=formula)
    c.font = font(NAVY, 12, True)
    c.alignment = align("center","center")
    c.border = border
    c.fill = fill(GREY_H)

# Freeze
ws.freeze_panes = "E5"

# ============================================================
# SHEET 3 — Konsens (Kopie-Vorlage für Team-Ergebnis)
# ============================================================
ws2 = wb.create_sheet("Konsens (Team)")
ws2["A1"] = "Konsens-Bewertung des Teams"
ws2["A1"].font = font(ORANGE, 16, True)
ws2["A2"] = "Nach dem individuellen Scoring hier den gemeinsamen Konsens eintragen."
ws2["A2"].font = font("6B7280", 11, italic=True)
ws2["A4"] = "→  Vorschlag: Sheet »Scorecard« duplizieren, umbenennen (Konsens-Team), Scores aus Diskussion eintragen."
ws2["A4"].font = font(DARK, 11)
ws2["A6"] = "Alternativ: Excel-Menü »Blatt verschieben oder kopieren« → Kopie erstellen → Werte im neuen Sheet überschreiben."
ws2["A6"].font = font(DARK, 11)
ws2.column_dimensions["A"].width = 130

# ============================================================
# SHEET 4 — Ergebnis / Sensitivität
# ============================================================
ws3 = wb.create_sheet("Ergebnis")
ws3.sheet_view.showGridLines = False
ws3["A1"] = "Ergebnis-Übersicht"
ws3["A1"].font = font(ORANGE, 18, True)
ws3["A2"] = "Zieht Werte aus Sheet »Scorecard«. Sensitivitäts-Zeilen zeigen ±10% Gewichts-Variation."
ws3["A2"].font = font("6B7280", 11, italic=True)

for i, w in enumerate([22, 14, 14, 14, 14, 40], start=1):
    ws3.column_dimensions[get_column_letter(i)].width = w

# Header
hd = ["Szenario", "SZ 1", "SZ 2a", "SZ 2", "SZ 3", "Kommentar"]
for ci, h in enumerate(hd, start=1):
    c = ws3.cell(row=4, column=ci, value=h)
    c.font = font(WHITE, 11, True); c.fill = fill(NAVY); c.alignment = align("center","center"); c.border=border

# Row: Gesamt-Score
r = 5
ws3.cell(row=r, column=1, value="Gesamt-Score (Basis-Gewichte)").font = font(DARK, 11, True)
ws3.cell(row=r, column=1).fill = fill(GREY_L)
ws3.cell(row=r, column=1).border = border
for si in range(4):
    col_letter = get_column_letter(3+si)
    c = ws3.cell(row=r, column=2+si,
                 value=f"=Scorecard!{col_letter}{total_row}")
    c.number_format = "0.00"
    c.font = font(DARK, 12, True); c.alignment = align("center","center"); c.border=border
ws3.cell(row=r, column=6, value="Konsens-Score aus Team-Bewertung").border = border

# Row: Rang
r = 6
ws3.cell(row=r, column=1, value="Rangfolge").font = font(NAVY, 11, True)
ws3.cell(row=r, column=1).fill = fill(GREY_L)
ws3.cell(row=r, column=1).border = border
for si in range(4):
    c = ws3.cell(row=r, column=2+si, value=f"=RANK(B5:E5,{'$B$5:$E$5'})" if False else f"=RANK({get_column_letter(2+si)}5,$B$5:$E$5)")
    c.font = font(NAVY, 12, True); c.alignment = align("center","center"); c.border=border
ws3.cell(row=r, column=6, value="1 = bester Score").border = border

# Row: Empfehlung heuristisch
r = 7
ws3.cell(row=r, column=1, value="Kurz-Empfehlung").font = font(NAVY, 11, True)
ws3.cell(row=r, column=1).fill = fill(GREY_L)
ws3.cell(row=r, column=1).border = border
for si in range(4):
    col_letter = get_column_letter(2+si)
    c = ws3.cell(row=r, column=2+si,
                 value=f'=IF({col_letter}6=1,"Empfohlen",IF({col_letter}6=2,"Bedingt","Nicht empfohlen"))')
    c.font = font(DARK, 11, True); c.alignment = align("center","center"); c.border=border

# Sensitivity block
ws3["A10"] = "Sensitivität — Wenn eine Dimension +10% Gewicht erhält, wer gewinnt?"
ws3["A10"].font = font(NAVY, 13, True)
ws3.merge_cells("A10:F10")
# Header
for ci, h in enumerate(["Dimension +10%","SZ 1","SZ 2a","SZ 2","SZ 3","Kommentar"], start=1):
    c = ws3.cell(row=11, column=ci, value=h)
    c.font = font(WHITE, 11, True); c.fill = fill(NAVY); c.alignment = align("center","center"); c.border=border

# Simple placeholder: user can note qualitative winners after adjusting weights in Scorecard
sens_rows = [
    "Strategisch", "Funktional", "Technisch",
    "Wirtschaftlich", "Organisatorisch", "Risiko/Zukunft",
]
for i, dim in enumerate(sens_rows):
    r = 12 + i
    ws3.cell(row=r, column=1, value=dim).font = font(DARK, 11, True)
    ws3.cell(row=r, column=1).border = border; ws3.cell(row=r, column=1).fill = fill(GREY_L)
    for j in range(4):
        ws3.cell(row=r, column=2+j, value="").border = border
    ws3.cell(row=r, column=6, value="Nach Gewichts-Änderung im Sheet Scorecard hier Ergebnis notieren").border = border

# Color scale on the Ergebnis-Score row
ws3.conditional_formatting.add(
    "B5:E5",
    ColorScaleRule(start_type="min", start_color="F8D7DA",
                   mid_type="percentile", mid_value=50, mid_color="FFF3C4",
                   end_type="max", end_color="D4EDDA")
)

# Row heights
for r in range(4, 20):
    ws3.row_dimensions[r].height = 22

wb.save(DST)
print(f"Saved: {DST}")
