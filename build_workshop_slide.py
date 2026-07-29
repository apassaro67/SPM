"""Einzelfolie: Team-Building & Scope Workshop der Abteilung Digital Workplace
& Support. Zum Zeigen in der Abteilung. Farbwelt: Navy (DWS-Identitaet)."""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

DST = "2026_07_08_Workshop_Digital_Workplace_Support.pptx"

NAVY   = RGBColor(0x1B, 0x3A, 0x6B)   # DWS-Primär
NAVY_D = RGBColor(0x12, 0x2A, 0x50)
BLUE   = RGBColor(0x2E, 0x6C, 0xB0)   # Scope
BLUE_LT= RGBColor(0xE4, 0xEC, 0xF6)
ORANGE = RGBColor(0xE0, 0x60, 0x0C)   # Team Building (Energie)
ORANGE_LT = RGBColor(0xFB, 0xE7, 0xD6)
DARK   = RGBColor(0x1F, 0x29, 0x37); GREY = RGBColor(0x4B, 0x55, 0x63)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF); LIGHT = RGBColor(0xF4, 0xF6, 0xF9)
LINE   = RGBColor(0xD1, 0xD7, 0xDF); MUTED_W = RGBColor(0xB9, 0xC6, 0xDB)

U = 76200
def ex(u): return int(u * U)
FONT = "Calibri"

prs = Presentation(); prs.slide_width = ex(160); prs.slide_height = ex(90)
s = prs.slides.add_slide(prs.slide_layouts[6])

def rect(left, top, w, h, fill, line=None, lw=1.0, round_=False):
    shp = MSO_SHAPE.ROUNDED_RECTANGLE if round_ else MSO_SHAPE.RECTANGLE
    o = s.shapes.add_shape(shp, ex(left), ex(top), ex(w), ex(h))
    if round_:
        try: o.adjustments[0] = 0.09
        except Exception: pass
    if fill is None: o.fill.background()
    else: o.fill.solid(); o.fill.fore_color.rgb = fill
    if line is None: o.line.fill.background()
    else: o.line.color.rgb = line; o.line.width = Pt(lw)
    o.shadow.inherit = False
    return o

def text(left, top, w, h, paras, *, anchor=MSO_ANCHOR.TOP, ls=None):
    tb = s.shapes.add_textbox(ex(left), ex(top), ex(w), ex(h))
    tb.fill.background(); tb.line.fill.background(); tb.text_frame.vertical_anchor = anchor
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = Emu(40000); tf.margin_right = Emu(40000); tf.margin_top = Emu(12000); tf.margin_bottom = Emu(12000)
    tf.clear()
    for i, item in enumerate(paras):
        t, st = (item, {}) if isinstance(item, str) else item
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = st.get("align", PP_ALIGN.LEFT)
        if ls: p.line_spacing = ls
        if "space_after" in st: p.space_after = Pt(st["space_after"])
        runs = st.get("runs")
        if runs:
            for rt, rs in runs:
                r = p.add_run(); r.text = rt; r.font.name = FONT
                r.font.size = Pt(rs.get("size", st.get("size", 11))); r.font.bold = rs.get("bold", False)
                r.font.italic = rs.get("italic", False); r.font.color.rgb = rs.get("color", DARK)
        else:
            r = p.add_run(); r.text = t; r.font.name = FONT
            r.font.size = Pt(st.get("size", 11)); r.font.bold = st.get("bold", False)
            r.font.italic = st.get("italic", False); r.font.color.rgb = st.get("color", DARK)
    return tb

# ---- Hintergrund ----
rect(0, 0, 160, 90, WHITE)

# ---- Titelband ----
rect(0, 0, 160, 15.5, NAVY)
rect(0, 0, 160, 1.6, ORANGE)
text(8, 3.2, 120, 6, [("Team-Building & Scope Workshop", {"size": 25, "bold": True, "color": WHITE})])
text(8, 10.2, 120, 4, [("", {"runs": [
    ("Digital Workplace & Support", {"size": 13, "bold": True, "color": WHITE}),
    ("     „Wir sind für Sie da.“", {"size": 13, "italic": True, "color": MUTED_W})]})])
text(120, 4.2, 33, 8, [("Abteilung", {"size": 9, "color": MUTED_W, "align": PP_ALIGN.RIGHT}),
                       ("Alexander Passaro", {"size": 12, "bold": True, "color": WHITE, "align": PP_ALIGN.RIGHT})])

# ---- Ziele ----
text(7, 17.2, 100, 3, [("ZIELE DES WORKSHOPS", {"size": 10, "bold": True, "color": NAVY})])
ziele = ["Über Teamgrenzen kennenlernen", "Gemeinsames Zielbild", "Scope schärfen", "Zusammenarbeit vereinbaren"]
zw = 35.6
for i, z in enumerate(ziele):
    lx = 7 + i * (zw + 1.2)
    rect(lx, 20.4, zw, 5.3, LIGHT, line=LINE, lw=0.75, round_=True)
    rect(lx, 20.4, 0.7, 5.3, NAVY)
    text(lx + 1.4, 20.4, zw - 2, 5.3, [(z, {"size": 9.8, "bold": True, "color": DARK})], anchor=MSO_ANCHOR.MIDDLE)

# ---- Zwei Spalten: Team Building / Scope ----
def column(lx, col, col_lt, kicker, title, bullets):
    rect(lx, 28.5, 72, 35, WHITE, line=LINE, lw=1.0, round_=True)
    rect(lx, 28.5, 72, 8, col, round_=True)
    rect(lx, 33, 72, 3.5, col)
    text(lx + 2, 29, 68, 3, [(kicker, {"size": 9, "bold": True, "color": col_lt})])
    text(lx + 2, 31.6, 68, 4.5, [(title, {"size": 14, "bold": True, "color": WHITE})])
    paras = [("", {"space_after": 4, "runs": [("▸  ", {"color": col, "bold": True, "size": 10.5}),
                                              (b, {"color": DARK, "size": 10.5})]}) for b in bullets]
    text(lx + 2, 38, 68, 24.5, paras, ls=1.08)

column(7, ORANGE, ORANGE_LT, "TEIL 1 · MITEINANDER", "Team Building", [
    "Vorstellung bewusst gemischt — Person, Rolle, Beitrag",
    "Skills- & Stärken-Landkarte der Abteilung",
    "Erwartungen: „was ich beitrage / was ich brauche“",
    "Gemeinsame Werte & Team-Charter erarbeiten",
    "Neue Kolleg:innen aktiv einbinden (Buddy-Prinzip)",
])
column(81, BLUE, BLUE_LT, "TEIL 2 · AUFTRAG", "Scope & Auftrag", [
    "Gemeinsame Mission: Digital Workplace aus einer Hand",
    "Drei Säulen & ihren Kernauftrag klären",
    "In-Scope vs. Out-of-Scope je Säule abgrenzen",
    "Sechs Leitmotive mit Leben füllen",
    "Schnittstellen zwischen Teams & nach außen",
])

# ---- Drei Teams ----
text(7, 65, 100, 3, [("UNSERE DREI TEAMS", {"size": 10, "bold": True, "color": NAVY})])
teams = [("IT Support & ITSM Platform Services", "Patrick Böhme"),
         ("Modern Workplace & Experience", "Christopher Hahn"),
         ("Client Services", "Moritz Heller")]
tw = 47.3
for i, (t, lead) in enumerate(teams):
    lx = 7 + i * (tw + 2.5)
    rect(lx, 68, tw, 8.5, NAVY, round_=True)
    text(lx + 2, 68.5, tw - 4, 4.5, [(t, {"size": 10, "bold": True, "color": WHITE})], ls=1.0)
    text(lx + 2, 72.6, tw - 4, 3.5, [("", {"runs": [("Leitung:  ", {"size": 8.5, "color": MUTED_W}),
                                                    (lead, {"size": 9.5, "bold": True, "color": WHITE})]})])

# ---- Ergebnisse / nächste Schritte ----
rect(7, 78.5, 146, 8, LIGHT, line=LINE, lw=0.75, round_=True)
rect(7, 78.5, 1.0, 8, ORANGE)
text(9.5, 78.5, 143, 8, [("", {"runs": [
    ("Was wir mitnehmen:  ", {"size": 10.5, "bold": True, "color": NAVY}),
    ("Team-Charter & Spielregeln  ·  Scope-Landkarte (in/out + Schnittstellen)  ·  konkrete "
     "Commitments  ·  Follow-up-Termin", {"size": 10.5, "color": GREY})]})], anchor=MSO_ANCHOR.MIDDLE, ls=1.12)

prs.save(DST)
print("saved", DST, "with", len(prs.slides._sldIdLst), "slides")
