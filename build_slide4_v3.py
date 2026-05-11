"""Slide 4 v15: compact table + click-triggered animations
- Top table visible immediately
- Click 1: TCO chart appears + dim overlay fades onto the top table
- Click 2: Gartner Executive Summary appears
"""
from copy import deepcopy
from lxml import etree
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn, nsmap

SRC = "/root/.claude/uploads/6f2327e1-58f0-4204-b00f-a2cbff9fd574/699f5bf9-2026_05_18_VS_Zielbild_SNOW_SPM_v13_1.pptx"
DST = "/tmp/2026_05_18_VS_Zielbild_SNOW_SPM_v15.pptx"
CHART_PNG = "/tmp/tco_chart.png"

ORANGE   = RGBColor(0xF0, 0x7F, 0x12)
DARK     = RGBColor(0x1F, 0x29, 0x37)
GREY_TXT = RGBColor(0x4B, 0x55, 0x63)
LIGHT_BG = RGBColor(0xF7, 0xF7, 0xF7)
ROW_ALT  = RGBColor(0xEE, 0xF1, 0xF4)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
GREEN    = RGBColor(0x2E, 0x7D, 0x32)
NAVY     = RGBColor(0x2C, 0x3E, 0x50)
LIGHT_GREY = RGBColor(0xD1, 0xD5, 0xDB)
BLACK    = RGBColor(0x00, 0x00, 0x00)

U = 76200
def ex(u): return int(u * U)

prs = Presentation(SRC)
slide = prs.slides[3]

# Wipe slide 4
for shp in list(slide.shapes):
    sp = shp._element
    sp.getparent().remove(sp)

def add_rect(left, top, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE,
             line_w=1.0, alpha=None):
    s = slide.shapes.add_shape(shape, ex(left), ex(top), ex(w), ex(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if alpha is not None:
        # add alpha to solidFill srgbClr
        sp = s.fill._xPr
        srgb = sp.find('.//' + qn('a:srgbClr'))
        if srgb is not None:
            srgb.set('val', "%02X%02X%02X" % (fill[0], fill[1], fill[2]))
            a = etree.SubElement(srgb, qn('a:alpha'))
            a.set('val', str(int(alpha * 100000)))
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line; s.line.width = Pt(line_w)
    s.shadow.inherit = False
    return s

def add_oval(left, top, w, h, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, ex(left), ex(top), ex(w), ex(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    s.line.fill.background()
    return s

def set_paragraphs(shape, paragraphs, *, default_size=10, default_color=DARK,
                   default_align=PP_ALIGN.LEFT):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(40000); tf.margin_right = Emu(40000)
    tf.margin_top = Emu(15000); tf.margin_bottom = Emu(15000)
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

def add_text(left, top, w, h, paragraphs, *, anchor=MSO_ANCHOR.MIDDLE, **kw):
    tb = slide.shapes.add_textbox(ex(left), ex(top), ex(w), ex(h))
    tb.fill.background(); tb.line.fill.background()
    tb.text_frame.vertical_anchor = anchor
    set_paragraphs(tb, paragraphs, **kw)
    return tb

# Container for tracking which shapes belong to which animation group
TCO_GROUP = []
GARTNER_GROUP = []
DIM_GROUP = []

def track(shape, group):
    group.append(shape.shape_id)
    return shape

# ========== HEADER ==========
add_rect(0, 0, 160, 6, DARK)
add_text(3, 0.4, 140, 2.6, [
    ("Bewertung der Plattformstrategie mit ServiceNow",
     dict(size=20, bold=True, color=WHITE)),
], anchor=MSO_ANCHOR.TOP)
add_text(3, 3.2, 140, 2.4, [
    ("Wirtschaftlichkeit (TCO) · Vergleich SPM-Plattform vs. Best-of-Breed · Gartner Executive Summary 2025",
     dict(size=12, italic=True, color=LIGHT_GREY)),
], anchor=MSO_ANCHOR.TOP)
add_rect(150, 0, 10, 6, ORANGE)
add_text(150, 0, 10, 6, [
    ("Folie 4", dict(size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)),
])

# ========== VERGLEICHSMATRIX ==========
COL_X = 4
TBL_W = 152
COL_W = [42, 50, 60]
TBL_TOTAL = sum(COL_W)
def colx(i): return COL_X + sum(COL_W[:i])

mat_top = 7.0
TABLE_X, TABLE_Y = COL_X, mat_top  # for dim overlay
add_text(COL_X, mat_top, TBL_W, 2.2, [
    ("Funktions-Vergleich – SPM-Plattform vs. Best-of-Breed",
     dict(size=13, bold=True, color=DARK)),
])

hdr_y = mat_top + 2.6
hdr_h = 3.5
add_rect(COL_X, hdr_y, COL_W[0], hdr_h, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(COL_X + 1.2, hdr_y, COL_W[0] - 2, hdr_h, [
    ("Funktion", dict(size=11, bold=True, color=WHITE)),
])
add_rect(colx(1), hdr_y, COL_W[1], hdr_h, ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(colx(1) + 1, hdr_y, COL_W[1] - 2, hdr_h, [
    ("SPM-Plattform (ServiceNow)",
     dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)),
], anchor=MSO_ANCHOR.MIDDLE)
add_rect(colx(2), hdr_y, COL_W[2], hdr_h, GREY_TXT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(colx(2) + 1, hdr_y, COL_W[2] - 2, hdr_h, [
    ("Best-of-Breed (Einzeltools)",
     dict(size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)),
], anchor=MSO_ANCHOR.MIDDLE)

rows = [
    ("Strategie (Ziele, KPI)",     "integriert", "Workpath"),
    ("Demand & Portfolio Management","integriert", "ServiceNow / SharePoint-Listen / Power-Automate"),
    ("Workflow Management",         "integriert", "SharePoint-Listen / Power-Automate"),
    ("Projekt-Mgmt. – Ressourcen-Planung","integriert", "Planisware"),
    ("Aufgaben-Management",         "integriert", "MS Planner · Jira"),
    ("AI-Fähigkeiten",              "integriert", "Je Tool separat, unterschiedliche Reife"),
    ("Reporting",                   "integriert + Power BI", "Power BI"),
]
row_y = hdr_y + hdr_h + 0.3
row_h = 3.2
row_gap = 0.2
for i, (krit, snow, bob) in enumerate(rows):
    y = row_y + i * (row_h + row_gap)
    bg = LIGHT_BG if i % 2 == 0 else ROW_ALT
    add_rect(COL_X, y, TBL_TOTAL, row_h, bg)
    add_rect(COL_X, y, COL_W[0], row_h, NAVY)
    add_text(COL_X + 1.2, y, COL_W[0] - 2, row_h, [
        (krit, dict(size=10.5, bold=True, color=WHITE)),
    ], anchor=MSO_ANCHOR.MIDDLE)
    add_rect(colx(1), y, 0.6, row_h, ORANGE)
    is_integ = "integ" in snow.lower()
    if is_integ:
        add_oval(colx(1) + 1.2, y + row_h/2 - 1.0, 2.0, 2.0, GREEN)
        add_text(colx(1) + 1.2, y + row_h/2 - 1.0, 2.0, 2.0, [
            ("✓", dict(size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)),
        ], anchor=MSO_ANCHOR.MIDDLE)
        add_text(colx(1) + 3.8, y, COL_W[1] - 4.5, row_h, [
            (snow, dict(size=10.5, bold=True, color=DARK)),
        ], anchor=MSO_ANCHOR.MIDDLE)
    else:
        add_text(colx(1) + 1.2, y, COL_W[1] - 2, row_h, [
            (snow, dict(size=10.5, color=DARK)),
        ], anchor=MSO_ANCHOR.MIDDLE)
    add_text(colx(2) + 1.2, y, COL_W[2] - 2, row_h, [
        (bob, dict(size=10.5, color=GREY_TXT)),
    ], anchor=MSO_ANCHOR.MIDDLE)

table_bot = row_y + len(rows) * (row_h + row_gap)
TABLE_H = table_bot - TABLE_Y - 0.5

# ========== BOTTOM SPLIT ==========
bot_top = table_bot + 1.0
left_w = 76
right_w = TBL_W - left_w - 2

# --- TCO panel (animated, click 1) ---
tco_x = COL_X
track(add_rect(tco_x, bot_top, left_w, 3.6, ORANGE,
               shape=MSO_SHAPE.ROUNDED_RECTANGLE), TCO_GROUP)
track(add_text(tco_x + 1.5, bot_top, left_w - 3, 3.6, [
    ("Wirtschaftlichkeitsbetrachtung – TCO-Analyse 5 Jahre",
     dict(size=12, bold=True, color=WHITE)),
], anchor=MSO_ANCHOR.MIDDLE), TCO_GROUP)

chart_top = bot_top + 4.0
chart_body_h = 86 - chart_top - 4
track(add_rect(tco_x, chart_top, left_w, chart_body_h, LIGHT_BG,
               shape=MSO_SHAPE.ROUNDED_RECTANGLE), TCO_GROUP)

pic_pad = 1.0
pic = slide.shapes.add_picture(
    CHART_PNG,
    ex(tco_x + pic_pad), ex(chart_top + pic_pad),
    width=ex(left_w - 2*pic_pad), height=ex(chart_body_h - 2*pic_pad)
)
TCO_GROUP.append(pic.shape_id)

# --- Gartner panel (animated, click 2) ---
gar_x = COL_X + left_w + 2
track(add_rect(gar_x, bot_top, right_w, 3.6, NAVY,
               shape=MSO_SHAPE.ROUNDED_RECTANGLE), GARTNER_GROUP)
track(add_text(gar_x + 1.5, bot_top, right_w - 3, 3.6, [
    ("Gartner Executive Summary 2025  (Pure-Play vs. Platform)",
     dict(size=12, bold=True, color=WHITE)),
], anchor=MSO_ANCHOR.MIDDLE), GARTNER_GROUP)

gar_top = bot_top + 4.0
gar_h = chart_body_h
track(add_rect(gar_x, gar_top, right_w, gar_h, LIGHT_BG,
               shape=MSO_SHAPE.ROUNDED_RECTANGLE), GARTNER_GROUP)

statements = [
    ("Integration schlägt Funktionsbreite",
     "Nur eine Plattform ermöglicht echte End-to-End-Governance über Strategie, Projekte, Produkte, IT & Finance hinweg."),
    ("TCO + Governance-Vorteile",
     "Weniger Tools · weniger Schnittstellen · konsistente KPIs."),
    ("Zukunftssicherheit",
     "Plattformen wie ServiceNow erweitern SPM nativ um AI und Workflow-Automatisierung."),
]
sb_h = (gar_h - 2.4) / 3
for i, (head, body) in enumerate(statements):
    y = gar_top + 0.8 + i * sb_h
    track(add_rect(gar_x + 1.5, y, right_w - 3, sb_h - 0.8, WHITE,
                   shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                   line=ORANGE, line_w=1.2), GARTNER_GROUP)
    track(add_oval(gar_x + 2.5, y + 0.9, 2.6, 2.6, NAVY), GARTNER_GROUP)
    track(add_text(gar_x + 2.5, y + 0.9, 2.6, 2.6, [
        (str(i + 1), dict(size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)),
    ], anchor=MSO_ANCHOR.MIDDLE), GARTNER_GROUP)
    track(add_text(gar_x + 6.0, y + 0.5, right_w - 8, 2.8, [
        (head, dict(size=11, bold=True, color=DARK)),
    ], anchor=MSO_ANCHOR.MIDDLE), GARTNER_GROUP)
    track(add_text(gar_x + 6.0, y + 3.0, right_w - 8, sb_h - 3.8, [
        (body, dict(size=9.5, color=GREY_TXT)),
    ], anchor=MSO_ANCHOR.TOP), GARTNER_GROUP)

track(add_text(gar_x + 2, gar_top + gar_h - 2.0, right_w - 4, 1.8, [
    ("Quelle: Gartner 2025 – Pure-Play Versus Platform, Comparing SPM and APMR Solutions",
     dict(size=8, italic=True, color=GREY_TXT)),
], anchor=MSO_ANCHOR.MIDDLE), GARTNER_GROUP)

# ========== DIM OVERLAY (animated, click 1) ==========
dim = add_rect(TABLE_X - 0.3, TABLE_Y - 0.3, TBL_W + 0.6, TABLE_H + 0.6, BLACK,
               shape=MSO_SHAPE.RECTANGLE, alpha=0.45)
DIM_GROUP.append(dim.shape_id)

# ========== FOOTER ==========
add_rect(0, 89.3, 160, 0.7, ORANGE)
add_text(4, 86.3, 152, 2.0, [
    ("Konsolidiert aus 'Bewertung_der_Plattformstrategie' + TCO_Vergleich_5J_SN_vs_BoB.xlsx (Tab: Auswertung).",
     dict(size=8, italic=True, color=GREY_TXT)),
])

# ============================================================
# Animation: build <p:timing> element and append to slide XML
# ============================================================
print("DIM_GROUP:", DIM_GROUP)
print("TCO_GROUP size:", len(TCO_GROUP))
print("GARTNER_GROUP size:", len(GARTNER_GROUP))

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

def make_appear_effect(tn_id, spid, click=False):
    """Build a <p:par> appear effect for one shape.
    click=True for the first shape in a click; otherwise withEffect."""
    nodeType = "clickEffect" if click else "withEffect"
    xml = f'''<p:par xmlns:p="{P_NS}">
      <p:cTn id="{tn_id}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="{nodeType}">
        <p:stCondLst><p:cond delay="0"/></p:stCondLst>
        <p:childTnLst>
          <p:set>
            <p:cBhvr>
              <p:cTn id="{tn_id+1}" dur="1" fill="hold">
                <p:stCondLst><p:cond delay="0"/></p:stCondLst>
              </p:cTn>
              <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
              <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
            </p:cBhvr>
            <p:to><p:strVal val="visible"/></p:to>
          </p:set>
        </p:childTnLst>
      </p:cTn>
    </p:par>'''
    return etree.fromstring(xml)

def make_click_group(click_inner_tn_id, spids):
    """A <p:par> wrapping all effects that play with one click."""
    xml = f'''<p:par xmlns:p="{P_NS}">
      <p:cTn id="{click_inner_tn_id}" fill="hold">
        <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
        <p:childTnLst>
          <p:par>
            <p:cTn id="{click_inner_tn_id+1}" fill="hold">
              <p:stCondLst><p:cond delay="0"/></p:stCondLst>
              <p:childTnLst></p:childTnLst>
            </p:cTn>
          </p:par>
        </p:childTnLst>
      </p:cTn>
    </p:par>'''
    par = etree.fromstring(xml)
    inner_child = par.find('.//' + qn('p:childTnLst'))
    inner_child = inner_child.find(qn('p:par')).find(qn('p:cTn')).find(qn('p:childTnLst'))
    tn = click_inner_tn_id + 2
    for i, spid in enumerate(spids):
        eff = make_appear_effect(tn, spid, click=(i == 0))
        inner_child.append(eff)
        tn += 2
    return par, tn

# Build complete timing element
timing_xml = f'''<p:timing xmlns:p="{P_NS}">
  <p:tnLst>
    <p:par>
      <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
        <p:childTnLst>
          <p:seq concurrent="1" nextAc="seek">
            <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
              <p:childTnLst></p:childTnLst>
            </p:cTn>
            <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
            <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
          </p:seq>
        </p:childTnLst>
      </p:cTn>
    </p:par>
  </p:tnLst>
</p:timing>'''
timing = etree.fromstring(timing_xml)
main_child = timing.find('.//' + qn('p:cTn') + '[@nodeType="mainSeq"]/' + qn('p:childTnLst'))

# Click 1: dim overlay + TCO group
click1_spids = DIM_GROUP + TCO_GROUP
click1, next_tn = make_click_group(3, click1_spids)
main_child.append(click1)

# Click 2: Gartner group
click2, _ = make_click_group(next_tn, GARTNER_GROUP)
main_child.append(click2)

# Insert <p:timing> into slide XML (right after cSld)
sld = slide._element
# Remove any existing timing
existing = sld.find(qn('p:timing'))
if existing is not None:
    sld.remove(existing)
sld.append(timing)

# Build <p:bldLst> per spec for the appearing shapes
bld_xml_parts = ['<p:bldLst xmlns:p="{0}">'.format(P_NS)]
for spid in click1_spids + GARTNER_GROUP:
    bld_xml_parts.append(f'<p:bldP spid="{spid}" grpId="0"/>')
bld_xml_parts.append('</p:bldLst>')
bld = etree.fromstring("".join(bld_xml_parts))
timing.append(bld)

prs.save(DST)
print("Saved:", DST)
