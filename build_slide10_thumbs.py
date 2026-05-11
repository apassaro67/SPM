"""Slide 10 v19: replace donut/percent charts with green 'thumbs up' shapes.
Thumb size scales with the percent value. Everything in green tones.
"""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

SRC = "/root/.claude/uploads/169d168f-c9e7-43eb-9b92-6f80b02d3b1f/d655d9ff-2026_05_18_VS_Zielbild_SNOW_SPM_v16.pptx"
DST = "/tmp/2026_05_18_VS_Zielbild_SNOW_SPM_v19.pptx"

# Green palette
GRN       = RGBColor(0x2E, 0x9B, 0x32)     # primary
GRN_LT    = RGBColor(0xDD, 0xF1, 0xDE)     # very light green
GRN_DK    = RGBColor(0x1F, 0x6D, 0x22)     # dark green for text
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
DARK      = RGBColor(0x1F, 0x29, 0x37)

prs = Presentation(SRC)
slide = prs.slides[9]

# Find and remove the existing chart + percent label pairs
to_remove = []
cells = []
shapes_list = list(slide.shapes)
i = 0
while i < len(shapes_list):
    sh = shapes_list[i]
    if sh.name.startswith("Chart") and abs(sh.width / 914400 - 1.30) < 0.05:
        pct_val = 80
        pct_shape = None
        if i + 1 < len(shapes_list):
            nxt = shapes_list[i + 1]
            if nxt.has_text_frame and "%" in nxt.text_frame.text:
                pct_val = int(nxt.text_frame.text.replace("%", "").strip())
                pct_shape = nxt
        cells.append({
            "left": sh.left / 914400,
            "top":  sh.top / 914400,
            "w":    sh.width / 914400,
            "h":    sh.height / 914400,
            "pct":  pct_val,
        })
        to_remove.append(sh)
        if pct_shape is not None:
            to_remove.append(pct_shape)
            i += 2
            continue
    i += 1

for sh in to_remove:
    sp = sh._element
    sp.getparent().remove(sp)

INCH = 914400

def add_oval_in(left, top, w, h, fill, line=None, line_w=0.0):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                               int(left * INCH), int(top * INCH),
                               int(w * INCH), int(h * INCH))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line; s.line.width = Pt(line_w)
    s.shadow.inherit = False
    return s

def add_rect_in(left, top, w, h, fill, line=None,
                shape=MSO_SHAPE.RECTANGLE, line_w=0.0):
    s = slide.shapes.add_shape(shape,
                               int(left * INCH), int(top * INCH),
                               int(w * INCH), int(h * INCH))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line; s.line.width = Pt(line_w)
    s.shadow.inherit = False
    return s

def add_text_in(left, top, w, h, text, *, size=10, bold=False, italic=False,
                color=DARK, font="Calibri", align=PP_ALIGN.CENTER,
                anchor=MSO_ANCHOR.MIDDLE):
    tb = slide.shapes.add_textbox(int(left * INCH), int(top * INCH),
                                  int(w * INCH), int(h * INCH))
    tb.fill.background(); tb.line.fill.background()
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(10000); tf.margin_right = Emu(10000)
    tf.margin_top = Emu(0); tf.margin_bottom = Emu(0)
    tf.vertical_anchor = anchor
    tf.text = ""
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run(); r.text = text
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    return tb

def thumbs_up_icon(cx, cy, scale, color):
    """Draw a stylized thumbs-up icon centered at (cx, cy)."""
    # All measurements in inches. 'scale' = max overall size in inches.
    # Hand fist proportions
    fist_w = scale * 0.62
    fist_h = scale * 0.50
    fist_x = cx - fist_w / 2
    fist_y = cy - fist_h / 2 + scale * 0.12

    # Thumb (vertical rounded rect rising from fist)
    thumb_w = scale * 0.25
    thumb_h = scale * 0.50
    thumb_x = cx - thumb_w / 2 - scale * 0.05
    thumb_y = cy - scale * 0.45

    # Cuff at bottom
    cuff_w = fist_w + scale * 0.10
    cuff_h = scale * 0.14
    cuff_x = cx - cuff_w / 2
    cuff_y = fist_y + fist_h - scale * 0.04

    # Cuff (rectangle - using a regular rounded rectangle)
    add_rect_in(cuff_x, cuff_y, cuff_w, cuff_h, color,
                shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    # Fist (rounded rectangle)
    add_rect_in(fist_x, fist_y, fist_w, fist_h, color,
                shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    # Thumb (rounded rectangle)
    add_rect_in(thumb_x, thumb_y, thumb_w, thumb_h, color,
                shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    # Small highlight oval at thumbnail tip
    nail_d = scale * 0.10
    add_oval_in(thumb_x + (thumb_w - nail_d) / 2,
                thumb_y + scale * 0.05,
                nail_d, nail_d, GRN_LT)

def render_cell(left, top, w, h, pct):
    """Render a green thumbs-up + percent for one cell at the chart slot."""
    # Background card (very light green tint) to visually unify the cell
    pad = 0.05
    card_w = w
    card_h = h + 0.45   # extend a bit to include the percent below
    card_x = left
    card_y = top - 0.05
    add_rect_in(card_x, card_y, card_w, card_h, GRN_LT,
                shape=MSO_SHAPE.ROUNDED_RECTANGLE)

    # Thumb size scales with percent: at 50% -> 0.65 of slot, at 100% -> 1.0
    scale_min = 0.55
    scale_max = 1.05
    t = max(0.0, min(1.0, (pct - 50) / 50.0))
    thumb_scale = scale_min + (scale_max - scale_min) * t   # in inches

    # Position thumb in upper portion of the slot
    icon_cx = left + w / 2
    icon_cy = top + (h * 0.46)
    thumbs_up_icon(icon_cx, icon_cy, thumb_scale, GRN)

    # Percent text below thumb
    pct_y = top + h - 0.05
    add_text_in(left, pct_y, w, 0.40, f"{pct} %",
                size=18, bold=True, color=GRN_DK,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

for c in cells:
    render_cell(c["left"], c["top"], c["w"], c["h"], c["pct"])

prs.save(DST)
print(f"Saved: {DST}  ({len(cells)} cells rendered)")
