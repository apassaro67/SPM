"""Fill slide 9 (Vorprojekt SNOW – Ergebnisse 1 von 2) with the
team-statement content from the workshop image.

Mapping per cell (matching the existing slide layout):
  Top row     | Entwicklung (EWW)        | Produktion (TGG)        | IT (VFI)
  Bottom row  | Strategie (SCD)          | Produktmanagement (MPM) | Finanzen (VFI)
"""
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

SRC = "/tmp/2026_05_18_VS_Zielbild_SNOW_SPM_v3.pptx"
DST = "/tmp/2026_05_18_VS_Zielbild_SNOW_SPM_v4.pptx"

DARK = RGBColor(0x1F, 0x29, 0x37)

# 3 kurze Bullets pro Feld (vorstandsgerecht verdichtet)
content = {
    21: [  # Entwicklung (EWW)
        "Szenario abbildbar (unter Umständen)",
        "Schwerpunkt Ressourcen- & Finanzplanung",
        "Anbietervergleich notwendig",
    ],
    22: [  # Produktion (TGG)
        "Szenario abbildbar",
        "Anforderungen weiter detaillieren",
        "Starker Implementierungspartner nötig",
    ],
    23: [  # IT (VFI)
        "IT-Szenario abbildbar",
        "Hohes Potenzial der SNOW-Plattform",
        "Detail-Klärung zur Jahresplanung",
    ],
    24: [  # Strategie (SCD)
        "SPM (Portfolio & OKR) abbildbar",
        "Sauberes Scope- & Rollout-Mgmt nötig",
        "Potenzial: zentrale Datenbasis",
    ],
    25: [  # Produktmanagement (MPM)
        "SPP Gantt & Bottom-up abbildbar",
        "Integration LH/CRD → Steckbrief",
        "Durchgängige Kette Strategie → Projekt",
    ],
    26: [  # Finanzen (VFI)
        "Szenario abbildbar",
        "Anforderungen weiter detaillieren",
        "Folge-Workshops zur Schärfung",
    ],
}

# Index of the percent text overlays (Textplatzhalter 5)
percent_updates = {
    28: "80%",  # Entwicklung
    30: "80%",  # Produktion
    32: "80%",  # IT
    34: "80%",  # Strategie
    36: "80%",  # Produktmanagement
    38: "60%",  # Finanzen
}

prs = Presentation(SRC)
slide = prs.slides[8]

shapes = list(slide.shapes)

def set_bullets(shape, bullets, *, size=10):
    tf = shape.text_frame
    tf.word_wrap = True
    # Reset
    tf.clear()
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.level = 0
        run = p.add_run()
        run.text = b
        run.font.name = "Calibri"
        run.font.size = Pt(size)
        run.font.color.rgb = DARK

for idx, bullets in content.items():
    set_bullets(shapes[idx], bullets, size=10)

# Update percent labels
for idx, val in percent_updates.items():
    sh = shapes[idx]
    tf = sh.text_frame
    # Preserve existing formatting; just swap text in first run
    if tf.paragraphs and tf.paragraphs[0].runs:
        tf.paragraphs[0].runs[0].text = val
    else:
        tf.text = val

prs.save(DST)
print("Saved:", DST)
