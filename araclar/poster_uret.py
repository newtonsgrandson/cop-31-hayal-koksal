"""Build the COP31 digital poster (9:16, English) for the Climate İMECE project.

Usage: python poster.py <COP31_Digital_Poster_Template.pptx> <media_dir> <logo_dir> <out.pptx>
"""
import math
import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402
from helpers import tb, card, circle, pill, pic, WHITE, INK, GREEN, RED  # noqa: E402

TEMPLATE, MEDIA, LOGOS, OUT = sys.argv[1:5]
helpers.set_media(MEDIA)
helpers.DEFAULT_FONT = "Arial"

# poster palette (template's own blue + deck accents)
BLUE = RGBColor(0x19, 0x4E, 0x9C)
LBLUE = RGBColor(0x1B, 0x75, 0xBB)
TINT = RGBColor(0xF2, 0xF6, 0xFA)
LINE = RGBColor(0xDC, 0xE6, 0xED)
SLATE = RGBColor(0x61, 0x74, 0x87)

L, R = 0.75, 12.75
W = R - L
GAP = 0.25
CW = (W - GAP) / 2          # column width
X2 = L + CW + GAP           # second column x
PAD = 0.35                  # card inner padding
IW = CW - 2 * PAD           # inner width of a half card
HEAD_H = 0.8                # card heading band


def section(s, name, x, y, w, h, title):
    card(s, f"{name} card", x, y, w, h, fill=WHITE, line=LINE, radius=0.15)
    tb(s, f"{name} heading", x + PAD, y + 0.28, w - 2 * PAD, 0.45, [title], size=20, bold=True,
       color=BLUE, margin=0)
    return y + HEAD_H


def numbered(s, name, x, y, w, items, size=12, d=0.32, gap=0.07, fill_a=BLUE, fill_b=GREEN):
    """Numbered rows with estimated heights; returns y after the last row."""
    cpl = max(10, int((w - d - 0.25) * 72 / (size * 0.52)))
    for i, it in enumerate(items):
        n, cur = 1, 0
        for word in it.split():
            if cur and cur + 1 + len(word) > cpl:
                n += 1; cur = len(word)
            else:
                cur += (1 if cur else 0) + len(word)
        hh = max(d, n * size * 1.22 / 72 + 0.04)
        circle(s, f"{name} num {i+1}", x + d / 2, y + hh / 2, d, fill=fill_a if i % 2 == 0 else fill_b,
               text=str(i + 1), size=size - 2)
        tb(s, f"{name} text {i+1}", x + d + 0.15, y, w - d - 0.15, hh, [it], size=size,
           anchor=MSO_ANCHOR.MIDDLE, margin=0.02)
        y += hh + gap
    return y


def build():
    prs = Presentation(TEMPLATE)
    s = prs.slides[0]

    # --- template placeholders -------------------------------------------
    for sh in list(s.shapes):
        if sh.name in ("Institution logo placeholder", "EDIT_UniversityLogo"):
            sh._element.getparent().remove(sh._element)
        elif sh.has_text_frame and "[INSTITUTION NAME]" in sh.text_frame.text:
            for par in sh.text_frame.paragraphs:
                for r in par.runs:
                    r.text = r.text.replace("[INSTITUTION NAME]", "Istanbul Kent University")
    logo = pic(s, "Istanbul Kent University logo", f"{LOGOS}/kent_logo_lacivert_kirmizi_crop.png",
               0, 1.5, h=1.35)
    logo.left = Inches(6.75) - logo.width // 2

    # --- title block (3.2 – 5.9) ---------------------------------------------
    y = 3.2
    tb(s, "Poster title", L, y, W, 0.85, ["CLIMATE İMECE"], size=52, bold=True, color=BLUE,
       align=PP_ALIGN.CENTER, margin=0)
    tb(s, "Poster subtitle", L + 0.5, y + 0.9, W - 1.0, 0.75,
       ["From Climate Action to Implementation with Young People, Educational Institutions and Local Communities"],
       size=20, color=INK, align=PP_ALIGN.CENTER, margin=0, line_spacing=1.05)
    tb(s, "Authors", L, y + 1.72, W, 0.3,
       [[("Prof. Dr. Hayal Köksal", dict(bold=True)), (" (Project Director)  ·  Dr. Memet Aslan  ·  Aleyna Çavdar  ·  Azad Kılıçarslan  ·  Ayça Selin Akyol", {})]],
       size=14, color=INK, align=PP_ALIGN.CENTER, margin=0)
    tb(s, "Affiliation", L, y + 2.04, W, 0.3,
       ["Istanbul Kent University, Türkiye  ·  hayal.koksal@kent.edu.tr"],
       size=13, color=SLATE, align=PP_ALIGN.CENTER, margin=0)
    tb(s, "Principle", L, y + 2.38, W, 0.32,
       [[("Guiding principle:  ", dict(bold=True)), ("Compete with Yourself, Collaborate with Others.", dict(italic=True))]],
       size=14, color=GREEN, align=PP_ALIGN.CENTER, margin=0)

    # --- row A: background / what is İMECE (6.1 – 9.8) ----------------------
    ya, ha = 6.1, 3.7
    cy = section(s, "Background", L, ya, CW, ha, "BACKGROUND & QUESTION")
    by = helpers.bullets(s, "Background", L + PAD, cy, IW, 1.6, [
        "Awareness alone is not enough: climate change is a multidimensional challenge.",
        "Local, experiential knowledge must meet academic knowledge.",
        "Young people need real opportunities to take part in solutions."], size=14, gap=4, dot=GREEN)
    qy = max(by + 0.1, ya + ha - 1.15)
    card(s, "Question card", L + PAD, qy, IW, ya + ha - 0.25 - qy, fill=BLUE, line=None, radius=0.12)
    tb(s, "Question", L + PAD + 0.15, qy, IW - 0.3, ya + ha - 0.25 - qy,
       ["How can we move young people from climate awareness to meaningful climate action?"],
       size=15, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)

    cy = section(s, "Imece", X2, ya, CW, ha, "WHAT IS İMECE?")
    tb(s, "Imece definition", X2 + PAD, cy, IW, 0.8,
       ["A traditional Turkish concept of collective effort: people coming together voluntarily to achieve a shared goal. In our educational model it represents:"],
       size=14, margin=0.02)
    vals = ["Belonging", "Participation", "Collaboration", "Collective Learning",
            "Collective Intelligence", "Collective Action", "Quality & Social Impact"]
    pw, ph_, pitch = (IW - 0.15) / 2, 0.38, 0.46
    for i, v in enumerate(vals):
        col, row = i % 2, i // 2
        pill(s, f"Value {i+1}", X2 + PAD + col * (pw + 0.15) + pw / 2, cy + 0.95 + row * pitch + ph_ / 2,
             pw, ph_, v, fill=[BLUE, GREEN, LBLUE][i % 3], size=12)

    # --- row B: İMECE circle / methodology (10.05 – 14.35) ------------------
    yb, hb = ya + ha + GAP, 4.3
    cy = section(s, "Circle", L, yb, CW, hb, "THE İMECE CIRCLE")
    tb(s, "Circle tag", L + PAD, cy - 0.1, IW, 0.35,
       ["A small, voluntary and collaborative learning community"], size=13, italic=True, color=SLATE, margin=0)
    ccx, ccy = L + CW / 2, cy + 1.95
    rx, ry = 2.3, 1.3
    ring = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(ccx - rx), Inches(ccy - ry), Inches(2 * rx), Inches(2 * ry))
    ring.name = "Ring"; ring.fill.background(); ring.line.color.rgb = LBLUE; ring.line.width = Pt(1.25)
    ring.shadow.inherit = False
    p = pic(s, "Circle photo", "image6.png", 0, 0, h=1.6)
    p.left = Inches(ccx) - p.width // 2; p.top = Inches(ccy - 0.8)
    labels = ["REAL PROBLEM", "PEOPLE AFFECTED", "SHARED GOAL", "COLLABORATION",
              "ACTION", "REFLECTION", "CONTINUOUS IMPROVEMENT"]
    for i, lab in enumerate(labels):
        a = -math.pi / 2 + i * 2 * math.pi / len(labels)
        pill(s, f"Ring label {i+1}", ccx + rx * math.cos(a), ccy + ry * math.sin(a), 1.6, 0.4, lab,
             fill=BLUE if i % 2 == 0 else GREEN, size=9)

    cy = section(s, "Method", X2, yb, CW, hb, "METHODOLOGY: İMECE + PDCA")
    tb(s, "Method lead", X2 + PAD, cy - 0.1, IW, 0.5,
       ["A participatory, action-oriented pilot study: İMECE Circles work through the PDCA quality-improvement cycle."],
       size=13, color=INK, margin=0)
    pdca = [("PLAN", "Identify problems, plan solutions", BLUE), ("DO", "Implement selected actions", GREEN),
            ("CHECK", "Monitor and evaluate results", LBLUE), ("ACT", "Improve, adapt and share", RED)]
    bw, bh = (IW - 0.15) / 2, 0.56
    for i, (k, v, col) in enumerate(pdca):
        c_, r_ = i % 2, i // 2
        bx, byy = X2 + PAD + c_ * (bw + 0.15), cy + 0.5 + r_ * (bh + 0.1)
        card(s, f"PDCA {k}", bx, byy, bw, bh, fill=TINT, line=None, radius=0.1)
        circle(s, f"PDCA letter {k}", bx + 0.33, byy + bh / 2, 0.4, fill=col, text=k[0], size=14)
        tb(s, f"PDCA text {k}", bx + 0.62, byy, bw - 0.67, bh,
           [[(k + "  ", dict(bold=True, color=col)), (v, {})]], size=12, anchor=MSO_ANCHOR.MIDDLE, margin=0.02)
    wy = cy + 0.5 + 2 * (bh + 0.1) + 0.05
    tb(s, "WP heading", X2 + PAD, wy, IW, 0.3, ["Four work packages"], size=14, bold=True, color=BLUE, margin=0)
    wps = ["WP1 Preparation: literature, instruments, participants",
           "WP2 İMECE Circles: local challenges, common problems",
           "WP3 International learning: online partner meetings",
           "WP4 Joint solutions: develop, pilot and evaluate"]
    helpers.bullets(s, "WP", X2 + PAD, wy + 0.35, IW, 1.3, wps, size=12, gap=3, dot=BLUE)

    # --- row C: aim & objectives / partners (14.6 – 18.4) --------------------
    yc, hc = yb + hb + GAP, 3.8
    cy = section(s, "Aim", L, yc, CW, hc, "AIM & OBJECTIVES")
    tb(s, "Aim text", L + PAD, cy - 0.1, IW, 0.55,
       ["A practical Climate İMECE model that enables young people to identify challenges, combine local and academic knowledge, and pilot solutions."],
       size=12, color=INK, margin=0)
    numbered(s, "Objective", L + PAD, cy + 0.55, IW, [
        "Pilot Climate İMECE application at Istanbul Kent University",
        "Collaborative İMECE Circles",
        "At least three applicable climate solutions",
        "International learning: Türkiye, Algeria and India",
        "A practical implementation guide",
        "Transferability to other institutions and communities"], size=12, d=0.32, gap=0.07)

    cy = section(s, "Partners", X2, yc, CW, hc, "PARTNERS & COOPERATION")
    for i, (cn, col) in enumerate([("Türkiye", BLUE), ("Algeria", GREEN), ("India", RED)]):
        pill(s, f"Country {i+1}", X2 + PAD + 0.7 + i * 1.55, cy + 0.1, 1.4, 0.38, cn, fill=col, size=12)
    tb(s, "Lead text", X2 + PAD, cy + 0.5, IW, 0.3,
       [[("Lead institution:  ", dict(bold=True, color=BLUE)), ("Istanbul Kent University, Türkiye", {})]],
       size=13, margin=0)
    tb(s, "Partners heading", X2 + PAD, cy + 0.9, IW, 0.3, ["International partners"], size=13, bold=True, color=BLUE, margin=0)
    helpers.bullets(s, "Partner", X2 + PAD, cy + 1.2, IW, 2.0, [
        "Kasdi Merbah University, Algeria",
        "World Council for Total Quality and Excellence in Education (WCTQEE)",
        "Little Angels High School, Gwalior · Suncity School, Gurgaon · Bhavan Vidyalaya, Panchkula · Gyan Heritage School, Lucknow (India) & more"],
        size=12, gap=3, dot=GREEN)

    # --- row D: expected outputs (18.65 – 21.0) ------------------------------
    yd, hd = yc + hc + GAP, 2.35
    cy = section(s, "Outputs", L, yd, W, hd, "EXPECTED OUTPUTS  ·  project targets, not yet measured results")
    outs = ["A Climate İMECE Youth Model", "At least three youth-developed climate solutions",
            "Pilot implementation activities", "A practical Climate İMECE Implementation Guide",
            "An International Youth Climate Learning Network", "A digital archive of experiences and solutions",
            "At least one scientific paper or conference presentation"]
    cols = 4
    gw = (W - 2 * PAD - 0.2 * (cols - 1)) / cols
    for i, it in enumerate(outs):
        c_, r_ = i % cols, i // cols
        ox, oy = L + PAD + c_ * (gw + 0.2), cy - 0.05 + r_ * 0.75
        circle(s, f"Output num {i+1}", ox + 0.17, oy + 0.17, 0.34, fill=BLUE if i % 2 == 0 else GREEN, text=str(i + 1), size=11)
        tb(s, f"Output text {i+1}", ox + 0.45, oy, gw - 0.45, 0.7, [it], size=12, margin=0.02)
    ox, oy = L + PAD + 3 * (gw + 0.2), cy - 0.05 + 0.75
    tb(s, "Evaluation note", ox, oy, gw, 0.7,
       [[("Evaluation: ", dict(bold=True)), ("pre/post assessment forms, İMECE process records, participant reflections.", {})]],
       size=11, color=SLATE, italic=True, margin=0.02)
    # the template's instruction notes must not ship with the submission
    if s.has_notes_slide:
        s.notes_slide.notes_text_frame.text = ""
    print("bottom", yd + hd)
    prs.save(OUT)
    print("wrote", OUT)


build()
