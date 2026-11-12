"""Shared drawing helpers for the COP31 deck and poster builders (python-pptx)."""
import math
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE

MEDIA = "."
DEFAULT_FONT = None


def set_media(path):
    global MEDIA
    MEDIA = path


# ---- template palette -------------------------------------------------------
BLUE = RGBColor(0x12, 0x4D, 0xA2)
LBLUE = RGBColor(0x1B, 0x75, 0xBB)
RED = RGBColor(0xE4, 0x1E, 0x26)
GREEN = RGBColor(0x44, 0xB9, 0x7E)
INK = RGBColor(0x1F, 0x29, 0x37)
MUTED = RGBColor(0x5B, 0x66, 0x75)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
CARD = RGBColor(0xFF, 0xFF, 0xFF)
TINT = RGBColor(0xE9, 0xF0, 0xFA)
GTINT = RGBColor(0xE6, 0xF5, 0xEC)
LINE = RGBColor(0xD5, 0xDC, 0xE6)
HEAD_FONT = "Mozaic GEO Bold"

# content area (inches)
L, T, R, B = 0.55, 1.35, 12.78, 6.35
W = R - L
H = B - T
CX = (L + R) / 2


# ---- helpers ----------------------------------------------------------------
def _runs(par, text, size, color, bold=False, italic=False, font=None):
    r = par.add_run()
    r.text = text
    f = r.font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = color
    if font or DEFAULT_FONT:
        f.name = font or DEFAULT_FONT
    return r


def tb(slide, name, x, y, w, h, paras, size=14, color=INK, bold=False,
       align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font=None, italic=False,
       space_after=4, margin=0.05, line_spacing=None):
    """paras: list of str or (str, dict) or list of runs [(text, dict), ...]."""
    s = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    s.name = name
    tf = s.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, Inches(margin))
    first = True
    for p in paras:
        par = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        par.alignment = align
        par.space_after = Pt(space_after)
        if line_spacing:
            par.line_spacing = line_spacing
        if isinstance(p, str):
            _runs(par, p, size, color, bold, italic, font)
        elif isinstance(p, tuple):
            o = dict(size=size, color=color, bold=bold, italic=italic, font=font)
            o.update(p[1])
            _runs(par, p[0], **o)
        else:  # list of runs
            for text, opts in p:
                o = dict(size=size, color=color, bold=bold, italic=italic, font=font)
                o.update(opts)
                _runs(par, text, **o)
    return s


def bullets(slide, name, x, y, w, h, items, size=14, color=INK, gap=6,
            dot=BLUE, title=None, title_color=BLUE, title_size=15):
    """Simple list with coloured dot markers (drawn as shapes)."""
    cy = y
    if title:
        tb(slide, name + " title", x, cy, w, 0.35, [title], size=title_size,
           color=title_color, bold=True)
        cy += 0.42
    line_h = size * 1.25 / 72
    for i, it in enumerate(items):
        # estimate lines for height
        chars_per_line = max(10, int((w - 0.3) * 72 / (size * 0.52)))
        n, cur = 1, 0
        for word in it.split():
            if cur and cur + 1 + len(word) > chars_per_line:
                n += 1; cur = len(word)
            else:
                cur += (1 if cur else 0) + len(word)
        hh = n * line_h + 0.06
        d = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.05),
                                   Inches(cy + line_h / 2 - 0.05), Inches(0.1), Inches(0.1))
        d.name = f"{name} dot {i+1}"
        d.fill.solid(); d.fill.fore_color.rgb = dot; d.line.fill.background()
        tb(slide, f"{name} item {i+1}", x + 0.25, cy, w - 0.25, hh, [it], size=size,
           color=color, margin=0.02)
        cy += hh + gap / 72
    return cy


def card(slide, name, x, y, w, h, fill=CARD, line=LINE, radius=0.06):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y),
                               Inches(w), Inches(h))
    s.name = name
    s.adjustments[0] = min(0.5, radius / min(w, h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line; s.line.width = Pt(0.75)
    s.shadow.inherit = False
    s.text_frame.text = ""
    return s


def circle(slide, name, cx, cy, d, fill=BLUE, text=None, size=14, color=WHITE):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - d / 2), Inches(cy - d / 2),
                               Inches(d), Inches(d))
    s.name = name
    s.fill.solid(); s.fill.fore_color.rgb = fill; s.line.fill.background()
    s.shadow.inherit = False
    tf = s.text_frame
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    if text is not None:
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        _runs(p, text, size, color, bold=True)
    return s


def pill(slide, name, cx, cy, w, h, text, fill=BLUE, color=WHITE, size=12):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx - w / 2),
                               Inches(cy - h / 2), Inches(w), Inches(h))
    s.name = name
    s.adjustments[0] = 0.5
    s.fill.solid(); s.fill.fore_color.rgb = fill; s.line.fill.background()
    s.shadow.inherit = False
    tf = s.text_frame
    tf.word_wrap = True
    for m in ("margin_left", "margin_right"):
        setattr(tf, m, Inches(0.1))
    for m in ("margin_top", "margin_bottom"):
        setattr(tf, m, 0)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    _runs(p, text, size, color, bold=True)
    return s


def pic(slide, name, fname, x, y, w=None, h=None):
    kw = {}
    if w: kw["width"] = Inches(w)
    if h: kw["height"] = Inches(h)
    import os
    path = fname if os.path.isabs(fname) else f"{MEDIA}/{fname}"
    p = slide.shapes.add_picture(path, Inches(x), Inches(y), **kw)
    p.name = name
    return p


def set_header(slide, title, size=24):
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip() == "TITLE":
            r = sh.text_frame.paragraphs[0].runs[0]
            r.text = title
            r.font.size = Pt(size)
            tf = sh.text_frame
            tf.word_wrap = True
            tf.auto_size = MSO_AUTO_SIZE.NONE
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            sh.top = Inches(0.18); sh.height = Inches(0.7)
            sh.width = Inches(7.5)
            return
    raise RuntimeError("header not found")


def fill_disclaimer(slide, tr_name, en_name):
    for sh in slide.shapes:
        if sh.has_text_frame and "[________]" in sh.text_frame.text:
            for par in sh.text_frame.paragraphs:
                for r in par.runs:
                    if "[________]" in r.text:
                        txt = par.runs[0].text + "".join(x.text for x in par.runs[1:])
                        is_en = txt.lstrip().startswith("[________] is")
                        r.text = r.text.replace("[________]", en_name if is_en else tr_name)
            return
    raise RuntimeError("disclaimer not found")


def subtitle(slide, text, color=GREEN, size=18, w=None):
    """Green sub-heading at the top of the content area; returns y below it."""
    tb(slide, "Subheading", L, T - 0.05, w or W, 0.45, [text], size=size, color=color,
       bold=True, font=HEAD_FONT)
    return T + 0.55


