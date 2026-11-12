"""Climate İMECE – 6-slide digital-screen deck (COP31 template) from one slide spec.

    python ds_build.py html   -> scene pages (ds/html/slideN.html) + illustration pages + QR
    python ds_build.py pptx   -> .pptx with fade-in builds, 13 s auto-advance and looping
    python ds_build.py words  -> visible word count per slide

Paths are taken from the environment (see CONFIG) so the script can run from the repo.
"""
import json
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ds_illust as I  # noqa: E402

# ---- CONFIG ---------------------------------------------------------------------
WORK = os.environ.get("DS_WORK", os.getcwd())                     # working dir (html/png/frames)
TPL = os.environ.get("DS_TPL", f"{WORK}/../tpl")                  # unpacked presentation template
BASE = os.environ.get("DS_BASE", f"{WORK}/../base6.pptx")         # cover + 4 content + closing
LOGOS = os.environ.get("DS_LOGOS", f"{WORK}/../logos")
COP31_LOGO = os.environ.get("DS_COP31", f"{WORK}/../ptpl/ppt/media/image3.png")
OUT = os.environ.get("DS_OUT", f"{WORK}/COP31_Climate_IMECE_Digital_Screen_EN.pptx")
QR_URL = os.environ.get("DS_QR_URL", "https://doi.org/10.32329/uad.788641")
SLIDE_SECONDS = 13

UNI_TR, UNI_EN, PI = "İstanbul Kent Üniversitesi", "Istanbul Kent University", "Prof. Dr. Hayal Köksal"
DISC_TR = ("Bu sunumda yer alan görüş, bilgi ve ifadelerden, sunumun hazırlanmasından ve içeriğinden yalnızca "
           "{who} sorumlu olup, sunum Yükseköğretim Kurulu’nun (YÖK) görüşlerini yansıtmaz ve Kurula herhangi "
           "bir sorumluluk yüklemez.")
DISC_EN = ("{who} is solely responsible for the views, information and statements expressed in this presentation, "
           "as well as for its preparation and content. The presentation does not necessarily reflect the views of "
           "the Council of Higher Education of Türkiye (YÖK), nor does it imply any responsibility or liability on its part.")
WHO_TR, WHO_EN = f"{UNI_TR} – {PI}", f"{UNI_EN} – {PI}"

BLUE, GREEN, LBLUE, RED = "124DA2", "44B97E", "1B75BB", "E41E26"
INK, MUTED, WHITE, LINE = "1F2937", "5B6675", "FFFFFF", "D5DCE6"
HEAD_FONT = "Mozaic GEO Bold"
PX = 144  # px per inch at 1920 px width


# ---- spec helpers -----------------------------------------------------------------
def R(text, sz, c=INK, b=False, head=False, i=False):
    return dict(text=text, sz=sz, c=c, b=b or head, head=head, i=i)


def T(x, y, w, h, paras, t, align="l", anchor="t", tpl=None, gap=4):
    return dict(k="text", x=x, y=y, w=w, h=h, paras=paras, t=t, align=align, anchor=anchor, tpl=tpl, gap=gap)


def C(x, y, w, h, paras, t, pad=(0.25, 0.2, 0.25, 0.2), align="l", anchor="m", radius=0.14, gap=4):
    return dict(k="card", x=x, y=y, w=w, h=h, paras=paras, t=t, pad=pad, align=align, anchor=anchor,
                radius=radius, gap=gap)


def S(name, x, y, w, h, t, words=0):
    return dict(k="svg", name=name, x=x, y=y, w=w, h=h, t=t, words=words)


def P(path, x, y, w, h, t, name="Picture"):
    return dict(k="pic", path=path, x=x, y=y, w=w, h=h, t=t, name=name)


KENT = f"{LOGOS}/kent_logo_lacivert_kirmizi_crop.png"
QR = f"{WORK}/png/qr.png"

X0, CW4, G4 = 0.55, 2.87, 0.25
TILES = [("1", ["Climate İMECE", "model"], BLUE, "icon_model"), ("3+", ["student climate", "solutions piloted"], GREEN, "icon_bulb"),
         ("1", ["how-to guide", "for schools"], LBLUE, "icon_guide"), ("3", ["countries", "learning together"], RED, "icon_network")]
STEPS = [("Form a volunteer circle", "icon_model"), ("Pick a local problem", "icon_pin"),
         ("Pilot a solution", "icon_bulb"), ("Measure, improve, share", "icon_chart")]
IMPACT = ["Real pilots where students live", "Solutions shared across three countries", "Any school can adopt it"]
RX, RW, RY, RH, RP = 5.75, 6.95, 2.2, 0.9, 1.02   # step rows on the method slide

SLIDES = [
    dict(layout="cover", elems=[
        T(1.14, 2.72, 6.9, 0.85, [[R("CLIMATE İMECE", 44, BLUE, head=True)]], 0.4, tpl="TITLE"),
        T(1.14, 3.52, 6.9, 1.15, [[R("From Climate Action to Implementation with Young People, Educational "
                                     "Institutions and Local Communities", 20, GREEN, head=True)]], 1.3, tpl="Subtitle"),
        P(KENT, 1.14, 4.92, 2.37, 1.0, 2.3, name="Istanbul Kent University logo"),
        T(3.8, 4.92, 4.2, 1.0, [[R("Prof. Dr. Hayal Köksal", 17, BLUE, head=True)],
                                [R("Dr. Memet Aslan · Aleyna Çavdar", 12.5)],
                                [R("Azad Kılıçarslan · Ayça Selin Akyol", 12.5)]], 2.9, anchor="m", gap=2),
    ]),
    dict(layout="content", title="PROBLEM AND AIM", elems=[
        T(0.6, 1.5, 5.9, 0.35, [[R("PROBLEM", 14, GREEN, head=True)]], 0.6),
        T(0.6, 1.85, 5.9, 1.65, [[R("Students learn about climate change but rarely get to act.", 30, INK, head=True)]], 0.9),
        S("awareness_action", 6.75, 1.95, 5.9, 3.13, 1.4),
        T(0.6, 3.8, 5.9, 0.35, [[R("AIM", 14, GREEN, head=True)]], 5.4),
        T(0.6, 4.15, 5.9, 2.0, [[R("Student teams solve real climate problems in their own communities.", 30, BLUE, head=True)]], 5.7),
    ]),
    dict(layout="content", title="METHOD", elems=[
        T(0.6, 1.36, 12.1, 0.6, [[R("İMECE Circles: ", 23, BLUE, head=True), R("a Turkish teamwork method used in ", 23, INK, head=True),
                                  R("5,000+ projects", 23, GREEN, head=True)]], 0.6, anchor="m"),
        S("imece_steps", 0.85, 2.05, 4.25, 4.25, 1.2),
    ] + [e for i, (txt, ic) in enumerate(STEPS) for e in (
        C(RX, RY + i * RP, RW, RH, [[R(txt, 23, INK, head=True)]], 3.0 + i * 1.2, pad=(1.2, 0.1, 1.2, 0.1)),
        S(f"badge{i + 1}", RX + 0.22, RY + i * RP + 0.15, 0.6, 0.6, 3.2 + i * 1.2),
        S(ic, RX + RW - 1.02, RY + i * RP + 0.07, 0.76, 0.76, 3.4 + i * 1.2),
    )]),
    dict(layout="content", title="EXPECTED OUTPUTS", elems=[e for i, (num, cap, col, ic) in enumerate(TILES) for e in (
        C(X0 + i * (CW4 + G4), 1.5, CW4, 3.75, [[R(num, 46, col, head=True)], [R(cap[0], 17)], [R(cap[1], 17)]], 0.6 + i * 1.1,
          pad=(0.15, 0.2, 0.15, 0.28), align="c", anchor="b", gap=0),
        S(ic, X0 + i * (CW4 + G4) + (CW4 - 1.85) / 2, 1.68, 1.85, 1.85, 0.9 + i * 1.1),
    )] + [
        T(0.55, 5.55, 12.23, 0.4, [[R("Targets, not results yet", 15, MUTED, i=True)]], 5.4, align="c"),
    ]),
    dict(layout="content", title="CLIMATE ACTION AND IMPACT", elems=[
        S("globe", 6.75, 1.22, 5.8, 4.97, 0.8, words=3),
    ] + [e for i, txt in enumerate(IMPACT) for e in (
        S("icon_check", 0.6, 1.95 + i * 1.3, 0.7, 0.7, 3.6 + i * 1.3),
        T(1.55, 1.75 + i * 1.3, 5.1, 1.1, [[R(txt, 25, BLUE if i == 0 else INK, head=True)]], 3.7 + i * 1.3, anchor="m"),
    )]),
    dict(layout="closing", elems=[
        T(1.14, 2.72, 6.9, 0.85, [[R("THANK YOU", 44, BLUE, head=True)]], 0.4, tpl="THANK YOU"),
        T(1.14, 3.6, 6.9, 0.5, [[R("Compete with Yourself, Collaborate with Others.", 19, GREEN, head=True)]], 1.3),
        P(QR, 1.14, 4.4, 1.75, 1.75, 2.2, name="QR code"),
        T(3.15, 4.4, 4.8, 1.75, [[R("Prof. Dr. Hayal Köksal", 17, BLUE, head=True)],
                                 [R("Istanbul Kent University, Türkiye", 13)],
                                 [R("hayal.koksal@kent.edu.tr", 13)],
                                 [R("Scan to read about İMECE Circles", 12, MUTED, i=True)]], 2.8, anchor="m", gap=3),
    ]),
]


def words():
    for n, sl in enumerate(SLIDES, 1):
        txt = [sl.get("title", "")]
        extra = 0
        for e in sl["elems"]:
            if e["k"] in ("text", "card"):
                txt += [r["text"] for p in e["paras"] for r in p]
            extra += e.get("words", 0)
        toks = [w for w in re.split(r"[\s·]+", " ".join(txt)) if re.search(r"\w", w)]
        print(f"slide {n}: {len(toks) + extra} words")


# ---- HTML scenes --------------------------------------------------------------------
PAGE_CSS = """
html,body{margin:0;background:transparent}
body{width:1920px;height:1080px;position:relative;overflow:hidden;font-family:'Inter',sans-serif;color:#000}
.el{position:absolute;box-sizing:border-box;display:flex;flex-direction:column}
.el p{margin:0;line-height:1.2}
.rise{animation:rise .7s cubic-bezier(.2,.8,.2,1) both}
.fadeel{animation:fade .8s ease both}
@keyframes rise{from{opacity:0;transform:translateY(20px)}to{opacity:1;transform:none}}
.card{background:#fff;border:1.5px solid #D5DCE6;box-shadow:0 6px 18px rgba(18,77,162,.08)}
"""


def px(v):
    return f"{v * PX:.1f}px"


def runs_html(paras, gap):
    out = []
    for k, p in enumerate(paras):
        spans = "".join(
            f'<span style="font-size:{r["sz"] * 2:.0f}px;color:#{r["c"]};font-weight:{800 if r["head"] else (700 if r["b"] else 400)};'
            f'font-style:{"italic" if r["i"] else "normal"}">{r["text"]}</span>' for r in p)
        mb = gap * 2 if k < len(paras) - 1 else 0
        out.append(f'<p style="margin-bottom:{mb}px">{spans}</p>')
    return "".join(out)


JUST = {"t": "flex-start", "m": "center", "b": "flex-end"}
ALIGN = {"l": "left", "c": "center", "r": "right"}


def chrome_html(sl):
    band = f"{TPL}/ppt/media/image3.png"
    yok = f"{TPL}/ppt/media/image5.png"
    disc = (f'<p style="font-size:16px;line-height:1.2;margin:0">{DISC_TR.format(who=WHO_TR)}</p>'
            f'<p style="font-size:16px;line-height:1.2;margin:0">{DISC_EN.format(who=WHO_EN)}</p>')
    if sl["layout"] == "content":
        return (
            f'<img src="file://{band}" style="position:absolute;left:{px(-0.005)};top:{px(-0.009)};width:{px(13.344)};height:{px(1.056)}">'
            f'<div class="el fadeel" style="left:{px(0.603)};top:{px(0.18)};width:{px(7.3)};height:{px(0.7)};justify-content:center;'
            f'animation-delay:.2s"><p style="font-size:48px;font-weight:800;color:#fff">{sl["title"]}</p></div>'
            f'<div class="el" style="left:{px(0.683)};top:{px(6.614)};width:{px(9.295)}">{disc}</div>'
            f'<img src="file://{COP31_LOGO}" style="position:absolute;left:{px(10.25)};top:{px(6.601)};width:{px(1.362)};height:{px(0.55)}">'
            f'<div style="position:absolute;left:{px(11.784)};top:{px(6.601)};width:3px;height:{px(0.55)};background:#124DA2"></div>'
            f'<img src="file://{yok}" style="position:absolute;left:{px(11.951)};top:{px(6.604)};width:{px(1.126)};height:{px(0.558)}">')
    return (
        f'<img src="file://{COP31_LOGO}" style="position:absolute;left:{px(1.139)};top:{px(1.047)};width:{px(3.285)};height:{px(1.326)}">'
        f'<div style="position:absolute;left:{px(4.837)};top:{px(1.047)};width:3px;height:{px(1.326)};background:#124DA2"></div>'
        f'<img src="file://{yok}" style="position:absolute;left:{px(5.192)};top:{px(1.08)};width:{px(2.543)};height:{px(1.261)}">'
        f'<div class="el" style="left:{px(0.683)};top:{px(6.614)};width:{px(11.97)}">{disc}</div>')


def elem_html(e):
    box = f'left:{px(e["x"])};top:{px(e["y"])};width:{px(e["w"])};height:{px(e["h"])}'
    d = f'animation-delay:{e["t"]:.2f}s'
    if e["k"] == "text":
        return (f'<div class="el rise" style="{box};justify-content:{JUST[e["anchor"]]};text-align:{ALIGN[e["align"]]};{d}">'
                f'{runs_html(e["paras"], e["gap"])}</div>')
    if e["k"] == "card":
        l, t, r, b = e["pad"]
        return (f'<div class="el card rise" style="{box};border-radius:{px(e["radius"])};padding:{px(t)} {px(r)} {px(b)} {px(l)};'
                f'justify-content:{JUST[e["anchor"]]};text-align:{ALIGN[e["align"]]};{d}">{runs_html(e["paras"], e["gap"])}</div>')
    if e["k"] == "svg":
        svg, _, _ = I.ILLUST[e["name"]](e["t"], False)
        return f'<div class="el" style="{box}">{svg}</div>'
    if e["k"] == "pic":
        return f'<img class="el fadeel" src="file://{e["path"]}" style="{box};{d}">'
    raise ValueError(e["k"])


def build_html():
    os.makedirs(f"{WORK}/html", exist_ok=True)
    os.makedirs(f"{WORK}/png", exist_ok=True)
    import segno
    segno.make(QR_URL, error="m").save(QR, scale=20, border=2, dark="#0B2A5B")
    for n, sl in enumerate(SLIDES, 1):
        body = chrome_html(sl) + "".join(elem_html(e) for e in sl["elems"])
        html = (f'<!doctype html><html><head><meta charset="utf-8"><style>{PAGE_CSS}{I.CSS}</style></head>'
                f'<body>{body}</body></html>')
        open(f"{WORK}/html/slide{n}.html", "w", encoding="utf8").write(html)
    jobs = []
    for name, f in I.ILLUST.items():
        svg, w, h = f(0.0, True)
        page = (f'<!doctype html><html><head><meta charset="utf-8"><style>html,body{{margin:0;background:transparent}}'
                f'{I.CSS}</style></head><body><div style="width:{w}px;height:{h}px">{svg}</div></body></html>')
        open(f"{WORK}/html/illust_{name}.html", "w", encoding="utf8").write(page)
        jobs.append(dict(html=f"{WORK}/html/illust_{name}.html", png=f"{WORK}/png/{name}.png", w=w, h=h, t=30000,
                         scale=3 if w >= 400 else 4))
    json.dump(jobs, open(f"{WORK}/illust_jobs.json", "w"))
    print("html written; QR ->", QR_URL)


# ---- PPTX ----------------------------------------------------------------------------
def build_pptx():
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE

    AL = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}
    AN = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}

    def fill_text(tf, e, margins):
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.vertical_anchor = AN[e["anchor"]]
        tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = [Inches(m) for m in margins]
        for k, p in enumerate(e["paras"]):
            par = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
            par.alignment = AL[e["align"]]
            par.space_after = Pt(e["gap"] if k < len(e["paras"]) - 1 else 0)
            for r in p:
                run = par.add_run()
                run.text = r["text"]
                f = run.font
                f.size = Pt(r["sz"]); f.bold = r["b"]; f.italic = r["i"]
                f.color.rgb = RGBColor.from_string(r["c"])
                if r["head"]:
                    f.name = HEAD_FONT

    prs = Presentation(BASE)
    anims = {}
    for n, (slide, sl) in enumerate(zip(prs.slides, SLIDES), 1):
        seq = []
        # disclaimer: institution + project lead in both languages
        for sh in slide.shapes:
            if sh.has_text_frame and "[________]" in sh.text_frame.text:
                for par in sh.text_frame.paragraphs:
                    is_en = par.runs and par.runs[0].text.lstrip().startswith("[________] is")
                    for r in par.runs:
                        if "[________]" in r.text:
                            r.text = r.text.replace("[________]", WHO_EN if is_en else WHO_TR)
        if sl["layout"] == "content":
            for sh in slide.shapes:
                if sh.has_text_frame and sh.text_frame.text.strip() == "TITLE":
                    sh.text_frame.paragraphs[0].runs[0].text = sl["title"]
                    tf = sh.text_frame
                    tf.word_wrap = True; tf.auto_size = MSO_AUTO_SIZE.NONE; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                    sh.top = Inches(0.18); sh.height = Inches(0.7); sh.width = Inches(7.5)
                    seq.append((sh.shape_id, 200, "text"))
        for e in sl["elems"]:
            box = [Inches(e[k]) for k in ("x", "y", "w", "h")]
            if e["k"] == "text" and e["tpl"]:
                sh = next(s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip() == e["tpl"])
                run = sh.text_frame.paragraphs[0].runs[0]
                run.text = e["paras"][0][0]["text"]
                run.font.size = Pt(e["paras"][0][0]["sz"])
                sh.text_frame.word_wrap = True
                sh.text_frame.auto_size = MSO_AUTO_SIZE.NONE
                sh.text_frame.margin_left = sh.text_frame.margin_right = 0
                sh.text_frame.margin_top = sh.text_frame.margin_bottom = 0
                sh.left, sh.top, sh.width, sh.height = box
                seq.append((sh.shape_id, int(e["t"] * 1000), "text"))
            elif e["k"] == "text":
                sh = slide.shapes.add_textbox(*box)
                sh.name = "Text: " + e["paras"][0][0]["text"][:30]
                fill_text(sh.text_frame, e, (0, 0, 0, 0))
                seq.append((sh.shape_id, int(e["t"] * 1000), "text"))
            elif e["k"] == "card":
                sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, *box)
                sh.name = "Card: " + e["paras"][-1][0]["text"][:30]
                sh.adjustments[0] = min(0.5, e["radius"] / min(e["w"], e["h"]))
                sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor.from_string(WHITE)
                sh.line.color.rgb = RGBColor.from_string(LINE); sh.line.width = Pt(0.75)
                sh.shadow.inherit = False
                fill_text(sh.text_frame, e, e["pad"])
                seq.append((sh.shape_id, int(e["t"] * 1000), "card"))
            elif e["k"] in ("svg", "pic"):
                path = e["path"] if e["k"] == "pic" else f"{WORK}/png/{e['name']}.png"
                sh = slide.shapes.add_picture(path, *box)
                sh.name = e.get("name", "Illustration")
                seq.append((sh.shape_id, int(e["t"] * 1000), "pic"))
        anims[str(slide.part.partname).lstrip("/")] = seq
    prs.save(OUT)
    inject(OUT, anims)
    print("wrote", OUT)


def effect(i, spid, delay, kind):
    grp = "" if kind == "pic" else ' grpId="0"'
    return (f'<p:par><p:cTn id="{i}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold"{grp} nodeType="withEffect">'
            f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>'
            f'<p:set><p:cBhvr><p:cTn id="{i + 1}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
            f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
            f'</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
            f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{i + 2}" dur="700"/>'
            f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>')


def timing_xml(seq):
    effs, i = [], 5
    for spid, delay, kind in seq:
        effs.append(effect(i, spid, delay, kind)); i += 3
    bld = "".join(f'<p:bldP spid="{s}" grpId="0"{" animBg=\"1\"" if k == "card" else ""}/>' for s, _, k in seq if k != "pic")
    return ('<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
            '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
            '<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>'
            '<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>'
            '<p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
            + "".join(effs) +
            '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>'
            '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
            '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
            '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>'
            f'<p:bldLst>{bld}</p:bldLst></p:timing>')


def inject(path, anims):
    """Add auto-advance transitions, fade-in builds and looping to the saved package."""
    zin = zipfile.ZipFile(path)
    items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    zin.close()
    trans = f'<p:transition spd="med" advTm="{SLIDE_SECONDS * 1000}"><p:fade/></p:transition>'
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, data in items:
            name = info.filename
            if name in anims:
                x = data.decode("utf8")
                x = re.sub(r"<p:transition.*?</p:transition>|<p:transition[^>]*/>", "", x, flags=re.S)
                x = re.sub(r"<p:timing>.*?</p:timing>", "", x, flags=re.S)
                assert "</p:clrMapOvr>" in x, name
                x = x.replace("</p:clrMapOvr>", "</p:clrMapOvr>" + trans + timing_xml(anims[name]), 1)
                data = x.encode("utf8")
            elif name == "ppt/presProps.xml":
                x = data.decode("utf8")
                if "<p:showPr" not in x:
                    x = x.replace("<p:clrMru>", '<p:showPr loop="1" showNarration="1"><p:present/><p:sldAll/></p:showPr><p:clrMru>', 1)
                data = x.encode("utf8")
            zout.writestr(info, data)


if __name__ == "__main__":
    cmd = sys.argv[1]
    {"html": build_html, "pptx": build_pptx, "words": words}[cmd]()
