"""A5 project card (148 x 210 mm) for the Climate İMECE project: page 1 English, page 2 Turkish.

Usage: python card.py <band.png> <logo_dir> <png_dir (illustrations + qr.png)> <out.pptx>
"""
import os
import sys
from pptx import Presentation
from pptx.util import Inches, Mm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402
from helpers import tb, card, pic, WHITE, INK, GREEN, MUTED  # noqa: E402

BAND, LOGOS, PNG, OUT = sys.argv[1:5]
QR_TEXT = os.environ.get("DS_QR_TEXT", "doi.org/10.32329/uad.788641")
helpers.DEFAULT_FONT = "Arial"
BLUE = RGBColor(0x12, 0x4D, 0xA2)
TINT = RGBColor(0xF2, 0xF6, 0xFA)

PW, PH = 148 / 25.4, 210 / 25.4       # 5.83 x 8.27 in
M = 0.42                               # side margin
W = PW - 2 * M

TXT = dict(
    en=dict(
        band="COP31 TÜRKİYE  ·  GREEN ZONE  ·  ANTALYA, 19 NOVEMBER 2026",
        title="CLIMATE İMECE",
        sub="From Climate Action to Implementation with Young People, Educational Institutions and Local Communities",
        summary=("Climate İMECE turns young people from passive learners into active partners in climate solutions. "
                 "In small, voluntary İMECE Circles, students combine local experience with academic knowledge and work "
                 "through the Plan–Do–Check–Act cycle. A pilot at Istanbul Kent University connects young people in "
                 "Türkiye, Algeria and India through online exchange. The model is designed so that other institutions "
                 "and communities can adopt it."),
        out_h="EXPECTED OUTPUTS (PROJECT TARGETS)",
        outs=["A Climate İMECE Youth Model", "At least three youth-led solutions",
              "A practical implementation guide", "A network across three countries"],
        team_h="PROJECT TEAM",
        team=[("Prof. Dr. Hayal Köksal", " (Project Director)"), "Dr. Memet Aslan  ·  Aleyna Çavdar  ·  Azad Kılıçarslan  ·  Ayça Selin Akyol"],
        partners="Partners: Kasdi Merbah University (Algeria)  ·  WCTQEE network schools (India)",
        contact_h="CONTACT",
        contact=["Prof. Dr. Hayal Köksal", "Istanbul Kent University, Türkiye", "hayal.koksal@kent.edu.tr"],
        qr="Scan to read the İMECE Circles method paper:",
        motto="Compete with Yourself, Collaborate with Others.",
    ),
    tr=dict(
        band="COP31 TÜRKİYE  ·  YEŞİL BÖLGE  ·  ANTALYA, 19 KASIM 2026",
        title="İKLİM İMECESİ",
        sub="Gençler, Eğitim Kurumları ve Yerel Topluluklarla İklim Eyleminden Uygulamaya",
        summary=("İklim İmecesi, gençleri pasif öğrenicilerden iklim çözümlerinin aktif ortaklarına dönüştürmeyi amaçlar. "
                 "Küçük ve gönüllü İMECE Halkaları'nda öğrenciler yerel deneyimi akademik bilgiyle birleştirir ve "
                 "Planla–Uygula–Kontrol Et–Önlem Al döngüsüyle çalışır. İstanbul Kent Üniversitesi'ndeki pilot uygulama; "
                 "Türkiye, Cezayir ve Hindistan'daki gençleri çevrim içi paylaşımlarla buluşturur. Model, başka kurum ve "
                 "toplulukların da uyarlayabileceği biçimde tasarlanmıştır."),
        out_h="BEKLENEN ÇIKTILAR (PROJE HEDEFLERİ)",
        outs=["İklim İmecesi Gençlik Modeli", "Gençlerden en az üç iklim çözümü",
              "Uygulamaya dönük bir rehber", "Üç ülkeyi kapsayan bir öğrenme ağı"],
        team_h="PROJE EKİBİ",
        team=[("Prof. Dr. Hayal Köksal", " (Proje Yürütücüsü)"), "Dr. Memet Aslan  ·  Aleyna Çavdar  ·  Azad Kılıçarslan  ·  Ayça Selin Akyol"],
        partners="Ortaklar: Kasdi Merbah Üniversitesi (Cezayir)  ·  WCTQEE ağı okulları (Hindistan)",
        contact_h="İLETİŞİM",
        contact=["Prof. Dr. Hayal Köksal", "İstanbul Kent Üniversitesi, Türkiye", "hayal.koksal@kent.edu.tr"],
        qr="İMECE Halkaları yöntem makalesi için okutun:",
        motto="Kendinle Yarış, Başkalarıyla İş Birliği Yap.",
    ),
)


def page(prs, c):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    # event band (full bleed)
    b = s.shapes.add_picture(BAND, 0, 0, width=prs.slide_width)
    b.name = "COP31 band"
    bh = b.height / 914400
    tb(s, "Event line", M, 0, W, bh, [c["band"]], size=8.5, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, margin=0)
    # identity row
    y = bh + 0.3
    logo = pic(s, "Istanbul Kent University logo", f"{LOGOS}/kent_logo_lacivert_kirmizi_crop.png", M, y, h=0.72)
    hero = 1.75
    pic(s, "İMECE circle illustration", f"{PNG}/imece_circle.png", PW - M - hero, y - 0.08, w=hero, h=hero)
    tw = W - hero - 0.1
    tb(s, "Title", M, y + 0.92, tw, 0.45, [c["title"]], size=24, bold=True, color=BLUE, margin=0)
    tb(s, "Subtitle", M, y + 1.36, tw, 0.6, [c["sub"]], size=10, bold=True, color=GREEN, margin=0, line_spacing=1.05)
    # summary
    y = 3.22
    tb(s, "Summary", M, y, W, 1.1, [c["summary"]], size=10, color=INK, margin=0, line_spacing=1.12)
    # expected outputs
    y = 4.5
    tb(s, "Outputs heading", M, y, W, 0.22, [c["out_h"]], size=9, bold=True, color=BLUE, margin=0)
    cw = (W - 0.2) / 2
    for i, it in enumerate(c["outs"]):
        ox, oy = M + (i % 2) * (cw + 0.2), y + 0.3 + (i // 2) * 0.33
        pic(s, f"Check {i+1}", f"{PNG}/icon_check.png", ox, oy + 0.02, w=0.2, h=0.2)
        tb(s, f"Output {i+1}", ox + 0.28, oy, cw - 0.28, 0.25, [it], size=9, color=INK, anchor=MSO_ANCHOR.MIDDLE, margin=0)
    # team
    y = 5.55
    tb(s, "Team heading", M, y, W, 0.22, [c["team_h"]], size=9, bold=True, color=BLUE, margin=0)
    tb(s, "Team", M, y + 0.26, W, 0.36, [[(c["team"][0][0], dict(bold=True)), (c["team"][0][1], {})], c["team"][1]], size=9, color=INK, margin=0, space_after=2)
    tb(s, "Partners", M, y + 0.66, W, 0.2, [c["partners"]], size=8.5, color=MUTED, margin=0)
    # contact + QR
    y = 6.5
    ch = 1.15
    card(s, "Contact card", M, y, W, ch, fill=TINT, line=None, radius=0.1)
    q = 0.98
    pic(s, "QR code", f"{PNG}/qr.png", M + 0.12, y + (ch - q) / 2, w=q, h=q)
    tx = M + 0.12 + q + 0.2
    tb(s, "Contact", tx, y + 0.1, W - (tx - M) - 0.1, ch - 0.2,
       [(c["contact_h"], dict(bold=True, color=BLUE, size=8.5)),
        (c["contact"][0], dict(bold=True, size=10.5)),
        c["contact"][1], c["contact"][2],
        [(c["qr"] + " ", dict(size=8, color=MUTED, italic=True)), (QR_TEXT, dict(size=8, color=BLUE, bold=True))]],
       size=9.5, color=INK, margin=0, space_after=2, anchor=MSO_ANCHOR.MIDDLE)
    tb(s, "Motto", M, PH - 0.45, W, 0.25, [c["motto"]], size=9, bold=True, italic=True, color=GREEN,
       align=PP_ALIGN.CENTER, margin=0)
    print("bottom of contact card:", round(y + ch, 2), "page height", round(PH, 2))


prs = Presentation()
prs.slide_width, prs.slide_height = Mm(148), Mm(210)
page(prs, TXT["en"])
page(prs, TXT["tr"])
prs.save(OUT)
print("wrote", OUT)
