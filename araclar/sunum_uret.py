"""Build COP31-template versions (EN / TR) of Prof. Hayal Köksal's Climate İMECE deck.

Usage: python build.py <base.pptx> <media_dir> <out_dir>
"""
import math
import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

BASE, MEDIA, OUT, LOGOS = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]

from helpers import *  # noqa
from helpers import _runs
import helpers
MEDIA = sys.argv[2]
helpers.set_media(MEDIA)

# ---- content ----------------------------------------------------------------
C = {}
C["en"] = dict(
    inst="Istanbul Kent University",
    title="CLIMATE İMECE",
    sub="From Climate Action to Implementation with Young People, Educational Institutions and Local Communities",
    presenter=["Prof. Dr. Hayal KÖKSAL",
               "Istanbul Kent University · Educational Scientist, Project Designer",
               "hayal.koksal@kent.edu.tr"],
    event="COP31 Türkiye · YÖK Green Zone · Antalya, 19 November 2026",
    s2=dict(h="INTRODUCTION", big="CLIMATE İMECE: FROM CLIMATE ACTION TO IMPLEMENTATION",
            tag="A human-centred, collaborative and participatory model for climate action",
            team_h="RESEARCH TEAM · Istanbul Kent University",
            team=[("Prof. Dr. Hayal Köksal", "Project Director"),
                  ("Res. Asst. Dr. Memet Aslan", "Researcher"),
                  ("Res. Asst. Aleyna Çavdar", "Researcher"),
                  ("Res. Asst. Azad Kılıçarslan", "Researcher"),
                  ("Res. Asst. Ayça Selin Akyol", "Researcher")]),
    s3=dict(h="OUR CENTRAL QUESTION",
            q="How can we move young people from climate awareness to meaningful climate action?",
            a="Young people should become active partners in identifying problems, developing solutions and implementing change."),
    s4=dict(h="THE MODEL", lead="A model that brings together",
            who=["Young people", "Educational institutions", "Universities", "Local communities"],
            how="through collaboration and collective action."),
    s5=dict(h="WHY CLIMATE İMECE?", items=[
        "Climate change is a multidimensional global challenge.",
        "Information and awareness alone are not sufficient.",
        "Local knowledge and academic knowledge need to meet.",
        "Young people need opportunities to participate in real solutions.",
        "Sustainable change requires collective responsibility."]),
    s6=dict(h="WHAT IS İMECE?", tag="A Turkish tradition of collective effort and action",
            defn="İMECE is a traditional Turkish concept referring to people coming together voluntarily to achieve a shared goal.",
            rep="In our educational model, İMECE represents:",
            vals=["Belonging", "Participation", "Collaboration", "Collective Learning",
                  "Collective Intelligence", "Collective Action", "Quality and Social Impact"]),
    s7=dict(h="THE POWER OF İMECE", tag="From individual effort to collective intelligence",
            cap1="Young people identifying real problems together",
            cap2="Learning and acting as a community"),
    s8=dict(h="WHAT IS AN İMECE CIRCLE?",
            tag="A small, voluntary and collaborative learning community",
            ring=["REAL PROBLEM", "PEOPLE AFFECTED", "SHARED GOAL", "COLLABORATION",
                  "ACTION", "REFLECTION", "CONTINUOUS IMPROVEMENT"]),
    s9=dict(h="OUR GUIDING PRINCIPLE", l1="Compete with Yourself,", l2="Collaborate with Others."),
    s10=dict(h="THE PROJECT'S ORIGINAL CONTRIBUTION", tag="From awareness to action",
             lead="Climate İMECE brings together:",
             items=["Young people's local and experiential knowledge",
                    "Academic and scientific knowledge",
                    "Educational institutions",
                    "Local communities and stakeholders",
                    "International learning and cooperation networks"],
             inn_h="THE INNOVATION",
             inn="Transforming young people from passive learners into active climate solution partners."),
    s11=dict(h="PARTNERS AND INTERNATIONAL COOPERATION", tag="A multi-country learning network",
             c1h="LEAD INSTITUTION", c1=["Istanbul Kent University, Türkiye"],
             c2h="INTERNATIONAL PARTNERS",
             c2=["Kasdi Merbah University, Algeria",
                 "World Council for Total Quality and Excellence in Education (WCTQEE)",
                 "Little Angels High School, Gwalior, India",
                 "Suncity School, Gurgaon, India",
                 "Bhavan Vidyalaya, Panchkula, India",
                 "Gyan Heritage School, Lucknow, India",
                 "& more…"],
             c3h="PROJECT RESEARCH TEAM",
             c3=["Prof. Dr. Hayal Köksal – Project Director",
                 "Dr. Memet Aslan – Researcher",
                 "Aleyna Çavdar – Researcher",
                 "Azad Kılıçarslan – Researcher",
                 "Ayça Selin Akyol – Researcher"],
             countries=["Türkiye", "Algeria", "India"]),
    s12=dict(h="MAIN AIM AND OBJECTIVES", tag="Our main aim",
             aim="To develop a practical Climate İMECE model that enables young people to:",
             steps=[("Identify", "climate challenges in their own environments"),
                    ("Combine", "local and academic knowledge"),
                    ("Develop", "collaborative and sustainable solutions"),
                    ("Implement", "small-scale pilot actions"),
                    ("Evaluate", "and share their results")]),
    s13=dict(h="MAIN AIM AND OBJECTIVES", tag="Specific objectives", items=[
        "Establish a pilot Climate İMECE application at Istanbul Kent University.",
        "Create collaborative İMECE Circles.",
        "Develop at least three applicable climate solutions.",
        "Encourage international learning between Türkiye, Algeria and India.",
        "Prepare a practical implementation guide.",
        "Evaluate the model's transferability to other institutions and communities."]),
    s14=dict(h="METHODOLOGY: İMECE CIRCLES AND PDCA",
             tag="A participatory and action-oriented research model",
             lead="The project integrates İMECE Circles with the PDCA mindset, the quality improvement cycle:",
             pdca=[("PLAN", "Identify problems and plan solutions."),
                   ("DO", "Implement selected actions."),
                   ("CHECK", "Monitor and evaluate results."),
                   ("ACT", "Improve, adapt and share successful practices.")]),
    s15=dict(h="THE RESEARCH PROCESS", tag="Four work packages (WPs)", wps=[
        ("WP1", "Preparation", ["Literature review", "Research instruments", "Participant identification"]),
        ("WP2", "İMECE Circles", ["Identifying local climate challenges", "Defining common problem areas", "Collaborative discussions"]),
        ("WP3", "International Collaborative Learning", ["Online meetings", "Sharing local experiences", "Comparing climate challenges"]),
        ("WP4", "Joint Solutions and Pilot Actions", ["Developing solutions", "Implementing selected actions", "Evaluating results"])]),
    s16=dict(h="RESEARCH AND EVALUATION", tag="How will we measure progress?", cols=[
        ("Three main sources of evidence", ["Initial and final assessment forms, including peer and jury assessment", "İMECE process records", "Participant reflections and evaluations"]),
        ("Data analysis", ["Frequencies, percentages and averages", "Comparison of initial and final assessments", "Thematic analysis of qualitative data"]),
        ("Key areas of evaluation", ["Climate awareness", "Participation in climate action", "Collaboration", "Collective problem-solving", "Confidence in implementing solutions"])]),
    s17=dict(h="EXPECTED OUTPUTS AND IMPACT", tag="From local solutions to international learning", items=[
        "A Climate İMECE Youth Model",
        "At least three climate solutions developed by young people",
        "Pilot implementation activities",
        "A practical Climate İMECE Implementation Guide",
        "An International Youth Climate Learning Network",
        "A digital archive documenting experiences and solutions",
        "At least one scientific paper or conference presentation"]),
    s18=dict(h="FROM LOCAL ACTION TO GLOBAL CONTRIBUTION", tag="Our vision", items=[
        "Empowering young people",
        "Connecting local and scientific knowledge",
        "Encouraging international cooperation",
        "Supporting sustainable educational practices",
        "Building a culture of shared responsibility"],
        quote="Climate action begins locally, but its impact can reach globally."),
    s19=dict(h="CONCLUSION", tag="Climate İMECE: learning together, acting together",
             lead="We believe that:", items=[
        "Young people are not merely the future; they are partners in today's climate action.",
        "Education should lead to meaningful action.",
        "Collaboration transforms individual knowledge into collective intelligence.",
        "Quality improvement makes action measurable and sustainable.",
        "İMECE creates opportunities for people to work together for a better future."]),
    s20=dict(thanks="THANK YOU",
             text="Today, I am presenting this research project on behalf of Istanbul Kent University and our international project partners.",
             motto="Compete with Yourself, Collaborate with Others.",
             sig=["Prof. Dr. Hayal Köksal", "Istanbul Kent University, Türkiye", "hayal.koksal@kent.edu.tr"]),
)

C["tr"] = dict(
    inst="İstanbul Kent Üniversitesi",
    title="İKLİM İMECESİ",
    sub="Gençler, Eğitim Kurumları ve Yerel Topluluklarla İklim Eyleminden Uygulamaya",
    presenter=["Prof. Dr. Hayal KÖKSAL",
               "İstanbul Kent Üniversitesi · Eğitim Bilimci, Proje Tasarımcısı",
               "hayal.koksal@kent.edu.tr"],
    event="COP31 Türkiye · YÖK Yeşil Bölge · Antalya, 19 Kasım 2026",
    s2=dict(h="GİRİŞ", big="İKLİM İMECESİ: İKLİM EYLEMİNDEN UYGULAMAYA",
            tag="İklim eylemi için insan odaklı, iş birliğine dayalı ve katılımcı bir model",
            team_h="ARAŞTIRMA EKİBİ · İstanbul Kent Üniversitesi",
            team=[("Prof. Dr. Hayal Köksal", "Proje Yürütücüsü"),
                  ("Arş. Gör. Dr. Memet Aslan", "Araştırmacı"),
                  ("Arş. Gör. Aleyna Çavdar", "Araştırmacı"),
                  ("Arş. Gör. Azad Kılıçarslan", "Araştırmacı"),
                  ("Arş. Gör. Ayça Selin Akyol", "Araştırmacı")]),
    s3=dict(h="TEMEL SORUMUZ",
            q="Gençleri iklim farkındalığından anlamlı iklim eylemine nasıl taşıyabiliriz?",
            a="Gençler; sorunları belirleme, çözüm geliştirme ve değişimi uygulama süreçlerinin aktif ortakları olmalıdır."),
    s4=dict(h="MODEL", lead="Bir araya getiren bir model:",
            who=["Gençler", "Eğitim kurumları", "Üniversiteler", "Yerel topluluklar"],
            how="İş birliği ve ortak eylem yoluyla."),
    s5=dict(h="NEDEN İKLİM İMECESİ?", items=[
        "İklim değişikliği çok boyutlu küresel bir sorundur.",
        "Bilgi ve farkındalık tek başına yeterli değildir.",
        "Yerel bilgi ile akademik bilginin buluşması gerekir.",
        "Gençlerin gerçek çözümlere katılmak için fırsatlara ihtiyacı vardır.",
        "Sürdürülebilir değişim ortak sorumluluk gerektirir."]),
    s6=dict(h="İMECE NEDİR?", tag="Ortak çaba ve ortak eylemin Türk geleneği",
            defn="İMECE, insanların ortak bir amaca ulaşmak için gönüllü olarak bir araya gelmesini ifade eden geleneksel bir Türk kavramıdır.",
            rep="Eğitim modelimizde İMECE şunları temsil eder:",
            vals=["Aidiyet", "Katılım", "İş Birliği", "Birlikte Öğrenme",
                  "Kolektif Zekâ", "Ortak Eylem", "Kalite ve Toplumsal Etki"]),
    s7=dict(h="İMECE'NİN GÜCÜ", tag="Bireysel çabadan kolektif zekâya",
            cap1="Gençler gerçek sorunları birlikte belirliyor",
            cap2="Bir topluluk olarak öğrenmek ve harekete geçmek"),
    s8=dict(h="İMECE HALKASI NEDİR?",
            tag="Küçük, gönüllü ve iş birliğine dayalı bir öğrenme topluluğu",
            ring=["GERÇEK SORUN", "ETKİLENEN İNSANLAR", "ORTAK HEDEF", "İŞ BİRLİĞİ",
                  "EYLEM", "YANSITMA", "SÜREKLİ İYİLEŞTİRME"]),
    s9=dict(h="YOL GÖSTERİCİ İLKEMİZ", l1="Kendinle Yarış,", l2="Başkalarıyla İş Birliği Yap."),
    s10=dict(h="PROJENİN ÖZGÜN KATKISI", tag="Farkındalıktan eyleme",
             lead="İklim İmecesi şunları bir araya getirir:",
             items=["Gençlerin yerel ve deneyimsel bilgisi",
                    "Akademik ve bilimsel bilgi",
                    "Eğitim kurumları",
                    "Yerel topluluklar ve paydaşlar",
                    "Uluslararası öğrenme ve iş birliği ağları"],
             inn_h="YENİLİK",
             inn="Gençleri pasif öğrenicilerden aktif iklim çözüm ortaklarına dönüştürmek."),
    s11=dict(h="ORTAKLAR VE ULUSLARARASI İŞ BİRLİĞİ", tag="Çok ülkeli bir öğrenme ağı",
             c1h="YÜRÜTÜCÜ KURUM", c1=["İstanbul Kent Üniversitesi, Türkiye"],
             c2h="ULUSLARARASI ORTAKLAR",
             c2=["Kasdi Merbah Üniversitesi, Cezayir",
                 "Dünya Toplam Kalite ve Eğitimde Mükemmellik Konseyi (WCTQEE)",
                 "Little Angels High School, Gwalior, Hindistan",
                 "Suncity School, Gurgaon, Hindistan",
                 "Bhavan Vidyalaya, Panchkula, Hindistan",
                 "Gyan Heritage School, Lucknow, Hindistan",
                 "ve diğerleri…"],
             c3h="PROJE ARAŞTIRMA EKİBİ",
             c3=["Prof. Dr. Hayal Köksal – Proje Yürütücüsü",
                 "Dr. Memet Aslan – Araştırmacı",
                 "Aleyna Çavdar – Araştırmacı",
                 "Azad Kılıçarslan – Araştırmacı",
                 "Ayça Selin Akyol – Araştırmacı"],
             countries=["Türkiye", "Cezayir", "Hindistan"]),
    s12=dict(h="AMAÇ VE HEDEFLER", tag="Projenin amacı",
             aim="Gençlerin aşağıdakileri yapmasını sağlayan uygulanabilir bir İklim İmecesi modeli geliştirmek:",
             steps=[("Belirleme", "kendi yaşam çevrelerindeki iklim sorunlarını"),
                    ("Birleştirme", "yerel ve akademik bilgiyi"),
                    ("Geliştirme", "iş birliğine dayalı, sürdürülebilir çözümleri"),
                    ("Uygulama", "küçük ölçekli pilot eylemleri"),
                    ("Değerlendirme", "ve sonuçları paylaşma")]),
    s13=dict(h="AMAÇ VE HEDEFLER", tag="Projenin hedefleri", items=[
        "İstanbul Kent Üniversitesi'nde bir İklim İmecesi pilot uygulaması oluşturmak.",
        "İş birliğine dayalı İMECE Halkaları kurmak.",
        "En az üç uygulanabilir iklim çözümü geliştirmek.",
        "Türkiye, Cezayir ve Hindistan arasında uluslararası öğrenmeyi teşvik etmek.",
        "Uygulamaya dönük bir uygulama rehberi hazırlamak.",
        "Modelin başka kurum ve topluluklara aktarılabilirliğini değerlendirmek."]),
    s14=dict(h="YÖNTEM: İMECE HALKALARI VE PUKÖ",
             tag="Katılımcı ve eylem odaklı bir araştırma modeli",
             lead="Proje, İMECE Halkaları'nı kalite iyileştirme döngüsü olan PUKÖ yaklaşımıyla bütünleştirir:",
             pdca=[("PLANLA", "Sorunları belirle ve çözümleri planla."),
                   ("UYGULA", "Seçilen eylemleri hayata geçir."),
                   ("KONTROL ET", "Sonuçları izle ve değerlendir."),
                   ("ÖNLEM AL", "İyileştir, uyarla ve başarılı uygulamaları paylaş.")]),
    s15=dict(h="ARAŞTIRMA SÜRECİ", tag="Dört iş paketi (İP)", wps=[
        ("İP1", "Hazırlık", ["Literatür taraması", "Araştırma araçları", "Katılımcıların belirlenmesi"]),
        ("İP2", "İMECE Halkaları", ["Yerel iklim sorunlarının belirlenmesi", "Ortak sorun alanlarının tanımlanması", "İş birliğine dayalı tartışmalar"]),
        ("İP3", "Uluslararası Ortak Öğrenme", ["Çevrim içi toplantılar", "Yerel deneyimlerin paylaşılması", "İklim sorunlarının karşılaştırılması"]),
        ("İP4", "Ortak Çözümler ve Pilot Eylemler", ["Çözümlerin geliştirilmesi", "Seçilen eylemlerin uygulanması", "Sonuçların değerlendirilmesi"])]),
    s16=dict(h="ARAŞTIRMA VE DEĞERLENDİRME", tag="İlerlemeyi nasıl ölçeceğiz?", cols=[
        ("Üç temel veri kaynağı", ["Akran ve jüri değerlendirmesini de içeren başlangıç ve sonuç değerlendirme formları", "İMECE süreç kayıtları", "Katılımcı yansıtmaları ve değerlendirmeleri"]),
        ("Veri analizi", ["Frekans, yüzde ve ortalamalar", "Başlangıç ve sonuç değerlendirmelerinin karşılaştırılması", "Nitel verilerin tematik analizi"]),
        ("Temel değerlendirme alanları", ["İklim farkındalığı", "İklim eylemine katılım", "İş birliği", "Ortak problem çözme", "Çözümleri uygulamaya yönelik öz güven"])]),
    s17=dict(h="BEKLENEN ÇIKTILAR VE ETKİ", tag="Yerel çözümlerden uluslararası öğrenmeye", items=[
        "İklim İmecesi Gençlik Modeli",
        "Gençler tarafından geliştirilen en az üç iklim çözümü",
        "Pilot uygulama etkinlikleri",
        "Uygulamaya dönük İklim İmecesi Uygulama Rehberi",
        "Uluslararası Gençlik İklim Öğrenme Ağı",
        "Deneyimleri ve çözümleri belgeleyen dijital arşiv",
        "En az bir bilimsel makale veya bildiri"]),
    s18=dict(h="YEREL EYLEMDEN KÜRESEL KATKIYA", tag="Vizyonumuz", items=[
        "Gençleri güçlendirmek",
        "Yerel ve bilimsel bilgiyi birbirine bağlamak",
        "Uluslararası iş birliğini teşvik etmek",
        "Sürdürülebilir eğitim uygulamalarını desteklemek",
        "Ortak sorumluluk kültürü oluşturmak"],
        quote="İklim eylemi yerelde başlar; etkisi ise küresele ulaşabilir."),
    s19=dict(h="SONUÇ", tag="İklim İmecesi: birlikte öğrenmek, birlikte harekete geçmek",
             lead="İnanıyoruz ki:", items=[
        "Gençler yalnızca gelecek değil, bugünün iklim eyleminin ortaklarıdır.",
        "Eğitim anlamlı eyleme yol açmalıdır.",
        "İş birliği bireysel bilgiyi kolektif zekâya dönüştürür.",
        "Kalite iyileştirme, eylemi ölçülebilir ve sürdürülebilir kılar.",
        "İMECE, insanların daha iyi bir gelecek için birlikte çalışmasına olanak sağlar."]),
    s20=dict(thanks="TEŞEKKÜRLER",
             text="Bu araştırma projesini bugün İstanbul Kent Üniversitesi ve uluslararası proje ortaklarımız adına sunuyorum.",
             motto="Kendinle Yarış, Başkalarıyla İş Birliği Yap.",
             sig=["Prof. Dr. Hayal Köksal", "İstanbul Kent Üniversitesi, Türkiye", "hayal.koksal@kent.edu.tr"]),
)


# ---- slide builders ---------------------------------------------------------
def kent_logo(s, x, y, h, name="Istanbul Kent University logo"):
    return pic(s, name, f"{LOGOS}/kent_logo_lacivert_kirmizi_crop.png", x, y, h=h)


def footer_logo(s):
    """Content slides: narrow the disclaimer and add the university logo beside the YÖK logos."""
    for sh in s.shapes:
        if sh.has_text_frame and "YÖK" in sh.text_frame.text and sh.width > Inches(8):
            sh.width = Inches(7.9)
    kent_logo(s, 8.75, 6.62, 0.5)


def build_title(s, c):
    for sh in s.shapes:
        if not sh.has_text_frame:
            continue
        t = sh.text_frame.text.strip()
        if t == "TITLE":
            sh.text_frame.paragraphs[0].runs[0].text = c["title"]
            sh.text_frame.word_wrap = True
            sh.text_frame.auto_size = MSO_AUTO_SIZE.NONE
            sh.top = Inches(2.8); sh.width = Inches(10.8); sh.height = Inches(0.9)
        elif t == "Subtitle":
            sh.text_frame.word_wrap = True
            sh.text_frame.auto_size = MSO_AUTO_SIZE.NONE
            sh.width = Inches(10.8); sh.top = Inches(3.75); sh.height = Inches(1.0)
            r = sh.text_frame.paragraphs[0].runs[0]
            r.text = c["sub"]; r.font.size = Pt(22)
    kent_logo(s, 8.3, 1.05, 1.32)
    tb(s, "Presenter", 1.14, 4.95, 10.8, 1.0,
       [(c["presenter"][0], dict(bold=True, size=18, color=BLUE)),
        c["presenter"][1], c["presenter"][2]], size=14, color=INK, space_after=2)
    tb(s, "Event", 1.14, 5.95, 10.8, 0.35, [c["event"]], size=12, color=MUTED)


def build_thanks(s, c):
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip() == "THANK YOU":
            sh.text_frame.paragraphs[0].runs[0].text = c["thanks"]
            sh.text_frame.word_wrap = True
            sh.text_frame.auto_size = MSO_AUTO_SIZE.NONE
            sh.top = Inches(2.8); sh.width = Inches(10.8); sh.height = Inches(0.9)
    kent_logo(s, 8.3, 1.05, 1.32)
    tb(s, "Closing text", 1.14, 3.8, 10.8, 0.8, [c["text"]], size=16, color=INK)
    tb(s, "Motto", 1.14, 4.6, 10.8, 0.5, [c["motto"]], size=20, color=GREEN, bold=True,
       font=HEAD_FONT)
    tb(s, "Signature", 1.14, 5.2, 10.8, 1.0,
       [(c["sig"][0], dict(bold=True, color=BLUE, size=16)), c["sig"][1], c["sig"][2]],
       size=14, space_after=2)


def build_s2(s, c):
    # left: project statement
    card(s, "Project card", L, T, 6.6, H, fill=BLUE, line=None)
    tb(s, "Project title", L + 0.4, T + 0.4, 5.8, 1.9, [c["big"]], size=26, color=WHITE,
       bold=True, font=HEAD_FONT, line_spacing=1.05)
    tb(s, "Project tagline", L + 0.4, T + 2.5, 5.8, 1.2, [c["tag"]], size=16,
       color=WHITE, italic=True)
    pill(s, "Country pill", L + 1.6, B - 0.55, 2.4, 0.42, "COP31 · Antalya 2026", fill=GREEN, size=12)
    # right: team
    x = L + 6.9; w = R - x
    tb(s, "Team heading", x, T, w, 0.4, [c["team_h"]], size=15, color=BLUE, bold=True)
    y = T + 0.55
    for i, (n, role) in enumerate(c["team"]):
        card(s, f"Team card {i+1}", x, y, w, 0.72)
        circle(s, f"Team num {i+1}", x + 0.45, y + 0.36, 0.44, fill=BLUE if i else GREEN,
               text=n.split()[-1][0] + n.split()[-2][0] if i else "HK", size=11)
        tb(s, f"Team name {i+1}", x + 0.85, y + 0.08, w - 0.95, 0.6,
           [(n, dict(bold=True, size=14)), (role, dict(size=11, color=MUTED))],
           space_after=0, anchor=MSO_ANCHOR.MIDDLE)
        y += 0.84


def build_s3(s, c):
    card(s, "Question card", L + 0.6, T + 0.35, W - 1.2, 2.4, fill=BLUE, line=None, radius=0.12)
    tb(s, "Question", L + 1.0, T + 0.55, W - 2.0, 2.0, [c["q"]], size=30, color=WHITE,
       bold=True, font=HEAD_FONT, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
       line_spacing=1.1)
    circle(s, "Q mark", L + 0.6, T + 0.35, 0.8, fill=GREEN, text="?", size=30)
    tb(s, "Answer", L + 1.5, T + 3.2, W - 3.0, 1.3, [c["a"]], size=20, color=INK,
       align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def build_s4(s, c):
    pic(s, "Paper people circle", "image3.jpg", 6.9, T + 0.35, w=5.7)
    tb(s, "Model lead", L, T + 0.2, 5.9, 0.6, [c["lead"]], size=22, color=BLUE, bold=True,
       font=HEAD_FONT)
    y = T + 1.0
    for i, w_ in enumerate(c["who"]):
        card(s, f"Who card {i+1}", L, y, 5.9, 0.62)
        circle(s, f"Who num {i+1}", L + 0.4, y + 0.31, 0.38, fill=GREEN if i % 2 else BLUE,
               text=str(i + 1), size=12)
        tb(s, f"Who text {i+1}", L + 0.75, y, 5.0, 0.62, [w_], size=17, bold=True,
           anchor=MSO_ANCHOR.MIDDLE)
        y += 0.74
    tb(s, "Model how", L, y + 0.1, 5.9, 0.8, [c["how"]], size=18, color=GREEN, bold=True,
       font=HEAD_FONT)


def build_s5(s, c):
    n = len(c["items"])
    gap = 0.22
    cw = (W - gap * (n - 1)) / n
    for i, it in enumerate(c["items"]):
        x = L + i * (cw + gap)
        card(s, f"Why card {i+1}", x, T + 0.35, cw, 4.3)
        circle(s, f"Why num {i+1}", x + cw / 2, T + 1.25, 0.9, fill=BLUE if i % 2 == 0 else GREEN,
               text=str(i + 1), size=24)
        tb(s, f"Why text {i+1}", x + 0.2, T + 2.0, cw - 0.4, 2.5, [it], size=16,
           align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)


def build_s6(s, c):
    y = subtitle(s, c["tag"])
    card(s, "Definition card", L, y, 5.6, 2.4, fill=BLUE, line=None)
    tb(s, "Definition", L + 0.3, y + 0.2, 5.0, 2.0, [c["defn"]], size=17, color=WHITE,
       anchor=MSO_ANCHOR.MIDDLE)
    tb(s, "Represents", L + 6.0, y, W - 6.0, 0.4, [c["rep"]], size=15, bold=True, color=BLUE)
    # value pills in two columns
    vx = L + 6.0; vy = y + 0.5; pw = 3.1; ph = 0.5
    for i, v in enumerate(c["vals"]):
        col = i % 2; row = i // 2
        pill(s, f"Value {i+1}", vx + col * (pw + 0.2) + pw / 2, vy + row * 0.62 + ph / 2,
             pw, ph, v, fill=GREEN if i % 3 == 1 else (LBLUE if i % 3 == 2 else BLUE), size=13)
    ph = B - (y + 2.6)
    p = pic(s, "Circle photo", "image6.png", L, y + 2.6, w=5.6)
    extra = 1 - Inches(ph) / p.height
    p.crop_top = extra * 0.45; p.crop_bottom = extra * 0.55
    p.height = Inches(ph)


def build_s7(s, c):
    y = subtitle(s, c["tag"])
    pw = (W - 0.3) / 2
    ph = 3.6
    pic(s, "Photo 1", "image5.jpg", L, y, w=pw)
    pic(s, "Photo 2", "image4.jpg", L + pw + 0.3, y, w=pw)
    # crop both to same height by adjusting crop
    for sh in s.shapes:
        if sh.name in ("Photo 1", "Photo 2"):
            if sh.height > Inches(ph):
                extra = 1 - Inches(ph) / sh.height
                sh.crop_top = extra / 2; sh.crop_bottom = extra / 2
                sh.height = Inches(ph)
    tb(s, "Caption 1", L, y + ph + 0.1, pw, 0.4, [c["cap1"]], size=12, color=MUTED, italic=True,
       align=PP_ALIGN.CENTER)
    tb(s, "Caption 2", L + pw + 0.3, y + ph + 0.1, pw, 0.4, [c["cap2"]], size=12, color=MUTED,
       italic=True, align=PP_ALIGN.CENTER)


def build_s8(s, c):
    subtitle(s, c["tag"])
    cx, cy = CX, 4.1
    rx, ry = 4.2, 1.75
    ring = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - rx), Inches(cy - ry), Inches(2 * rx), Inches(2 * ry))
    ring.name = "Ring"; ring.fill.background(); ring.line.color.rgb = LBLUE; ring.line.width = Pt(1.5)
    ring.shadow.inherit = False
    ih = 2.2
    p = pic(s, "Circle photo", "image6.png", 0, 0, h=ih)
    p.left = Inches(cx) - p.width // 2; p.top = Inches(cy - ih / 2)
    n = len(c["ring"])
    for i, lab in enumerate(c["ring"]):
        a = -math.pi / 2 + i * 2 * math.pi / n
        px, py = cx + rx * math.cos(a), cy + ry * math.sin(a)
        pill(s, f"Ring label {i+1}", px, py, 2.2, 0.5, lab,
             fill=BLUE if i % 2 == 0 else GREEN, size=11)


def build_s9(s, c):
    card(s, "Principle card", L + 1.2, T + 0.6, W - 2.4, 3.6, fill=CARD)
    tb(s, "Line 1", L + 1.5, T + 1.0, W - 3.0, 1.3, [c["l1"]], size=40, color=BLUE, bold=True,
       font=HEAD_FONT, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    tb(s, "Line 2", L + 1.5, T + 2.3, W - 3.0, 1.3, [c["l2"]], size=40, color=GREEN, bold=True,
       font=HEAD_FONT, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def build_s10(s, c):
    y = subtitle(s, c["tag"])
    tb(s, "Lead", L, y, 6.6, 0.4, [c["lead"]], size=16, bold=True, color=INK)
    yy = y + 0.55
    for i, it in enumerate(c["items"]):
        card(s, f"Item card {i+1}", L, yy, 6.6, 0.6)
        circle(s, f"Item num {i+1}", L + 0.38, yy + 0.3, 0.36, fill=BLUE, text=str(i + 1), size=12)
        tb(s, f"Item text {i+1}", L + 0.7, yy, 5.8, 0.6, [it], size=14, anchor=MSO_ANCHOR.MIDDLE)
        yy += 0.7
    card(s, "Innovation card", L + 7.0, y, W - 7.0, 3.9, fill=GREEN, line=None, radius=0.12)
    tb(s, "Innovation heading", L + 7.35, y + 0.35, W - 7.7, 0.5, [c["inn_h"]], size=16,
       color=WHITE, bold=True, font=HEAD_FONT)
    tb(s, "Innovation text", L + 7.35, y + 0.95, W - 7.7, 2.7, [c["inn"]], size=22, color=WHITE,
       bold=True, line_spacing=1.1)


def build_s11(s, c):
    y = subtitle(s, c["tag"])
    # country pills on the right of the subtitle
    for i, cn in enumerate(c["countries"]):
        pill(s, f"Country {i+1}", R - 0.8 - i * 1.65, T + 0.17, 1.5, 0.38, cn,
             fill=[BLUE, GREEN, RED][i], size=11)
    cols = [(c["c1h"], c["c1"], 3.2), (c["c2h"], c["c2"], 4.9), (c["c3h"], c["c3"], 3.7)]
    x = L
    for i, (hd, items, cw) in enumerate(cols):
        card(s, f"Col card {i+1}", x, y, cw, B - y)
        tb(s, f"Col head {i+1}", x + 0.25, y + 0.2, cw - 0.5, 0.4, [hd], size=13, bold=True, color=BLUE)
        bullets(s, f"Col {i+1}", x + 0.25, y + 0.7, cw - 0.5, 3.0, items, size=13 if i else 15,
                gap=5, dot=GREEN if i == 1 else BLUE)
        x += cw + 0.215


def build_s12(s, c):
    y = subtitle(s, c["tag"])
    tb(s, "Aim", L, y + 0.1, W, 0.9, [c["aim"]], size=20, color=INK, bold=True)
    n = len(c["steps"])
    gap = 0.15
    cw = (W - gap * (n - 1)) / n
    yy = y + 1.3
    for i, (verb, rest) in enumerate(c["steps"]):
        x = L + i * (cw + gap)
        sh = s.shapes.add_shape(MSO_SHAPE.CHEVRON if i else MSO_SHAPE.PENTAGON,
                                Inches(x), Inches(yy), Inches(cw), Inches(0.7))
        sh.name = f"Step arrow {i+1}"
        sh.fill.solid(); sh.fill.fore_color.rgb = BLUE if i % 2 == 0 else GREEN
        sh.line.fill.background(); sh.shadow.inherit = False
        sh.adjustments[0] = 0.25
        tf = sh.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        _runs(p, str(i + 1), 18, WHITE, bold=True)
        card(s, f"Step card {i+1}", x, yy + 0.85, cw, 1.9)
        tb(s, f"Step text {i+1}", x + 0.15, yy + 1.0, cw - 0.3, 1.9,
           [(verb, dict(bold=True, color=BLUE, size=17, font=HEAD_FONT)), rest],
           size=14, space_after=6)


def build_s13(s, c):
    y = subtitle(s, c["tag"])
    n = len(c["items"]); cols = 3; rows = 2
    gap = 0.2
    cw = (W - gap * (cols - 1)) / cols
    ch = (B - y - gap) / rows
    for i, it in enumerate(c["items"]):
        col, row = i % cols, i // cols
        x = L + col * (cw + gap); yy = y + row * (ch + gap)
        card(s, f"Obj card {i+1}", x, yy, cw, ch)
        circle(s, f"Obj num {i+1}", x + 0.5, yy + 0.5, 0.56, fill=BLUE if i % 2 == 0 else GREEN,
               text=str(i + 1), size=16)
        tb(s, f"Obj text {i+1}", x + 0.25, yy + 0.95, cw - 0.5, ch - 1.1, [it], size=15)


def build_s14(s, c):
    y = subtitle(s, c["tag"])
    tb(s, "Lead", L, y, W, 0.7, [c["lead"]], size=16, color=INK)
    yy = y + 0.85
    gap = 0.2
    cw = (W - gap * 3) / 4
    colors = [BLUE, GREEN, LBLUE, RED]
    for i, (k, v) in enumerate(c["pdca"]):
        x = L + i * (cw + gap)
        card(s, f"PDCA card {i+1}", x, yy, cw, B - yy)
        band = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(yy), Inches(cw), Inches(1.1))
        band.name = f"PDCA band {i+1}"; band.adjustments[0] = 0.12
        band.fill.solid(); band.fill.fore_color.rgb = colors[i]; band.line.fill.background()
        band.shadow.inherit = False
        tb(s, f"PDCA key {i+1}", x, yy, cw, 1.1, [k], size=22, color=WHITE, bold=True,
           font=HEAD_FONT, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        circle(s, f"PDCA letter {i+1}", x + cw / 2, yy + 1.75, 0.7, fill=colors[i], text=k[0], size=24)
        tb(s, f"PDCA text {i+1}", x + 0.2, yy + 2.25, cw - 0.4, B - yy - 2.4, [v], size=15,
           align=PP_ALIGN.CENTER)


def build_s15(s, c):
    y = subtitle(s, c["tag"])
    gap = 0.2
    cw = (W - gap * 3) / 4
    for i, (code, name, items) in enumerate(c["wps"]):
        x = L + i * (cw + gap)
        card(s, f"WP card {i+1}", x, y, cw, B - y)
        pill(s, f"WP code {i+1}", x + 0.65, y + 0.4, 0.9, 0.42, code, fill=BLUE if i % 2 == 0 else GREEN, size=13)
        tb(s, f"WP name {i+1}", x + 0.2, y + 0.75, cw - 0.4, 0.9, [name], size=16, bold=True,
           color=BLUE, anchor=MSO_ANCHOR.MIDDLE)
        bullets(s, f"WP {i+1}", x + 0.2, y + 1.75, cw - 0.4, 2.5, items, size=14, gap=6,
                dot=GREEN if i % 2 == 0 else BLUE)
        if i < 3:
            ar = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x + cw - 0.02), Inches(y + 0.47),
                                    Inches(gap + 0.04), Inches(0.28))
            ar.name = f"WP arrow {i+1}"; ar.fill.solid(); ar.fill.fore_color.rgb = LBLUE
            ar.line.fill.background(); ar.shadow.inherit = False


def build_s16(s, c):
    y = subtitle(s, c["tag"])
    gap = 0.25
    cw = (W - gap * 2) / 3
    fills = [BLUE, GREEN, LBLUE]
    for i, (hd, items) in enumerate(c["cols"]):
        x = L + i * (cw + gap)
        card(s, f"Eval card {i+1}", x, y, cw, B - y)
        band = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(cw), Inches(0.8))
        band.name = f"Eval band {i+1}"; band.adjustments[0] = 0.15
        band.fill.solid(); band.fill.fore_color.rgb = fills[i]; band.line.fill.background()
        band.shadow.inherit = False
        tb(s, f"Eval head {i+1}", x + 0.2, y, cw - 0.4, 0.8, [hd], size=15, color=WHITE, bold=True,
           anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
        bullets(s, f"Eval {i+1}", x + 0.25, y + 1.0, cw - 0.5, 3.0, items, size=14, gap=6, dot=fills[i])


def build_s17(s, c):
    y = subtitle(s, c["tag"])
    items = c["items"]
    gap = 0.2
    # row 1: 4 cards, row 2: 3 cards
    rows = [items[:4], items[4:]]
    ch = (B - y - gap) / 2
    cw = (W - gap * 3) / 4
    for r_i, row in enumerate(rows):
        n = len(row)
        x0 = L + (W - (n * cw + (n - 1) * gap)) / 2
        for j, it in enumerate(row):
            i = r_i * 4 + j
            x = x0 + j * (cw + gap); yy = y + r_i * (ch + gap)
            card(s, f"Out card {i+1}", x, yy, cw, ch)
            circle(s, f"Out num {i+1}", x + 0.45, yy + 0.45, 0.5, fill=GREEN if i % 2 else BLUE,
                   text=str(i + 1), size=14)
            tb(s, f"Out text {i+1}", x + 0.2, yy + 0.85, cw - 0.4, ch - 1.0, [it], size=14)


def build_s18(s, c):
    y = subtitle(s, c["tag"])
    yy = y
    for i, it in enumerate(c["items"]):
        card(s, f"Vision card {i+1}", L, yy, 6.4, 0.66)
        circle(s, f"Vision num {i+1}", L + 0.4, yy + 0.33, 0.4, fill=BLUE if i % 2 == 0 else GREEN,
               text=str(i + 1), size=12)
        tb(s, f"Vision text {i+1}", L + 0.75, yy, 5.5, 0.66, [it], size=15, anchor=MSO_ANCHOR.MIDDLE)
        yy += 0.78
    card(s, "Quote card", L + 6.8, y, W - 6.8, 3.9, fill=BLUE, line=None, radius=0.12)
    tb(s, "Quote mark", L + 7.1, y + 0.15, 1.0, 1.0, ["“"], size=60, color=GREEN, bold=True, font=HEAD_FONT)
    tb(s, "Quote", L + 7.2, y + 0.9, W - 7.6, 2.8, [c["quote"]], size=26, color=WHITE, bold=True,
       line_spacing=1.1, anchor=MSO_ANCHOR.TOP)


def build_s19(s, c):
    y = subtitle(s, c["tag"])
    tb(s, "Lead", L, y, W, 0.4, [c["lead"]], size=16, bold=True, color=INK)
    yy = y + 0.5
    items = c["items"]
    gap = 0.15
    ch = (B - yy - gap * (len(items) - 1)) / len(items)
    for i, it in enumerate(items):
        card(s, f"Belief card {i+1}", L, yy, W, ch)
        circle(s, f"Belief num {i+1}", L + 0.45, yy + ch / 2, 0.42, fill=BLUE if i % 2 == 0 else GREEN,
               text=str(i + 1), size=13)
        tb(s, f"Belief text {i+1}", L + 0.85, yy, W - 1.0, ch, [it], size=15, anchor=MSO_ANCHOR.MIDDLE)
        yy += ch + gap


BUILDERS = {2: build_s2, 3: build_s3, 4: build_s4, 5: build_s5, 6: build_s6, 7: build_s7,
            8: build_s8, 9: build_s9, 10: build_s10, 11: build_s11, 12: build_s12,
            13: build_s13, 14: build_s14, 15: build_s15, 16: build_s16, 17: build_s17,
            18: build_s18, 19: build_s19}
LONG_HEADERS = 34


def build(lang, outfile):
    c = C[lang]
    prs = Presentation(BASE)
    slides = list(prs.slides)
    assert len(slides) == 20
    for i, s in enumerate(slides, 1):
        fill_disclaimer(s, C["tr"]["inst"], C["en"]["inst"])
        if i == 1:
            build_title(s, c)
        elif i == 20:
            build_thanks(s, c["s20"])
        else:
            cc = c[f"s{i}"]
            set_header(s, cc["h"], size=24 if len(cc["h"]) <= LONG_HEADERS else 20)
            BUILDERS[i](s, cc)
            footer_logo(s)
    prs.save(outfile)
    print("wrote", outfile)


build("en", f"{OUT}/COP31_Climate_IMECE_Hayal_Koksal_EN.pptx")
build("tr", f"{OUT}/COP31_Iklim_Imecesi_Hayal_Koksal_TR.pptx")
