# COP31 — Hayal Köksal

İstanbul Kent Üniversitesi üzerinden sunulan COP31 projesinin slayt ve poster revizyon çalışması.

## Klasörler

| Klasör | İçerik |
|---|---|
| `kaynak/` | Proje metni (`COP31_KENT_Son_Hayal_Koksal.docx`), üniversite belgesi (`Istanbul_Kent_Universitesi.pdf`) ve üniversite logoları (`logo/`) |
| `sablonlar/` | COP31'in resmî sunum ve dijital poster şablonları |
| `sunum/` | Mevcut sunum (`COP31_PROJECT_by_HAYAL_KOKSAL.pptx`) ve revizeleri |
| `poster/` | Poster revizeleri |
| `dijital_ekran/` | 6 slaytlık sessiz dijital ekran sunumu (.pptx + MP4) ve özel çizim görseller |
| `kart/` | A5 proje tanıtım kartı (ön yüz İngilizce, arka yüz Türkçe) |
| `araclar/` | Üretim betikleri: `sunum_uret.py`, `poster_uret.py`, `dijital_ekran_uret.py` (+ `ds_illust.py`, `ds_shots.js`, `ds_video.sh`), `kart_uret.py`, `render_copy.py`, ortak `helpers.py` |

## Çalışma düzeni

Özgün dosyalara dokunulmaz; her revize yeni bir dosya adıyla (`..._v2.pptx` gibi) ilgili klasöre eklenir.

## Sunum revizeleri (COP31 şablonu)

| Dosya | Açıklama |
|---|---|
| `sunum/COP31_Climate_IMECE_Hayal_Koksal_EN.pptx` | İngilizce, COP31/YÖK sunum şablonuna uyarlanmış (20 slayt) |
| `sunum/COP31_Iklim_Imecesi_Hayal_Koksal_TR.pptx` | Türkçe sürüm (aynı yapı) |
| `sunum/*.pdf` | Hızlı önizleme için PDF dışa aktarımları (yazı tipi yerine konulmuş olabilir; asıl dosya .pptx) |

Üretim: `araclar/sunum_uret.py <taban.pptx> <medya_klasörü> <çıktı_klasörü> <logo_klasörü> <fotoğraf_klasörü>`. Taban dosya, şablondaki
"Başlık ve İçerik" slaydının 18 kez çoğaltılmasıyla (kapak + 18 içerik + kapanış) oluşturulur;
medya klasörü hocanın özgün sunumundaki görsellerdir (`ppt/media`).

İstanbul Kent Üniversitesi logosu kapak ve kapanış slaytlarında YÖK logosunun yanında, içerik
slaytlarında ise alt bilgide (sorumluluk notu ile YÖK logosu arasında) yer alır.

## Poster

| Dosya | Açıklama |
|---|---|
| `poster/COP31_Climate_IMECE_Poster_EN.pptx` | COP31 dijital poster şablonu (9:16, 13,5×24 inç) üzerine İngilizce poster |
| `poster/COP31_Climate_IMECE_Poster_EN.pdf` | Önizleme |

Şablon notları gereği poster İngilizce hazırlandı; beklenen çıktılar "hedef" olarak işaretlendi,
QR yer tutucusu kullanılmadı (proje sayfası yok). Üretim:
`araclar/poster_uret.py <poster_şablonu.pptx> <medya_klasörü> <logo_klasörü> <çıktı.pptx>`.

**Program notu:** `kaynak/Istanbul_Kent_Universitesi.pdf` (COP31 Yeşil Bölge katılım programı, 19 Kasım 2026)
içinde İklim İmecesi projesi yalnızca **Panel 2, 14:00–15:00** oturumunda yer alıyor; stant projeleri ve
ortak dijital ekran listelerinde yok. Poster resmî programda istenmese de hazır bulundurulmak üzere eklendi.

## Ekran ve teslim bilgileri (YÖK e-postası)

YÖK'ün programla birlikte gönderdiği e-postadaki teknik koşullar ve bu projeye düşen kısım:

| Konu | Koşul | Bu proje için anlamı |
|---|---|---|
| Metin dili | Sunum dosyaları ve posterlerde metin **İngilizce** olmalı; sözlü sunum Türkçe veya İngilizce olabilir | Teslim edilecek dosya `sunum/COP31_Climate_IMECE_Hayal_Koksal_EN.pptx`. Türkçe sürüm yalnızca konuşmacı desteği içindir, teslim edilmez |
| Projenin yeri | Panel 2, 19 Kasım 2026, 14:00–15:00 (`kaynak/Istanbul_Kent_Universitesi.pdf`) | Panel alanı ekranı: **75 inç yatay TV**; 16:9 sunum şablonu doğru seçim |
| Bağlantı | Her TV'nin altında USB girişi, 3 m HDMI kablosu | HDMI ile dizüstünden .pptx sunulabilir; USB için ekran dosyaları `sunum/ekran/` altında |
| Ortak dijital ekran | Dikey 75 inç TV (9:16) | Proje bu listede olmadığından poster zorunlu değil; gerekirse `poster/ekran/` dosyaları hazır |
| Stant | 65 inç TV | Projenin stant tahsisi yok |

USB'den oynatma için hazırlanan dosyalar (LibreOffice ile çizildi; şablonun Mozaic GEO fontu
yerine benzer kalın sans-serif kullanıldı, .pptx dosyaları fontu korur):

- `sunum/ekran/slaytlar_png/slide-01.png … slide-20.png` — 1920×1080, slayt başına bir görsel
- `sunum/ekran/COP31_Climate_IMECE_Hayal_Koksal_EN_1920x1080.mp4` — H.264, slayt başına 20 sn (yedek)
- `poster/ekran/COP31_Climate_IMECE_Poster_EN_1080x1920.png` / `.jpg` — dikey ekran için
- `poster/ekran/COP31_Climate_IMECE_Poster_EN_1080x1920.mp4` — 60 sn sabit video (TV'nin görsel oynatmadığı durum için)

Hangi formatın oynatılacağı etkinlik teknik ekibiyle teyit edilmeli (şablon notunda da istenir).

## Dijital ekran sunumu ve A5 kart (üniversite e-postası, son teslim 17 Kasım 2026 17.00)

Üniversitenin COP31 koordinasyonundan gelen e-postaya göre hazırlanan materyaller
(`cop31@kent.edu.tr` adresine iletilecek):

| İstenen | Dosya | Durum |
|---|---|---|
| Dijital ekran sunumu (.pptx) | `dijital_ekran/COP31_Climate_IMECE_Digital_Screen_EN.pptx` | 6 slayt, İngilizce, slayt başına en çok 25 kelime (kapak 29: proje adı ve ekip zorunlu) |
| Aynı sunumun videosu | `dijital_ekran/COP31_Climate_IMECE_Digital_Screen_EN_1920x1080.mp4` | 1920×1080, 75,5 sn, slayt başına 13 sn, sessiz, döngüye uygun |
| Proje tanıtım kartı (A5) | `kart/COP31_Climate_IMECE_Card_A5.pdf` (+ `.pptx`, `_EN.png`, `_TR.png`) | 148×210 mm; 1. sayfa İngilizce, 2. sayfa Türkçe |
| Bilgi notu (1 sayfa) | — | Hazırlanmadı |

Akış: Kapak → Problem and Aim → Method → Expected Outputs → Climate Action and Impact → Thank You.
İçerik projede fiilen ne yapıldığını anlatır: öğrenciler gönüllü bir halka kurar, kendi çevrelerinden bir iklim
sorunu seçer, çözümü küçük ölçekte dener, ölçüp paylaşır (Türkiye, Cezayir, Hindistan).
Proje henüz başlamadığı için 4. slayt "bulgular" değil **beklenen çıktılar (hedefler)** olarak adlandırıldı.

- Sorumluluk reddi metnindeki boşluklara, e-postadaki talimat gereği hem Türkçe hem İngilizce metinde
  "İstanbul Kent Üniversitesi – Prof. Dr. Hayal Köksal" yazıldı.
- Görseller projeye özel çizildi (`araclar/ds_illust.py`): İMECE halkası, PDCA döngüsü, farkındalık→eylem,
  ülke ağı küresi ve ikonlar. Şeffaf PNG kopyaları `dijital_ekran/gorseller/` altında.
- .pptx içinde öğeler sırayla belirir, slaytlar 13 sn'de otomatik geçer ve gösteri döngüdedir. Video aynı
  slayt tanımından üretilir; çizimler videoda hareketlidir ve arka planda şablonun kendi videoları oynar.
- **QR kod** şu an projenin yöntem dayanağı olan yayına gider: `https://doi.org/10.32329/uad.788641`
  (başvuru formundaki [3] numaralı kaynak). Proje sayfası ya da sunum bağlantısı verilirse
  `DS_QR_URL=<adres>` ile yeniden üretilir. Bağlantı bu ortamdan açılıp doğrulanamadı.

Şablonla ilgili iki teknik not (tüm sunumları etkiler):

- Şablondaki ikinci logo **COP31 Türkiye** logosudur ve EMF+ biçimindedir; LibreOffice bunu çizemez.
  PDF/PNG önizlemeleri bu yüzden `araclar/render_copy.py` ile üretilir (yalnızca önizleme kopyasında logo
  PNG ile değiştirilir, teslim edilen .pptx şablondaki haliyle kalır).
- Kapak ve kapanış düzeninin arka planında, slayt gösterisinde sağ tarafta büyüyen kırmızı lale videosu
  oynar. Bu yüzden kapak/kapanış içeriği slaydın sol bölümünde (x < 8 inç) tutulur; 20 slaytlık panel
  sunumunun kapağı da buna göre düzeltildi.

## Hocanın revizeleri (13 Kasım 2026)

Hocanın mor vurguyla işaretlediği dosya `kaynak/revize/` altında, eklediği dört fotoğraf `kaynak/fotograflar/`
altında. Tüm revizeler Türkçe ve İngilizce panel sunumuna işlendi. Slogan slaydı ("Kendinle Yarış,
Başkalarıyla İş Birliği Yap") hocanın isteğiyle sona, kapanıştan hemen önceye alındı ve dört fotoğraf bu
slayda yerleştirildi.
