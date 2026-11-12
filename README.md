# COP31 — Hayal Köksal

İstanbul Kent Üniversitesi üzerinden sunulan COP31 projesinin slayt ve poster revizyon çalışması.

## Klasörler

| Klasör | İçerik |
|---|---|
| `kaynak/` | Proje metni (`COP31_KENT_Son_Hayal_Koksal.docx`), üniversite belgesi (`Istanbul_Kent_Universitesi.pdf`) ve üniversite logoları (`logo/`) |
| `sablonlar/` | COP31'in resmî sunum ve dijital poster şablonları |
| `sunum/` | Mevcut sunum (`COP31_PROJECT_by_HAYAL_KOKSAL.pptx`) ve revizeleri |
| `poster/` | Poster revizeleri |
| `araclar/` | Sunum ve posteri şablondan üreten betikler (`sunum_uret.py`, `poster_uret.py`, ortak `helpers.py`) |

## Çalışma düzeni

Özgün dosyalara dokunulmaz; her revize yeni bir dosya adıyla (`..._v2.pptx` gibi) ilgili klasöre eklenir.

## Sunum revizeleri (COP31 şablonu)

| Dosya | Açıklama |
|---|---|
| `sunum/COP31_Climate_IMECE_Hayal_Koksal_EN.pptx` | İngilizce, COP31/YÖK sunum şablonuna uyarlanmış (20 slayt) |
| `sunum/COP31_Iklim_Imecesi_Hayal_Koksal_TR.pptx` | Türkçe sürüm (aynı yapı) |
| `sunum/*.pdf` | Hızlı önizleme için PDF dışa aktarımları (yazı tipi yerine konulmuş olabilir; asıl dosya .pptx) |

Üretim: `araclar/sunum_uret.py <taban.pptx> <medya_klasörü> <çıktı_klasörü> <logo_klasörü>`. Taban dosya, şablondaki
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
