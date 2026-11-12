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
