# COP31 — Hayal Köksal

İstanbul Kent Üniversitesi üzerinden sunulan COP31 projesinin slayt ve poster revizyon çalışması.

## Klasörler

| Klasör | İçerik |
|---|---|
| `kaynak/` | Proje metni (`COP31_KENT_Son_Hayal_Koksal.docx`) ve üniversite belgesi (`Istanbul_Kent_Universitesi.pdf`) |
| `sablonlar/` | COP31'in resmî sunum ve dijital poster şablonları |
| `sunum/` | Mevcut sunum (`COP31_PROJECT_by_HAYAL_KOKSAL.pptx`) ve revizeleri |
| `poster/` | Poster revizeleri |
| `araclar/` | Sunumları şablondan üreten betik (`sunum_uret.py`) |

## Çalışma düzeni

Özgün dosyalara dokunulmaz; her revize yeni bir dosya adıyla (`..._v2.pptx` gibi) ilgili klasöre eklenir.

## Sunum revizeleri (COP31 şablonu)

| Dosya | Açıklama |
|---|---|
| `sunum/COP31_Climate_IMECE_Hayal_Koksal_EN.pptx` | İngilizce, COP31/YÖK sunum şablonuna uyarlanmış (20 slayt) |
| `sunum/COP31_Iklim_Imecesi_Hayal_Koksal_TR.pptx` | Türkçe sürüm (aynı yapı) |
| `sunum/*.pdf` | Hızlı önizleme için PDF dışa aktarımları (yazı tipi yerine konulmuş olabilir; asıl dosya .pptx) |

Üretim: `araclar/sunum_uret.py <taban.pptx> <medya_klasörü> <çıktı_klasörü>`. Taban dosya, şablondaki
"Başlık ve İçerik" slaydının 18 kez çoğaltılmasıyla (kapak + 18 içerik + kapanış) oluşturulur;
medya klasörü hocanın özgün sunumundaki görsellerdir (`ppt/media`).

## Poster durumu

`kaynak/Istanbul_Kent_Universitesi.pdf` (COP31 Yeşil Bölge katılım programı, 19 Kasım 2026) içinde
İklim İmecesi projesi yalnızca **Panel 2, 14:00–15:00** oturumunda yer alıyor. Proje ne stant projeleri
listesinde ne de dijital poster şablonunun kullanıldığı **ortak dijital ekran** programında bulunuyor.
Bu nedenle bu proje için poster istenmiyor; `poster/` klasörü şimdilik boş.
