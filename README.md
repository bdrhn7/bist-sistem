# BIST Sistem

Google Colab'da çalışan, Borsa İstanbul hisselerini tarayan kişisel analiz sistemi.
Kodlar bu depoda durur; Colab her çalıştırmada en güncel halini buradan çeker.

> Eğitim amaçlıdır, yatırım tavsiyesi değildir.

## Colab kurulumu (bir kez)

1. Colab'da yeni bir defter aç.
2. İlk hücreye [`colab_baslat.py`](colab_baslat.py) dosyasının içeriğini yapıştır ve çalıştır.
   Drive bağlantısını onayla.
3. Diğer hücreler tek satırdır:

```python
calistir(2)    # temel tarama → BIST_Sistem.xlsx
calistir(10)   # Zirve Stratejisi: bu ayın listesi + gerçek sonuç takibi
```

Ayar değiştirmek için kodu düzenlemene gerek yok:

```python
calistir(10, STRATEJI_PARASI=20000)
calistir(10, HARIC=['ABCDE'], YENIDEN_SEC=True)
```

## Aylık rutin (ayın ilk işlem günü)

BAŞLAT hücresini çalıştır (hücre 1'i de o çalıştırır), sonra:

```python
calistir(2)
calistir(10)
```

Hücre 10'un yazdığı "listeden çıkan" hisseleri sat, "listeye giren"leri al.
Ay içinde bir şey yapma. Kurallar hücre 10'un başında ve Excel'deki "🧭 Zirve Stratejisi" sayfasında.

## Hücreler

| Dosya | Ne yapar | Yapay zekâ kredisi |
|---|---|---|
| `hucre1.py` | Drive bağlantısı, Türkiye saati, dosya yaşları | — |
| `hucre2.py` | Temel tarama, puanlama, hedef fiyat → `BIST_Sistem.xlsx` | — |
| `hucre3.py` | AL listesindeki hisselerin son 14 gün haberleri | — |
| `hucre4.py` | Teknik analiz ve Excel sayfaları | — |
| `hucre5.py` | Yapay zekâ kütüphanesini kurar (hücre 6 ve 8 için) | — |
| `hucre6.py` | Haberlerin yapay zekâ yorumu (önbellekli) | harcar |
| `hucre7.py` | Kapsamlı analiz: güven, DCF, giriş/çıkış, portföy | — |
| `hucre8.py` | Makro radar | harcar |
| `hucre9.py` | Backtest (geçmiş tarihlerde sistemi dener) | — |
| `hucre10.py` | **Zirve Stratejisi** — kullanılan strateji | — |

Zirve Stratejisi sadece **1 → 2 → 10** ister. 3–9 isteğe bağlı.

## Drive'daki dosyalar (`MyDrive/isyatirim`)

- İsyatırım'dan indirilenler: `temelozet.xlsx`, `temelfinansal.xlsx`, `temeltarihsel.xlsx`, `temelyabancioran.xlsx`, `temelsermaye.xlsx`, `temelperformans.xlsx`
- Sistemin yazdıkları: `BIST_Sistem.xlsx`, `BIST_Tam_KAP_v9_11mart.xlsx`, `zirve_kayit.csv`, `sistem_kayit.csv`
- Gerçek alım-satımlar: `gercek_portfoy.csv` (`Tarih,Kod,Islem,Adet,Fiyat`; Islem = AL / SAT)

API anahtarı koda yazılmaz; Colab Secrets'ta `ANTHROPIC_API_KEY` adıyla durur.
