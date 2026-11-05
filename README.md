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
| `hucre1.py` | Drive bağlantısı, Türkiye saati, İsyatırım dosyalarını otomatik indirme, dosya yaşları | — |
| `hucre2.py` | Temel tarama, puanlama, hedef fiyat → `BIST_Sistem.xlsx` (fiyat ve işlem hacmi İş Yatırım'dan, bilançolar Yahoo'dan) | — |
| `hucre3.py` | AL listesindeki hisselerin son 14 gün haberleri | — |
| `hucre4.py` | Teknik analiz ve Excel sayfaları | — |
| `hucre5.py` | Yapay zekâ kütüphanesini kurar (hücre 6 ve 8 için) | — |
| `hucre6.py` | Haberlerin yapay zekâ yorumu (önbellekli) | harcar |
| `hucre7.py` | Kapsamlı analiz: güven, DCF, giriş/çıkış, portföy | — |
| `hucre8.py` | Makro radar | harcar |
| `hucre9.py` | Backtest (geçmiş tarihlerde sistemi dener) | — |
| `hucre10.py` | **Zirve Stratejisi** — kullanılan strateji | — |

Zirve Stratejisi sadece **1 → 2 → 10** ister. 3–9 isteğe bağlı.

## Veri kaynakları

- **Fiyat geçmişi ve TL işlem hacmi (hücre 2, v2.6'dan beri):** İş Yatırım `HisseTekil` servisi — Borsa İstanbul
  resmi verisi, bölünme ve temettüye göre düzeltilmiş (`HGDG_` alanları). İş Yatırım bir hisse için cevap vermezse
  o hisse Yahoo'dan alınır; hangi kaynağın kullanıldığı `BIST_Sistem.xlsx`'teki **Fiyat Kaynağı** sütununda yazar.
  Yahoo bazı günleri bütün hisseler için boş bırakıyordu (15.10 ve 04.11.2026).
- **Bilanço, gelir tablosu, hisse bilgileri:** Yahoo (yfinance) + İş Yatırım temel tablolar (hücre 1).
- Hücre 4, 7 ve 9 fiyat geçmişini hâlâ Yahoo'dan alıyor.

## Drive'daki dosyalar (`MyDrive/isyatirim`)

- İsyatırım'dan gelenler: `temelozet.xlsx`, `temelfinansal.xlsx`, `temeltarihsel.xlsx`, `temelyabancioran.xlsx`, `temelsermaye.xlsx`, `temelperformans.xlsx` — hücre 1 bunları 12 saatten eskiyse "Temel Değerler ve Oranlar" sayfasından kendisi indirir (zorla: `calistir(1, ISY_ZORLA=True)`)
- Sistemin yazdıkları: `BIST_Sistem.xlsx`, `BIST_Tam_KAP_v9_11mart.xlsx`, `zirve_kayit.csv`, `sistem_kayit.csv`
- Gerçek alım-satımlar: `gercek_portfoy.csv` (`Tarih,Kod,Islem,Adet,Fiyat`; Islem = AL / SAT)

API anahtarı koda yazılmaz; Colab Secrets'ta `ANTHROPIC_API_KEY` adıyla durur.
