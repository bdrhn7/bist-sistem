# ═══ HÜCRE 1 — DRIVE BAĞLANTISI + İSYATIRIM DOSYALARI + DOSYA KONTROLÜ  (v3 — 16.10.2026) ═══
#  v3: İsyatırım "Temel Değerler ve Oranlar" sayfasındaki 6 tablo 12 saatten eskiyse
#      otomatik indirilip Drive'a aynı isimlerle kaydediliyor (elle indirmeye gerek yok).
#      Zorla yenilemek için: calistir(1, ISY_ZORLA=True)
from google.colab import drive
drive.mount('/content/drive')

import os, re, time
from datetime import datetime
import requests
import pandas as pd
from bs4 import BeautifulSoup

# Saatler Türkiye saatiyle yazılsın (Colab normalde 3 saat geride)
os.environ['TZ'] = 'Europe/Istanbul'
time.tzset()

DRIVE = '/content/drive/MyDrive/isyatirim'
print('✅ Drive bağlandı!')
print(f'🕐 Saat: {datetime.now().strftime("%d.%m.%Y %H:%M")} (Türkiye)')

# ── İSYATIRIM DOSYALARI ──────────────────────────────────────
ISY_URL = 'https://www.isyatirim.com.tr/tr-tr/analiz/hisse/Sayfalar/Temel-Degerler-Ve-Oranlar.aspx'
ISY_DOSYALAR = ['temelozet', 'temelfinansal', 'temeltarihsel', 'temelyabancioran', 'temelsermaye', 'temelperformans']
ISY_YENILE_SAAT = 12      # dosyalar bundan eskiyse yeniden indir
ISY_ZORLA = False         # True: yaşına bakmadan indir

def _isy_deger(metin):
    """'2.693,3' → 2693.3 · '600.000.000' → 600000000 · 'A/D', '-', '08.05.2026', '6/2026' olduğu gibi."""
    m = metin.strip()
    if re.fullmatch(r'-?\d{1,3}(\.\d{3})+(,\d+)?|-?\d+(,\d+)?', m):
        return float(m.replace('.', '').replace(',', '.'))
    return m if m else None

def _isy_tablolar(html):
    soup = BeautifulSoup(html, 'lxml')
    tablolar = {}
    for t in soup.find_all('table', attrs={'data-csvname': True}):
        ad = t['data-csvname']
        if ad not in ISY_DOSYALAR or t.find('thead') is None or t.find('tbody') is None:
            continue
        basliklar = []
        for th in t.find('thead').find_all('tr')[-1].find_all('th'):
            sec = th.find('select')
            if sec is not None:                        # yabancı oranı: seçili tarih sütun adı olur
                o = sec.find('option', selected=True) or sec.find('option')
                basliklar.append(o.get_text(strip=True) if o else '')
            else:                                      # 'Kapanış<br/>(TL)' → 'Kapanış(TL)' (Excel düğmesiyle aynı)
                basliklar.append(''.join(x.strip() for x in th.find_all(string=True)))
        satirlar = [[_isy_deger(td.get_text()) for td in tr.find_all('td')] for tr in t.find('tbody').find_all('tr')]
        satirlar = [s for s in satirlar if len(s) == len(basliklar)]
        tablolar[ad] = pd.DataFrame(satirlar, columns=basliklar).sort_values('Kod', kind='stable').reset_index(drop=True)
    return tablolar

def isyatirim_indir(zorla=False):
    yollar = {a: f'{DRIVE}/{a}.xlsx' for a in ISY_DOSYALAR}
    yas = max((time.time() - os.path.getmtime(y)) / 3600 if os.path.exists(y) else 1e9 for y in yollar.values())
    if not zorla and yas < ISY_YENILE_SAAT:
        print(f'✅ İsyatırım dosyaları güncel (en eskisi {yas:.0f} saat önce indirilmiş)')
        return
    print('🔄 İsyatırım temel veriler indiriliyor...')
    try:
        r = requests.get(ISY_URL, timeout=60, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36',
            'Accept-Language': 'tr-TR,tr;q=0.9'})
        r.raise_for_status()
        tablolar = _isy_tablolar(r.content)
        eksik = [a for a in ISY_DOSYALAR if a not in tablolar or len(tablolar[a]) < 20 or 'Kod' not in tablolar[a].columns]
        if eksik:
            raise ValueError(f'sayfada bulunamayan ya da boş tablo: {", ".join(eksik)}')
    except Exception as e:
        print(f'⚠️ İsyatırım indirilemedi ({e}).')
        print('   Eski dosyalar kullanılacak. Sorun sürerse sayfadan Excel olarak elle indirip Drive\'a koy.')
        return
    for ad, d in tablolar.items():
        d.to_excel(yollar[ad], index=False, sheet_name='isyatirim')
    print('✅ İsyatırım dosyaları güncellendi: ' + ', '.join(f'{a} ({len(tablolar[a])})' for a in ISY_DOSYALAR))

isyatirim_indir(zorla=ISY_ZORLA)

# ── Klasördeki dosyalar ne kadar eski? ───────────────────────
print('\n📂 Klasördeki Excel dosyaları:')
dosyalar = sorted(f for f in os.listdir(DRIVE) if f.lower().endswith('.xlsx'))
for f in dosyalar:
    yol = os.path.join(DRIVE, f)
    gun = (time.time() - os.path.getmtime(yol)) / 86400
    tarih = datetime.fromtimestamp(os.path.getmtime(yol)).strftime('%d.%m.%Y')
    isaret = '✅' if gun <= 7 else '🟠' if gun <= 30 else '🔴'
    print(f'  {isaret} {f:40s} {tarih}  ({gun:.0f} gün önce)')

# Hücre 2 ve 7'nin okuduğu İsyatırım dosyaları yerinde mi?
eksik = [f'{a}.xlsx' for a in ISY_DOSYALAR if f'{a}.xlsx' not in dosyalar]
if eksik:
    print('\n⚠️ Eksik dosyalar:', ', '.join(eksik))
