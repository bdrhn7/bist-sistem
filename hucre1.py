# ═══ HÜCRE 1 — DRIVE BAĞLANTISI + DOSYA KONTROLÜ ═══
from google.colab import drive
drive.mount('/content/drive')

import os, time
from datetime import datetime

# Saatler Türkiye saatiyle yazılsın (Colab normalde 3 saat geride)
os.environ['TZ'] = 'Europe/Istanbul'
time.tzset()

DRIVE = '/content/drive/MyDrive/isyatirim'
print('✅ Drive bağlandı!')
print(f'🕐 Saat: {datetime.now().strftime("%d.%m.%Y %H:%M")} (Türkiye)')

# Klasördeki dosyalar ne kadar eski?
print('\n📂 Klasördeki Excel dosyaları:')
dosyalar = sorted(f for f in os.listdir(DRIVE) if f.lower().endswith('.xlsx'))
for f in dosyalar:
    yol = os.path.join(DRIVE, f)
    gun = (time.time() - os.path.getmtime(yol)) / 86400
    tarih = datetime.fromtimestamp(os.path.getmtime(yol)).strftime('%d.%m.%Y')
    isaret = '✅' if gun <= 7 else '🟠' if gun <= 30 else '🔴'
    print(f'  {isaret} {f:40s} {tarih}  ({gun:.0f} gün önce)')

# Hücre 7'nin okuduğu İsyatırım dosyaları yerinde mi?
gerekli = ['temelozet.xlsx', 'temelyabancioran.xlsx', 'temeltarihsel.xlsx',
           'temelfinansal.xlsx', 'temelsermaye.xlsx', 'temelperformans.xlsx']
eksik = [f for f in gerekli if f not in dosyalar]
if eksik:
    print('\n⚠️ Eksik dosyalar:', ', '.join(eksik))
print('\n🟠 = 7-30 gün, 🔴 = 30 günden eski. Eski İsyatırım dosyalarını yeniden indirip yükle.')
