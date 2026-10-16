# ═══════════════════════════════════════════════════════════════
#  HÜCRE 2 — TEMEL TARAMA  (v2.5 — 16.10.2026)
#  v2 değişiklikleri:
#   - Hedef fiyat gerçek tarihi F/K ile: o yılın fiyatı ÷ o yılın HBK
#     (eskiden bugünkü fiyat ÷ eski HBK kullanılıyordu; hedef fiyatı
#      bugünkü fiyattan türetiyordu)
#   - Güncel kâr: son 4 çeyrek (yoksa son yıllık)
#   - Kâr ve ciro büyümesi enflasyondan arındırılıyor (reel büyüme)
#   - Likidite puanı hatası düzeltildi (50-200 mn TL artık ceza almıyor)
#   - Borç = finansal borç (ticari borçlar, karşılıklar dahil değil)
#   - Banka/sigorta/GYO'da borç, FCF ve faiz karşılama puanı yok
#   - KAP sinyali doğrulanana kadar puana katılmıyor (sütunda görünüyor)
#   v2.1:
#   - Hisse sayısı İsyatırım piyasa değeriyle kontrol ediliyor
#     (bedelsiz sonrası yfinance'ta eski kalan hisse sayısı F/K'yı bozuyordu)
#   - F/K 2'nin altındaysa veri şüpheli sayılıyor, hedef verilmiyor
#   - Hedef fiyat daha temkinli: F/K tavanları 15/20, gelecek yıl reel büyüme en fazla %10
#   - Seans açıkken bugünün yarım günlük verisi kullanılmıyor
#   v2.2:
#   - Dolar/euro ile rapor veren şirketlerin (THYAO, ENKAI, TAVHL...) rakamları
#     TL'ye çevriliyor; eskiden TL fiyat ÷ dolar kâr yapılıyordu (F/K 142, 556 gibi)
#   v2.3:
#   - Hedef F/K = şirketin kendi geçmiş F/K'sı ile bugünkü sektör F/K'sının ortalaması
#     Kötü: en düşük F/K / sektörün alt çeyreği · Orta: medyan / sektör medyanı ·
#     İyi: en yüksek F/K / sektörün üst çeyreği
#     (eskiden "Orta" hedef en kötü ile medyanın ortasıydı; hedefler sistematik
#      olarak düşük çıkıyordu)
#   - Değerleme puanı hedefin enflasyonu geçip geçmediğine bakıyor: 12 ay sonraki
#     hedef bugünkü paraya çevrilip fiyatla karşılaştırılıyor (Reel Potansiyel)
#   v2.4:
#   - Gelecek yıl büyümesi son döneme de bakıyor: son 4 çeyrek kârı son yıllığın
#     reel olarak altındaysa büyüme varsayılmıyor (eskiden sadece 2022-2025 ortalamasına
#     bakıyordu; kârı yeni düşmeye başlayan şirkete büyüme biçiyordu, ör. LMKDC)
#   - Sektör F/K'sı: bugünkü F/K medyanı ile geçmiş F/K medyanından DÜŞÜK olanı.
#     Bugünkü F/K kârı çöken şirketlerde şişiyor (çimentoda 16,6; geçmişte 12,4),
#     bazı sektörlerde de geçmiş F/K yüksek; iki ölçünün düşüğü temkinli taraf
#   - Yeni sütunlar: Son Dönem Reel Değişim (%), İleri Reel Büyüme (%)
#   v2.5:
#   - Yahoo fiyatı boş (NaN) satır gönderirse o satır atılıyor. 16.10 sabahı bütün
#     hisselerin Son Fiyat'ı boş geldi, puanlar çöktü (hiçbir hisse 55'i geçmedi)
#   - Yeni sütun: Veri Tarihi (fiyatın hangi günün kapanışı olduğu); sonda
#     fiyatların hangi günden olduğu ve beklenen günden eskiyse uyarı yazılıyor
# ═══════════════════════════════════════════════════════════════
import yfinance as yf
import pandas as pd
import numpy as np
import os
import time
import requests
from tqdm.notebook import tqdm
from datetime import datetime

DRIVE = '/content/drive/MyDrive/isyatirim'

# ── AYARLAR ──────────────────────────────────────────────────
# TÜİK yıllık TÜFE (Aralık, %). Her ocakta yeni yılı ekle.
YILLIK_TUFE = {2020: 14.60, 2021: 36.08, 2022: 64.27, 2023: 64.77,
               2024: 44.38, 2025: 30.89}
VARSAYILAN_TUFE    = 30.0   # tabloda olmayan yıl için (%)
BEKLENEN_ENFLASYON = 0.25   # önümüzdeki 12 ay için beklenen enflasyon (hedef fiyatta)
ENFLASYON_DUZELT   = True   # büyümeyi enflasyondan arındır
KAP_PUANA_KAT      = False  # KAP verisi doğrulanana kadar puana katma
ILERI_REEL_TAVAN   = 10     # hedef fiyatta gelecek yıl için en fazla reel büyüme (%)
FK_MIN_TAVAN       = 15     # kötü senaryo F/K'sı en fazla
FK_ORT_TAVAN       = 20     # orta senaryo F/K'sı en fazla
FK_IYI_TAVAN       = 25     # iyi senaryo F/K'sı en fazla
SEKTOR_MIN_HISSE   = 5      # sektör F/K'sı için en az kaç şirket (azsa tüm piyasa)
FK_SUPHELI         = 2      # F/K bunun altındaysa veri şüpheli

# ── İSYATIRIM DOSYALARI ──────────────────────────────────────
ozet     = pd.read_excel(f'{DRIVE}/temelozet.xlsx')
yabanci  = pd.read_excel(f'{DRIVE}/temelyabancioran.xlsx')
tarihsel = pd.read_excel(f'{DRIVE}/temeltarihsel.xlsx')
finansal = pd.read_excel(f'{DRIVE}/temelfinansal.xlsx')
print('✅ Drive dosyaları okundu!')

# ── HİSSE LİSTESİ — temelozet'ten otomatik ──────────────────
HISSELER = ozet['Kod'].dropna().astype(str).str.upper().str.strip().tolist()
HISSELER = list(dict.fromkeys(HISSELER))
print(f'✅ {len(HISSELER)} hisse temelozet.xlsx\'ten yüklendi')

# ── PİYASA DEĞERİ (İsyatırım) — hisse sayısı kontrolü için ─────
PD_ISY = {}
pd_kolon = next((c for c in ozet.columns if 'Piyasa Değeri' in str(c)), None)
if pd_kolon is not None:
    for k, v in zip(ozet['Kod'], pd.to_numeric(ozet[pd_kolon], errors='coerce')):
        if pd.notna(k) and pd.notna(v) and v > 0:
            PD_ISY[str(k).upper().strip()] = float(v) * 1_000_000   # mn TL → TL
print(f'✅ İsyatırım piyasa değeri: {len(PD_ISY)} hisse' if PD_ISY else '⚠️ İsyatırım piyasa değeri sütunu bulunamadı')

def tamamlanmis_gunler(df_h):
    """Fiyatı boş gelen satırları atar; seans kapanmadan (18:15'ten önce) bugünün
    satırlarını (yarım ya da boş mum) çıkarır."""
    try:
        if 'Close' in df_h.columns:
            df_h = df_h[df_h['Close'].notna() & (df_h['Close'] > 0)]
        simdi = datetime.now()
        if len(df_h) and (simdi.hour, simdi.minute) < (18, 15):
            df_h = df_h[[d.date() != simdi.date() for d in df_h.index]]
    except: pass
    return df_h

# ── DÖVİZ: dolar/euro ile rapor veren şirketler için ─────────
_KUR_CACHE = {}
def kur_serisi(para):
    """'USD' → USD/TRY günlük kapanış serisi (6 yıl). TRY için None."""
    para = (para or 'TRY').upper()
    if para in ('TRY', 'TL'):
        return None
    if para not in _KUR_CACHE:
        k = yf.Ticker(f'{para}TRY=X').history(period='6y')['Close']
        if getattr(k.index, 'tz', None) is not None:
            k.index = k.index.tz_localize(None)
        _KUR_CACHE[para] = k
    return _KUR_CACHE[para]

def kur(seri, tarih):
    """Verilen tarihteki (ya da öncesindeki son) kur. seri None ise 1."""
    if seri is None:
        return 1.0
    o = seri[seri.index <= pd.Timestamp(tarih)]
    return float(o.iloc[-1]) if len(o) else float(seri.iloc[0])

# ── SEKTÖR (İsyatırım) ──────────────────────────────────────
SEKTOR = {}
if 'Sektör' in ozet.columns:
    for k, s in zip(ozet['Kod'], ozet['Sektör']):
        if pd.notna(k):
            SEKTOR[str(k).upper().strip()] = str(s) if pd.notna(s) else ''

def sektor_turu(kod):
    s = SEKTOR.get(kod, '').lower()
    if 'gyo' in s or 'gayrimenkul' in s:
        return 'gyo'
    if any(k in s for k in ['banka', 'sigorta', 'kiralama', 'faktoring',
                             'aracı kurum', 'menkul', 'finansman', 'emeklilik']):
        return 'finans'
    return ''

# ── ENFLASYON YARDIMCILARI ──────────────────────────────────
def tufe(yil):
    return YILLIK_TUFE.get(int(yil), VARSAYILAN_TUFE) / 100

def enflasyon_carpani(yil_bas, yil_son):
    """yil_bas sonundan yil_son sonuna kadar fiyatlar kaç kat arttı."""
    c = 1.0
    for y in range(int(yil_bas) + 1, int(yil_son) + 1):
        c *= 1 + tufe(y)
    return c

def yillik_buyume(degerler, yillar):
    """degerler/yillar: en yeni başta. (nominal, reel) yıllık bileşik büyüme, %."""
    if len(degerler) < 2 or degerler[0] <= 0 or degerler[-1] <= 0:
        return None, None
    n = max(int(yillar[0]) - int(yillar[-1]), 1)
    oran = degerler[0] / degerler[-1]
    nominal = (oran ** (1 / n) - 1) * 100
    reel = ((oran / enflasyon_carpani(yillar[-1], yillar[0])) ** (1 / n) - 1) * 100
    return nominal, reel

# ── KAP İÇERİDEN ALIM — OTOMATİK ÇEK ───────────────────────
kap_alimlar = {}
KAP_ARSIV   = f'{DRIVE}/kap_alim_arsiv.xlsx'

def kap_otomatik_cek():
    """KAP.org.tr'den içeriden alım bildirimlerini otomatik çeker."""
    print('🔄 KAP\'tan içeriden alım çekiliyor...')
    url = 'https://www.kap.org.tr/tr/api/memberDisclosureQuery'
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Accept': 'application/json',
        'Referer': 'https://www.kap.org.tr/',
    }
    params = {
        'disclosureType': 'ODA',
        'subject': '2007',
        'yearFrom': '2024',
        'yearTo':   '2026',
    }
    try:
        r = requests.get(url, headers=headers, params=params, timeout=15)
        if r.status_code == 200:
            data = r.json()
            if isinstance(data, list) and len(data) > 0:
                print(f'✅ KAP API: {len(data)} bildirim geldi')
                return data
            print('⚠️ KAP API boş döndü, arşiv kullanılacak')
            return None
        print(f'⚠️ KAP API {r.status_code}, arşiv kullanılacak')
        return None
    except Exception as e:
        print(f'⚠️ KAP API bağlantı hatası: {e}, arşiv kullanılacak')
        return None

def kap_api_isle(data):
    """KAP API verisinden sembol bazlı bildirim sayısı çıkarır."""
    sonuc = {}
    for item in data:
        try:
            sembol = None
            if isinstance(item, dict):
                for alan in ['stockCodes', 'memberCode', 'companyCode', 'code']:
                    if alan in item and item[alan]:
                        s = str(item[alan]).strip().upper()
                        if len(s) >= 3:
                            sembol = s
                            break
            if sembol:
                sonuc.setdefault(sembol, {'islem_sayisi': 0})
                sonuc[sembol]['islem_sayisi'] += 1
        except:
            continue
    return sonuc

def kap_dosya_oku(yol):
    """İki biçimi de okur: (1) 'İlgili Şirketler' sütunu, (2) etiket | değer satırları."""
    sonuc = {}
    df_k = pd.read_excel(yol)
    if 'İlgili Şirketler' in df_k.columns:
        semboller = df_k['İlgili Şirketler'].dropna().astype(str)
    else:
        df_raw = pd.read_excel(yol, header=None)
        semboller = []
        if df_raw.shape[1] >= 2:
            for i in range(len(df_raw)):
                if str(df_raw.iloc[i, 0]).strip() == 'İlgili Şirketler':
                    semboller.append(str(df_raw.iloc[i, 1]))
    for s in semboller:
        s = s.strip().upper()
        if len(s) >= 3 and s.isalpha():
            sonuc.setdefault(s, {'islem_sayisi': 0})
            sonuc[s]['islem_sayisi'] += 1
    return sonuc

kap_data = kap_otomatik_cek()
if kap_data:
    kap_alimlar = kap_api_isle(kap_data)
    try:
        satirlar = [{'İlgili Şirketler': s} for s, v in kap_alimlar.items()
                    for _ in range(v['islem_sayisi'])]
        if satirlar:
            pd.DataFrame(satirlar).to_excel(KAP_ARSIV, index=False)
            print(f'✅ KAP arşiv güncellendi: {len(kap_alimlar)} hisse')
    except Exception as e:
        print(f'⚠️ KAP arşiv kayıt hatası: {e}')
else:
    KAP_DOSYA = f'{DRIVE}/kap_alim.xlsx'
    KAP_OKUMA = KAP_ARSIV if os.path.exists(KAP_ARSIV) else KAP_DOSYA
    if os.path.exists(KAP_OKUMA):
        try:
            kap_alimlar = kap_dosya_oku(KAP_OKUMA)
            print(f'✅ KAP dosyadan yüklendi: {len(kap_alimlar)} hisse ({os.path.basename(KAP_OKUMA)})')
        except Exception as e:
            print(f'⚠️ KAP dosyası okunamadı: {e}')
    else:
        print('ℹ️  KAP verisi bulunamadı (API başarısız, dosya yok)')

print(f'✅ KAP: {len(kap_alimlar)} hissede bildirim'
      + ('' if KAP_PUANA_KAT else '  (puana katılmıyor, sadece bilgi)'))
print(f'\n⏳ Tarama başlıyor ({len(HISSELER)} hisse)...\n')

sonuc           = []
hatali_hisseler = []

for h in tqdm(HISSELER, desc='Taranıyor'):
    try:
        t    = yf.Ticker(f'{h}.IS')
        info = t.info
        inc  = t.income_stmt
        bal  = t.balance_sheet
        cf   = t.cashflow
        hist = t.history(period='5y')     # tarihi F/K için 5 yıl
        tur  = sektor_turu(h)

        if hist is None or len(hist) < 51:
            hatali_hisseler.append({'Kod': h, 'Sebep': 'Fiyat verisi yetersiz'}); continue
        hist = tamamlanmis_gunler(hist)
        if inc is None or inc.empty:
            hatali_hisseler.append({'Kod': h, 'Sebep': 'Finansal veri yok'}); continue

        # Raporlama para birimi (THYAO dolar, TAVHL euro gibi)
        rapor_para = (info.get('financialCurrency') or 'TRY').upper()
        try:
            kur_s = kur_serisi(rapor_para)
        except Exception:
            hatali_hisseler.append({'Kod': h, 'Sebep': f'{rapor_para}/TRY kuru alınamadı'}); continue

        # Hisse sayısı: yfinance, İsyatırım piyasa değeriyle karşılaştırılır
        shares = info.get('sharesOutstanding', 0) or 0
        hisse_kaynak = 'yfinance'
        fiyat_simdi = hist['Close'].iloc[-1]
        if h in PD_ISY and fiyat_simdi > 0:
            isy_hisse = PD_ISY[h] / fiyat_simdi
            if shares <= 0:
                shares = isy_hisse; hisse_kaynak = 'İsyatırım'
            else:
                oran = isy_hisse / shares
                # %25'ten fazla fark varsa (bedelsiz/bedelli sonrası eski veri) İsyatırım'ı kullan;
                # 20 kattan fazla farkta birim hatası olabilir, dokunma
                if 0.05 <= oran <= 20 and abs(oran - 1) > 0.25:
                    shares = isy_hisse; hisse_kaynak = f'İsyatırım (yf x{oran:.1f})'
        if shares < 1000:
            hatali_hisseler.append({'Kod': h, 'Sebep': 'Hisse sayısı hatalı'}); continue

        kapanis = hist['Close'].copy()
        if getattr(kapanis.index, 'tz', None) is not None:
            kapanis.index = kapanis.index.tz_localize(None)

        son_fiyat  = hist['Close'].iloc[-1]
        yuksek_52h = hist['High'].iloc[-252:].max()
        dusuk_52h  = hist['Low'].iloc[-252:].min()
        zirveden   = ((son_fiyat - yuksek_52h) / yuksek_52h) * 100
        dipten     = ((son_fiyat - dusuk_52h)  / dusuk_52h)  * 100
        yil_ici    = ((son_fiyat / hist['Close'].iloc[-252]) - 1) * 100 if len(hist) >= 252 else 0
        aylik      = ((son_fiyat / hist['Close'].iloc[-21])  - 1) * 100 if len(hist) >= 21  else 0
        haftalik   = ((son_fiyat / hist['Close'].iloc[-5])   - 1) * 100 if len(hist) >= 5   else 0
        gunluk     = ((son_fiyat / hist['Close'].iloc[-2])   - 1) * 100 if len(hist) >= 2   else 0
        son3       = ((son_fiyat / hist['Close'].iloc[-4])   - 1) * 100 if len(hist) >= 4   else 0

        son_1ay_adet = hist['Volume'].iloc[-20:].mean()
        onceki_adet  = hist['Volume'].iloc[-60:-20].mean()
        hacim_artis  = ((son_1ay_adet - onceki_adet) / onceki_adet * 100) if onceki_adet > 0 else 0
        tl_hacim_ort = (hist['Close'] * hist['Volume']).iloc[-20:].mean() / 1_000_000
        if tl_hacim_ort >= 200:  likidite = '🟢 Yüksek'
        elif tl_hacim_ort >= 50: likidite = '🟢 Normal'
        elif tl_hacim_ort >= 10: likidite = '🟡 Düşük'
        else:                    likidite = '🔴 Çok Düşük'

        # ── Yıllık net kârlar (en yeni başta) ──
        net_karlar = []; kar_tarihleri = []
        if 'Net Income' in inc.index:
            for col in inc.columns[:4]:
                v = inc.loc['Net Income', col]
                if pd.notna(v) and float(v) != 0:
                    net_karlar.append(float(v) * kur(kur_s, col)); kar_tarihleri.append(pd.Timestamp(col))
        if len(net_karlar) < 2:
            hatali_hisseler.append({'Kod': h, 'Sebep': 'Kar verisi yetersiz'}); continue
        kar_yillari = [d.year for d in kar_tarihleri]
        hbk_gecmis  = [nk / shares for nk in net_karlar]

        # ── Güncel HBK: son 4 çeyrek, olmazsa son yıllık ──
        hbk_guncel = hbk_gecmis[0]; hbk_kaynak = 'Yıllık'
        q_son_tarih = None
        try:
            q = t.quarterly_income_stmt
            if q is not None and not q.empty and 'Net Income' in q.index:
                qs = q.loc['Net Income'].iloc[:4]
                qv = [float(v) * kur(kur_s, d) for d, v in qs.items() if pd.notna(v)]
                if len(qv) == 4:
                    ttm = sum(qv) / shares
                    q_son_tarih = pd.Timestamp(qs.index[0])
                    # Akıl kontrolü: yıllıktan çok saparsa çeyrek verisi bozuk say
                    if hbk_gecmis[0] > 0 and 0.2 <= ttm / hbk_gecmis[0] <= 2.5:
                        hbk_guncel = ttm; hbk_kaynak = 'Son 4 çeyrek'
                    elif hbk_gecmis[0] <= 0:
                        hbk_guncel = ttm; hbk_kaynak = 'Son 4 çeyrek'
        except: pass

        if hbk_guncel <= 0:
            hatali_hisseler.append({'Kod': h, 'Sebep': 'Negatif HBK'}); continue
        guncel_fk = son_fiyat / hbk_guncel

        # ── Gerçek tarihi F/K: o yılın kapanış fiyatı ÷ o yılın HBK ──
        fk_tarihi = []
        for tarih, hbk in zip(kar_tarihleri, hbk_gecmis):
            if hbk <= 0:
                continue
            o_gun = kapanis[kapanis.index <= tarih]
            if len(o_gun) == 0 or (tarih - o_gun.index[-1]).days > 10:
                continue          # o tarihte fiyat verisi yok
            fk_tarihi.append(o_gun.iloc[-1] / hbk)

        if len(fk_tarihi) >= 2:
            fk_min    = min(min(fk_tarihi), FK_MIN_TAVAN)
            fk_medyan = min(float(np.median(fk_tarihi)), FK_ORT_TAVAN)
            fk_med_ham = float(np.median(fk_tarihi))          # sektör istatistiği için (tavansız)
            fk_max    = min(max(fk_tarihi), 35)
        else:
            fk_min = fk_medyan = fk_max = fk_med_ham = None

        # ── Büyüme (nominal ve reel) ──
        buyume, reel_buyume = yillik_buyume(net_karlar, kar_yillari)
        buyume_yok = buyume is None          # baş ya da son yıl zararda
        if buyume_yok:
            buyume, reel_buyume = 0.0, 0.0
        # puanlamada kullanılan büyüme (hesaplanamıyorsa ceza alsın)
        b_puan = -100.0 if buyume_yok else (reel_buyume if ENFLASYON_DUZELT else buyume)

        buyume_tutarli_skor = 'Veri Yok'
        if len(net_karlar) >= 3:
            b_yil = 0
            for i in range(len(net_karlar) - 1):
                esik = net_karlar[i + 1] * (1 + tufe(kar_yillari[i])) if ENFLASYON_DUZELT and net_karlar[i + 1] > 0 \
                       else net_karlar[i + 1]
                if net_karlar[i] > esik:
                    b_yil += 1
            oran = b_yil / (len(net_karlar) - 1)
            if oran == 1.0:    buyume_tutarli_skor = '🟢 Sürekli Büyüme'
            elif oran >= 0.75: buyume_tutarli_skor = '🟢 Güçlü Büyüme'
            elif oran >= 0.5:  buyume_tutarli_skor = '🟡 Kararsız'
            else:              buyume_tutarli_skor = '🔴 Tutarsız'

        roe_tutarli_skor = 'Veri Yok'; roe_ort = None
        try:
            if bal is not None and not bal.empty and 'Net Income' in inc.index:
                roe_l = []
                for i, col in enumerate(inc.columns[:4]):
                    ni = inc.loc['Net Income', col]
                    if i < len(bal.columns) and 'Common Stock Equity' in bal.index:
                        eq = bal.loc['Common Stock Equity', bal.columns[i]]
                        if pd.notna(ni) and pd.notna(eq) and float(eq) > 0:
                            roe_l.append(float(ni) / float(eq) * 100)
                if len(roe_l) >= 2:
                    roe_ort = np.mean(roe_l)
                    roe_cv  = np.std(roe_l) / abs(roe_ort) if roe_ort != 0 else 999
                    if roe_ort >= 15 and roe_cv < 0.3:   roe_tutarli_skor = '🟢 Yüksek & Tutarlı'
                    elif roe_ort >= 10 and roe_cv < 0.5: roe_tutarli_skor = '🟢 İyi ROE'
                    elif roe_ort >= 5:                   roe_tutarli_skor = '🟡 Orta ROE'
                    elif roe_ort > 0:                    roe_tutarli_skor = '🟠 Düşük ROE'
                    else:                                roe_tutarli_skor = '🔴 Negatif ROE'
        except: pass

        peg = None; peg_skor = 'Veri Yok'
        if guncel_fk > 0 and b_puan > 5:
            peg = round(guncel_fk / b_puan, 2)
            if peg < 0.5:    peg_skor = '🟢 Çok Ucuz'
            elif peg < 1.0:  peg_skor = '🟢 Ucuz'
            elif peg < 1.5:  peg_skor = '🟡 Makul'
            elif peg < 2.0:  peg_skor = '🟠 Pahalı'
            else:            peg_skor = '🔴 Çok Pahalı'

        kap_islem_sayisi = kap_alimlar.get(h, {}).get('islem_sayisi', 0)
        if kap_islem_sayisi >= 3:   kap_sinyal = '🟢 Güçlü Alım'
        elif kap_islem_sayisi >= 1: kap_sinyal = '🟡 Alım Var'
        else:                       kap_sinyal = '-'

        # ── HEDEF FİYAT: 12 ay sonraki HBK × tarihi F/K ──
        # Son dönem: son 4 çeyrek kârı, son yıllık kârla reel olarak karşılaştırılır
        reel_son = None
        if hbk_kaynak == 'Son 4 çeyrek' and q_son_tarih is not None and hbk_gecmis[0] > 0:
            ara_yil = (q_son_tarih - kar_tarihleri[0]).days / 365.25
            if ara_yil >= 0.2:
                enf = (1 + tufe(kar_tarihleri[0].year + 1)) ** ara_yil
                reel_son = ((hbk_guncel / hbk_gecmis[0]) / enf - 1) * 100
        # Gelecek yıl reel büyüme: uzun dönem ortalaması ile son dönemden düşük olanı
        reel_ileri = reel_buyume if reel_son is None else min(reel_buyume, reel_son)
        reel_ileri = max(min(reel_ileri, ILERI_REEL_TAVAN), -20)
        if ENFLASYON_DUZELT:
            g_ileri = (1 + reel_ileri / 100) * (1 + BEKLENEN_ENFLASYON) - 1
        else:
            g_ileri = max(min(buyume, 25), -20) / 100
        tahmini_hbk = hbk_guncel * (1 + g_ileri)
        veri_uyari = ''
        if guncel_fk < FK_SUPHELI:
            veri_uyari = f'F/K {guncel_fk:.1f}: veri şüpheli (tek seferlik kâr ya da hisse sayısı hatası)'
        # Hedef fiyat döngüden sonra (sektör F/K'ları belli olunca) hesaplanıyor

        pd_dd = info.get('priceToBook', None)
        roe   = (info.get('returnOnEquity', 0) or 0) * 100

        ciro_buyume = None; reel_ciro = None
        try:
            if 'Total Revenue' in inc.index:
                cirolar = []; ciro_yil = []
                for col in inc.columns[:4]:
                    v = inc.loc['Total Revenue', col]
                    if pd.notna(v) and float(v) > 0:
                        cirolar.append(float(v) * kur(kur_s, col)); ciro_yil.append(pd.Timestamp(col).year)
                nom, reel = yillik_buyume(cirolar, ciro_yil)
                if nom is not None:
                    ciro_buyume = round(nom, 1); reel_ciro = round(reel, 1)
        except: pass

        faiz_karsilama = None; faiz_kategori = 'Veri Yok'
        if tur:
            faiz_kategori = 'Uygulanmaz'
        else:
            try:
                favok = None; faiz_gider = None
                for k in ['EBITDA', 'Operating Income']:
                    if k in inc.index:
                        v = inc.loc[k].iloc[0]
                        if pd.notna(v): favok = float(v); break
                for k in ['Interest Expense', 'Interest Expense Non Operating']:
                    if k in inc.index:
                        v = inc.loc[k].iloc[0]
                        if pd.notna(v) and float(v) != 0:
                            faiz_gider = abs(float(v)); break
                if favok and faiz_gider and faiz_gider > 0:
                    faiz_karsilama = round(favok / faiz_gider, 1)
                    if faiz_karsilama >= 5:   faiz_kategori = '🟢 Güçlü'
                    elif faiz_karsilama >= 3: faiz_kategori = '🟡 Yeterli'
                    elif faiz_karsilama >= 1: faiz_kategori = '🟠 Riskli'
                    else:                    faiz_kategori = '🔴 Kritik'
            except: pass

        fcf = None; fcf_verim = None; fcf_kategori = 'Veri Yok'
        if tur:
            fcf_kategori = 'Uygulanmaz'
        else:
            try:
                if cf is not None and not cf.empty:
                    op_cf = None; capex = None
                    for k in ['Operating Cash Flow', 'Cash From Operations']:
                        if k in cf.index:
                            v = cf.loc[k].iloc[0]
                            if pd.notna(v): op_cf = float(v); break
                    for k in ['Capital Expenditure', 'Purchase Of PPE', 'Capital Expenditures']:
                        if k in cf.index:
                            v = cf.loc[k].iloc[0]
                            if pd.notna(v): capex = abs(float(v)); break
                    if op_cf is not None and capex is not None:
                        fcf_raw       = (op_cf - capex) * kur(kur_s, cf.columns[0])
                        piyasa_degeri = son_fiyat * shares
                        if piyasa_degeri > 0:
                            fcf_verim = round((fcf_raw / piyasa_degeri) * 100, 1)
                        fcf = round(fcf_raw / 1_000_000, 1)
                        if fcf_verim is not None:
                            if fcf_verim >= 8:   fcf_kategori = '🟢 Güçlü FCF'
                            elif fcf_verim >= 3: fcf_kategori = '🟢 Pozitif FCF'
                            elif fcf_verim >= 0: fcf_kategori = '🟡 Zayıf FCF'
                            else:                fcf_kategori = '🔴 Negatif FCF'
            except: pass

        # ── Borç: sadece finansal borç (Total Debt) ──
        borc_ozsermaye = None; net_borc_ozserm = None; borc_riski = 'Veri Yok'
        if tur:
            borc_riski = 'Uygulanmaz'
        else:
            try:
                if bal is not None and not bal.empty:
                    tb = None; nakit = 0.0; oz = None
                    if 'Total Debt' in bal.index and pd.notna(bal.loc['Total Debt'].iloc[0]):
                        tb = float(bal.loc['Total Debt'].iloc[0])
                    if 'Common Stock Equity' in bal.index and pd.notna(bal.loc['Common Stock Equity'].iloc[0]):
                        oz = float(bal.loc['Common Stock Equity'].iloc[0])
                    for k in ['Cash Cash Equivalents And Short Term Investments', 'Cash And Cash Equivalents']:
                        if k in bal.index and pd.notna(bal.loc[k].iloc[0]):
                            nakit = float(bal.loc[k].iloc[0]); break
                    if tb is not None and oz and oz > 0:
                        borc_ozsermaye  = round(tb / oz, 2)
                        net_borc_ozserm = round((tb - nakit) / oz, 2)
                    if net_borc_ozserm is not None:
                        if net_borc_ozserm < 0:      borc_riski = '🟢 Net Nakit'
                        elif net_borc_ozserm < 0.5:  borc_riski = '🟢 Düşük Borç'
                        elif net_borc_ozserm < 1.0:  borc_riski = '🟡 Orta Borç'
                        elif net_borc_ozserm < 2.0:  borc_riski = '🟠 Yüksek Borç'
                        else:                         borc_riski = '🔴 Tehlikeli Borç'
            except: pass

        # ── SKOR ───────────────────────────────────────────
        skor = 0      # değerleme puanı (+30 / −15) döngüden sonra eklenecek
        if zirveden >= -5:    skor += 12
        elif zirveden >= -15: skor += 8
        elif zirveden >= -25: skor += 4
        elif zirveden >= -40: skor += 1
        else:                 skor -= 6
        if aylik >= 15:     skor += 6
        elif aylik >= 8:    skor += 4
        elif aylik >= 3:    skor += 2
        elif aylik < 0:     skor -= 3
        if haftalik >= 5:   skor += 3
        elif haftalik >= 2: skor += 2
        elif haftalik < 0:  skor -= 1
        if hacim_artis >= 60:   skor += 6
        elif hacim_artis >= 25: skor += 4
        elif hacim_artis >= 0:  skor += 1
        else:                   skor -= 2
        # Büyüme: reel ise 20/10/0 eşikleri, nominal ise eski 40/20/5
        e1, e2, e3 = (20, 10, 0) if ENFLASYON_DUZELT else (40, 20, 5)
        if b_puan >= e1:   skor += 15
        elif b_puan >= e2: skor += 10
        elif b_puan >= e3: skor += 5
        else:              skor -= 8
        if buyume_tutarli_skor == '🟢 Sürekli Büyüme':  skor += 15
        elif buyume_tutarli_skor == '🟢 Güçlü Büyüme':  skor += 10
        elif buyume_tutarli_skor == '🔴 Tutarsız':       skor -= 12
        if peg is not None:
            if peg < 0.5:    skor += 8
            elif peg < 1.0:  skor += 5
            elif peg < 1.5:  skor += 2
            elif peg < 2.0:  skor -= 3
            else:            skor -= 8
        if roe_tutarli_skor == '🟢 Yüksek & Tutarlı':  skor += 8
        elif roe_tutarli_skor == '🟢 İyi ROE':          skor += 6
        elif roe_tutarli_skor == '🟡 Orta ROE':         skor += 2
        elif roe_tutarli_skor == '🟠 Düşük ROE':        skor -= 4
        elif roe_tutarli_skor == '🔴 Negatif ROE':      skor -= 10
        c_puan = reel_ciro if ENFLASYON_DUZELT else ciro_buyume
        c1, c2 = (15, 5) if ENFLASYON_DUZELT else (30, 15)
        if c_puan is not None:
            if c_puan >= c1:   skor += 6
            elif c_puan >= c2: skor += 4
            elif c_puan >= 0:  skor += 1
            else:              skor -= 4
        if net_borc_ozserm is not None:
            if net_borc_ozserm < 0:      skor += 10
            elif net_borc_ozserm < 0.5:  skor += 6
            elif net_borc_ozserm < 1.0:  skor += 0
            elif net_borc_ozserm < 2.0:  skor -= 12
            else:                         skor -= 25
        if faiz_karsilama is not None:
            if faiz_karsilama >= 5:   skor += 10
            elif faiz_karsilama >= 3: skor += 6
            elif faiz_karsilama >= 1: skor += 0
            elif faiz_karsilama > 0:  skor -= 8
            else:                     skor -= 15
        if fcf_verim is not None:
            if fcf_verim >= 8:   skor += 15
            elif fcf_verim >= 3: skor += 10
            elif fcf_verim >= 0: skor += 2
            else:                skor -= 12
        # Likidite (düzeltildi: 50-200 mn TL artık ceza almıyor)
        if tl_hacim_ort >= 200:  skor += 5
        elif tl_hacim_ort >= 50: skor += 0
        elif tl_hacim_ort >= 10: skor -= 10
        else:                    skor -= 25
        if son3 <= -5:   skor -= 4
        elif son3 <= -3: skor -= 2
        if KAP_PUANA_KAT:
            if kap_islem_sayisi >= 3:   skor += 12
            elif kap_islem_sayisi >= 1: skor += 6
        tavan5 = zirveden <= -60 and aylik < 0      # sert düşen hisse: skor en fazla 5

        sonuc.append({
            'Sembol':               h,
            'Son Fiyat':            round(son_fiyat, 2),
            'Veri Tarihi':          pd.Timestamp(hist.index[-1]).strftime('%Y-%m-%d'),
            'Hedef (Kötü)':         None,
            'Hedef (Ort.)':         None,
            'Hedef (İyi)':          None,
            'Hedefe Uzaklık (%)':   0.0,
            'Toplam Skor':          skor,
            'Güncel F/K':           round(guncel_fk, 1) if guncel_fk > 0 else None,
            'HBK Kaynağı':          hbk_kaynak,
            'Rapor Para Birimi':    rapor_para,
            'Hisse Sayısı Kaynağı': hisse_kaynak,
            'Veri Uyarısı':         veri_uyari,
            'Tarihi F/K Ort.':      round(fk_medyan, 1) if fk_medyan else None,
            'Tarihi F/K Min':       round(fk_min, 1) if fk_min else None,
            'Tarihi F/K Max':       round(fk_max, 1) if fk_max else None,
            'PD/DD':                round(pd_dd, 2) if pd_dd else None,
            'ROE (%)':              round(roe, 1) if roe else None,
            'Kar Büyümesi (%)':     round(buyume, 1),
            'Reel Kar Büyümesi (%)': round(reel_buyume, 1),
            'Son Dönem Reel Değişim (%)': round(reel_son, 1) if reel_son is not None else None,
            'İleri Reel Büyüme (%)': round(reel_ileri, 1),
            'Büyüme Tutarlılığı':   buyume_tutarli_skor,
            'PEG Oranı':            peg,
            'PEG Skor':             peg_skor,
            'ROE Tutarlılığı':      roe_tutarli_skor,
            'ROE Ort. (%)':         round(roe_ort, 1) if roe_ort else None,
            'KAP Alım Sinyali':     kap_sinyal,
            'KAP Alım Sayısı':      kap_islem_sayisi if kap_islem_sayisi > 0 else None,
            'Ciro Büyümesi (%)':    ciro_buyume,
            'Reel Ciro Büyümesi (%)': reel_ciro,
            'Faiz Karşılama':       faiz_karsilama,
            'Faiz Karşılama Kat.':  faiz_kategori,
            'FCF (mn TL)':          fcf,
            'FCF Verimi (%)':       fcf_verim,
            'FCF Kategori':         fcf_kategori,
            'Yıl İçi (%)':          round(yil_ici, 1),
            'Aylık (%)':            round(aylik, 1),
            'Haftalık (%)':         round(haftalik, 1),
            'Günlük (%)':           round(gunluk, 1),
            'Son 3 Gün (%)':        round(son3, 1),
            'Zirveden (%)':         round(zirveden, 1),
            'Dipten (%)':           round(dipten, 1),
            'Hacim Artışı (%)':     round(hacim_artis, 1),
            'Ort. TL Hacim (mn)':   round(tl_hacim_ort, 1),
            'Likidite':             likidite,
            'Borç/Özsermaye':       borc_ozsermaye,
            'Net Borç/Özsermaye':   net_borc_ozserm,
            'Borç Riski':           borc_riski,
            'İSY Sektör':           SEKTOR.get(h, ''),
            '_tahmini_hbk': tahmini_hbk, '_fk_min_t': fk_min, '_fk_med_t': fk_medyan, '_fk_max_t': fk_max, '_fk_med_ham': fk_med_ham,
            '_veri_uyari': veri_uyari, '_tavan5': tavan5,
        })
        time.sleep(0.35)

    except Exception as e:
        hatali_hisseler.append({'Kod': h, 'Sebep': str(e)[:80]})
        continue

# ── SONUÇ ────────────────────────────────────────────────────
df = pd.DataFrame(sonuc)

# ── HEDEF FİYAT (2. aşama): kendi geçmiş F/K'sı + bugünkü sektör F/K'sı ──
# Sektör F/K'sı: bugünkü F/K ile geçmiş F/K medyanlarından düşük olanı (temkinli)
def _fk_istat(seri):
    seri = pd.to_numeric(seri, errors='coerce').dropna()
    return (float(seri.quantile(0.25)), float(seri.median()), float(seri.quantile(0.75)))
_temiz = df['_veri_uyari'] == ''
g_bugun  = df[_temiz & df['Güncel F/K'].between(2, 60)]
g_gecmis = df[_temiz & pd.to_numeric(df['_fk_med_ham'], errors='coerce').between(2, 60)]
def _sektor_istat(grup_b, grup_g, alt=SEKTOR_MIN_HISSE):
    adaylar = []
    if len(grup_b) >= alt: adaylar.append(_fk_istat(grup_b['Güncel F/K']))
    if len(grup_g) >= alt: adaylar.append(_fk_istat(grup_g['_fk_med_ham']))
    if not adaylar:
        return None
    return tuple(min(a[j] for a in adaylar) for j in range(3))
PIYASA_FK = _sektor_istat(g_bugun, g_gecmis, 1) or (7.0, 10.0, 14.0)
sek = {}
for ad in set(df['İSY Sektör'].dropna()):
    if not ad:
        continue
    st = _sektor_istat(g_bugun[g_bugun['İSY Sektör'] == ad], g_gecmis[g_gecmis['İSY Sektör'] == ad])
    if st:
        sek[ad] = st
print(f'\n📐 Piyasa F/K (bugünkü ve geçmişin düşüğü): alt çeyrek {PIYASA_FK[0]:.1f} | medyan {PIYASA_FK[1]:.1f}'
      f' | üst çeyrek {PIYASA_FK[2]:.1f} | kendi sektörüyle karşılaştırılan sektör: {len(sek)}')
if 'Çimento' in sek:
    print(f"   örnek: Çimento sektörü F/K medyanı {sek['Çimento'][1]:.1f}")

def _ort(*xs):
    xs = [float(x) for x in xs if x is not None and pd.notna(x) and x > 0]
    return float(np.mean(xs)) if xs else None

for col in ['Reel Potansiyel (%)', 'Sektör F/K Medyan', 'Hedef F/K (Kötü)', 'Hedef F/K (Orta)', 'Hedef F/K (İyi)']:
    df[col] = np.nan
for i, r in df.iterrows():
    s_p25, s_med, s_p75 = sek.get(r['İSY Sektör'], PIYASA_FK)
    df.at[i, 'Sektör F/K Medyan'] = round(s_med, 1)
    skor = r['Toplam Skor']
    thbk = r['_tahmini_hbk']; son_fiyat = r['Son Fiyat']
    if (not r['_veri_uyari'] and pd.notna(thbk) and thbk > 0
            and pd.notna(r['_fk_min_t']) and r['_fk_min_t']):
        fk_orta = min(_ort(r['_fk_med_t'], s_med), FK_ORT_TAVAN)
        fk_kotu = min(_ort(r['_fk_min_t'], s_p25), FK_MIN_TAVAN, fk_orta)
        fk_iyi  = max(min(_ort(r['_fk_max_t'], s_p75), FK_IYI_TAVAN), fk_orta)
        hedef_kotu = min(thbk * fk_kotu, son_fiyat * 3)
        hedef_ort  = min(thbk * fk_orta, son_fiyat * 3)
        hedef_iyi  = min(thbk * fk_iyi,  son_fiyat * 3)
        df.at[i, 'Hedef (Kötü)'] = round(hedef_kotu, 2)
        df.at[i, 'Hedef (Ort.)'] = round(hedef_ort, 2)
        df.at[i, 'Hedef (İyi)']  = round(hedef_iyi, 2)
        df.at[i, 'Hedefe Uzaklık (%)'] = round((hedef_ort - son_fiyat) / son_fiyat * 100, 1)
        df.at[i, 'Hedef F/K (Kötü)'] = round(fk_kotu, 1)
        df.at[i, 'Hedef F/K (Orta)'] = round(fk_orta, 1)
        df.at[i, 'Hedef F/K (İyi)']  = round(fk_iyi, 1)
        # Puan: 12 ay sonraki hedefler bugünkü paraya çevrilip fiyatla karşılaştırılıyor.
        # Sadece enflasyon kadar artış beklenen hisse "ucuz" sayılmıyor.
        ind = 1 + BEKLENEN_ENFLASYON if ENFLASYON_DUZELT else 1
        df.at[i, 'Reel Potansiyel (%)'] = round((hedef_ort / ind - son_fiyat) / son_fiyat * 100, 1)
        if son_fiyat < hedef_kotu / ind:   skor += 30
        elif son_fiyat < hedef_ort / ind:  skor += 22
        elif son_fiyat < hedef_iyi / ind:  skor += 10
        else:                              skor -= 15
    if r['_tavan5']:
        skor = min(skor, 5)
    df.at[i, 'Toplam Skor'] = int(skor)

df = df.drop(columns=['_tahmini_hbk', '_fk_min_t', '_fk_med_t', '_fk_max_t', '_fk_med_ham', '_veri_uyari', '_tavan5'])
_k = [c for c in df.columns if c != 'Reel Potansiyel (%)']
_k.insert(_k.index('Hedefe Uzaklık (%)') + 1, 'Reel Potansiyel (%)')
df = df[_k]
df = df.sort_values('Toplam Skor', ascending=False).reset_index(drop=True)

print(f'\n✅ {len(df)} hisse tamamlandı!')
print(f'⚠️  {len(hatali_hisseler)} hisse hata verdi')

if hatali_hisseler:
    df_hatali = pd.DataFrame(hatali_hisseler)
    print('\n--- HATA SEBEPLERİ ---')
    print(df_hatali['Sebep'].value_counts().head(10).to_string())
    df_hatali.to_excel(f'{DRIVE}/BIST_Hatali.xlsx', index=False)
    print(f'\n✅ Hatalı hisseler BIST_Hatali.xlsx\'e kaydedildi')

print(f'\n--- KAP ---')
print(df['KAP Alım Sinyali'].value_counts().to_string())
print(f'\n--- HBK KAYNAĞI ---')
print(df['HBK Kaynağı'].value_counts().to_string())
dovizli = df[df['Rapor Para Birimi'] != 'TRY']
if len(dovizli):
    print(f'\n💱 Döviz ile rapor veren ({len(dovizli)}), TL\'ye çevrildi:',
          ', '.join(f"{a}({b})" for a, b in zip(dovizli['Sembol'], dovizli['Rapor Para Birimi'])))
print(f'\n--- HİSSE SAYISI KAYNAĞI ---')
print(df['Hisse Sayısı Kaynağı'].str.split(' ').str[0].value_counts().to_string())
duzeltilen = df[df['Hisse Sayısı Kaynağı'].str.contains('yf x', na=False)]
if len(duzeltilen):
    print('   Hisse sayısı düzeltilenler:', ', '.join(duzeltilen['Sembol'].head(15)))
supheli = df[df['Veri Uyarısı'] != '']
if len(supheli):
    print(f'⚠️ Verisi şüpheli ({len(supheli)}):', ', '.join(supheli['Sembol'].head(15)))
# Fiyatlar hangi günün kapanışı?
_simdi = datetime.now()
_beklenen = pd.Timestamp(_simdi.date())
if not (_beklenen.weekday() < 5 and (_simdi.hour, _simdi.minute) >= (18, 15)):
    _beklenen = _beklenen - pd.offsets.BDay(1)
_vt = pd.to_datetime(df['Veri Tarihi'])
_en_cok = _vt.mode().iloc[0]
print(f"\n📅 Fiyatlar {_en_cok:%d.%m.%Y} kapanışından ({int((_vt == _en_cok).sum())}/{len(df)} hisse)")
if _en_cok < _beklenen:
    print(f"⚠️  Beklenen son kapanış {_beklenen:%d.%m.%Y}. Yahoo o günün fiyatlarını henüz vermemiş olabilir "
          f"(bayram tatili değilse). 1-2 saat sonra hücre 2'yi tekrar çalıştır.")
_bos = int(df['Son Fiyat'].isna().sum())
if _bos:
    print(f"⚠️  {_bos} hissenin fiyatı boş geldi — bu hisselerin puanı eksik hesaplandı")
print('\n--- TOP 10 ---')
print(df.head(10)[['Sembol', 'Toplam Skor', 'Son Fiyat', 'Hedef (Ort.)', 'Reel Potansiyel (%)',
                   'Güncel F/K', 'Hedef F/K (Orta)', 'İleri Reel Büyüme (%)']].to_string(index=False))
hedefli = df[df['Hedef (Ort.)'].notna()]
if len(hedefli):
    print(f"\n🎯 Hedefi olan {len(hedefli)} hisse | medyan hedefe uzaklık %{hedefli['Hedefe Uzaklık (%)'].median():+.1f}"
          f" | medyan reel potansiyel %{hedefli['Reel Potansiyel (%)'].median():+.1f}")

df.to_excel(f'{DRIVE}/BIST_Sistem.xlsx', index=False)
print(f"\n✅ BIST_Sistem.xlsx kaydedildi! {len(df)} hisse | {len(hatali_hisseler)} hata")
