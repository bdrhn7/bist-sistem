# ═══════════════════════════════════════════════════════════════
#  HÜCRE 8 — MAKRO RADAR  (v2.3 — 15.10.2026)
#  Global ekonomik göstergeler + haber analizi
#  Piyasa ortamını ölçer, AL/BEKLE kararı verir
#  ⚠️ HÜCRE 7'DEN ÖNCE ÇALIŞTIR (hücre 7 bu sayfadaki skoru okuyor)
#  v2 değişiklikleri:
#   - API anahtarı koddan çıkarıldı (Colab Secrets); güncel model + bağlantı testi
#   - "Altın (TL)" satırı aslında ons/dolar fiyatıydı: adı düzeltildi,
#     ayrıca gram altın (TL) satırı eklendi
#   - Brent için aynı anda "stabil" ve "eşik aşıldı" yazması düzeltildi
#   - Kapanmış Reuters kaynakları yerine Google Haberler (TR + global) eklendi
#   - Kombine AL sayfasına satır ekleyip tabloyu kaydıran bölüm kaldırıldı
#   - v2.1: AI cevabı yarıda kesiliyordu (TAVSİYE boş kalıyordu): sınır artırıldı, kısa cevap istendi
#   - v2.2: Verisi gelmeyen gösterge artık puan almıyor (USD/TL gelmeyince "stabil" sayılıp
#           +2 veriyordu). Her gösterge 3 kez deneniyor, USD/TL için yedek sembol var.
#   - v2.2: Son 5 işlem gününün makro skoru fiyat geçmişinden hesaplanıp gösteriliyor.
#           MAKRO_GUN = 3 yapılırsa karar son 3 günün ortalamasına göre verilir.
#   - v2.3: yapay zekâ yorumu alınamazsa 3 kez deneniyor; yine olmazsa sebebi sayfada yazıyor
#           (eskiden yorum satırları boş kalıyordu). Skor yapay zekâdan bağımsız.
# ═══════════════════════════════════════════════════════════════

import yfinance as yf
import pandas as pd
import numpy as np
import requests
from bs4 import BeautifulSoup
import anthropic
import time
import warnings
warnings.filterwarnings('ignore')
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
from urllib.parse import quote
import re

# ── AYARLAR ──────────────────────────────────────────────────
MODEL      = "claude-sonnet-5-5"
MAKRO_GUN  = 1      # 1 = sadece bugünün skoru | 3 = son 3 işlem gününün ortalaması
GECMIS_GUN = 5      # tabloda gösterilecek geçmiş gün sayısı
EXCEL_YOLU = '/content/drive/MyDrive/isyatirim/BIST_Tam_KAP_v9_11mart.xlsx'

# API anahtarı: Colab Secrets → ANTHROPIC_API_KEY
client = None
AI_AKTIF = False
try:
    from google.colab import userdata
    client = anthropic.Anthropic(api_key=userdata.get('ANTHROPIC_API_KEY'))
    client.messages.create(model=MODEL, max_tokens=5,
                           messages=[{"role": "user", "content": "Sadece OK yaz."}])
    AI_AKTIF = True
    print(f"✅ Yapay zekâ bağlantısı çalışıyor (model: {MODEL})")
except Exception as e:
    print(f"❌ Yapay zekâ bağlantısı yok: {type(e).__name__}: {str(e)[:150]}")
    print("   Göstergeler ve skor yine hesaplanacak; sadece AI yorumu olmayacak.")


BORDER = Border(
    left=Side(style='thin', color='1E3A5F'),
    right=Side(style='thin', color='1E3A5F'),
    top=Side(style='thin', color='1E3A5F'),
    bottom=Side(style='thin', color='1E3A5F')
)

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Accept-Language': 'tr-TR,tr;q=0.9,en;q=0.8',
}

def hc(ws, r, c, v, bg='051020', fg='00FF88', sz=10):
    x = ws.cell(r, c, v)
    x.font = Font(name='Arial', bold=True, color=fg, size=sz)
    x.fill = PatternFill('solid', fgColor=bg)
    x.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    x.border = BORDER
    return x

def dc(ws, r, c, v, bg='091422', fg='E8F4F8', bold=False, sz=9, wrap=False, align='center'):
    try:
        val = None if isinstance(v, float) and np.isnan(v) else v
    except:
        val = v
    x = ws.cell(r, c, val)
    x.font = Font(name='Arial', bold=bold, color=fg, size=sz)
    x.fill = PatternFill('solid', fgColor=bg)
    x.alignment = Alignment(horizontal=align, vertical='center', wrap_text=wrap)
    x.border = BORDER
    return x

# ════════════════════════════════════════════════════════════════
#  BÖLÜM 1: FİNANSAL GÖSTERGELER
# ════════════════════════════════════════════════════════════════

SEMBOLLER = {
    'USD/TL':        ['USDTRY=X', 'TRY=X'],
    'EUR/TL':        ['EURTRY=X'],
    'Altın (ons $)': ['GC=F'],
    'Brent Petrol':  ['BZ=F'],
    'WTI Petrol':    ['CL=F'],
    'VIX':           ['^VIX'],
    'S&P500':        ['^GSPC'],
    'Nasdaq':        ['^IXIC'],
    'BIST100':       ['XU100.IS'],
    'Dolar Endeksi': ['DX-Y.NYB'],
}

def seri_cek(semboller, deneme=3):
    """Kapanış serisini çeker (indeks = işlem günü). Olmazsa yedek sembolü dener."""
    for sembol in semboller:
        for _ in range(deneme):
            try:
                h = yf.Ticker(sembol).history(period='3mo', interval='1d')
                if h is not None and len(h) >= 5:
                    k = h['Close'].dropna()
                    k = k[k > 0]
                    k.index = pd.to_datetime(k.index.date)      # her barın kendi tarihi
                    k = k[~k.index.duplicated(keep='last')]
                    if len(k) >= 5:
                        return k
            except Exception:
                pass
            time.sleep(1)
    return None

def gosterge_hesapla(seriler, tarih=None):
    """Seri sözlüğünden göstergeleri hesaplar. tarih verilirse o güne kadarki veriyle."""
    gostergeler = {}
    for ad, k in seriler.items():
        if tarih is not None:
            k = k[k.index <= tarih]
        if len(k) < 5:
            continue
        son     = float(k.iloc[-1])
        once_1g = float(k.iloc[-2]) if len(k) >= 2 else son
        once_1h = float(k.iloc[-6]) if len(k) >= 6 else son
        once_1a = float(k.iloc[-22]) if len(k) >= 22 else float(k.iloc[0])
        gostergeler[ad] = {
            'Son Değer':    round(son, 2),
            'Günlük (%)':   round((son / once_1g - 1) * 100, 2),
            'Haftalık (%)': round((son / once_1h - 1) * 100, 2),
            'Aylık (%)':    round((son / once_1a - 1) * 100, 2),
        }

    # Gram altın (TL) ≈ ons ($) × USD/TL ÷ 31,1035 (yaklaşık; değişimler bileşik)
    try:
        g, u = gostergeler['Altın (ons $)'], gostergeler['USD/TL']
        def bilesik(a, b): return round(((1 + a / 100) * (1 + b / 100) - 1) * 100, 2)
        gostergeler['Gram Altın (TL)'] = {
            'Son Değer':    round(g['Son Değer'] * u['Son Değer'] / 31.1035, 0),
            'Günlük (%)':   bilesik(g['Günlük (%)'],   u['Günlük (%)']),
            'Haftalık (%)': bilesik(g['Haftalık (%)'], u['Haftalık (%)']),
            'Aylık (%)':    bilesik(g['Aylık (%)'],    u['Aylık (%)']),
        }
    except: pass
    return gostergeler

def finansal_gosterge_cek():
    """Tüm göstergelerin serilerini çeker; bugünkü değerleri ve serileri döndürür."""
    seriler, eksik = {}, []
    for ad, semboller in SEMBOLLER.items():
        k = seri_cek(semboller)
        if k is None:
            eksik.append(ad)
        else:
            seriler[ad] = k
        time.sleep(0.2)
    return gosterge_hesapla(seriler), seriler, eksik

def karar_ver(skor):
    if skor >= 30:    return '🟢 UYGUN ORTAM', '00FF88', '001A10'
    elif skor >= 5:   return '🟡 DİKKATLİ OL', 'FFB300', '1A1400'
    elif skor >= -20: return '🟠 RİSKLİ ORTAM', 'FF8C00', '1A0A00'
    return '🔴 RİSK YÜKSEK', 'FF4757', '2D0000'

def makro_skor_hesapla(gostergeler):
    """
    Göstergelerden makro risk skoru üretir
    Negatif = riskli ortam, Pozitif = iyi ortam
    -100 ile +100 arası
    """
    skor = 0
    sinyaller = []
    risk_seviyesi = []

    # ── DOLAR/TL ──
    try:
        usd = gostergeler['USD/TL']          # veri yoksa KeyError → puan verilmez
        gunluk = usd.get('Günlük (%)', 0)
        aylik  = usd.get('Aylık (%)', 0)
        if gunluk >= 2:      skor -= 15; risk_seviyesi.append('🔴'); sinyaller.append(f'USD/TL günlük +%{gunluk:.1f} — TL değer kaybı')
        elif gunluk >= 1:    skor -= 8;  risk_seviyesi.append('🟠'); sinyaller.append(f'USD/TL günlük +%{gunluk:.1f} — TL baskı altında')
        elif gunluk <= -1:   skor += 8;  risk_seviyesi.append('🟢'); sinyaller.append(f'USD/TL günlük %{gunluk:.1f} — TL güçleniyor')
        else:                skor += 2;  risk_seviyesi.append('🟡'); sinyaller.append(f'USD/TL stabil ({usd.get("Son Değer",0):.2f})')

        if aylik >= 5:       skor -= 10; sinyaller.append(f'USD/TL aylık +%{aylik:.1f} — TL trend kaybı')
        elif aylik <= -3:    skor += 5;  sinyaller.append(f'USD/TL aylık %{aylik:.1f} — TL güçlü trend')
    except: pass

    # ── BRENT PETROL ──
    try:
        brent = gostergeler['Brent Petrol']          # veri yoksa KeyError → puan verilmez
        gunluk = brent.get('Günlük (%)', 0)
        aylik  = brent.get('Aylık (%)', 0)
        son    = brent.get('Son Değer', 0)

        if aylik >= 30:      skor -= 20; risk_seviyesi.append('🔴'); sinyaller.append(f'Brent petrol aylık +%{aylik:.1f} — jeopolitik kriz!')
        elif aylik >= 15:    skor -= 12; risk_seviyesi.append('🔴'); sinyaller.append(f'Brent petrol aylık +%{aylik:.1f} — ciddi yükseliş')
        elif aylik >= 5:     skor -= 5;  risk_seviyesi.append('🟠'); sinyaller.append(f'Brent petrol aylık +%{aylik:.1f} — yükseliş baskısı')
        elif aylik <= -10:   skor += 8;  risk_seviyesi.append('🟢'); sinyaller.append(f'Brent petrol aylık %{aylik:.1f} — düşüş, enflasyon azalır')
        else:                skor += 1;  risk_seviyesi.append('🟡'); sinyaller.append(f'Brent petrol aylık %{aylik:+.1f} — aylık değişim sınırlı')

        if son >= 100:       skor -= 10; sinyaller.append(f'Brent {son:.1f}$ — 100$ eşiğinin üstünde!')
        elif son >= 85:      skor -= 5;  sinyaller.append(f'Brent {son:.1f}$ — yüksek seviye')
    except: pass

    # ── VIX (KORKU ENDEKSİ) ──
    try:
        vix = gostergeler['VIX']          # veri yoksa KeyError → puan verilmez
        son = vix.get('Son Değer', 20)
        gunluk = vix.get('Günlük (%)', 0)

        if son >= 35:        skor -= 20; risk_seviyesi.append('🔴'); sinyaller.append(f'VIX {son:.1f} — PANIK! Global korku çok yüksek')
        elif son >= 25:      skor -= 12; risk_seviyesi.append('🔴'); sinyaller.append(f'VIX {son:.1f} — Global belirsizlik yüksek')
        elif son >= 20:      skor -= 5;  risk_seviyesi.append('🟠'); sinyaller.append(f'VIX {son:.1f} — Orta belirsizlik')
        elif son <= 15:      skor += 10; risk_seviyesi.append('🟢'); sinyaller.append(f'VIX {son:.1f} — Global piyasalar sakin')
        else:                skor += 3;  risk_seviyesi.append('🟡'); sinyaller.append(f'VIX {son:.1f} — Normal seviye')

        if gunluk >= 20:     skor -= 10; sinyaller.append(f'VIX günlük +%{gunluk:.1f} — ani korku artışı!')
    except: pass

    # ── S&P500 ──
    try:
        sp = gostergeler['S&P500']          # veri yoksa KeyError → puan verilmez
        gunluk = sp.get('Günlük (%)', 0)
        haftalik = sp.get('Haftalık (%)', 0)

        if gunluk <= -2:     skor -= 15; risk_seviyesi.append('🔴'); sinyaller.append(f'S&P500 günlük %{gunluk:.1f} — sert düşüş')
        elif gunluk <= -1:   skor -= 8;  risk_seviyesi.append('🟠'); sinyaller.append(f'S&P500 günlük %{gunluk:.1f} — düşüş')
        elif gunluk >= 1:    skor += 8;  risk_seviyesi.append('🟢'); sinyaller.append(f'S&P500 günlük +%{gunluk:.1f} — yükseliş')
        else:                skor += 2;  risk_seviyesi.append('🟡'); sinyaller.append(f'S&P500 stabil (%{gunluk:.1f})')

        if haftalik <= -3:   skor -= 8;  sinyaller.append(f'S&P500 haftalık %{haftalik:.1f} — haftalık trend kötü')
        elif haftalik >= 2:  skor += 5;  sinyaller.append(f'S&P500 haftalık +%{haftalik:.1f} — haftalık trend iyi')
    except: pass

    # ── ALTIN ──
    try:
        altin = gostergeler['Altın (ons $)']          # veri yoksa KeyError → puan verilmez
        gunluk = altin.get('Günlük (%)', 0)
        aylik  = altin.get('Aylık (%)', 0)

        if gunluk >= 2:      skor -= 8;  sinyaller.append(f'Altın günlük +%{gunluk:.1f} — güvenli liman talebi artıyor')
        elif gunluk <= -2:   skor += 5;  sinyaller.append(f'Altın günlük %{gunluk:.1f} — risk iştahı artıyor')
        else:                sinyaller.append(f'Altın stabil (%{gunluk:.1f})')

        if aylik >= 10:      skor -= 10; sinyaller.append(f'Altın aylık +%{aylik:.1f} — uzun süreli kaçış!')
        elif aylik <= -8:    skor += 8;  sinyaller.append(f'Altın aylık %{aylik:.1f} — risk iştahı güçlü')
    except: pass

    # ── BIST100 ──
    try:
        bist = gostergeler['BIST100']          # veri yoksa KeyError → puan verilmez
        gunluk = bist.get('Günlük (%)', 0)
        haftalik = bist.get('Haftalık (%)', 0)
        aylik = bist.get('Aylık (%)', 0)

        if gunluk <= -2:     skor -= 12; risk_seviyesi.append('🔴'); sinyaller.append(f'BIST100 günlük %{gunluk:.1f} — sert düşüş')
        elif gunluk <= -1:   skor -= 6;  risk_seviyesi.append('🟠'); sinyaller.append(f'BIST100 günlük %{gunluk:.1f} — düşüş')
        elif gunluk >= 1:    skor += 6;  risk_seviyesi.append('🟢'); sinyaller.append(f'BIST100 günlük +%{gunluk:.1f} — yükseliş')
        else:                skor += 1;  risk_seviyesi.append('🟡'); sinyaller.append(f'BIST100 stabil (%{gunluk:.1f})')

        if aylik <= -5:      skor -= 10; sinyaller.append(f'BIST100 aylık %{aylik:.1f} — trend kötü')
        elif aylik >= 5:     skor += 8;  sinyaller.append(f'BIST100 aylık +%{aylik:.1f} — trend iyi')
    except: pass

    # Verisi gelmeyen göstergeler puana katılmadı
    for ad in ['USD/TL', 'Brent Petrol', 'VIX', 'S&P500', 'Altın (ons $)', 'BIST100']:
        if ad not in gostergeler:
            sinyaller.append(f'⚠️ {ad} verisi alınamadı — puana katılmadı')

    # Skoru -100 ile +100 arasında tut
    skor = max(min(skor, 100), -100)
    karar, karar_renk, karar_bg = karar_ver(skor)
    return skor, karar, karar_renk, karar_bg, sinyaller

# ════════════════════════════════════════════════════════════════
#  BÖLÜM 2: GLOBAL HABER TARAYICI
# ════════════════════════════════════════════════════════════════

def global_haber_cek():
    """
    Reuters, BBC, Bloomberg RSS'lerinden global haberler çeker
    """
    haberler = []
    tr_q = quote('borsa istanbul OR merkez bankası OR enflasyon OR dolar kuru when:2d')
    en_q = quote('stock market OR federal reserve OR oil prices OR geopolitical when:2d')
    kaynaklar = [
        ('Google TR',   f'https://news.google.com/rss/search?q={tr_q}&hl=tr&gl=TR&ceid=TR:tr'),
        ('Google Global', f'https://news.google.com/rss/search?q={en_q}&hl=en-US&gl=US&ceid=US:en'),
        ('BBC Ekonomi', 'https://feeds.bbci.co.uk/news/business/rss.xml'),
    ]
    sinir = datetime.now(timezone.utc) - timedelta(days=3)

    for kaynak_ad, url in kaynaklar:
        try:
            r = requests.get(url, headers=HEADERS, timeout=10)
            soup = BeautifulSoup(r.content, 'xml')
            alinan = 0
            for item in soup.find_all('item')[:25]:
                if alinan >= 6:
                    break
                try:
                    baslik = item.find('title').text.strip()
                    tarih_raw = item.find('pubDate').text.strip() if item.find('pubDate') else ''
                    try:
                        dt = parsedate_to_datetime(tarih_raw)
                        if dt.tzinfo is None:
                            dt = dt.replace(tzinfo=timezone.utc)
                    except:
                        continue
                    if dt < sinir or len(baslik) <= 10:
                        continue
                    haberler.append({
                        'Kaynak': kaynak_ad,
                        'Tarih':  dt.astimezone().strftime('%d.%m.%Y %H:%M'),
                        'Başlık': baslik[:200],
                        '_dt':    dt,
                    })
                    alinan += 1
                except:
                    continue
            time.sleep(0.5)
        except:
            continue

    haberler.sort(key=lambda h: h['_dt'], reverse=True)
    return haberler

def ai_makro_analiz(gostergeler, haberler, makro_skor, makro_karar):
    """
    Claude API ile makro ortamı analiz eder
    """
    # Gösterge özeti
    gosterge_ozet = []
    for ad, veri in gostergeler.items():
        gosterge_ozet.append(
            f"- {ad}: {veri.get('Son Değer','?')} | "
            f"Günlük:{veri.get('Günlük (%)',0):+.1f}% | "
            f"Haftalık:{veri.get('Haftalık (%)',0):+.1f}% | "
            f"Aylık:{veri.get('Aylık (%)',0):+.1f}%"
        )

    # Son 10 haber
    haber_ozet = [f"- {h['Tarih']} [{h['Kaynak']}] {h['Başlık']}" for h in haberler[:12]]
    if not AI_AKTIF:
        return "AI analizi yok: yapay zekâ bağlantısı kurulamadı"

    prompt = f"""Türkiye borsası (BIST) yatırımcısı için makro ortam analizi yap.

FİNANSAL GÖSTERGELER:
{chr(10).join(gosterge_ozet)}

SON GLOBAL HABERLER:
{chr(10).join(haber_ozet)}

SİSTEM MAKRO SKORU: {makro_skor} | KARAR: {makro_karar}

Sadece yukarıdaki göstergelere ve haberlere dayan; bunlarda olmayan bilgi ekleme.
Şu formatta, her başlık TEK SATIR ve en fazla 2 kısa cümle olacak şekilde cevap ver:
ÖZET: 2-3 cümle genel durum
JEOPOLİTİK: Varsa jeopolitik riskler (savaş, gerilim, ambargo)
BIST ETKİSİ: Bu ortamın BIST'e etkisi
TAVSİYE: BEKLE veya AL veya KISMI AL
GEREKÇE: Tek cümle"""

    son_hata = ''
    for deneme in range(3):
        try:
            mesaj = client.messages.create(
                model=MODEL,
                max_tokens=1200,
                messages=[{"role": "user", "content": prompt}]
            )
            return mesaj.content[0].text.strip()
        except Exception as e:
            son_hata = f"{type(e).__name__}: {str(e)[:120]}"
            time.sleep(5 * (deneme + 1))
    return f"AI analizi alınamadı (3 deneme): {son_hata}"

def ai_yanit_parse(yanit):
    """AI yanıtını parse eder"""
    sonuc = {
        'Özet': '',
        'Jeopolitik': '',
        'BIST Etkisi': '',
        'Tavsiye': '⚪ BELİRSİZ',
        'Gerekçe': '',
    }
    for satir in yanit.replace('*', '').split('\n'):
        satir = satir.strip().lstrip('-# ').strip()
        if satir.startswith('ÖZET:'):
            sonuc['Özet'] = satir.replace('ÖZET:', '').strip()
        elif satir.startswith('JEOPOLİTİK:'):
            sonuc['Jeopolitik'] = satir.replace('JEOPOLİTİK:', '').strip()
        elif satir.startswith('BIST ETKİSİ:'):
            sonuc['BIST Etkisi'] = satir.replace('BIST ETKİSİ:', '').strip()
        elif satir.startswith('TAVSİYE:'):
            t = satir.replace('TAVSİYE:', '').strip().upper()
            if 'BEKLE' in t:                       sonuc['Tavsiye'] = '🔴 BEKLE'
            elif 'KISMI' in t or 'KISMİ' in t:     sonuc['Tavsiye'] = '🟡 KISMI AL'
            elif re.search(r'\bAL\b', t):          sonuc['Tavsiye'] = '🟢 AL'
        elif satir.startswith('GEREKÇE:'):
            sonuc['Gerekçe'] = satir.replace('GEREKÇE:', '').strip()
    return sonuc

# ════════════════════════════════════════════════════════════════
#  ANA KOD
# ════════════════════════════════════════════════════════════════

print("=" * 65)
print("  📡 MAKRO RADAR — Global Piyasa & Jeopolitik Analiz")
print(f"  {datetime.now().strftime('%d.%m.%Y %H:%M')}")
print("=" * 65)

# 1. Finansal göstergeleri çek
print("\n📊 Finansal göstergeler çekiliyor...")
gostergeler, seriler, eksik = finansal_gosterge_cek()
print(f"✅ {len(gostergeler)} gösterge alındı")
if eksik:
    print(f"⚠️ Alınamayan: {', '.join(eksik)} — bunlar puana katılmayacak")

for ad, veri in gostergeler.items():
    gunluk = veri.get('Günlük (%)', 0)
    emoji = '🔴' if gunluk < -1 else '🟢' if gunluk > 1 else '🟡'
    print(f"  {emoji} {ad:20s} {veri.get('Son Değer',0):10.2f} | "
          f"G:{gunluk:+.1f}% H:{veri.get('Haftalık (%)',0):+.1f}% A:{veri.get('Aylık (%)',0):+.1f}%")

# 2. Makro skor hesapla
print("\n🧮 Makro skor hesaplanıyor...")
makro_skor, makro_karar, karar_renk, karar_bg, sinyaller = makro_skor_hesapla(gostergeler)
bugun_skor = makro_skor

# Son günlerin skoru: aynı kurallar, o güne kadarki fiyatlarla
takvim_seri = seriler.get('BIST100', seriler.get('S&P500'))
gecmis = []          # [(tarih, skor), ...] eskiden yeniye, son eleman bugün
if takvim_seri is not None:
    for t in list(takvim_seri.index[-GECMIS_GUN:-1]):
        s_t = makro_skor_hesapla(gosterge_hesapla(seriler, t))[0]
        gecmis.append((t.strftime('%d.%m.%Y'), s_t))
gecmis.append((datetime.now().strftime('%d.%m.%Y') + ' (şimdi)', bugun_skor))

if MAKRO_GUN > 1 and len(gecmis) >= 2:
    son_n = [x[1] for x in gecmis[-MAKRO_GUN:]]
    makro_skor = int(round(sum(son_n) / len(son_n)))
    makro_karar, karar_renk, karar_bg = karar_ver(makro_skor)
    skor_aciklama = f"SKOR: {makro_skor:+d} (son {len(son_n)} gün ort.; bugün {bugun_skor:+d})"
else:
    skor_aciklama = f"SKOR: {makro_skor:+d}"

print(f"\n{'='*50}")
print(f"  MAKRO {skor_aciklama}")
print(f"  KARAR: {makro_karar}")
print(f"{'='*50}")
print("\n📅 Son günlerin makro skoru:")
for t, sk in gecmis:
    print(f"  {t:22s} {sk:+4d}  {karar_ver(sk)[0]}")
print("\nSinyaller:")
for s in sinyaller:
    print(f"  {s}")

# 3. Global haberler
print("\n📰 Global haberler çekiliyor...")
haberler = global_haber_cek()
print(f"✅ {len(haberler)} haber alındı")
for h in haberler[:5]:
    print(f"  [{h['Kaynak']}] {h['Başlık'][:80]}")

# 4. AI makro analizi
print("\n🤖 AI makro analizi yapılıyor...")
ai_yanit = ai_makro_analiz(gostergeler, haberler, makro_skor, makro_karar)
ai_sonuc = ai_yanit_parse(ai_yanit)
if not ai_sonuc['Özet']:                       # yorum okunamadıysa sebebi sayfada görünsün
    ai_sonuc['Özet'] = '⚠️ ' + ai_yanit.strip().replace('\n', ' ')[:250]
print(f"\n📋 AI ANALİZİ:")
print(ai_yanit)

TARIH = datetime.now().strftime('%d.%m.%Y %H:%M')

# ── EXCEL'E YAZ ──────────────────────────────────────────────
print(f"\n💾 Excel güncelleniyor...")
wb = load_workbook(EXCEL_YOLU)

# ── 📡 MAKRO RADAR SAYFASI ──
if '📡 Makro Radar' in wb.sheetnames: del wb['📡 Makro Radar']
ws = wb.create_sheet('📡 Makro Radar', 0)
ws.sheet_tab_color = 'FF4757' if '🔴' in makro_karar else 'FF8C00' if '🟠' in makro_karar else 'FFB300' if '🟡' in makro_karar else '00FF88'
ws.sheet_view.showGridLines = False

# Ana başlık
ws.merge_cells('A1:H1')
ws['A1'] = f"📡 MAKRO RADAR — {TARIH}"
ws['A1'].font = Font(name='Arial', bold=True, size=14, color='00D4FF')
ws['A1'].fill = PatternFill('solid', fgColor='020D1A')
ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
ws.row_dimensions[1].height = 40

# Makro karar kutusu
ws.merge_cells('A2:H2')
ws['A2'] = f"MAKRO KARAR: {makro_karar}   |   {skor_aciklama}   |   AI: {ai_sonuc['Tavsiye']}"
ws['A2'].font = Font(name='Arial', bold=True, size=13, color=karar_renk)
ws['A2'].fill = PatternFill('solid', fgColor=karar_bg)
ws['A2'].alignment = Alignment(horizontal='center', vertical='center')
ws.row_dimensions[2].height = 35

# AI özet
ws.merge_cells('A3:H3')
ws['A3'] = f"📋 {ai_sonuc['Özet']}"
ws['A3'].font = Font(name='Arial', size=10, color='E8F4F8', italic=True)
ws['A3'].fill = PatternFill('solid', fgColor='051020')
ws['A3'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
ws.row_dimensions[3].height = 30

# Jeopolitik uyarı
if ai_sonuc['Jeopolitik']:
    ws.merge_cells('A4:H4')
    ws['A4'] = f"⚠️ JEOPOLİTİK: {ai_sonuc['Jeopolitik']}"
    ws['A4'].font = Font(name='Arial', size=10, color='FF6B35', bold=True)
    ws['A4'].fill = PatternFill('solid', fgColor='1A0800')
    ws['A4'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws.row_dimensions[4].height = 30
    baslik_satir = 6
else:
    baslik_satir = 5

# BIST etkisi
ws.merge_cells(f'A{baslik_satir-1}:H{baslik_satir-1}')
ws[f'A{baslik_satir-1}'] = f"📈 BIST ETKİSİ: {ai_sonuc['BIST Etkisi']} | {ai_sonuc['Gerekçe']}"
ws[f'A{baslik_satir-1}'].font = Font(name='Arial', size=9, color='B2DFDB', italic=True)
ws[f'A{baslik_satir-1}'].fill = PatternFill('solid', fgColor='051A10')
ws[f'A{baslik_satir-1}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
ws.row_dimensions[baslik_satir-1].height = 25

# ── FİNANSAL GÖSTERGELER TABLOSU ──
g_cols = ['Gösterge', 'Son Değer', 'Günlük (%)', 'Haftalık (%)', 'Aylık (%)', 'Sinyal']
for ci, col in enumerate(g_cols, 1):
    hc(ws, baslik_satir, ci, col, fg='00D4FF')
ws.row_dimensions[baslik_satir].height = 25

# Gösterge sinyali hesapla
def gosterge_sinyal(ad, veri):
    gunluk = veri.get('Günlük (%)', 0)
    aylik  = veri.get('Aylık (%)', 0)
    son    = veri.get('Son Değer', 0)

    if ad == 'USD/TL':
        if gunluk >= 1: return '🔴 TL baskı'
        elif gunluk <= -0.5: return '🟢 TL güçlü'
        return '🟡 Stabil'
    elif 'Petrol' in ad:
        if aylik >= 20: return '🔴 Kriz!'
        elif aylik >= 10: return '🟠 Yükseliş'
        elif aylik <= -10: return '🟢 Düşüş'
        return '🟡 Normal'
    elif ad == 'VIX':
        if son >= 30: return '🔴 Panik!'
        elif son >= 20: return '🟠 Gergin'
        elif son <= 15: return '🟢 Sakin'
        return '🟡 Normal'
    elif 'S&P' in ad or 'Nasdaq' in ad:
        if gunluk <= -2: return '🔴 Düşüş'
        elif gunluk >= 1: return '🟢 Yükseliş'
        return '🟡 Stabil'
    elif ad == 'BIST100':
        if gunluk <= -2: return '🔴 Düşüş'
        elif gunluk >= 1: return '🟢 Yükseliş'
        return '🟡 Stabil'
    elif 'Altın' in ad:
        if aylik >= 8: return '🔴 Güvenli liman'
        elif aylik <= -5: return '🟢 Risk iştahı'
        return '🟡 Normal'
    return '🟡 Normal'

SINYAL_R = {
    '🔴': ('2D0000', 'FF4757'),
    '🟠': ('1A0A00', 'FF8C00'),
    '🟢': ('001A10', '00FF88'),
    '🟡': ('1A1400', 'FFB300'),
}

for ri, (ad, veri) in enumerate(gostergeler.items(), baslik_satir + 1):
    bg = '0D1B2E' if ri % 2 == 0 else '091422'
    sinyal = gosterge_sinyal(ad, veri)
    gunluk = veri.get('Günlük (%)', 0)
    haftalik = veri.get('Haftalık (%)', 0)
    aylik = veri.get('Aylık (%)', 0)

    dc(ws, ri, 1, ad, bg=bg, bold=True, align='left')
    dc(ws, ri, 2, veri.get('Son Değer', ''), bg=bg, bold=True, fg='FFD700')

    for ci, deger in [(3, gunluk), (4, haftalik), (5, aylik)]:
        c = dc(ws, ri, ci, f"{deger:+.2f}%", bg=bg, bold=True)
        c.font = Font(name='Arial', bold=True, size=10,
            color='00FF88' if deger > 0 else 'FF4757' if deger < 0 else 'AAAAAA')

    # Sinyal rengi
    ilk_emoji = sinyal[:2]
    bk, fk = SINYAL_R.get(ilk_emoji, ('091422', 'E8F4F8'))
    c = ws.cell(ri, 6, sinyal)
    c.fill = PatternFill('solid', fgColor=bk)
    c.font = Font(name='Arial', bold=True, size=9, color=fk)
    c.alignment = Alignment(horizontal='center', vertical='center')
    c.border = BORDER
    ws.row_dimensions[ri].height = 22

son_gosterge_satir = baslik_satir + len(gostergeler) + 1

# ── SON GÜNLERİN MAKRO SKORU ──
ws.merge_cells(f'A{son_gosterge_satir}:H{son_gosterge_satir}')
ws[f'A{son_gosterge_satir}'] = ('📅 SON GÜNLERİN MAKRO SKORU' +
    (f' — karar son {MAKRO_GUN} günün ortalamasına göre' if MAKRO_GUN > 1 else ' — karar sadece bugüne göre'))
ws[f'A{son_gosterge_satir}'].font = Font(name='Arial', bold=True, size=11, color='00D4FF')
ws[f'A{son_gosterge_satir}'].fill = PatternFill('solid', fgColor='051020')
ws[f'A{son_gosterge_satir}'].alignment = Alignment(horizontal='center')
for i, (t, sk) in enumerate(gecmis):
    r = son_gosterge_satir + 1 + i
    k_ad, k_fg, k_bg = karar_ver(sk)
    dc(ws, r, 1, t, bg='091422', bold=True, align='left')
    dc(ws, r, 2, sk, bg=k_bg, fg=k_fg, bold=True)
    dc(ws, r, 3, k_ad, bg=k_bg, fg=k_fg)
son_gosterge_satir += len(gecmis) + 2

# ── SİNYALLER ──
ws.merge_cells(f'A{son_gosterge_satir}:H{son_gosterge_satir}')
ws[f'A{son_gosterge_satir}'] = '📊 DETAYLI SİNYALLER'
ws[f'A{son_gosterge_satir}'].font = Font(name='Arial', bold=True, size=11, color='00D4FF')
ws[f'A{son_gosterge_satir}'].fill = PatternFill('solid', fgColor='051020')
ws[f'A{son_gosterge_satir}'].alignment = Alignment(horizontal='center')
ws.row_dimensions[son_gosterge_satir].height = 25

for i, sinyal in enumerate(sinyaller):
    r = son_gosterge_satir + i + 1
    ws.merge_cells(f'A{r}:H{r}')
    ilk = sinyal[:2]
    bk = '2D0000' if '🔴' in ilk else '1A0A00' if '🟠' in ilk else '001A10' if '🟢' in ilk else '1A1400'
    fk = 'FF4757' if '🔴' in ilk else 'FF8C00' if '🟠' in ilk else '00FF88' if '🟢' in ilk else 'FFB300'
    ws[f'A{r}'] = sinyal
    ws[f'A{r}'].font = Font(name='Arial', size=9, color=fk)
    ws[f'A{r}'].fill = PatternFill('solid', fgColor=bk)
    ws[f'A{r}'].alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[r].height = 20

son_sinyal_satir = son_gosterge_satir + len(sinyaller) + 2

# ── GLOBAL HABERLER ──
ws.merge_cells(f'A{son_sinyal_satir}:H{son_sinyal_satir}')
ws[f'A{son_sinyal_satir}'] = f'📰 GLOBAL HABERLER ({len(haberler)} haber)'
ws[f'A{son_sinyal_satir}'].font = Font(name='Arial', bold=True, size=11, color='00D4FF')
ws[f'A{son_sinyal_satir}'].fill = PatternFill('solid', fgColor='051020')
ws[f'A{son_sinyal_satir}'].alignment = Alignment(horizontal='center')
ws.row_dimensions[son_sinyal_satir].height = 25

haber_baslik_satir = son_sinyal_satir + 1
for ci, col in enumerate(['Kaynak', 'Tarih', 'Başlık'], 1):
    hc(ws, haber_baslik_satir, ci, col, fg='FFB300')
ws.row_dimensions[haber_baslik_satir].height = 22

for i, h in enumerate(haberler):
    r = haber_baslik_satir + i + 1
    bg = '0D1B2E' if i % 2 == 0 else '091422'
    dc(ws, r, 1, h['Kaynak'], bg=bg, fg='FFB300', bold=True)
    dc(ws, r, 2, h['Tarih'], bg=bg, fg='AAAAAA')
    c = ws.cell(r, 3, h['Başlık'])
    c.font = Font(name='Arial', size=9, color='E8F4F8')
    c.fill = PatternFill('solid', fgColor=bg)
    c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    c.border = BORDER
    ws.row_dimensions[r].height = 28

# Kolon genişlikleri
ws.column_dimensions['A'].width = 20
ws.column_dimensions['B'].width = 14
ws.column_dimensions['C'].width = 12
ws.column_dimensions['D'].width = 12
ws.column_dimensions['E'].width = 12
ws.column_dimensions['F'].width = 16
ws.column_dimensions['G'].width = 12
ws.column_dimensions['H'].width = 12

# Haberler için C kolonunu genişlet (3. kolon)
ws.column_dimensions['C'].width = 80

wb.save(EXCEL_YOLU)
print(f"✅ 📡 Makro Radar sayfası güncellendi!")

# (Eski 'Kombine AL'a makro filtre ekle' bölümü kaldırıldı: hücre 7 makro skorunu bu sayfadan okuyor.)

# ── ÖZET ──
print(f"\n{'='*60}")
print(f"  📡 MAKRO RADAR SONUCU")
print(f"  Skor    : {skor_aciklama.replace('SKOR: ', '')}")
print(f"  Karar   : {makro_karar}")
print(f"  AI      : {ai_sonuc['Tavsiye']}")
print(f"  Özet    : {ai_sonuc['Özet'][:80]}")
if ai_sonuc['Jeopolitik']:
    print(f"  Jeopol. : {ai_sonuc['Jeopolitik'][:80]}")
print(f"  BIST    : {ai_sonuc['BIST Etkisi'][:80]}")
print(f"{'='*60}")

if '🔴' in makro_karar:
    print("\n⚠️  UYARI: RİSK YÜKSEK — Yeni pozisyon açmaktan kaçının!")
    print("   Mevcut pozisyonlarda stop seviyelerini kontrol edin.")
elif '🟠' in makro_karar:
    print("\n⚠️  DİKKAT: Riskli ortam — Küçük pozisyonlarla devam edin.")
elif '🟡' in makro_karar:
    print("\n🟡  DİKKATLİ OL: Seçici olun, sadece en güçlü sinyallere girin.")
else:
    print("\n✅  UYGUN ORTAM: Normal pozisyon büyüklükleriyle devam edebilirsiniz.")

print(f"\n🎉 HÜCRE 8 TAMAMLANDI! Şimdi hücre 7'yi çalıştır.")