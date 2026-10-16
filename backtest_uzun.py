"""Zirve Stratejisi — uzun dönem backtest (2015–2026), hücre 10 kurallarıyla.
Sinyal: t-1 kapanışına kadar veri. İşlem: t (ayın ilk işlem günü) açılışı → sonraki ayın ilk işlem günü açılışı.
"""
import numpy as np, pandas as pd, sys

MIN_TL_HACIM = 50e6        # 2026-10 TL'si; geçmişe TÜFE ile indirgenir
MIN_FIYAT    = 5.0         # 2026-10 TL'si; düzeltilmiş fiyat üzerinde yaklaşık
MAX_HAFTALIK = 0.20
ARG          = sys.argv[1] if len(sys.argv) > 1 else '30'
ORAN         = float(ARG[1:]) / 100 if ARG.startswith('p') else None
HISSE        = None if ORAN else int(ARG)
FIYAT_FILTRE = '--fiyatsiz' not in sys.argv
MALIYET      = 0.002       # tek yön (komisyon + kayma)
MIN_GECMIS   = 200         # işlem günü

TUFE = {2014: 8.17, 2015: 8.81, 2016: 8.53, 2017: 11.92, 2018: 20.30, 2019: 11.84, 2020: 14.60,
        2021: 36.08, 2022: 64.27, 2023: 64.77, 2024: 44.38, 2025: 30.89, 2026: 25.0}

f = pd.read_pickle('fiyat.pkl')
f['Tarih'] = pd.to_datetime(f['Tarih']).dt.tz_localize(None).dt.normalize()
f = f.sort_values(['Kod', 'Tarih'])

# ── Bedelsiz / bölünme düzeltmesi: BIST'te günlük fiyat limiti var; |getiri| > %30 = sermaye işlemi ──
def duzelt(g):
    g = g.copy()
    r = g['Close'] / g['Close'].shift(1)
    olay = (r < 0.70) | (r > 1.35)
    for i in np.where(olay.values)[0]:
        k = r.iloc[i]
        for c in ['Open', 'High', 'Low', 'Close']:
            g.iloc[:i, g.columns.get_loc(c)] *= k
        g.iloc[:i, g.columns.get_loc('Volume')] /= k
    return g, int(olay.sum())
parca, olay_say = [], 0
for kod, g in f.groupby('Kod', sort=False):
    g2, n = duzelt(g) if kod != 'XU100' else (g, 0)
    parca.append(g2); olay_say += n
f = pd.concat(parca)

C = f.pivot(index='Tarih', columns='Kod', values='Close')
O = f.pivot(index='Tarih', columns='Kod', values='Open')
H = f.pivot(index='Tarih', columns='Kod', values='High')
V = f.pivot(index='Tarih', columns='Kod', values='Volume')
xu = C.pop('XU100'); O.pop('XU100'); H.pop('XU100'); V.pop('XU100')
gunler = xu.dropna().index
C, O, H, V = [x.reindex(gunler) for x in (C, O, H, V)]

# TÜFE endeksi (aylık, geometrik), 2026-10 = 1
ay = pd.period_range('2014-01', '2026-12', freq='M')
idx = []; v = 1.0
for p in ay:
    v *= (1 + TUFE[p.year] / 100) ** (1 / 12); idx.append(v)
cpi = pd.Series(idx, index=ay); cpi = cpi / cpi[pd.Period('2026-10', 'M')]

tlh = (C * V).rolling(20, min_periods=15).mean()
zirve = H.rolling(252, min_periods=MIN_GECMIS).max()
gecmis = C.notna().cumsum()

ilk_gun = pd.Series(gunler, index=gunler).groupby(gunler.to_period('M')).first()
ilk_gun = ilk_gun[ilk_gun >= pd.Timestamp('2015-02-01')]

kayit, onceki = [], set()
for j in range(len(ilk_gun) - 1):
    t, t2 = ilk_gun.iloc[j], ilk_gun.iloc[j + 1]
    i = gunler.get_loc(t); dun = gunler[i - 1]
    d = cpi[t.to_period('M')]
    c1 = C.loc[dun]
    likit = (tlh.loc[dun] >= MIN_TL_HACIM * d) & (gecmis.loc[dun] >= MIN_GECMIS) & c1.notna()
    hafta = c1 / C.iloc[i - 6] - 1
    aday = likit & ((c1 >= MIN_FIYAT * d) | (not FIYAT_FILTRE)) & (hafta <= MAX_HAFTALIK) & O.loc[t].notna()
    yakin = (c1 / zirve.loc[dun] - 1)[aday].dropna().sort_values(ascending=False)
    n_sec = HISSE if HISSE else max(10, round(ORAN * int(likit.sum())))
    sec = list(yakin.index[:n_sec])
    # getiri: t açılışı → t2 açılışı (t2'de açılış yoksa son kapanış)
    acilis = O.loc[t]
    son = O.loc[t2].where(O.loc[t2].notna(), C.loc[:t2].ffill().iloc[-1])
    r = son / acilis - 1
    evren = [k for k in likit[likit & acilis.notna()].index]
    r_sec = r[sec].dropna(); r_ev = r[evren].dropna()
    devir = len(set(sec) - onceki) / max(len(sec), 1) if onceki else 1.0
    net = r_sec.mean() - devir * 2 * MALIYET
    kayit.append({'Ay': t.to_period('M'), 'Strateji': net, 'Brüt': r_sec.mean(), 'Likit': r_ev.mean(),
                  'XU100': xu.loc[t2] / xu.loc[dun] - 1 if pd.notna(xu.loc[t2]) else np.nan,
                  'Evren': len(evren), 'Seçilen': len(r_sec), 'Devir': devir,
                  'Top2': r[sec[:2]].mean() if len(sec) >= 2 else np.nan, 'Liste': ','.join(sec)})
    onceki = set(sec)

k = pd.DataFrame(kayit).set_index('Ay')
k['Fark'] = k['Strateji'] - k['Likit']
k.to_pickle(f'sonuc_{ARG}{"_fs" if not FIYAT_FILTRE else ""}.pkl')

def ozet(x, ad):
    n = len(x); yil = n / 12
    def bil(s): return (1 + s).prod() ** (1 / yil) - 1
    defl = cpi.reindex(x.index); enf = (defl.iloc[-1] / defl.iloc[0]) ** (12 / max(n - 1, 1)) - 1
    dd = ((1 + x['Strateji']).cumprod() / (1 + x['Strateji']).cumprod().cummax() - 1).min()
    t = x['Fark'].mean() / (x['Fark'].std(ddof=1) / np.sqrt(n))
    return {'Dönem': ad, 'Ay': n, 'Strateji yıllık': bil(x['Strateji']), 'Likit yıllık': bil(x['Likit']),
            'XU100 yıllık': bil(x['XU100'].fillna(0)), 'Aylık fark ort.': x['Fark'].mean(), 'Fark medyan': x['Fark'].median(),
            't': t, 'Fark>0 ay oranı': (x['Fark'] > 0).mean(), 'En kötü ay': x['Strateji'].min(),
            'Maks. düşüş': dd, 'Ort. evren': x['Evren'].mean(), 'Ort. devir': x['Devir'].mean()}

satir = [ozet(k, 'TÜMÜ'),
         ozet(k[k.index < pd.Period('2020-01', 'M')], '2015-2019'),
         ozet(k[(k.index >= pd.Period('2020-01', 'M')) & (k.index < pd.Period('2024-04', 'M'))], '2020-03.2024'),
         ozet(k[k.index < pd.Period('2024-04', 'M')], 'ÖNCESİ (test dışı)'),
         ozet(k[k.index >= pd.Period('2024-04', 'M')], '04.2024+ (seçildiği dönem)')]
o = pd.DataFrame(satir).set_index('Dönem')
pd.set_option('display.width', 250)
print(f'HİSSE={ARG} fiyat filtresi={FIYAT_FILTRE} | bölünme/bedelsiz düzeltmesi: {olay_say} olay | ay sayısı {len(k)}')
fmt = o.copy()
for c in fmt.columns:
    if c not in ('Ay', 't', 'Ort. evren'):
        fmt[c] = (fmt[c] * 100).round(2)
fmt['t'] = fmt['t'].round(2); fmt['Ort. evren'] = fmt['Ort. evren'].round(0)
print(fmt.T.to_string())
yillik = k.groupby(k.index.year).apply(lambda x: pd.Series({'Strateji': (1 + x['Strateji']).prod() - 1,
                                                          'Likit': (1 + x['Likit']).prod() - 1,
                                                          'XU100': (1 + x['XU100'].fillna(0)).prod() - 1}))
yillik['Fark (puan)'] = yillik['Strateji'] - yillik['Likit']
print('\nYILLARA GÖRE (%):'); print((yillik * 100).round(1).to_string())
