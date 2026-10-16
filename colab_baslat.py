# ═══════════════════════════════════════════════════════════════
#  HÜCRE 1 — BAŞLAT  (bir kez yapıştır, bir daha değişmez)
#  Bütün hücrelerin kodu GitHub'daki bdrhn7/bist-sistem deposundan,
#  her çalıştırmada EN GÜNCEL haliyle gelir. Kopyala-yapıştır yok.
#
#  Diğer hücreler tek satır:
#     calistir(2)                              → hücre 2
#     calistir(10)                             → hücre 10
#     calistir(10, STRATEJI_PARASI=20000)      → ayar değiştirerek
#     calistir(10, HARIC=['ABCDE'], YENIDEN_SEC=True)
# ═══════════════════════════════════════════════════════════════
import re, time, requests

DEPO = 'bdrhn7/bist-sistem'

def _indir(dosya):
    """Önce GitHub API (anında güncel), olmazsa raw adres (birkaç dk gecikmeli olabilir)."""
    adresler = [
        (f'https://api.github.com/repos/{DEPO}/contents/{dosya}', {'Accept': 'application/vnd.github.raw'}),
        (f'https://raw.githubusercontent.com/{DEPO}/main/{dosya}?t={int(time.time())}', {}),
    ]
    son_hata = None
    for url, baslik in adresler:
        for deneme in range(2):
            try:
                r = requests.get(url, headers=baslik, timeout=30)
                if r.status_code == 200 and r.text.strip():
                    return r.text
                son_hata = f'HTTP {r.status_code}'
            except Exception as e:
                son_hata = type(e).__name__
            time.sleep(2)
    raise SystemExit(f'❌ {dosya} GitHub\'dan indirilemedi ({son_hata}). İnternet bağlantısını kontrol edip tekrar dene.')

def calistir(no, **ayar):
    """GitHub'daki hucre{no}.py dosyasını indirip çalıştırır. ayar: kodun başındaki AYAR = değer satırlarını değiştirir."""
    kod = _indir(f'hucre{no}.py')
    for ad, deger in ayar.items():
        m = re.search(rf'^{re.escape(ad)}\s*=.*$', kod, re.M)
        if not m:
            raise SystemExit(f'❌ Hücre {no}\'da "{ad}" diye bir ayar yok.')
        try:
            compile(m.group(0), 'ayar', 'exec')          # tek satırlık ayar mı?
        except SyntaxError:
            raise SystemExit(f'❌ "{ad}" birden çok satırlık bir ayar; bunu buradan değiştiremezsin.')
        kod = kod[:m.start()] + f'{ad} = {deger!r}   # calistir() ile verildi' + kod[m.end():]
    try:
        kod = get_ipython().transform_cell(kod)          # !pip gibi Colab komutları da çalışsın
    except NameError:
        pass
    baslik = next((s.strip('#═ ').strip() for s in kod.splitlines()[:4] if 'HÜCRE' in s), f'Hücre {no}')
    print(f'📥 {baslik}' + (f'   | ayar: {ayar}' if ayar else ''))
    exec(compile(kod, f'hucre{no}.py', 'exec'), globals())

calistir(1)
