#!/usr/bin/env python3
"""
Membangun data.json (baseline + prakiraan hujan Sulbar berbasis ENSO/IOD) untuk dashboard GitHub Pages.

Struktur repo (root): index.html, build_data.py, Data_Curah_Hujan_-_Mamuju.xlsx, .github/workflows/update.yml
Lokal : pip install numpy openpyxl requests && python build_data.py

PERUBAHAN UTAMA (v2)
 1. LOYO "nested": klimatologi (mean, sd) dihitung ulang tanpa tahun uji di tiap fold, sehingga baseline
    pembanding (prakiraan klimatologi) dan target distandardisasi benar-benar di luar sampel.
    Skill = MSESS terhadap baseline nested (bukan baseline in-sample).
 2. Uji signifikansi: permutasi tahun (blok) pada prediktor, PERMS kali -> p-value per model.
    Model dipakai bila skill > SKMIN dan p <= PMAX; selain itu kembali ke klimatologi.
 3. Curah hujan dimodelkan pada skala akar pangkat tiga (distribusi miring kanan -> mendekati normal),
    sehingga peluang kategori tidak bias. Hari hujan tetap skala linear.
 4. Sumber indeks: bila BoM mencakup periode latih dipakai untuk melatih DAN memprakirakan. Kenyataannya
    berkas BoM baru mulai 2008, sehingga yang dipakai CPC (Nino 3.4 mingguan) + NOAA PSL (DMI bulanan);
    BoM hanya mengisi bulan DMI terbaru yang belum dirilis PSL (dikalibrasi ke skala PSL).
 5. MIN_WEEKS = 3 (sama dengan process_enso.py / process_iod.py). Bulan terakhir = bulan lengkap terakhir
    yang tersedia di KEDUA indeks.
 6. data.json hanya ditulis ulang bila isi (selain stempel waktu) berubah -> tidak ada commit harian sia-sia.
"""
import datetime as dt, json, math, os, pathlib, re, warnings
import numpy as np, requests
from openpyxl import load_workbook

warnings.simplefilter('ignore')
XLSX = 'Data_Curah_Hujan_-_Mamuju.xlsx'
Y0, Y1, LEADS, MIN_WEEKS = 1996, 2025, 6, 3
SKMIN = float(os.environ.get('SKMIN', 0.03))   # ambang MSESS (nested)
PMAX = float(os.environ.get('PMAX', 0.10))     # ambang p-value permutasi
PERMS = int(os.environ.get('PERMS', 99))       # 0 = lewati uji permutasi
BOM_ENSO_URL = 'https://www.bom.gov.au/clim_data/IDCK000072/rnino_3.4.txt'
BOM_URL = 'https://www.bom.gov.au/clim_data/IDCK000072/iod_1.txt'
ENSO_URL = 'https://www.cpc.ncep.noaa.gov/data/indices/wksst9120.for'
DMI_URLS = ['https://psl.noaa.gov/gcos_wgsp/Timeseries/Data/dmi.had.long.data',
            'https://psl.noaa.gov/data/timeseries/month/data/dmi.had.long.data']
UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'
N, T0 = Y1 - Y0 + 1, Y0 * 12
CACHE = pathlib.Path('cache'); CACHE.mkdir(exist_ok=True)
MON = {m: i for i, m in enumerate('JAN FEB MAR APR MAY JUN JUL AUG SEP OCT NOV DEC'.split())}


# ---------- Unduh & parsing indeks ----------
def fetch(urls, name):
    err = ''
    for u in urls:
        try:
            r = requests.get(u, timeout=60, headers={'User-Agent': UA, 'Referer': 'https://www.bom.gov.au/climate/enso/'})
            r.raise_for_status()
            (CACHE / name).write_text(r.text, encoding='utf-8')
            return r.text, 'unduh langsung OK'
        except Exception as e:
            err = f'{type(e).__name__}: {str(e)[:90]}'
    f = CACHE / name
    if f.exists():
        return f.read_text(encoding='utf-8'), f'cache lokal ({err})'
    return None, f'gagal: {err}'


def monthly(acc):
    return {t: float(np.mean(v)) for t, v in acc.items() if len(v) >= MIN_WEEKS}


def parse_weekly(txt):  # CPC: 03JAN1990  23.4-0.4  25.1-0.3  26.6 0.0  28.6 0.3 -> SSTA Nino3.4
    acc = {}
    for l in txt.splitlines():
        d = re.match(r'^\s*(\d{1,2})([A-Z]{3})(\d{4})(.*)$', l)
        if not d or d[2] not in MON:
            continue
        v = [float(x) for x in re.findall(r'-?\d+\.\d', d[4])]
        x = v[5] if len(v) >= 8 else v[2] if len(v) >= 4 else None
        if x is not None:
            acc.setdefault(int(d[3]) * 12 + MON[d[2]], []).append(x)
    return monthly(acc)


def parse_psl(txt):  # baris: tahun + 12 nilai bulanan; missing <= -90
    o = {}
    for l in txt.splitlines():
        p = l.split()
        if len(p) >= 13 and re.fullmatch(r'\d{4}', p[0]):
            for m, x in enumerate(p[1:13]):
                try:
                    v = float(x)
                except ValueError:
                    continue
                if v > -90:
                    o[int(p[0]) * 12 + m] = v
    return o


def parse_bom(txt):  # baris: yyyymmdd mulai, yyyymmdd akhir, nilai; bulan = titik tengah minggu
    acc = {}
    for l in txt.splitlines():
        p = l.strip().split(',')
        if len(p) < 3 or not re.fullmatch(r'\d{8}', p[0]):
            continue
        try:
            v = float(p[2])
        except ValueError:
            continue
        d = dt.datetime.strptime(p[0], '%Y%m%d') + dt.timedelta(days=3)
        acc.setdefault(d.year * 12 + d.month - 1, []).append(v)
    return monthly(acc)


def merge_bom(dmi, bom):  # cadangan: isi bulan kosong PSL dengan BoM terkalibrasi
    both = [t for t in bom if t in dmi]
    a, b = 0.0, 1.0
    if len(both) >= 24:
        sl, ic = np.polyfit([bom[t] for t in both], [dmi[t] for t in both], 1)
        if 0.3 < sl < 3:
            a, b = ic, sl
    for t, v in bom.items():
        if t not in dmi:
            dmi[t] = a + b * v


def get_indices():
    # (1) Utama: BoM untuk latih & prakiraan (satu skala, tanpa kalibrasi)
    te, me = fetch([BOM_ENSO_URL], 'rnino_3.4.txt')
    td, md = fetch([BOM_URL], 'iod_1.txt')
    oni = parse_bom(te) if te else {}
    dmi = parse_bom(td) if td else {}
    need = T0 - 8  # prediktor lead 6 + pooling +-1 bulan
    if oni and dmi and min(oni) <= need and min(dmi) <= need:
        return oni, dmi, 'BoM', f'ENSO: {me}; IOD: {md}'
    why = f'BoM hanya mulai {min(oni) // 12 if oni else "-"}/{min(dmi) // 12 if dmi else "-"}, periode latih mulai {Y0}'
    print('INFO:', why, '-> pelatihan memakai CPC (Nino 3.4) + NOAA PSL (DMI); BoM hanya mengisi bulan DMI terbaru')
    # (2) Cadangan: CPC + PSL (+ BoM untuk bulan terbaru)
    txt, m1 = fetch([ENSO_URL], 'wksst9120.for')
    oni = parse_weekly(txt) if txt else {}
    if len(oni) < 20:
        raise SystemExit('Data Nino 3.4 tidak terbaca: ' + m1)
    txt, m2 = fetch(DMI_URLS, 'dmi_psl.data')
    dmi = parse_psl(txt) if txt else {}
    if len(dmi) < 100:
        raise SystemExit('Data DMI PSL tidak terbaca: ' + m2)
    if td:
        b = parse_bom(td)
        if b:
            merge_bom(dmi, b)
    return oni, dmi, 'CPC Nino3.4 + PSL DMI (BoM mengisi bulan terbaru)', f'{why}; {m1}; {m2}'


# ---------- Data hujan ----------
def num(v):
    try:
        return float(str(v).replace(',', '.'))
    except ValueError:
        return np.nan


def find_xlsx():
    """Pakai nama baku; bila tidak ada, cari berkas Data_Curah_Hujan*.xlsx di root repo."""
    if pathlib.Path(XLSX).exists():
        return XLSX
    c = sorted(pathlib.Path('.').glob('Data_Curah_Hujan*.xlsx'))
    if c:
        print('Memakai berkas data hujan:', c[0])
        return str(c[0])
    raise SystemExit('Data hujan tidak ditemukan. Pilih salah satu: (1) unggah berkas xlsx data hujan ke root repo '
                     f'(nama "{XLSX}" atau diawali "Data_Curah_Hujan"), atau (2) isi secret GitHub SHEET_XLSX_URL '
                     'dengan tautan ekspor Google Sheets (.../export?format=xlsx, akses "siapa saja yang memiliki tautan").')


def read_stations():
    out = []
    for ws in load_workbook(find_xlsx(), data_only=True).worksheets:
        v = [list(r) for r in ws.iter_rows(values_only=True)]
        meta, r1, r2 = {}, -1, -1
        for i, r in enumerate(v):
            a = str(r[0]) if r and r[0] is not None else ''
            if a == 'Nama Stasiun': meta['name'] = r[1]
            elif a == 'Lokasi': meta['loc'] = r[1]
            elif a.startswith('Longitude'):
                p = str(r[1]).split('/'); meta['lon'], meta['lat'] = num(p[0]), num(p[1])
            elif a.startswith('CURAH HUJAN'): r1 = i + 2
            elif a.startswith('JUMLAH HARI HUJAN'): r2 = i + 2
        if r1 < 0 or r2 < 0:
            continue

        def grab(s):
            o = np.full(N * 12, np.nan)
            for r in v[s:]:
                y = num(r[0])
                if np.isnan(y):
                    break
                if Y0 <= y <= Y1:
                    vals = [num(x) for x in r[1:13]]
                    vals += [np.nan] * (12 - len(vals))
                    k = (int(y) - Y0) * 12
                    o[k:k + 12] = vals
            return o

        ch, hh = grab(r1), grab(r2)
        ch[(ch == 0) & np.isnan(hh)] = np.nan  # nol tanpa data hari hujan = data hilang
        out.append(dict(id=ws.title, name=meta.get('name') or ws.title, loc=meta.get('loc'),
                        lat=meta.get('lat'), lon=meta.get('lon'), ch=ch, hh=hh))
    return out


# ---------- Statistik ----------
def stat(v, tf=0):
    """Statistik mentah + statistik pada skala model (tf=1: akar pangkat tiga)."""
    v = v[~np.isnan(v)]
    if len(v) < 5:
        return dict(mean=0, sd=1, p33=0, p67=0, med=0, min=0, max=0, tm=0, ts=1, tf=tf)
    p33, med, p67 = np.percentile(v, [100 / 3, 50, 200 / 3])
    tv = np.cbrt(v) if tf else v
    return dict(mean=v.mean(), sd=v.std(ddof=1) or 1.0, p33=p33, p67=p67, med=med, min=v.min(), max=v.max(),
                tm=tv.mean(), ts=tv.std(ddof=1) or 1.0, tf=tf)


def clim(s, tf=0):
    return [stat(s[m::12], tf) for m in range(12)]


def corr(x, y):
    if len(x) < 5:
        return 0.0
    c = np.corrcoef(x, y)[0, 1]
    return float(c) if np.isfinite(c) else 0.0


def cor(s, idx):  # korelasi bersamaan (lag 0) CH vs indeks per bulan
    r = []
    for m in range(12):
        pr = [(s[i * 12 + m], idx[(Y0 + i) * 12 + m]) for i in range(N)
              if not np.isnan(s[i * 12 + m]) and (Y0 + i) * 12 + m in idx]
        r.append(corr([p[0] for p in pr], [p[1] for p in pr]))
    return r


# ---------- Model: ridge ENSO+IOD, LOYO nested, uji permutasi ----------
def rows(s, m, h, I, shift=None):
    """Baris latih: (kelompok-tahun, bulan kalender, indeks tahun observasi, nilai mentah, ENSO(t-h), DMI(t-h)).
    shift = permutasi tahun untuk prediktor (uji permutasi blok); None = data asli."""
    R = []
    for y in range(N):
        sh = 12 * (int(shift[y]) - y) if shift is not None else 0
        for d in (-1, 0, 1):
            k = m + d; t = (Y0 + y) * 12 + k; i = t - T0
            if i < 0 or i >= len(s) or np.isnan(s[i]):
                continue
            a, b = I['oni'].get(t - h + sh), I['dmi'].get(t - h + sh)
            if a is None or b is None:
                continue
            R.append((y, k % 12, t // 12 - Y0, s[i], a, b))
    return np.array(R)


def cstats(V, tf, cs, skip=None):
    """(mean, sd) per bulan kalender pada skala model; skip={bulan: indeks tahun} = observasi uji dikeluarkan."""
    out = {}
    for c in cs:
        col = V[:, c].copy()
        if skip and c in skip:
            col[skip[c]] = np.nan
        col = col[~np.isnan(col)]
        if tf:
            col = np.cbrt(col)
        out[c] = (col.mean(), col.std(ddof=1) or 1.0) if len(col) >= 5 else (0.0, 1.0)
    return out


def zvec(R, cl, tf):
    mu = np.array([cl[int(c)][0] for c in R[:, 1]])
    sd = np.array([cl[int(c)][1] for c in R[:, 1]])
    x = np.cbrt(R[:, 3]) if tf else R[:, 3]
    return (x - mu) / sd


def zmat(R, cl, tf):
    return np.column_stack([R[:, 0], zvec(R, cl, tf), R[:, 4], R[:, 5]])


def fit(R):  # kolom R: tahun, z, ENSO, DMI
    n = len(R)
    if n < 20:
        return None
    _, z, a, b = R.T
    ma, mb, mz = a.mean(), b.mean(), z.mean()
    sa, sb = a.std() or 1.0, b.std() or 1.0
    A, B, Z = (a - ma) / sa, (b - mb) / sb, z - mz
    aa, ab, bb, az, bz = A @ A, A @ B, B @ B, A @ Z, B @ Z
    l = 0.05 * n
    D = (aa + l) * (bb + l) - ab * ab
    return dict(ma=ma, sa=sa, mb=mb, sb=sb, mz=mz, b1=((bb + l) * az - ab * bz) / D, b2=((aa + l) * bz - ab * az) / D)


def pred(M, a, b):
    return M['mz'] + M['b1'] * (a - M['ma']) / M['sa'] + M['b2'] * (b - M['mb']) / M['sb']


def loyo(R, V, tf, cs):
    """Leave-one-year-out nested: seluruh baris satu tahun dikeluarkan; klimatologi dihitung ulang tanpa tahun uji.
    Baseline = prakiraan klimatologi (z = 0) pada skala yang sama."""
    if len(R) < 20:
        return None
    sse = sst = 0.0; P, Zs = [], []
    for y in range(N):
        te = R[R[:, 0] == y]
        if len(te) == 0:
            continue
        tr = R[R[:, 0] != y]
        skip = {int(c): int(yo) for c, yo in zip(te[:, 1], te[:, 2])}
        cl = cstats(V, tf, cs, skip)
        Mc = fit(zmat(tr, cl, tf))
        if Mc is None:
            continue
        zt = zvec(te, cl, tf)
        p = pred(Mc, te[:, 4], te[:, 5])
        sse += ((zt - p) ** 2).sum(); sst += (zt ** 2).sum()
        P += list(p); Zs += list(zt)
    if not P or sst == 0:
        return None
    return sse, sst, P, Zs


def model(s, m, h, I, tf, rng):
    V = s.reshape(N, 12)
    cs = sorted({(m + d) % 12 for d in (-1, 0, 1)})
    R = rows(s, m, h, I)
    if len(R) < 20:
        return None
    M = fit(zmat(R, cstats(V, tf, cs), tf))   # model akhir: seluruh data
    out = loyo(R, V, tf, cs)
    if M is None or out is None:
        return None
    sse, sst, P, Zs = out
    sk = 1 - sse / sst
    p = None
    if PERMS > 0:
        ge = 0
        for _ in range(PERMS):
            o2 = loyo(rows(s, m, h, I, rng.permutation(N)), V, tf, cs)
            if o2 and 1 - o2[0] / o2[1] >= sk:
                ge += 1
        p = (1 + ge) / (1 + PERMS)
    use = bool(sk > SKMIN and (p is None or p <= PMAX))
    M.update(rmse=math.sqrt(sse / len(P)), r=corr(P, Zs), skill=sk, n=len(R), p=p, use=use)
    return M


def build(s, I, t0, rng):
    cc, hc = clim(s['ch'], 1), clim(s['hh'], 0)
    fc = []
    for h in range(1, LEADS + 1):
        t = t0 + h; m = t % 12
        fc.append(dict(y=t // 12, m=m, h=h,
                       ch={**cc[m], 'mod': model(s['ch'], m, h, I, 1, rng)},
                       hh={**hc[m], 'mod': model(s['hh'], m, h, I, 0, rng)}))
    return dict(id=s['id'], name=s['name'], loc=s.get('loc'), lat=s.get('lat'), lon=s.get('lon'),
                clim=dict(ch=cc, hh=hc), cor=dict(oni=cor(s['ch'], I['oni']), dmi=cor(s['ch'], I['dmi'])), fc=fc)


def clean(o):  # numpy -> JSON, NaN -> null
    if isinstance(o, dict): return {k: clean(v) for k, v in o.items()}
    if isinstance(o, (list, tuple)): return [clean(v) for v in o]
    if isinstance(o, (bool, np.bool_)): return bool(o)
    if isinstance(o, (float, np.floating)): return round(float(o), 4) if math.isfinite(o) else None
    if isinstance(o, np.integer): return int(o)
    return o


def main():
    u = (os.environ.get('SHEET_XLSX_URL') or '').strip().strip('"\'')  # data hujan dari Google Sheets (link export xlsx)
    print('Secret SHEET_XLSX_URL:', 'terisi' if u else 'KOSONG (memakai berkas xlsx di repo)')
    if u:
        try:
            r = requests.get(u, timeout=120); r.raise_for_status()
            if r.content[:2] != b'PK':  # xlsx = berkas zip; halaman login/HTML tidak diawali 'PK'
                raise ValueError('hasil unduhan bukan berkas xlsx (sheet belum dibagikan "siapa saja yang memiliki tautan", '
                                 'atau tautan bukan format .../export?format=xlsx)')
            pathlib.Path(XLSX).write_bytes(r.content)
            print(f'Data hujan diambil dari Google Sheets ({len(r.content) // 1024} KB)')
        except Exception as e:
            print('Gagal ambil Google Sheets, memakai file xlsx di repo:', e)
    oni, dmi, src, msg = get_indices()
    I = dict(oni=oni, dmi=dmi)
    t0 = min(max(oni), max(dmi))               # bulan terakhir yang lengkap di KEDUA indeks
    st = read_stations()
    reg = lambda k: np.nanmean(np.vstack([s[k] for s in st]), axis=0)
    allst = [dict(id='SULBAR', name=f'SULBAR (rata-rata {len(st)} pos)', loc='Rata-rata seluruh pos',
                  ch=reg('ch'), hh=reg('hh'))] + st
    rng = np.random.default_rng(20240101)      # deterministik -> hasil sama bila data sama
    hist = [dict(y=t // 12, m=t % 12, oni=oni.get(t), dmi=dmi.get(t)) for t in range(t0 - 35, t0 + 1)]
    S = [build(s, I, t0, rng) for s in allst]
    mods = [f[k]['mod'] for s in S for f in s['fc'] for k in ('ch', 'hh') if f[k]['mod']]
    out = dict(stations=S,
               idx=dict(y=t0 // 12, m=t0 % 12, oni=oni[t0], dmi=dmi[t0], dy=t0 // 12, dm=t0 % 12,
                        src=src, bomMsg=msg, hist=hist),
               meta=dict(y0=Y0, y1=Y1, n=len(st), perms=PERMS, skmin=SKMIN, pmax=PMAX,
                         nModels=len(mods), nUse=sum(1 for M in mods if M['use'])),
               updated=dt.datetime.now(dt.timezone.utc).strftime('%Y-%m-%d %H:%M UTC'))
    new = clean(out)
    p = pathlib.Path('data.json')
    if p.exists():
        try:
            old = json.loads(p.read_text(encoding='utf-8')); old.pop('updated', None)
            cmp = dict(new); cmp.pop('updated')
            if old == cmp:
                print('Tidak ada perubahan data -> data.json tidak ditulis ulang.')
                return
        except Exception:
            pass
    p.write_text(json.dumps(new, ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
    print('data.json ditulis | sumber indeks:', src, '| bulan terakhir', t0 // 12, t0 % 12 + 1,
          '| model lolos', out['meta']['nUse'], '/', out['meta']['nModels'], '|', msg)


if __name__ == '__main__':
    main()
