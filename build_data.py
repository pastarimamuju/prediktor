#!/usr/bin/env python3
"""
Membangun data.json (baseline + prakiraan hujan Sulbar berbasis ENSO/IOD) untuk dashboard GitHub Pages.

Struktur repo (root): index.html, build_data.py, Data_Curah_Hujan_-_Mamuju.xlsx, .github/workflows/update.yml
Lokal : pip install numpy openpyxl requests netCDF4 && python build_data.py

PERUBAHAN UTAMA (v3: seluruh indeks dari NOAA, BoM dihapus)
 A. Niño 3.4 : NOAA CPC, berkas mingguan wksst9120.for (OISST, basis 1991-2020), dirata-ratakan bulanan.
 B. DMI      : dihitung sendiri dari NOAA ERSSTv6 bulanan (sst.mnmean.nc di NOAA PSL) dengan kotak Saji et al. (1999):
               WTIO 50-70E, 10S-10N dikurangi SETIO 90-110E, 10S-0; anomali terhadap klimatologi 1991-2020
               (basis sama dengan Niño 3.4 CPC). Alasan: seri DMI siap-pakai NOAA PSL (HadISST, dmi.had.long.data)
               berhenti di 2025 sehingga tidak bisa dipakai untuk prakiraan terbaru. Tidak ada lagi kalibrasi/penambalan BoM.
 C. Berkas .nc (~90 MB) diunduh ke folder sementara dan TIDAK disimpan di repo; yang di-cache hanya
    cache/dmi_ersst.csv (kecil) sebagai cadangan bila unduhan gagal.

Perubahan v2 yang tetap berlaku
 1. LOYO "nested": klimatologi (mean, sd) dihitung ulang tanpa tahun uji di tiap fold, sehingga baseline
    pembanding (prakiraan klimatologi) dan target distandardisasi benar-benar di luar sampel.
    Skill = MSESS terhadap baseline nested (bukan baseline in-sample).
 2. Uji signifikansi: permutasi tahun (blok) pada prediktor, PERMS kali -> p-value per model.
    Model dipakai bila skill > SKMIN dan p <= PMAX; selain itu kembali ke klimatologi.
 3. Curah hujan dimodelkan pada skala akar pangkat tiga (distribusi miring kanan -> mendekati normal),
    sehingga peluang kategori tidak bias. Hari hujan tetap skala linear.
 4. MIN_WEEKS = 3. Bulan terakhir = bulan lengkap terakhir yang tersedia di KEDUA indeks.
 5. data.json hanya ditulis ulang bila isi (selain stempel waktu) berubah -> tidak ada commit harian sia-sia.
 6. meta.oniRange / meta.dmiRange = rentang indeks pada periode latih; dashboard memakainya untuk membatasi
    skenario ekstrem (mis. El Nino sangat kuat) dan menampilkan peringatan.
 7. Bulan terakhir (t0) tidak pernah melebihi bulan kalender penuh terakhir.
"""
import datetime as dt, json, math, os, pathlib, re, tempfile, warnings
import numpy as np, requests
from openpyxl import load_workbook

warnings.simplefilter('ignore')
XLSX = 'Data_Curah_Hujan_-_Mamuju.xlsx'
Y0, Y1, LEADS, MIN_WEEKS = 1996, 2025, 6, 3
SKMIN = float(os.environ.get('SKMIN', 0.03))   # ambang MSESS (nested)
PMAX = float(os.environ.get('PMAX', 0.10))     # ambang p-value permutasi
PERMS = int(os.environ.get('PERMS', 99))       # 0 = lewati uji permutasi
ENSO_URL = 'https://www.cpc.ncep.noaa.gov/data/indices/wksst9120.for'
# ERSSTv6 bulanan (NOAA PSL). Dua alamat unduh yang setara; dicoba berurutan.
ERSST_URLS = ['https://psl.noaa.gov/thredds/fileServer/Datasets/noaa.ersst.v6/sst.mnmean.nc',
              'https://downloads.psl.noaa.gov/Datasets/noaa.ersst.v6/sst.mnmean.nc']
BASE0, BASE1 = 1991, 2020                      # periode klimatologi anomali DMI (sama dengan wksst9120)
WTIO = (50.0, 70.0, -10.0, 10.0)               # (lon0, lon1, lat0, lat1) kutub barat
SETIO = (90.0, 110.0, -10.0, 0.0)              # kutub timur
UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'
N, T0 = Y1 - Y0 + 1, Y0 * 12
CACHE = pathlib.Path('cache'); CACHE.mkdir(exist_ok=True)
DMI_CACHE = CACHE / 'dmi_ersst.csv'
MON = {m: i for i, m in enumerate('JAN FEB MAR APR MAY JUN JUL AUG SEP OCT NOV DEC'.split())}


# ---------- Unduh & parsing indeks ----------
def fetch(urls, name):
    err = ''
    for u in urls:
        try:
            r = requests.get(u, timeout=60, headers={'User-Agent': UA})
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


# ---------- DMI dari NOAA ERSSTv6 ----------
def box_mean(sst, lat, lon, box):
    """Rerata kotak berbobot cos(lat); sst berbentuk (waktu, lat, lon); NaN dilewati. Hasil: vektor (waktu)."""
    lo0, lo1, la0, la1 = box
    li = (lat >= la0) & (lat <= la1)
    lj = (lon >= lo0) & (lon <= lo1)
    if not li.any() or not lj.any():
        raise ValueError(f'kotak {box} tidak mengenai grid (lat {lat.min()}..{lat.max()}, lon {lon.min()}..{lon.max()})')
    sub = sst[:, li][:, :, lj]
    w = np.broadcast_to(np.cos(np.deg2rad(lat[li]))[None, :, None], sub.shape)
    ok = np.isfinite(sub)
    num = (np.where(ok, sub, 0.0) * w).sum(axis=(1, 2))
    den = np.where(ok, w, 0.0).sum(axis=(1, 2))
    out = np.full(sub.shape[0], np.nan)
    m = den > 0
    out[m] = num[m] / den[m]
    return out


def dmi_from_nc(path):
    """Baca sst.mnmean.nc -> {t: DMI}, t = tahun*12 + bulan(0-11). Anomali basis BASE0-BASE1, DMI = WTIO' - SETIO'."""
    from netCDF4 import Dataset, num2date
    with Dataset(path) as nc:
        vn = 'sst' if 'sst' in nc.variables else next(
            (k for k, v in nc.variables.items() if v.ndim >= 3 and k not in ('time_bnds',)), None)
        if vn is None:
            raise ValueError('variabel SST tidak ditemukan: ' + ', '.join(nc.variables))
        lat = np.array(nc.variables['lat'][:], dtype=float)
        lon = np.mod(np.array(nc.variables['lon'][:], dtype=float), 360.0)
        tv = nc.variables['time']
        dates = num2date(tv[:], tv.units, getattr(tv, 'calendar', 'standard'))
        raw = nc.variables[vn][...]
        sst = np.ma.filled(np.ma.asarray(raw, dtype=float), np.nan)
    sst = sst.reshape(len(dates), len(lat), len(lon))   # buang dimensi tunggal (mis. zlev)
    sst[(sst < -5) | (sst > 45)] = np.nan               # jaga-jaga nilai isian yang lolos
    t = np.array([d.year * 12 + d.month - 1 for d in dates])
    w, e = box_mean(sst, lat, lon, WTIO), box_mean(sst, lat, lon, SETIO)
    out = {}
    cw, ce = {}, {}
    for m in range(12):
        sel = (t % 12 == m) & (t >= BASE0 * 12) & (t < (BASE1 + 1) * 12) & np.isfinite(w) & np.isfinite(e)
        if sel.sum() < 20:
            raise ValueError(f'klimatologi bulan {m + 1} hanya {int(sel.sum())} tahun')
        cw[m], ce[m] = w[sel].mean(), e[sel].mean()
    for i, ti in enumerate(t):
        if np.isfinite(w[i]) and np.isfinite(e[i]):
            m = int(ti % 12)
            out[int(ti)] = float((w[i] - cw[m]) - (e[i] - ce[m]))
    if len(out) < 600:
        raise ValueError(f'seri DMI terlalu pendek ({len(out)} bulan)')
    return out


def download_nc(urls):
    err = ''
    for u in urls:
        tmp = None
        try:
            with requests.get(u, stream=True, timeout=300, headers={'User-Agent': UA}) as r:
                r.raise_for_status()
                with tempfile.NamedTemporaryFile(suffix='.nc', delete=False) as f:
                    tmp = f.name
                    for ch in r.iter_content(1 << 20):
                        f.write(ch)
            head = open(tmp, 'rb').read(4)
            if not (head[:3] == b'CDF' or head == b'\x89HDF'):   # bukan netCDF (mis. halaman HTML)
                raise ValueError('berkas unduhan bukan netCDF')
            return tmp, u
        except Exception as e:
            err = f'{type(e).__name__}: {str(e)[:90]}'
            if tmp and os.path.exists(tmp):
                os.unlink(tmp)
    return None, err


def read_dmi_cache():
    o = {}
    if DMI_CACHE.exists():
        for l in DMI_CACHE.read_text(encoding='utf-8').splitlines():
            p = l.split(',')
            if len(p) == 3 and p[0].isdigit():
                o[int(p[0]) * 12 + int(p[1]) - 1] = float(p[2])
    return o


def write_dmi_cache(d):
    DMI_CACHE.write_text('\n'.join(f'{t // 12},{t % 12 + 1},{v:.4f}' for t, v in sorted(d.items())) + '\n',
                         encoding='utf-8')


def get_dmi():
    path, info = download_nc(ERSST_URLS)
    err = info
    if path:
        try:
            d = dmi_from_nc(path)
            write_dmi_cache(d)
            return d, f'ERSSTv6 unduh langsung OK ({info})'
        except Exception as e:
            err = f'{type(e).__name__}: {str(e)[:120]}'
        finally:
            if os.path.exists(path):
                os.unlink(path)
    d = read_dmi_cache()
    if len(d) >= 600:
        return d, f'cache lokal ({err})'
    raise SystemExit('Data DMI (NOAA ERSSTv6) tidak bisa dibaca dan cache kosong: ' + err)


def get_indices():
    txt, m1 = fetch([ENSO_URL], 'wksst9120.for')
    oni = parse_weekly(txt) if txt else {}
    if len(oni) < 20:
        raise SystemExit('Data Nino 3.4 (NOAA CPC) tidak terbaca: ' + m1)
    dmi, m2 = get_dmi()
    need = T0 - 8  # prediktor lead 6 + pooling +-1 bulan
    if min(oni) > need or min(dmi) > need:
        raise SystemExit(f'Indeks tidak mencakup periode latih: ONI mulai {min(oni) // 12}, DMI mulai {min(dmi) // 12}')
    return oni, dmi, 'CPC Niño 3.4 (OISST) + DMI ERSSTv6 (NOAA PSL)', f'ENSO: {m1}; IOD: {m2}'


def idx_range(d):
    """Rentang [min, max] indeks pada periode latih (Y0..Y1); dipakai dashboard untuk membatasi skenario ekstrem."""
    v = [x for t, x in d.items() if T0 <= t < T0 + N * 12]
    return [round(min(v), 3), round(max(v), 3)] if v else None


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
    today = dt.datetime.now(dt.timezone.utc)
    last_full = today.year * 12 + today.month - 2   # bulan kalender penuh terakhir (bulan berjalan belum lengkap)
    t0 = min(max(oni), max(dmi), last_full)         # bulan terakhir yang lengkap di KEDUA indeks
    if t0 not in oni or t0 not in dmi:              # bulan itu bisa kosong bila data terlambat: mundur ke yang ada
        t0 = max(t for t in oni if t <= t0 and t in dmi)
    if t0 < last_full - 1:
        print(f'PERINGATAN: indeks tertinggal {last_full - t0} bulan dari bulan penuh terakhir '
              f'(ONI sampai {max(oni) // 12}-{max(oni) % 12 + 1:02d}, DMI sampai {max(dmi) // 12}-{max(dmi) % 12 + 1:02d}). '
              'Awal prakiraan ikut mundur; periksa sumber data.')
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
                        src=src, srcMsg=msg, hist=hist),
               meta=dict(y0=Y0, y1=Y1, n=len(st), perms=PERMS, skmin=SKMIN, pmax=PMAX,
                         nModels=len(mods), nUse=sum(1 for M in mods if M['use']),
                         oniRange=idx_range(oni), dmiRange=idx_range(dmi)),
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
          '| model lolos', out['meta']['nUse'], '/', out['meta']['nModels'], '| rentang ONI', out['meta']['oniRange'],
          'DMI', out['meta']['dmiRange'], '|', msg)


if __name__ == '__main__':
    main()
