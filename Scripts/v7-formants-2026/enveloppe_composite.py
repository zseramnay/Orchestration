#!/usr/bin/env python3
"""
Methode 8 : enveloppe composite par agregation des amplitudes harmoniques.
Methode 6 : enveloppe vraie (True Envelope) a ordre cepstral adapte a f0.
Source : fichiers Orchidea `spectrum` = spectre moyen NON lisse (1024 bins, 0-11015 Hz).
"""
import re, os, math, json
import numpy as np

FS    = 44100.0
NFFT  = 4096
BW    = FS / NFFT          # 10.7666 Hz / bin
NBINS = 1024
FMAX  = (NBINS - 1) * BW   # ~11015 Hz
QCOMB = 2 * (NBINS - 1)    # longueur effective pour le cepstre demi-spectre

NOTE_RE = re.compile(r'^([A-G])(#?)(-?\d{1,2})$')
SEMI = {'C':0,'D':2,'E':4,'F':5,'G':7,'A':9,'B':11}
DYN_SOFT = {'p','pp','ppp','mp'}
DYN_LOUD = {'f','ff','fff','mf'}

def note_to_f0(tok):
    m = NOTE_RE.match(tok)
    if not m: return None
    L, sharp, octv = m.group(1), m.group(2), int(m.group(3))
    midi = 12*(octv+1) + SEMI[L] + (1 if sharp else 0)
    if not (12 <= midi <= 120): return None
    return 440.0 * 2**((midi-69)/12.0), midi

def parse_name(path):
    parts_path = path.split('/')
    tech = parts_path[3] if len(parts_path) > 3 else '?'
    base = parts_path[-1].replace('.wav','')
    f0 = midi = None; dyn = None
    toks = base.split('-')
    for i, p in enumerate(toks):
        r = note_to_f0(p)
        if r:
            f0, midi = r
            if i+1 < len(toks): dyn = toks[i+1].lower()
            break
    return f0, midi, dyn, tech

def load(fn, techniques):
    out = []
    with open(fn, 'r', errors='ignore') as fh:
        fh.readline()
        for line in fh:
            if not line.startswith('/'): continue
            i = line.index(';')
            f0, midi, dyn, tech = parse_name(line[:i])
            if tech not in techniques or f0 is None: continue
            vals = np.fromstring(line[i+1:], sep=';')
            if vals.size < NBINS: continue
            out.append((f0, midi, dyn, vals[:NBINS]))
    return out

def peak_near(spec, f, tol):
    b = int(round(f / BW))
    lo, hi = max(1, b-tol), min(NBINS, b+tol+1)
    if lo >= hi: return None
    j = lo + int(np.argmax(spec[lo:hi]))
    return j*BW, spec[j]

def harmonics(spec, f0, kmax=90):
    fr, am = [], []
    for k in range(1, kmax+1):
        f = k*f0
        if f > FMAX - 3*BW: break
        r = peak_near(spec, f, 2 if k < 20 else 3)
        if r is None: break
        fr.append(r[0]); am.append(20*math.log10(r[1] + 1e-12))
    return np.array(fr), np.array(am)

# ---------- grille 1/12 d'octave --------------------------------------------
F_LO, F_HI = 80.0, 10000.0
NB = int(round(12*math.log2(F_HI/F_LO)))
EDGES   = F_LO * 2**(np.arange(NB+1)/12.0)
CENTERS = np.sqrt(EDGES[:-1]*EDGES[1:])

HALFWIN = 0.75   # +/- 3/4 d'octave pour le redressement local

def local_residual(fr, am):
    """Retire a chaque harmonique la mediane locale (+/- 3/4 d'octave) de la
       meme note. Insensible a la pente globale, quelle que soit sa forme :
       seules les resonances plus etroites que ~1.5 octave survivent."""
    lf = np.log2(fr)
    res = np.empty_like(am)
    npts = np.empty(am.size, dtype=int)
    for i in range(am.size):
        m = np.abs(lf - lf[i]) <= HALFWIN
        npts[i] = m.sum()
        res[i] = am[i] - np.median(am[m])
    return res, npts

def composite(samples, min_local=3):
    """Residus locaux agreges par bande de 1/12 d'octave."""
    acc  = [[] for _ in range(NB)]
    seen = [set() for _ in range(NB)]
    nd = nm = 0
    for f0, midi, dyn, spec in samples:
        fr, am = harmonics(spec, f0)
        if fr.size == 0: continue
        res, npts = local_residual(fr, am)
        if fr.size >= 6: nd += 1
        else:            nm += 1
        idx = np.digitize(fr, EDGES) - 1
        for i, r, nl in zip(idx, res, npts):
            if 0 <= i < NB and nl >= min_local:
                acc[i].append(r); seen[i].add(midi)
    med = np.array([np.median(a) if len(a) >= 5 else np.nan for a in acc])
    nn  = np.array([len(s) for s in seen])
    npt = np.array([len(a) for a in acc])
    return med, nn, nd, nm, npt

def smooth(y, w=2):
    out = np.full_like(y, np.nan)
    for i in range(len(y)):
        seg = y[max(0,i-w):min(len(y),i+w+1)]
        seg = seg[~np.isnan(seg)]
        if seg.size: out[i] = seg.mean()
    return out

def find_formants(med, nn, npt=None, min_notes=5, min_pts=10, min_res=1.0,
                  max_n=4, fhi=6000.0):
    s = smooth(med)
    ok = (~np.isnan(s)) & (nn >= min_notes) & (CENTERS <= fhi)
    if npt is not None: ok &= (npt >= min_pts)
    peaks = []
    for i in range(1, NB-1):
        if not (ok[i] and ok[i-1] and ok[i+1]): continue
        if s[i] >= s[i-1] and s[i] > s[i+1] and s[i] >= min_res:
            lo = hi = i
            while lo > 0  and ok[lo-1] and s[lo-1] > s[i]-3: lo -= 1
            while hi < NB-1 and ok[hi+1] and s[hi+1] > s[i]-3: hi += 1
            peaks.append((CENTERS[i], s[i], CENTERS[lo], CENTERS[hi], int(nn[i])))
    peaks.sort(key=lambda p: -p[1])
    keep = []
    for p in peaks:
        if all(abs(math.log2(p[0]/q[0])) > 1/3 for q in keep): keep.append(p)
    return keep[:max_n]

# ---------- methode 6 : True Envelope a ordre adaptatif ---------------------
def true_envelope(spec, f0, iters=20):
    q_comb = QCOMB * BW / f0            # indice cepstral du peigne harmonique
    C = max(3, int(0.7 * q_comb))
    logX = np.log(np.maximum(spec, 1e-12))
    V = logX.copy()
    for _ in range(iters):
        V = np.maximum(logX, V)
        cep = np.fft.irfft(V)
        cep[C:len(cep)-C+1] = 0.0
        V = np.fft.rfft(cep).real[:NBINS]
    return V, C

def te_peak(spec, f0, fhi=6000.0):
    env, C = true_envelope(spec, f0)
    hi = int(fhi/BW)
    seg = env[3:hi]
    cand = [i for i in range(1, len(seg)-1) if seg[i] >= seg[i-1] and seg[i] > seg[i+1]]
    if not cand: return float('nan'), C
    j = max(cand, key=lambda i: seg[i])
    return (j+3)*BW, C

# ---------- pipeline --------------------------------------------------------
INSTRUMENTS = [
    ("Petite flûte",           "Yan_Adds-Divers.spectrum_par_instrument/Piccolo_spectrum.txt",               ["ordinario"]),
    ("Flûte",                  "FullSOL2020.spectrum_par_instrument/Flute_spectrum.txt",                     ["ordinario"]),
    ("Flûte basse",            "Yan_Adds-Divers.spectrum_par_instrument/Bass-Flute_spectrum.txt",            ["ordinario"]),
    ("Flûte contrebasse",      "Yan_Adds-Divers.spectrum_par_instrument/Contrabass-Flute_spectrum.txt",      ["ordinario"]),
    ("Hautbois",               "FullSOL2020.spectrum_par_instrument/Oboe_spectrum.txt",                      ["ordinario"]),
    ("Cor anglais",            "Yan_Adds-Divers.spectrum_par_instrument/EnglishHorn_spectrum.txt",           ["ordinario"]),
    ("Basson",                 "FullSOL2020.spectrum_par_instrument/Bassoon_spectrum.txt",                   ["ordinario"]),
    ("Contrebasson",           "Yan_Adds-Divers.spectrum_par_instrument/Contrebasson_spectrum.txt",          ["non-vibrato"]),
    ("Clarinette en Mib",      "Yan_Adds-Divers.spectrum_par_instrument/Clarinet-Eb_spectrum.txt",           ["ordinario"]),
    ("Clarinette en Sib",      "FullSOL2020.spectrum_par_instrument/Clarinet_Bb_spectrum.txt",               ["ordinario"]),
    ("Clarinette basse",       "Yan_Adds-Divers.spectrum_par_instrument/Bass-Clarinet-Bb_spectrum.txt",      ["ordinario"]),
    ("Clarinette contrebasse", "Yan_Adds-Divers.spectrum_par_instrument/Contrabass-Clarinet-Bb_spectrum.txt",["ordinario"]),
]

def spearman(a, b):
    ra = np.argsort(np.argsort(a)).astype(float)
    rb = np.argsort(np.argsort(b)).astype(float)
    ra -= ra.mean(); rb -= rb.mean()
    d = math.sqrt((ra**2).sum() * (rb**2).sum())
    return float((ra*rb).sum()/d) if d else float('nan')

BASE = "/home/claude/Orchestration/Data/"
OUT  = "/tmp/claude-0/-home-claude/f9fbd600-866a-5b45-8e67-755d352e197e/scratchpad/"
results = {}

for name, fn, techs in INSTRUMENTS:
    path = BASE + fn
    if not os.path.exists(path):
        print(f"!! absent : {fn}"); continue
    s = load(path, techs)
    if len(s) < 10:
        print(f"!! trop peu d'echantillons : {name} ({len(s)})"); continue

    # controle de nomenclature : hauteur annoncee vs pic reel
    ratios = [peak_near(sp, f0, 4)[0]/f0 for f0, mi, dy, sp in s if peak_near(sp, f0, 4)]
    nomen = float(np.median(ratios))

    f0s = np.array([x[0] for x in s])
    med, nn, nd, nm, npt = composite(s)
    form = find_formants(med, nn, npt)

    cut  = float(np.median(f0s))
    lo_s = [x for x in s if x[0] <= cut]
    hi_s = [x for x in s if x[0] >  cut]
    m,n2,_,_,p2 = composite(lo_s); f_lo = find_formants(m, n2, p2, min_notes=3, min_pts=5)
    m,n2,_,_,p2 = composite(hi_s); f_hi = find_formants(m, n2, p2, min_notes=3, min_pts=5)

    soft = [x for x in s if (x[2] or '') in DYN_SOFT]
    loud = [x for x in s if (x[2] or '') in DYN_LOUD]
    f_soft = f_loud = []
    if len(soft) > 10:
        m,n2,_,_,p2 = composite(soft); f_soft = find_formants(m, n2, p2, min_notes=3, min_pts=5)
    if len(loud) > 10:
        m,n2,_,_,p2 = composite(loud); f_loud = find_formants(m, n2, p2, min_notes=3, min_pts=5)

    te, orders = [], []
    for f0, mi, dy, sp in s:
        v, C = te_peak(sp, f0); te.append(v); orders.append(C)
    te = np.array(te, float); orders = np.array(orders)
    g  = ~np.isnan(te)
    rho = spearman(f0s[g], te[g]) if g.sum() > 5 else float('nan')
    onh = float(np.mean([min(abs(math.log2(t/(k*f))) for k in (1,2,3,4)) < 0.04
                         for t, f in zip(te[g], f0s[g])])) if g.sum() else float('nan')

    results[name] = dict(
        n=len(s), tech=techs[0], f0min=float(f0s.min()), f0max=float(f0s.max()),
        nomen=nomen, n_detrended=nd, n_meanonly=nm,
        formants=[(round(a,1), round(b,2), round(c,1), round(d,1), e) for a,b,c,d,e in form],
        f_lo=[(round(a,1), round(b,2)) for a,b,*_ in f_lo],
        f_hi=[(round(a,1), round(b,2)) for a,b,*_ in f_hi],
        f_soft=[(round(a,1), round(b,2)) for a,b,*_ in f_soft],
        f_loud=[(round(a,1), round(b,2)) for a,b,*_ in f_loud],
        te_median=float(np.nanmedian(te)), te_sd=float(np.nanstd(te)),
        te_rho=rho, te_onharm=onh,
        te_order_min=int(orders.min()), te_order_max=int(orders.max()),
        med=[None if np.isnan(v) else round(float(v),2) for v in med],
        nn=[int(v) for v in nn], npt=[int(v) for v in npt],
    )
    print(f"OK {name:24s} n={len(s):4d}  nomen={nomen:.3f}  "
          f"detrend={nd}/{nd+nm}  F={[f'{a:.0f}({b:+.1f})' for a,b,c,d,e in form]}")

with open(OUT+"res.json","w") as fh:
    json.dump({"centers":[round(c,1) for c in CENTERS], "results":results}, fh)
print("\nterminé —", len(results), "instruments")
