#!/usr/bin/env python3
"""Tests de rejet : independance au rang harmonique + robustesse a la fenetre."""
import sys, math, json
import numpy as np
sys.path.insert(0, "/tmp/claude-0/-home-claude/f9fbd600-866a-5b45-8e67-755d352e197e/scratchpad")
import importlib.util
spec = importlib.util.spec_from_file_location("comp",
    "/tmp/claude-0/-home-claude/f9fbd600-866a-5b45-8e67-755d352e197e/scratchpad/composite.py")
# on recharge seulement les fonctions, sans relancer le pipeline
src = open("/tmp/claude-0/-home-claude/f9fbd600-866a-5b45-8e67-755d352e197e/scratchpad/composite.py").read()
src = src.split("# ---------- pipeline")[0]
ns = {}
exec(compile(src, "comp", "exec"), ns)
load, harmonics, local_residual = ns['load'], ns['harmonics'], ns['local_residual']
composite, find_formants, smooth = ns['composite'], ns['find_formants'], ns['smooth']
CENTERS, EDGES, NB = ns['CENTERS'], ns['EDGES'], ns['NB']
BASE = "/home/claude/Orchestration/Data/"

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

KGROUPS = [(1,1,"k=1"), (2,3,"k=2–3"), (4,7,"k=4–7"), (8,15,"k=8–15"), (16,99,"k≥16")]

def collect(samples, halfwin=0.75):
    """Retourne la liste (freq, residu, k, midi) pour tous les harmoniques."""
    ns['HALFWIN'] = halfwin
    pts = []
    for f0, midi, dyn, spec in samples:
        fr, am = harmonics(spec, f0)
        if fr.size == 0: continue
        res, npl = local_residual(fr, am)
        for k in range(fr.size):
            if npl[k] >= 3:
                pts.append((fr[k], res[k], k+1, midi))
    return pts

def band_by_k(pts, fc, tol_oct=1/12):
    """Residu median par groupe de rang harmonique dans une bande autour de fc."""
    out = []
    for lo, hi, lab in KGROUPS:
        v = [r for f, r, k, m in pts if lo <= k <= hi and abs(math.log2(f/fc)) <= tol_oct]
        nm = len({m for f, r, k, m in pts if lo <= k <= hi and abs(math.log2(f/fc)) <= tol_oct})
        out.append((lab, (float(np.median(v)) if len(v) >= 5 else None), len(v), nm))
    return out

report = {}
for name, fn, techs in INSTRUMENTS:
    s = load(BASE+fn, techs)
    ns['HALFWIN'] = 0.75
    med, nn, nd, nm, npt = composite(s)
    form = find_formants(med, nn, npt)
    pts = collect(s, 0.75)

    # robustesse a la fenetre de redressement
    robust = {}
    for hw in (0.5, 1.0):
        ns['HALFWIN'] = hw
        m2, n2, _, _, p2 = composite(s)
        robust[hw] = [a for a, b, c, d, e in find_formants(m2, n2, p2)]
    ns['HALFWIN'] = 0.75

    rows = []
    for fc, amp, blo, bhi, nnotes in form:
        byk = band_by_k(pts, fc)
        # critere : >= 2 groupes avec k>=2, residu >= +1 dB, >= 5 notes distinctes
        ok_groups = [lab for lab, v, npts_, nmd in byk[1:]
                     if v is not None and v >= 1.0 and nmd >= 5]
        # robustesse : un pic a moins de 1/3 d'octave dans les deux variantes
        rob = all(any(abs(math.log2(p/fc)) <= 1/3 for p in robust[hw]) for hw in (0.5, 1.0))
        verdict = "RETENU" if (len(ok_groups) >= 2 and rob) else \
                  ("fragile" if len(ok_groups) >= 2 or rob else "REJETE")
        rows.append(dict(fc=round(fc,1), amp=round(amp,2), blo=round(blo,1), bhi=round(bhi,1),
                         nnotes=nnotes, byk=[(l, None if v is None else round(v,2), c, d) for l,v,c,d in byk],
                         ok_groups=ok_groups, robust=rob, verdict=verdict,
                         rob05=[round(p) for p in robust[0.5]], rob10=[round(p) for p in robust[1.0]]))
    report[name] = dict(n=len(s), tech=techs[0], rows=rows)
    print(f"\n{name}  (n={len(s)}, {techs[0]})")
    for r in rows:
        bk = "  ".join(f"{l}:{'—' if v is None else f'{v:+.1f}'}({d})" for l,v,c,d in r['byk'])
        print(f"  {r['fc']:7.0f} Hz  {r['amp']:+5.1f} dB  [{r['blo']:.0f}–{r['bhi']:.0f}]  "
              f"{r['verdict']:8s} rob={r['robust']}")
        print(f"            {bk}")

json.dump(report, open("/tmp/claude-0/-home-claude/f9fbd600-866a-5b45-8e67-755d352e197e/scratchpad/valid.json","w"))
print("\nfait")
