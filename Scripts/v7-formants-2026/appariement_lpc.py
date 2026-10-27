# -*- coding: utf-8 -*-
"""
appariement_lpc.py — tâche C4 : transformer la corroboration LPC en test.

Ce que la corroboration valait jusqu'ici
----------------------------------------
La route B disait : « corroboré par deux familles de méthodes indépendantes ». En pratique on
comparait des listes — la méthode 8 trouve un formant vers 2 216 Hz, le LPC trouve un pôle dans
la même région, donc les deux se confirment. C'est un **indice**, pas un test : avec six à douze
pôles répartis sur 6 kHz, un pôle tombe près de n'importe quelle fréquence par hasard assez
souvent pour que la coïncidence globale ne prouve rien.

Le test
-------
On apparie **note par note**, par nom de fichier — les deux analyses portent sur les mêmes
échantillons. Pour un formant candidat à fc, et pour chaque note n :

    A(n) = 1  si la méthode 8 donne un résidu local POSITIF à l'harmonique le plus proche de fc
    B(n) = 1  si le LPC de cette même note place un pôle à moins de TOL de fc

La question n'est plus « les deux méthodes trouvent-elles la même chose en moyenne » mais
« trouvent-elles la même chose **sur les mêmes notes** ». On lit donc le tableau 2×2 :

    P(B | A=1)  contre  P(B | A=0)

P(B | A=0) **est** le niveau de hasard, mesuré sur les notes où la méthode 8 ne dit rien, au lieu
d'être supposé. Le rapport des cotes et un test exact de Fisher unilatéral chiffrent l'écart.

Témoin négatif
--------------
Le même test est refait à un triton de fc. Si l'association à fc tient mais que celle du témoin
ne tient pas, l'accord est localisé ; si le témoin passe aussi, c'est que les deux méthodes
partagent un biais large bande et que l'accord ne prouve rien.

Choix fixés d'avance, pour ne pas les ajuster après coup
--------------------------------------------------------
  · ordre LPC 12 en primaire — le moins de pôles, donc le hasard le plus bas. 16, 20 et 24 sont
    rapportés pour montrer la dépendance à l'ordre.
  · tolérance ±2 demi-tons en primaire. Les largeurs de pôle valant des centaines de hertz, une
    fenêtre plus serrée serait injuste envers le LPC. ±1 et ±4 sont rapportés.
  · segment v7 (attaque exclue), le même que celui de la méthode 8. Le LPC de mai 2026 est à
    44,1 kHz et ordre 18 sur un autre segment : l'apparier à la méthode 8 note à note
    mélangerait deux différences à la fois.

Usage
    python3 appariement_lpc.py --instruments Alto,Violoncelle,Basson
"""
import os, sys, csv, math, json, argparse
from collections import defaultdict
import numpy as np

ICI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ICI)
import formants_lib as FL
from fenetre_redressement import NOMS

RACINE = os.path.dirname(os.path.dirname(ICI))
ETUDE = os.path.join(RACINE, 'Etude-Formants')
SOL_OUT = os.path.expanduser('~/work/out')
LPC = os.path.join(ETUDE, 'Largeurs', 'LPC-v7')

ORDRE_PRIMAIRE = '12'
TOL_PRIMAIRE = 2.0          # demi-tons
TOLS = (1.0, 2.0, 4.0)
ORDRES = ('12', '16', '20', '24')
TRITON = 2 ** (6.0 / 12)

# ───────────────────── lecture des spectres AVEC le nom de fichier ──────────
def lire_avec_nom(chemin):
    """[(nom_fichier, f0, midi, dyn, spectre)] — `FL.lire_spectres` jette le nom,
    or c'est la seule clé d'appariement fiable : plusieurs échantillons partagent
    la même hauteur et la même dynamique (cordes différentes, reprises)."""
    out = []
    with open(chemin, errors='ignore') as fh:
        fh.readline()
        for L in fh:
            if not L.startswith('/'):
                continue
            i = L.index(';')
            chemin_ech = L[:i]
            f0, midi, dyn = FL.analyser_nom(chemin_ech)
            if f0 is None:
                continue
            v = np.fromstring(L[i + 1:], sep=';')
            if v.size >= FL.NBINS:
                out.append((os.path.basename(chemin_ech), f0, midi, dyn, v[:FL.NBINS]))
    return out

def poles_par_fichier(chemin_csv):
    """{nom_fichier: {ordre: [freq_hz]}}"""
    d = defaultdict(lambda: defaultdict(list))
    with open(chemin_csv, encoding='utf-8') as f:
        for r in csv.DictReader(f):
            try:
                d[r['filename']][r['ordre']].append(float(r['freq_hz']))
            except (ValueError, KeyError):
                continue
    return d

# ──────────────────────────── test exact de Fisher ─────────────────────────
def fisher_unilateral(a, b, c, d):
    """p-value unilatérale (association positive) du tableau [[a,b],[c,d]].
    Écrit à la main : scipy n'est pas disponible sur ce poste."""
    n = a + b + c + d
    if n == 0 or (a + b) == 0 or (a + c) == 0:
        return None
    p = 0.0
    lim = min(a + b, a + c)
    for k in range(a, lim + 1):
        p += (math.comb(a + b, k) * math.comb(c + d, a + c - k)
              / math.comb(n, a + c))
    return min(1.0, p)

def tableau(A, B):
    a = sum(1 for x, y in zip(A, B) if x and y)
    b = sum(1 for x, y in zip(A, B) if x and not y)
    c = sum(1 for x, y in zip(A, B) if not x and y)
    d = sum(1 for x, y in zip(A, B) if not x and not y)
    pa = a / (a + b) if a + b else None
    pc = c / (c + d) if c + d else None
    orr = ((a * d) / (b * c)) if b and c else (float('inf') if a and d else None)
    return dict(a=a, b=b, c=c, d=d, p_B_si_A=pa, p_B_sinon=pc,
                cote=None if orr is None else (round(orr, 2) if orr != float('inf') else 'inf'),
                p_fisher=fisher_unilateral(a, b, c, d))

# ────────────────────────────── le test, par formant ───────────────────────
def preparer(ech):
    """Harmoniques et résidus calculés UNE fois par note. Sans cela le test de rang,
    qui évalue des dizaines de fréquences témoins, serait hors de portée."""
    out = []
    for nom, f0, midi, dyn, sp in ech:
        fr, am = FL.harmoniques(sp, f0)
        if fr.size == 0:
            continue
        res, npts = FL.residu_local(fr, am)
        out.append((nom, fr, res, npts))
    return out

def signaux(notes, poles, fc, ordre, tol_dt):
    """(A, B) note par note pour une fréquence fc."""
    A, B = [], []
    lfc = math.log2(fc)
    for nom, fr, res, npts in notes:
        ec = np.abs(12 * (np.log2(np.maximum(fr, 1e-9)) - lfc))
        i = int(np.argmin(ec))
        if ec[i] > 1.0 or npts[i] < 3:
            continue
        A.append(1 if res[i] > 0 else 0)
        pl = poles.get(nom, {}).get(ordre, [])
        B.append(1 if any(abs(12 * (math.log2(max(p, 1e-9)) - lfc)) <= tol_dt
                          for p in pl) else 0)
    return A, B

def rang_de_fc(notes, poles, fc, ordre, tol_dt, exclure, f_bas=150.0, f_haut=6500.0,
               garde_dt=2.0, min_marge=5):
    """Le cœur du test corrigé.

    Le témoin au triton a montré que l'association note-à-note n'est PAS localisée :
    les deux méthodes s'accordent sur les notes qui ont de l'énergie dans une large
    région, pas sur l'endroit de la résonance. Un test juste doit donc comparer fc
    **au reste du spectre**, et non à une seule fréquence témoin.

    On évalue la cote à fc et sur toutes les fréquences de la grille tempérée du
    domaine, en écartant une bande de garde autour de fc et autour des autres formants
    établis du même instrument. Le rang de fc dans cette distribution dit si l'accord
    est propre à fc. Toutes les fréquences témoins partagent la même variation de
    contenu spectral d'une note à l'autre : le facteur de confusion est donc neutralisé.
    """
    temoins = []
    f = f_bas
    while f <= f_haut:
        if all(abs(12 * math.log2(f / g)) > garde_dt for g in list(exclure) + [fc]):
            t = tableau(*signaux(notes, poles, f, ordre, tol_dt))
            if (t['a'] + t['b'] >= min_marge and t['c'] + t['d'] >= min_marge
                    and t['cote'] is not None):
                c = t['cote']
                temoins.append(float('inf') if c == 'inf' else float(c))
        f *= 2 ** (1.0 / 12)
    t = tableau(*signaux(notes, poles, fc, ordre, tol_dt))
    c = t['cote']
    if c is None or not temoins:
        return None, len(temoins), None
    cf = float('inf') if c == 'inf' else float(c)
    plus_bas = sum(1 for x in temoins if x < cf)
    return round(100.0 * plus_bas / len(temoins), 1), len(temoins), \
        round(float(np.median([x for x in temoins if x != float('inf')] or [0])), 2)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--instruments', default='')
    ap.add_argument('--sortie', default=os.path.join(ETUDE, 'Mesures', 'Appariement-LPC'))
    a = ap.parse_args()
    os.makedirs(a.sortie, exist_ok=True)

    with open(os.path.join(ETUDE, 'Resultats-v7', 'formants-v7.csv'), encoding='utf-8') as f:
        FORM = [r for r in csv.DictReader(f) if r['statut'] in ('route A', '1 corpus')]
    voulus = {x.strip() for x in a.instruments.split(',') if x.strip()}
    par_inst = defaultdict(list)
    for r in FORM:
        if not voulus or r['instrument'] in voulus:
            par_inst[r['instrument']].append((float(r['hz']), r['note'], r['statut']))

    lignes = []
    for nom in sorted(par_inst):
        fich = NOMS.get(nom)
        psp = os.path.join(SOL_OUT, (fich or '') + '.txt')
        pcsv = os.path.join(LPC, (fich or '') + '_largeurs.csv')
        if not fich or not os.path.exists(psp) or not os.path.exists(pcsv):
            print('{:28s} — pas de LPC ({})'.format(nom, fich or '?'))
            continue
        notes = preparer(lire_avec_nom(psp))
        poles = poles_par_fichier(pcsv)
        communs = len({e[0] for e in notes} & set(poles))
        autres = [x[0] for x in par_inst[nom]]
        print('\n=== {}   {} notes, {} appariées au LPC'.format(nom, len(notes), communs))
        print('{:>7s} {:>7s} {:>4s} {:>10s} {:>10s} {:>7s} {:>9s} {:>8s} {:>9s}  {}'.format(
            'formant', 'statut', 'n', 'P(B|A=1)', 'P(B|A=0)', 'cote', 'Fisher',
            'triton', 'rang fc', 'cote médiane des témoins'))
        for fc, note, statut in sorted(par_inst[nom]):
            t = tableau(*signaux(notes, poles, fc, ORDRE_PRIMAIRE, TOL_PRIMAIRE))
            tt = tableau(*signaux(notes, poles, fc * TRITON, ORDRE_PRIMAIRE, TOL_PRIMAIRE))
            rang, n_t, med_t = rang_de_fc(notes, poles, fc, ORDRE_PRIMAIRE, TOL_PRIMAIRE,
                                          [g for g in autres if g != fc])
            n = t['a'] + t['b'] + t['c'] + t['d']
            d = dict(instrument=nom, formant_hz=int(fc), note=note, statut=statut,
                     n=n, **{k: t[k] for k in ('a', 'b', 'c', 'd', 'cote', 'p_fisher')},
                     p_B_si_A=None if t['p_B_si_A'] is None else round(t['p_B_si_A'], 3),
                     p_B_sinon=None if t['p_B_sinon'] is None else round(t['p_B_sinon'], 3),
                     temoin_triton_cote=tt['cote'],
                     rang_percentile=rang, n_temoins=n_t, cote_mediane_temoins=med_t)
            for o in ORDRES:
                d['cote_ordre_' + o] = tableau(*signaux(notes, poles, fc, o,
                                                        TOL_PRIMAIRE))['cote']
            for tl in TOLS:
                d['cote_tol_%g' % tl] = tableau(*signaux(notes, poles, fc,
                                                         ORDRE_PRIMAIRE, tl))['cote']
            lignes.append(d)
            def s(x, d_='—'):
                return d_ if x is None else (('%.4f' % x) if isinstance(x, float)
                                             and x < 0.001 else str(x))
            print('{:>6}Hz {:>7s} {:4d} {:>10} {:>10} {:>7} {:>9} {:>8} {:>7}%  '
                  '{} (n={})'.format(
                      int(fc), statut[:7], n, s(d['p_B_si_A']), s(d['p_B_sinon']),
                      s(d['cote']), s(d['p_fisher']), s(tt['cote']),
                      s(rang), s(med_t), n_t))

    with open(os.path.join(a.sortie, 'appariement-lpc.json'), 'w', encoding='utf-8') as f:
        json.dump(lignes, f, ensure_ascii=False, indent=1, default=str)
    if lignes:
        with open(os.path.join(a.sortie, 'appariement-lpc.csv'), 'w', encoding='utf-8',
                  newline='') as f:
            w = csv.DictWriter(f, fieldnames=list(lignes[0]))
            w.writeheader(); w.writerows(lignes)
    print('\n{} formants testés -> {}'.format(len(lignes), a.sortie))

if __name__ == '__main__':
    main()
