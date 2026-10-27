# -*- coding: utf-8 -*-
"""
impact_regle_f0.py — tâche C3 : impact global de la règle « F1 ≥ f0 » sur la carte vocalique.

La règle est proposée pour corriger l'instabilité du violon au suraigu. Elle est globale, donc
avant de l'adopter il faut savoir ce qu'elle change ailleurs : combien de cellules changent de
voyelle, et dans quel sens va la dispersion.

Le script recalcule la carte cellule par cellule, avec et sans la règle, et compare. Il n'écrit
aucune carte : il ne produit qu'un rapport.

Usage
    python3 impact_regle_f0.py --corpus SOL
    python3 impact_regle_f0.py --corpus ZSamples,Xsamples,VO,VN
"""
import os, sys, math, json, glob, argparse
from collections import defaultdict
import numpy as np

ICI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ICI)
import formants_lib as FL
import carte_vocalique as CV
from fenetre_redressement import NOMS as NOMS_SOL

RACINE = os.path.dirname(os.path.dirname(ICI))
ETUDE = os.path.join(RACINE, 'Etude-Formants')
SOL_OUT = os.path.expanduser('~/work/out')
TOL = 2 ** (-1.0 / 12)

def cellules(chemin, regs, borne):
    """{(registre, dynamique): (F1_hz, F2_hz, IQR1, voyelle, statut, n)}"""
    acc = defaultdict(list)
    for f0, midi, dyn, sp in FL.lire_spectres(chemin):
        if not f0:
            continue
        m = CV.maxima_fusionnes(CV.motif_excitation(sp))
        hz = [CV.bark_vers_hz(z) for z, lv in m]
        if borne:
            hz = [h for h in hz if h >= f0 * TOL]
        if not hz:
            continue
        z1 = CV.hz_vers_bark(hz[0])
        z2 = CV.hz_vers_bark(hz[1]) if len(hz) > 1 else None
        dy = CV.DYN.get((dyn or '').lower(), 'autre')
        acc[(CV.registre_de(regs, midi), dy)].append((z1, z2, hz[0]))
    out = {}
    for k, v in acc.items():
        if len(v) < CV.MIN_N_CELLULE:
            continue
        z1 = float(np.median([x[0] for x in v]))
        z2v = [x[1] for x in v if x[1] is not None]
        z2 = float(np.median(z2v)) if len(z2v) >= len(v) / 2 else None
        iq1 = float(np.percentile([x[0] for x in v], 75)
                    - np.percentile([x[0] for x in v], 25))
        voy, dist = CV.voyelle_proche(z1, z2)
        out[k] = dict(F1=int(round(CV.bark_vers_hz(z1))),
                      F2=None if z2 is None else int(round(CV.bark_vers_hz(z2))),
                      IQR1=round(iq1, 2), voyelle=voy,
                      statut=CV.statut_distance(dist), n=len(v))
    return out

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--corpus', default='SOL')
    ap.add_argument('--sortie', default=os.path.join(ETUDE, 'Mesures', 'Suraigu'))
    a = ap.parse_args()
    os.makedirs(a.sortie, exist_ok=True)
    regs_tous = CV.lire_registres(os.path.join(RACINE, 'Data', 'registres.md'))

    rapport = []
    for corpus in [x.strip() for x in a.corpus.split(',') if x.strip()]:
        if corpus == 'SOL':
            src = [(n, os.path.join(SOL_OUT, f + '.txt')) for n, f in NOMS_SOL.items()]
        else:
            src = [(os.path.basename(p)[:-len('.spectres.txt')], p) for p in
                   sorted(glob.glob(os.path.join(ETUDE, 'Mesures', corpus,
                                                 '*.spectres.txt')))]
        for nom, p in src:
            if not os.path.exists(p):
                continue
            cle = None
            for k in CV.INSTRUMENTS:
                if k.lower().startswith(nom.lower()[:6]) or nom.lower().startswith(k.lower()[:6]):
                    cle = CV.INSTRUMENTS[k][1]; break
            if cle is None:
                continue
            try:
                regs = next(regs_tous[k] for k in regs_tous
                            if k == cle or k.startswith(cle + ' ') or cle in k.split(' et '))
            except StopIteration:
                continue
            av = cellules(p, regs, False)
            ap_ = cellules(p, regs, True)
            for k in sorted(set(av) | set(ap_)):
                A, B = av.get(k), ap_.get(k)
                if not A or not B:
                    rapport.append(dict(corpus=corpus, instrument=nom, registre=k[0],
                                        dynamique=k[1], cas='cellule apparue/disparue'))
                    continue
                dt = 12 * math.log2(B['F1'] / A['F1'])
                rapport.append(dict(corpus=corpus, instrument=nom, registre=k[0],
                                    dynamique=k[1], n=A['n'],
                                    F1_avant=A['F1'], F1_apres=B['F1'], ecart_dt=round(dt, 1),
                                    IQR_avant=A['IQR1'], IQR_apres=B['IQR1'],
                                    voy_avant=A['voyelle'], voy_apres=B['voyelle'],
                                    statut_avant=A['statut'], statut_apres=B['statut']))

    chg = [r for r in rapport if r.get('ecart_dt') and abs(r['ecart_dt']) > 1.0]
    voy = [r for r in rapport if r.get('voy_avant') != r.get('voy_apres') and 'cas' not in r]
    iqr = [r for r in rapport if r.get('IQR_avant') is not None]
    mieux = [r for r in iqr if r['IQR_apres'] < r['IQR_avant'] - 0.05]
    pire = [r for r in iqr if r['IQR_apres'] > r['IQR_avant'] + 0.05]
    print('{} cellules comparées'.format(len(rapport)))
    print('  F1 déplacé de plus d\'un demi-ton : {} ({:.0f} %)'.format(
        len(chg), 100 * len(chg) / max(len(rapport), 1)))
    print('  voyelle changée                   : {} ({:.0f} %)'.format(
        len(voy), 100 * len(voy) / max(len(rapport), 1)))
    print('  dispersion améliorée              : {}'.format(len(mieux)))
    print('  dispersion dégradée               : {}'.format(len(pire)))
    print('\npar registre — cellules dont F1 bouge de plus d\'un demi-ton :')
    par = defaultdict(lambda: [0, 0])
    for r in rapport:
        if 'ecart_dt' not in r:
            continue
        par[r['registre']][1] += 1
        if abs(r['ecart_dt']) > 1.0:
            par[r['registre']][0] += 1
    for reg, (c, t) in sorted(par.items(), key=lambda kv: -kv[1][0]):
        print('   {:14s} {:3d}/{:<3d}  {:3.0f} %'.format(reg, c, t, 100 * c / max(t, 1)))
    with open(os.path.join(a.sortie, 'impact-regle-f0.json'), 'w', encoding='utf-8') as f:
        json.dump(rapport, f, ensure_ascii=False, indent=1)

if __name__ == '__main__':
    main()
