# -*- coding: utf-8 -*-
"""
verification_vocalique.py — tâche C13 : la lecture vocalique réplique-t-elle d'un corpus à
l'autre, maintenant que les 22 instruments en ont au moins deux ?

Ce que la tâche a d'abord trouvé
--------------------------------
Trois instruments étaient réputés sans second corpus : clarinette basse, flûte contrebasse,
clarinette contrebasse. **Les spectres existaient pourtant**, et ZSamples — vingt
instruments — n'avait aucune carte du tout. La cause n'était pas un manque de données mais
un **échec d'appariement de noms** : les corpus écrivent « Flute basse », « Trombone tenor »,
« Clarinette basse en Sib » là où le script attendait l'orthographe exacte des fiches. Il
imprimait « absent » et passait. Trente-cinq cartes mesurables ont ainsi été écartées en
silence. Voir `carte_vocalique.resoudre_fichier`.

Ce qui est mesuré ici
---------------------
L'unité est la cellule : un registre à une nuance. Pour chaque cellule présente dans au
moins deux corpus de **provenances distinctes**, on compare

    F1   écart maximal entre corpus, en demi-tons
    F2   idem
    voyelle   les corpus nomment-ils la même ?

La règle de provenance de `_CORPUS-CROISES.md` s'applique : la paire (SOL, VN) ne compte pas
quand la source SOL de l'instrument est `Yan_Adds-Divers`, qui est du VSL comme VN.

Usage
    python3 verification_vocalique.py
    python3 verification_vocalique.py --seuil 1.0
"""
import os, sys, csv, json, math, glob, argparse, itertools
from collections import defaultdict, Counter
import numpy as np

ICI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ICI)
import carte_vocalique as CV

RACINE = os.path.dirname(os.path.dirname(ICI))
ETUDE = os.path.join(RACINE, 'Etude-Formants')
MES = os.path.join(ETUDE, 'Mesures')

# Les dix instruments dont le côté « SOL » vient de `Yan_Adds-Divers`, donc du VSL.
# Pour eux, la paire (SOL, VN) ne compte pas. Source : `_CORPUS-CROISES.md`, 25/10/2026.
VSL = {'Petite flûte', 'Flûte basse', 'Flûte contrebasse', 'Cor anglais', 'Contrebasson',
       'Clarinette en Mib', 'Clarinette basse', 'Clarinette contrebasse',
       'Trombone basse', 'Tuba contrebasse'}
ORDRE = ['SOL', 'ZSamples', 'Xsamples', 'VO', 'VN']

def paire_valide(inst, a, b):
    return not ({a, b} == {'SOL', 'VN'} and inst in VSL)

def charger():
    """{instrument: {corpus: {(registre, dynamique): cellule}}}"""
    C = defaultdict(dict)
    for p in sorted(glob.glob(os.path.join(MES, '*', 'carte-vocalique-*.json'))):
        corpus = os.path.basename(os.path.dirname(p))
        for inst, bloc in json.load(open(p, encoding='utf-8')).items():
            d = {(c['registre'], c['dynamique']): c for c in bloc.get('cellules', [])
                 if c.get('F1_bark') is not None}
            if d:
                C[inst][corpus] = d
    return C

def dt(a, b):
    return abs(12.0 * math.log2(a / b))

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--seuil', type=float, default=0.5)
    ap.add_argument('--sortie', default=os.path.join(MES, 'Verification-Vocalique'))
    a = ap.parse_args()
    os.makedirs(a.sortie, exist_ok=True)
    C = charger()

    print('=' * 92)
    print('COUVERTURE')
    print('=' * 92)
    sans = [i for i, d in C.items()
            if not any(paire_valide(i, x, y) for x, y in itertools.combinations(d, 2))]
    print('  {} instruments cartographiés ; {} ont au moins deux corpus ; '
          '{} au moins deux de provenances distinctes'.format(
              len(C), sum(1 for d in C.values() if len(d) >= 2), len(C) - len(sans)))
    if sans:
        print('  sans paire valide : {}'.format(', '.join(sorted(sans))))

    print()
    print('=' * 92)
    print("ACCORD CELLULE PAR CELLULE — écart maximal entre corpus de provenances distinctes")
    print('=' * 92)
    print('{:24s} {:>5s} {:>7s} {:>7s} {:>8s} {:>8s} {:>9s}'.format(
        'instrument', 'cell', 'F1 méd.', 'F1 max', 'F2 méd.', '≤ seuil', 'voyelle'))
    lignes, detail = [], []
    for inst in sorted(C):
        D = C[inst]
        e1, e2, acc_v, n_v = [], [], 0, 0
        for cle in set().union(*[set(d) for d in D.values()]):
            pres = [c for c in ORDRE if c in D and cle in D[c]]
            pairs = [(x, y) for x, y in itertools.combinations(pres, 2)
                     if paire_valide(inst, x, y)]
            if not pairs:
                continue
            f1 = max(dt(D[x][cle]['F1_hz'], D[y][cle]['F1_hz']) for x, y in pairs)
            e1.append(f1)
            f2s = [(x, y) for x, y in pairs
                   if D[x][cle].get('F2_hz') and D[y][cle].get('F2_hz')]
            f2 = (max(dt(D[x][cle]['F2_hz'], D[y][cle]['F2_hz']) for x, y in f2s)
                  if f2s else None)
            if f2 is not None:
                e2.append(f2)
            vs = [D[c][cle].get('voyelle') for c in pres if D[c][cle].get('voyelle')]
            if len(vs) >= 2:
                n_v += 1
                acc_v += 1 if len(set(vs)) == 1 else 0
            detail.append(dict(instrument=inst, registre=cle[0], dynamique=cle[1],
                               corpus=' '.join(pres), ecart_F1=round(f1, 2),
                               ecart_F2=None if f2 is None else round(f2, 2),
                               voyelles=' '.join(vs), accord_voyelle=len(set(vs)) == 1
                               if len(vs) >= 2 else None))
        if not e1:
            continue
        ok = 100.0 * sum(1 for x in e1 if x <= a.seuil) / len(e1)
        d = dict(instrument=inst, n_cellules=len(e1),
                 F1_median=round(float(np.median(e1)), 2), F1_max=round(max(e1), 2),
                 F2_median=round(float(np.median(e2)), 2) if e2 else None,
                 pct_sous_seuil=round(ok, 0),
                 accord_voyelle=(round(100.0 * acc_v / n_v) if n_v else None), n_voyelle=n_v)
        lignes.append(d)
        print('{:24s} {:5d} {:6.2f}  {:6.2f}  {:>8} {:7.0f} % {:>8}'.format(
            inst[:24], len(e1), d['F1_median'], d['F1_max'],
            d['F2_median'] if d['F2_median'] is not None else '—', ok,
            ('%d %%' % d['accord_voyelle']) if d['accord_voyelle'] is not None else '—'))

    e1 = [d['F1_median'] for d in lignes]
    print('\n  ENSEMBLE : {} instruments, écart médian de F1 entre corpus = {:.2f} demi-ton ; '
          '{} instruments sous {:.1f} dt'.format(
              len(lignes), float(np.median(e1)), sum(1 for x in e1 if x <= a.seuil), a.seuil))
    v = [d['accord_voyelle'] for d in lignes if d['accord_voyelle'] is not None]
    if v:
        print('  accord de VOYELLE entre corpus : {:.0f} % en médiane'.format(float(np.median(v))))

    for f, L in (('par-instrument.csv', lignes), ('par-cellule.csv', detail)):
        with open(os.path.join(a.sortie, f), 'w', encoding='utf-8', newline='') as fh:
            w = csv.DictWriter(fh, fieldnames=list(L[0])); w.writeheader(); w.writerows(L)
    print('\n-> {}'.format(a.sortie))

if __name__ == '__main__':
    main()
