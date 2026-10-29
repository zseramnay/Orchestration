# -*- coding: utf-8 -*-
"""
hors_espace.py — tâche C11 : un critère de « hors espace par le bas » qui tient
d'un corpus à l'autre.

L'ancien critère comptait les maxima : un seul maximum survivant à la fusion de
Chistovich valait « hors espace par le bas ». Il dépend de la brillance de la prise de
son — 17 points d'étendue médiane entre corpus, jusqu'à 50 pour le cor anglais.

Le remplaçant compare **F1 au plancher de l'espace vocalique**, c'est-à-dire au plus bas
F1 des dix voyelles de Calliope, celui de /y/ à 300 Hz. Deux avantages : F1 est la
grandeur la plus stable de l'étude, et il existe même dans les cellules que l'ancien
critère jetait — une cellule « 1 max » a bien un F1, c'est son F2 qui manque.

Le script rend les deux critères côte à côte, leur étendue entre corpus, et le
classement final. Voir `_HORS-ESPACE.md`.
"""
import os, sys, json, glob, math, argparse, statistics as st
from collections import defaultdict

ICI = os.path.dirname(os.path.abspath(__file__))
RACINE = os.path.dirname(os.path.dirname(ICI))
MES = os.path.join(RACINE, 'Etude-Formants', 'Mesures')

# Calliope, voix d'homme — le plancher est le plus bas F1 des dix voyelles, celui de /y/
VOYELLES = {'i': 308, 'y': 300, 'u': 315, 'e': 365, 'ø': 381, 'o': 383,
            'ɛ': 530, 'œ': 517, 'ɔ': 531, 'a': 684}
SEUIL_HORS = 50.0      # part de cellules sous le plancher : au-delà, hors espace
SEUIL_REGISTRE = 20.0  # entre les deux : hors espace dans le grave seulement
SEUIL_SANS_F2 = 40.0   # part de cellules « 1 max »

def bark(f):
    return 13.0 * math.atan(0.00076 * f) + 3.5 * math.atan((f / 7500.0) ** 2)

PLANCHER = bark(min(VOYELLES.values()))

def charger():
    """{instrument: {corpus: [cellules]}} depuis toutes les cartes vocaliques."""
    C = defaultdict(dict)
    for p in glob.glob(os.path.join(MES, '*', 'carte-vocalique-*.json')):
        corpus = os.path.basename(os.path.dirname(p))
        for nom, v in json.load(open(p, encoding='utf-8')).items():
            cells = [x for x in (v if isinstance(v, list) else v.get('cellules', []))
                     if x.get('F1_bark')]
            if cells:
                C[nom][corpus] = cells
    return C

def parts(cells):
    sous = 100.0 * sum(1 for x in cells if x['F1_bark'] < PLANCHER) / len(cells)
    un = 100.0 * sum(1 for x in cells if '1 max' in (x.get('marques') or [])) / len(cells)
    return sous, un

def classer(sous, un, par_dyn):
    if sous >= SEUIL_HORS:
        return 'hors espace par le bas'
    if sous >= SEUIL_REGISTRE:
        return 'hors espace dans le grave seulement'
    if un >= SEUIL_SANS_F2:
        d, f = par_dyn.get('doux'), par_dyn.get('fort')
        if d is not None and f is not None and d - f >= 25:
            return 'sans second maximum au doux seulement'
        return 'dans l\'espace, sans second maximum'
    return 'dans l\'espace'

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--sortie', default=os.path.join(MES, 'Hors-espace'))
    a = ap.parse_args()
    os.makedirs(a.sortie, exist_ok=True)
    C = charger()
    print('plancher de l\'espace vocalique : {:.0f} Hz = {:.2f} Bark (/y/)\n'
          .format(min(VOYELLES.values()), PLANCHER))

    # ── stabilité des deux critères entre corpus ──────────────────────────
    e_un, e_sous = [], []
    for nom, d in C.items():
        if len(d) < 2:
            continue
        v = [parts(c) for c in d.values()]
        e_sous.append(max(x[0] for x in v) - min(x[0] for x in v))
        e_un.append(max(x[1] for x in v) - min(x[1] for x in v))
    # La MÉDIANE vaut zéro des deux côtés — la majorité des instruments sont à 0 %
    # sur les deux critères — et ne distingue donc rien. C'est la moyenne, et le pire
    # cas, qui montrent le gain.
    print('étendue entre corpus, sur {} instruments à plusieurs corpus :'.format(len(e_un)))
    print('{:24s} {:>9s} {:>9s} {:>9s}'.format('', 'moyenne', 'médiane', 'pire cas'))
    print('{:24s} {:8.1f}  {:8.0f}  {:8.0f}'.format(
        '« 1 max »', st.mean(e_un), st.median(e_un), max(e_un)))
    print('{:24s} {:8.1f}  {:8.0f}  {:8.0f}'.format(
        'F1 sous le plancher', st.mean(e_sous), st.median(e_sous), max(e_sous)))

    # ── classement ────────────────────────────────────────────────────────
    L = []
    for nom, d in C.items():
        tout = [x for c in d.values() for x in c]
        sous, un = parts(tout)
        par_dyn = {}
        for dy in ('doux', 'moyen', 'fort'):
            g = [x for x in tout if x.get('dynamique') == dy]
            if g:
                par_dyn[dy] = 100.0 * sum(
                    1 for x in g if '1 max' in (x.get('marques') or [])) / len(g)
        L.append(dict(instrument=nom, n_corpus=len(d), n_cellules=len(tout),
                      pct_sous_plancher=round(sous, 1), pct_un_max=round(un, 1),
                      F1_median_hz=int(st.median([x['F1_hz'] for x in tout])),
                      un_max_doux=par_dyn.get('doux'), un_max_moyen=par_dyn.get('moyen'),
                      un_max_fort=par_dyn.get('fort'),
                      classement=classer(sous, un, par_dyn)))
    L.sort(key=lambda r: -r['pct_sous_plancher'])
    print('\n{:28s} {:>2s} {:>4s} {:>9s} {:>9s} {:>8s}  {}'.format(
        'instrument', 'nc', 'n', 'sous le', '« 1 max »', 'F1 méd.', 'classement'))
    for r in L:
        print('{:28s} {:2d} {:4d} {:8.0f}% {:8.0f}% {:7d}Hz  {}'.format(
            r['instrument'][:28], r['n_corpus'], r['n_cellules'],
            r['pct_sous_plancher'], r['pct_un_max'], r['F1_median_hz'], r['classement']))
    import csv as _csv
    with open(os.path.join(a.sortie, 'hors-espace.csv'), 'w', encoding='utf-8',
              newline='') as f:
        w = _csv.DictWriter(f, fieldnames=list(L[0]))
        w.writeheader(); w.writerows(L)
    print('\n{} instruments -> {}'.format(len(L), a.sortie))

if __name__ == '__main__':
    main()
