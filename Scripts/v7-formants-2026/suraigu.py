# -*- coding: utf-8 -*-
"""
suraigu.py — tâche C3 : l'instabilité du violon au suraigu.

Le défaut
---------
`_CARTE-VOCALIQUE.md` note, comme seul défaut de méthode reproductible identifié : « Le violon
est instable au suraigu dans les deux corpus. SOL donne F1 = 496 Hz au doux et 1 478 Hz au
moyen ; ZSamples donne 977 Hz avec un IQR de 8,6 Bark, soit un tiers de l'échelle. »

Le diagnostic avancé était : « un basculement du détecteur entre fondamentale et résonance de
caisse ». Il est juste, et on peut en tirer une règle.

La règle
--------
`f1_f2()` prend pour F1 **le maximum survivant le plus bas**, sans borne inférieure. Or un
spectre harmonique **n'a aucun partiel en dessous de f0**. Ce qui se trouve là est du bruit
d'archet, la réponse de la caisse à une excitation large bande, ou la jupe de la fonction
d'étalement — jamais un formant au sens vocalique, puisque la lecture (F1, F2) suppose une
source harmonique filtrée par des résonances.

Tant que f0 reste sous la zone de résonance, la question ne se pose pas : à l'alto grave
(f0 131–196 Hz) la caisse est à 375–395 Hz, donc au-dessus. Mais au suraigu du violon
(f0 ≥ 1 175 Hz) la fondamentale a dépassé la caisse (~530 Hz) : le détecteur a alors le choix
entre deux maxima sans rapport, et bascule de l'un à l'autre selon la dynamique.

**Correctif proposé : chercher F1 parmi les maxima situés à f0 ou au-dessus**, à un demi-ton de
tolérance près. Le script mesure ce que la règle change, registre par registre et corpus par
corpus, avant de toucher à `carte_vocalique.py`.

Ce qu'il faut vérifier avant d'adopter la règle
-----------------------------------------------
1. qu'elle réduise bien la dispersion du violon au suraigu ;
2. qu'elle **ne déplace pas** les résultats acquis — la caisse de l'alto au grave (374–395 Hz,
   quatre corpus) et celle du violoncelle au médium (209–223 Hz, quatre corpus) ;
3. ce qu'elle fait du reste de la carte : combien de cellules changent de voyelle.

Usage
    python3 suraigu.py --instruments Violon,Alto,Violoncelle --registres tous
"""
import os, sys, math, json, argparse
from collections import defaultdict
import numpy as np

ICI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ICI)
import formants_lib as FL
import carte_vocalique as CV
from resonance_caisse import CORDES, chemins, resoudre_registres
from fenetre_redressement import NOMS as NOMS_SOL

RACINE = os.path.dirname(os.path.dirname(ICI))
ETUDE = os.path.join(RACINE, 'Etude-Formants')
TOL = 2 ** (-1.0 / 12)          # un demi-ton de tolérance sous f0

def maxima(spectre):
    return CV.maxima_fusionnes(CV.motif_excitation(spectre))

def f1f2(m, f0, borne=False):
    """(F1, F2) en Hz. borne=True : on ne retient que les maxima à f0 ou au-dessus,
    à un demi-ton près."""
    hz = [CV.bark_vers_hz(z) for z, lv in m]
    if borne:
        hz = [h for h in hz if h >= f0 * TOL]
    if not hz:
        return None, None
    return hz[0], (hz[1] if len(hz) > 1 else None)

def analyser(nom, corpus, chemin, regs):
    par = defaultdict(lambda: {'avant': [], 'apres': [], 'sous_f0': 0, 'vide': 0, 'n': 0})
    for f0, midi, dyn, sp in FL.lire_spectres(chemin):
        if not f0:
            continue
        r = CV.registre_de(regs, midi)
        m = maxima(sp)
        if not m:
            continue
        a1, a2 = f1f2(m, f0, False)
        b1, b2 = f1f2(m, f0, True)
        d = par[r]
        d['n'] += 1
        if a1:
            d['avant'].append(a1)
        if a1 and a1 < f0 * TOL:
            d['sous_f0'] += 1
        if b1:
            d['apres'].append(b1)
        else:
            d['vide'] += 1
    out = []
    for r, d in par.items():
        def stat(v):
            if len(v) < 3:
                return None, None
            z = CV.hz_vers_bark(np.asarray(v))
            return (int(round(float(np.median(v)))),
                    round(float(np.percentile(z, 75) - np.percentile(z, 25)), 2))
        ma, ia = stat(d['avant']); mb, ib = stat(d['apres'])
        out.append({'instrument': nom, 'corpus': corpus, 'registre': r, 'n': d['n'],
                    'F1_avant_hz': ma, 'IQR_avant_bark': ia,
                    'F1_apres_hz': mb, 'IQR_apres_bark': ib,
                    'pct_sous_f0': int(round(100 * d['sous_f0'] / max(d['n'], 1))),
                    'pct_sans_maximum': int(round(100 * d['vide'] / max(d['n'], 1)))})
    ordre = ['sons pédale', 'extrême grave', 'grave', 'chalumeau', 'bas médium', 'gorge',
             'médium', 'clarine', 'haut médium', 'aigu', 'suraigu']
    out.sort(key=lambda d: ordre.index(d['registre']) if d['registre'] in ordre else 99)
    return out

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--instruments', default='Violon,Alto,Violoncelle,Contrebasse')
    ap.add_argument('--registres', default='tous')
    ap.add_argument('--sortie', default=os.path.join(ETUDE, 'Mesures', 'Suraigu'))
    a = ap.parse_args()
    voulus = (set() if a.registres == 'tous'
              else {x.strip() for x in a.registres.split(',') if x.strip()})
    os.makedirs(a.sortie, exist_ok=True)
    regs_tous = CV.lire_registres(os.path.join(RACINE, 'Data', 'registres.md'))

    tout = []
    for nom in [x.strip() for x in a.instruments.split(',') if x.strip()]:
        regs = resoudre_registres(regs_tous, CORDES[nom][0] if nom in CORDES
                                  else nom.lower())
        print('\n' + '=' * 92)
        print(nom)
        print('=' * 92)
        print('{:9s} {:9s} {:>4s} {:>8s} {:>8s} {:>8s} {:>8s} {:>8s} {:>8s}'.format(
            'corpus', 'registre', 'n', 'F1 avant', 'IQR av.', 'F1 après', 'IQR ap.',
            '<f0', 'sans max'))
        for corpus, p in chemins(nom).items():
            for d in analyser(nom, corpus, p, regs):
                if voulus and d['registre'] not in voulus:
                    continue
                tout.append(d)
                print('{:9s} {:9s} {:4d} {:>8} {:>8} {:>8} {:>8} {:7d}% {:7d}%'.format(
                    d['corpus'], d['registre'], d['n'],
                    d['F1_avant_hz'] or '—', d['IQR_avant_bark'] or '—',
                    d['F1_apres_hz'] or '—', d['IQR_apres_bark'] or '—',
                    d['pct_sous_f0'], d['pct_sans_maximum']))
    with open(os.path.join(a.sortie, 'suraigu.json'), 'w', encoding='utf-8') as f:
        json.dump(tout, f, ensure_ascii=False, indent=1)
    print('\n{} lignes -> {}'.format(len(tout), a.sortie))

if __name__ == '__main__':
    main()
