# -*- coding: utf-8 -*-
"""
reference_feminine.py — tâche C12 : la carte vocalique change-t-elle de lecture si l'on
prend la référence féminine de Calliope au lieu de la masculine ?

Pourquoi la question
--------------------
Toute la carte vocalique est lue contre les dix voyelles orales du français mesurées par
Calliope (Tubach 1989) en **voix d'homme**. Ce choix n'a jamais été justifié : il a été fait
parce que c'était le tableau sous la main. Or il décide de trois choses, et pas des moindres.

  1. La VOYELLE attribuée à chaque cellule — « le basson sonne /o/ » dépend du tableau.
  2. Le PLANCHER de l'espace vocalique, c'est-à-dire le plus bas F1 des dix voyelles. C'est
     le critère de la tâche C11, qui classe « hors espace par le bas ».
  3. Le PLAFOND, le plus haut F1. Le chapitre en tire que la frontière de l'espace tombe
     vers 700 Hz, « exactement le plafond de /a/ chez l'homme ».

Les valeurs des deux voix sont dans le dépôt depuis le 23 octobre —
`Formants-Francais/donnees/formants_calliope.csv` — et n'avaient jamais été croisées.

Ce qui est mesuré
-----------------
Les 453 cellules des quatre cartes sont relues contre les deux références : voyelle la plus
proche dans le plan (F1, F2) en Bark, distance, statut, et position par rapport au plancher.
On compte ce qui change, et on regarde laquelle des deux décrit le mieux l'orchestre —
par la distance médiane à la voyelle la plus proche, qui est une mesure d'ajustement et non
une préférence.

Usage
    python3 reference_feminine.py
    python3 reference_feminine.py --voix enfant_estime
"""
import os, sys, csv, json, math, glob, argparse
from collections import defaultdict, Counter
import numpy as np

ICI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ICI)
import carte_vocalique as CV

RACINE = os.path.dirname(os.path.dirname(ICI))
ETUDE = os.path.join(RACINE, 'Etude-Formants')
MES = os.path.join(ETUDE, 'Mesures')
CALLIOPE = os.path.join(ETUDE, 'Formants-Francais', 'donnees', 'formants_calliope.csv')

SEUIL_ASSIGNABLE = CV.SEUIL_ASSIGNABLE      # 2,0 Bark
SEUIL_HORS = CV.SEUIL_HORS                  # 3,5 Bark — seuil de Chistovich

def voyelles(voix):
    """[(voyelle, F1_hz, F2_hz)] pour une voix de `formants_calliope.csv`."""
    out = []
    with open(CALLIOPE, encoding='utf-8') as f:
        for r in csv.DictReader(f):
            if r['voix'] == voix:
                out.append((r['voyelle'], float(r['F1_Hz']), float(r['F2_Hz'])))
    return out

def en_bark(V):
    return [(v, float(CV.hz_vers_bark(a)), float(CV.hz_vers_bark(b))) for v, a, b in V]

def proche(z1, z2, VB):
    d = sorted((math.hypot(z1 - a, z2 - b), v) for v, a, b in VB)
    return d[0][1], d[0][0]

def statut(d):
    return ('assignable' if d <= SEUIL_ASSIGNABLE else
            'approximative' if d <= SEUIL_HORS else 'HORS ESPACE')

def cellules():
    out = []
    for p in sorted(glob.glob(os.path.join(MES, '*', 'carte-vocalique-*.json'))):
        corpus = os.path.basename(os.path.dirname(p))
        for inst, bloc in json.load(open(p, encoding='utf-8')).items():
            for c in bloc.get('cellules', []):
                if c.get('F1_bark') is None:
                    continue
                out.append((corpus, inst, c))
    return out

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--voix', default='femme')
    ap.add_argument('--sortie', default=os.path.join(MES, 'Reference-Feminine'))
    a = ap.parse_args()
    os.makedirs(a.sortie, exist_ok=True)
    H, F = voyelles('homme'), voyelles(a.voix)
    HB, FB = en_bark(H), en_bark(F)
    cel = cellules()
    avec_f2 = [x for x in cel if x[2].get('F2_bark') is not None]

    # ── 1. les deux espaces ─────────────────────────────────────────────────
    print('=' * 80)
    print('1. LES DEUX ESPACES — ce que la référence décide avant toute mesure')
    print('=' * 80)
    print('{:9s} {:>10s} {:>10s} {:>12s} {:>12s}'.format(
        'voyelle', 'homme F1', a.voix[:10] + ' F1', 'écart', 'écart F2'))
    for (v, h1, h2), (w, f1, f2) in zip(sorted(H), sorted(F)):
        print('{:9s} {:7.0f} Hz {:7.0f} Hz {:9.2f} dt {:9.2f} dt'.format(
            v, h1, f1, 12 * math.log2(f1 / h1), 12 * math.log2(f2 / h2)))
    bh, bf = min(x[1] for x in H), min(x[1] for x in F)
    th, tf = max(x[1] for x in H), max(x[1] for x in F)
    globals()['HB'], globals()['FB'] = HB, FB
    print()
    print('  PLANCHER (plus bas F1)  homme {:.0f} Hz /{}/   {} {:.0f} Hz /{}/   -> {:+.2f} dt'
          .format(bh, min(H, key=lambda x: x[1])[0], a.voix, bf,
                  min(F, key=lambda x: x[1])[0], 12 * math.log2(bf / bh)))
    print('  PLAFOND  (plus haut F1) homme {:.0f} Hz /{}/   {} {:.0f} Hz /{}/   -> {:+.2f} dt'
          .format(th, max(H, key=lambda x: x[1])[0], a.voix, tf,
                  max(F, key=lambda x: x[1])[0], 12 * math.log2(tf / th)))

    # ── 2. ce qui change cellule par cellule ────────────────────────────────
    print()
    print('=' * 80)
    print('2. LES {} CELLULES AYANT UN COUPLE (F1, F2), RELUES'.format(len(avec_f2)))
    print('=' * 80)
    chg_v = chg_s = 0
    dh, df = [], []
    paires = Counter(); lignes = []
    for corpus, inst, c in avec_f2:
        z1, z2 = c['F1_bark'], c['F2_bark']
        vh, eh = proche(z1, z2, HB)
        vf, ef = proche(z1, z2, FB)
        dh.append(eh); df.append(ef)
        if vh != vf:
            chg_v += 1; paires[(vh, vf)] += 1
        if statut(eh) != statut(ef):
            chg_s += 1
        lignes.append(dict(corpus=corpus, instrument=inst, registre=c['registre'],
                           dynamique=c['dynamique'], n=c['n'],
                           F1_hz=c['F1_hz'], F2_hz=c['F2_hz'],
                           voyelle_homme=vh, distance_homme=round(eh, 2),
                           statut_homme=statut(eh),
                           voyelle_autre=vf, distance_autre=round(ef, 2),
                           statut_autre=statut(ef), change=vh != vf))
    print('  voyelle qui change : {} cellules sur {}  ({:.0f} %)'.format(
        chg_v, len(avec_f2), 100 * chg_v / len(avec_f2)))
    print('  statut qui change  : {} cellules  ({:.0f} %)'.format(
        chg_s, 100 * chg_s / len(avec_f2)))
    print()
    print('  ajustement — distance médiane à la voyelle la plus proche :')
    print('     voix d\'homme : {:.2f} Bark      {} : {:.2f} Bark'.format(
        float(np.median(dh)), a.voix, float(np.median(df))))
    print('     cellules « assignable » (≤ {:.1f} Bark) : {} contre {}'.format(
        SEUIL_ASSIGNABLE, sum(1 for x in dh if x <= SEUIL_ASSIGNABLE),
        sum(1 for x in df if x <= SEUIL_ASSIGNABLE)))
    print('     cellules « HORS ESPACE » (> {:.1f} Bark) : {} contre {}'.format(
        SEUIL_HORS, sum(1 for x in dh if x > SEUIL_HORS),
        sum(1 for x in df if x > SEUIL_HORS)))
    if paires:
        print()
        print('  les reclassements, du plus fréquent au moins :')
        for (vh, vf), n in paires.most_common():
            print('     /{}/ -> /{}/   {} cellules'.format(vh, vf, n))

    # ── 3. le plancher de la tâche C11 ──────────────────────────────────────
    print()
    print('=' * 80)
    print('3. LE PLANCHER — le critère « hors espace par le bas » de la tâche C11')
    print('=' * 80)
    ph, pf = float(CV.hz_vers_bark(bh)), float(CV.hz_vers_bark(bf))
    par = defaultdict(lambda: [0, 0, 0])
    for corpus, inst, c in cel:
        par[inst][0] += 1
        par[inst][1] += 1 if c['F1_bark'] < ph else 0
        par[inst][2] += 1 if c['F1_bark'] < pf else 0
    print('{:28s} {:>5s} {:>12s} {:>12s} {:>8s}'.format(
        'instrument', 'cell', 'sous homme', 'sous ' + a.voix[:6], 'écart'))
    cl = []
    for inst, (n, a1, a2) in sorted(par.items(), key=lambda kv: -kv[1][1] / max(kv[1][0], 1)):
        p1, p2 = 100.0 * a1 / n, 100.0 * a2 / n
        if a1 == 0 and a2 == 0:
            continue
        cl.append(dict(instrument=inst, n=n, pct_homme=round(p1, 1), pct_autre=round(p2, 1)))
        print('{:28s} {:5d} {:11.0f} % {:11.0f} % {:+7.0f} pt'.format(
            inst[:28], n, p1, p2, p2 - p1))
    tot_h = sum(1 for x in cel if x[2]['F1_bark'] < ph)
    tot_f = sum(1 for x in cel if x[2]['F1_bark'] < pf)
    print('\n  ENSEMBLE : {} cellules sous le plancher masculin, {} sous le {} — {:+d}'
          .format(tot_h, tot_f, a.voix, tot_f - tot_h))

    # ── 4. le témoin : « mieux ajusté » veut-il dire quelque chose ? ────────
    print()
    print('=' * 80)
    print("4. TÉMOIN — un espace plus grand ajuste-t-il mieux par construction ?")
    print('=' * 80)
    print("   On dilate l'espace MASCULIN d'un facteur commun sur F1 et F2, et on relit.")
    print("   Si la distance médiane décroît continûment, « mieux ajusté » ne distingue")
    print("   aucune voix : il dit seulement que l'orchestre est au-dessus de l'espace.")
    print()
    Z = [(c['F1_bark'], c['F2_bark']) for _, _, c in avec_f2]
    def ajuste(VB):
        return float(np.median([proche(z1, z2, VB)[1] for z1, z2 in Z]))
    print('{:>10s} {:>14s} {:>12s}'.format('facteur', 'distance méd.', 'assignables'))
    sweep = []
    for k in (1.0, 1.05, 1.1, 1.15, 1.2, 1.3, 1.4, 1.5, 1.6, 1.8, 2.0):
        VB = en_bark([(v, f1 * k, f2 * k) for v, f1, f2 in H])
        d = [proche(z1, z2, VB)[1] for z1, z2 in Z]
        sweep.append(dict(facteur=k, distance_mediane=round(float(np.median(d)), 3),
                          assignables=sum(1 for x in d if x <= SEUIL_ASSIGNABLE)))
        print('{:10.2f} {:11.2f} Bark {:11d}'.format(
            k, sweep[-1]['distance_mediane'], sweep[-1]['assignables']))
    best = min(sweep, key=lambda s: s['distance_mediane'])
    print('\n   minimum à ×{:.2f} ({:.2f} Bark). Les trois voix de Calliope valent, elles :'
          .format(best['facteur'], best['distance_mediane']))
    for voix in ('homme', 'femme', 'enfant_estime'):
        V = voyelles(voix)
        if not V:
            continue
        print('      {:14s} {:.2f} Bark   (F1 moyen {:.0f} Hz)'.format(
            voix, ajuste(en_bark(V)), float(np.mean([x[1] for x in V]))))

    # ── 5. le plafond, et la voyelle dominante par instrument ──────────────
    print()
    print('=' * 80)
    print("5. LE PLAFOND ET LES NOMS DE VOYELLES")
    print('=' * 80)
    haut_h = [x for x in cel if x[2]['F1_hz'] > th]
    rentrent = [x for x in haut_h if x[2]['F1_hz'] <= tf]
    print('  plafond {:.0f} Hz (homme, /a/) -> {:.0f} Hz ({}, /a/)'.format(th, tf, a.voix))
    print('  {} cellules sont au-dessus du plafond masculin, dont {} repassent dessous :'
          .format(len(haut_h), len(rentrent)))
    print('     {}'.format(', '.join(sorted({x[1] for x in rentrent}))))
    print()
    dom = defaultdict(lambda: (Counter(), Counter()))
    for corpus, inst, c in avec_f2:
        z1, z2 = c['F1_bark'], c['F2_bark']
        dom[inst][0][proche(z1, z2, HB)[0]] += 1
        dom[inst][1][proche(z1, z2, FB)[0]] += 1
    print('{:28s} {:>5s}  {:>8s} {:>8s}'.format('instrument', 'cell', 'homme', a.voix[:8]))
    chg = []
    for inst, (ch, cf) in sorted(dom.items()):
        vh, vf = ch.most_common(1)[0][0], cf.most_common(1)[0][0]
        if vh != vf:
            chg.append(dict(instrument=inst, voyelle_homme=vh, voyelle_autre=vf,
                            n=sum(ch.values())))
        print('{:28s} {:5d}  {:>8s} {:>8s}{}'.format(
            inst[:28], sum(ch.values()), '/%s/' % vh, '/%s/' % vf,
            '   <-- change' if vh != vf else ''))
    print('\n  {} instruments sur {} changent de voyelle dominante.'
          .format(len(chg), len(dom)))

    for f, L in (('cellules.csv', lignes), ('plancher.csv', cl), ('dilatation.csv', sweep),
                 ('voyelle-dominante.csv', chg)):
        if L:
            with open(os.path.join(a.sortie, f), 'w', encoding='utf-8', newline='') as fh:
                w = csv.DictWriter(fh, fieldnames=list(L[0])); w.writeheader(); w.writerows(L)
    print('\n-> {}'.format(a.sortie))

if __name__ == '__main__':
    main()
