# -*- coding: utf-8 -*-
"""
melanges.py — tâche C9 : fusion contre renforcement, par la méthode 12.

C'est la question d'orchestration qui a lancé la reprise : **quand deux instruments jouent
ensemble, le mélange produit-il une couleur nouvelle, ou l'un ne fait-il que renforcer
l'autre ?**

Ce qu'on peut mesurer, et ce qu'on ne peut pas
----------------------------------------------
Nous n'avons pas d'enregistrements de mélanges d'instruments différents. Mais nous n'en avons
pas besoin pour ce calcul-là, et ce n'est pas une simulation : **le motif d'excitation auditive
est linéaire en puissance.** Il vaut `ETALEMENT @ spectre²`, et deux sources jouant ensemble
s'ajoutent en puissance — leur phase relative varie d'une note à l'autre et au cours du son.
Donc

    motif(A + B) = motif(A) + motif(B)     exactement, en échelle linéaire

Le mélange se calcule donc par une addition, sans relire ni recombiner le moindre échantillon.
C'est aussi ce qui rend la matrice complète abordable : 231 paires sans un seul produit
matriciel de plus.

**Ce que le calcul ignore**, et qu'il faut écrire avec les résultats : la cohérence de phase à
l'unisson et les battements, la séparation spatiale des pupitres, et le fait que des musiciens
qui jouent ensemble s'ajustent — justesse, nuance, vibrato. Le résultat porte sur le mélange
**spectral**, qui est une condition nécessaire de la fusion, pas sa totalité.

L'équilibre
-----------
Les niveaux d'enregistrement des banques sont arbitraires : sommer les spectres bruts
laisserait l'échantillon le plus fort dominer pour une raison qui n'a rien de musical. On
normalise donc chaque parent à énergie égale avant de sommer, et on mesure la sensibilité à ce
choix en refaisant à 1:2 et 2:1.

Le critère
----------
La fusion de Chistovich réunit les maxima distants de moins de 3,5 Bark. C'est le mécanisme même
de la fusion timbrale. Pour une paire à une hauteur donnée, on compare les maxima de A, de B et
du mélange, puis on lit :

    dA = distance en Bark entre le (F1, F2) du mélange et celui de A
    dB = idem pour B
    Δn = nombre de maxima du mélange − nombre de maxima du parent le plus riche

    FUSION         dA ≥ 1 et dB ≥ 1 et Δn ≤ 0
                   la couleur du mélange n'appartient à aucun des deux, sans maximum de plus
    RENFORCEMENT   min(dA, dB) < 1 et max(dA, dB) ≥ 1
                   le mélange se lit comme l'un des deux ; l'autre apporte de l'énergie
    JUXTAPOSITION  Δn > 0
                   le mélange porte plus de maxima distincts que son parent le plus riche :
                   on entend les deux
    INDISTINCT     le reste — les deux parents ont déjà presque la même couleur

Usage
    python3 melanges.py --paires 'Hautbois+Cor en Fa,Flûte+Violon'
    python3 melanges.py --tout
"""
import os, sys, math, json, csv, argparse, itertools
from collections import defaultdict
import numpy as np

ICI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ICI)
import formants_lib as FL
import carte_vocalique as CV
from fenetre_redressement import NOMS
from largeur_par_registre import FAMILLE

RACINE = os.path.dirname(os.path.dirname(ICI))
ETUDE = os.path.join(RACINE, 'Etude-Formants')
SOL_OUT = os.path.expanduser('~/work/out')

DYN = {'ppp': 'doux', 'pp': 'doux', 'p': 'doux', 'mp': 'doux',
       'mf': 'moyen', 'f': 'fort', 'ff': 'fort', 'fff': 'fort'}
SEUIL_BARK = 1.0          # distance au-delà de laquelle la couleur « n'est plus celle de »
MIN_HAUTEURS = 8

def motifs(nom):
    """{(midi, dyn): motif d'excitation LINÉAIRE}, moyenné sur les échantillons.
    Linéaire et non en dB, parce que c'est en linéaire que les mélanges s'additionnent."""
    p = os.path.join(SOL_OUT, (NOMS.get(nom) or '') + '.txt')
    if not os.path.exists(p):
        return None
    acc = defaultdict(list)
    for f0, midi, dyn, sp in FL.lire_spectres(p):
        d = DYN.get((dyn or '').lower())
        if midi is None or d is None:
            continue
        pw = np.asarray(sp, dtype=float) ** 2
        e = CV._SPREAD @ pw
        s = e.sum()
        if s > 0:
            acc[(midi, d)].append(e / s)      # énergie unité : l'équilibre est explicite
    return {k: np.mean(v, axis=0) for k, v in acc.items()}

def lire(e):
    """(n_maxima, z1, z2) d'un motif linéaire."""
    m = CV.maxima_fusionnes(10.0 * np.log10(np.maximum(e, 1e-30)))
    if not m:
        return 0, None, None
    return len(m), m[0][0], (m[1][0] if len(m) > 1 else None)

def distance(a, b):
    """Distance en Bark dans le plan (F1, F2). F2 manquant des deux côtés : on
    compare sur F1 seul plutôt que de jeter le cas."""
    (n1, z1, z2), (m1, y1, y2) = a, b
    if z1 is None or y1 is None:
        return None
    if z2 is None or y2 is None:
        return abs(z1 - y1)
    return math.hypot(z1 - y1, z2 - y2)

def chimere(la, lb, lm, seuil=0.5):
    """Vrai si le mélange prend son F1 d'un parent et son F2 de l'autre.

    Ce n'est PAS une couleur nouvelle née de la fusion de deux maxima : c'est un
    assemblage, chaque parent gardant son formant. La distinction compte, parce que
    le critère de distance seul la manque — un point à plus d'un Bark des deux
    parents peut n'être qu'un F1 emprunté à l'un et un F2 à l'autre."""
    _, a1, a2 = la; _, b1, b2 = lb; _, m1, m2 = lm
    if None in (a1, b1, m1, a2, b2, m2):
        return False
    return ((abs(m1 - a1) < seuil and abs(m2 - b2) < seuil) or
            (abs(m1 - b1) < seuil and abs(m2 - a2) < seuil))

def classer(dA, dB, dn, dAB, chim=False):
    """dAB = distance entre les deux parents eux-mêmes. Sans elle, deux instruments
    qui ont DÉJÀ la même couleur tombaient dans « indistinct » au même titre qu'un cas
    sans mesure — alors que c'est au contraire la forme la plus forte du fondu, par
    identité et non par émergence. Les deux doivent être nommées séparément."""
    if dA is None or dB is None:
        return 'indistinct'
    if dn > 0:
        return 'juxtaposition'
    if dAB is not None and dAB < SEUIL_BARK:
        return 'homogène'
    if dA >= SEUIL_BARK and dB >= SEUIL_BARK:
        return 'chimère' if chim else 'fusion'
    if min(dA, dB) < SEUIL_BARK <= max(dA, dB):
        return 'renforcement'
    return 'indistinct'

def paire(MA, MB, poids=(0.5, 0.5), decalage=0):
    """decalage : nombre de demi-tons dont B est transposé par rapport à A.
    0 = unisson. 12 = B double A à l'octave supérieure.

    À l'unisson la paire est symétrique et un seul calcul suffit. **Dès que le
    décalage n'est pas nul elle cesse de l'être** : flûte au-dessus du violon n'est
    pas flûte en dessous, et les deux sens doivent être calculés séparément. C'est
    pourquoi le pilote énumère des permutations et non des combinaisons quand
    `--octave` est donné."""
    communes = sorted(k for k in MA if (k[0] + decalage, k[1]) in MB)
    cas = []
    for k in communes:
        a, b = MA[k], MB[(k[0] + decalage, k[1])]
        mix = poids[0] * a + poids[1] * b
        la, lb, lm = lire(a), lire(b), lire(mix)
        dA, dB = distance(lm, la), distance(lm, lb)
        dAB = distance(la, lb)
        dn = lm[0] - max(la[0], lb[0])
        cas.append(dict(midi=k[0], midi_B=k[0] + decalage, dyn=k[1],
                        n_A=la[0], n_B=lb[0], n_mix=lm[0],
                        dA=None if dA is None else round(dA, 2),
                        dB=None if dB is None else round(dB, 2),
                        dAB=None if dAB is None else round(dAB, 2),
                        delta_n=dn, chimere=int(chimere(la, lb, lm)),
                        classe=classer(dA, dB, dn, dAB, chimere(la, lb, lm)),
                        F1_mix=None if lm[1] is None else int(round(CV.bark_vers_hz(lm[1]))),
                        vers=None if dA is None or dB is None
                        else ('A' if dA < dB else 'B')))
    return cas

def resumer(nom_a, nom_b, cas):
    n = len(cas)
    c = defaultdict(int)
    for x in cas:
        c[x['classe']] += 1
    dom = defaultdict(int)
    for x in cas:
        if x['classe'] == 'renforcement' and x['vers']:
            dom[x['vers']] += 1
    dA = [x['dA'] for x in cas if x['dA'] is not None]
    dB = [x['dB'] for x in cas if x['dB'] is not None]
    dAB = [x['dAB'] for x in cas if x.get('dAB') is not None]
    pr = max(c, key=lambda k: c[k]) if c else '—'
    return dict(A=nom_a, B=nom_b, famille_A=FAMILLE.get(nom_a, '?'),
                famille_B=FAMILLE.get(nom_b, '?'), n_hauteurs=n,
                pct_fusion=int(round(100 * c['fusion'] / n)) if n else None,
                pct_renforcement=int(round(100 * c['renforcement'] / n)) if n else None,
                pct_juxtaposition=int(round(100 * c['juxtaposition'] / n)) if n else None,
                pct_homogene=int(round(100 * c['homogène'] / n)) if n else None,
                pct_chimere=int(round(100 * c['chimère'] / n)) if n else None,
                pct_indistinct=int(round(100 * c['indistinct'] / n)) if n else None,
                ecart_parents_bark=round(float(np.median(dAB)), 2) if dAB else None,
                dominant=(nom_a if dom['A'] > dom['B'] else nom_b) if sum(dom.values()) else None,
                dA_median=round(float(np.median(dA)), 2) if dA else None,
                dB_median=round(float(np.median(dB)), 2) if dB else None,
                classe_dominante=pr)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--paires', default='')
    ap.add_argument('--tout', action='store_true')
    ap.add_argument('--instruments', default='')
    ap.add_argument('--sortie', default=os.path.join(ETUDE, 'Mesures', 'Melanges'))
    ap.add_argument('--recalculer', action='store_true')
    ap.add_argument('--octave', type=int, default=0, metavar='DEMI-TONS',
                    help='transposition de B par rapport à A ; 12 = doublage à '
                         "l'octave supérieure. Rend la paire ORIENTÉE : les deux sens "
                         'sont alors calculés séparément.')
    a = ap.parse_args()
    os.makedirs(a.sortie, exist_ok=True)

    noms = ([x.strip() for x in a.instruments.split(',') if x.strip()]
            or sorted(NOMS))
    cache = os.path.join(a.sortie, 'motifs.npz')
    M = {}
    if os.path.exists(cache) and not a.recalculer:
        z = np.load(cache, allow_pickle=True)
        for k in z.files:
            nom, midi, dy = k.split('\x1f')
            M.setdefault(nom, {})[(int(midi), dy)] = z[k]
        print('{} instruments relus du cache'.format(len(M)), file=sys.stderr)
    else:
        for n in noms:
            m = motifs(n)
            if m:
                M[n] = m
            print('  {} chargé'.format(n), file=sys.stderr)
        np.savez_compressed(cache, **{'\x1f'.join((n, str(k[0]), k[1])): v
                                      for n, d in M.items() for k, v in d.items()})
        print('{} instruments chargés et mis en cache'.format(len(M)), file=sys.stderr)

    if a.paires:
        couples = []
        for p in a.paires.split(','):
            x, y = p.split('+')
            couples.append((x.strip(), y.strip()))
    elif a.octave:
        couples = list(itertools.permutations(sorted(M), 2))
    else:
        couples = list(itertools.combinations(sorted(M), 2))

    res, detail = [], []
    for x, y in couples:
        if x not in M or y not in M:
            continue
        cas = paire(M[x], M[y], decalage=a.octave)
        if len(cas) < MIN_HAUTEURS:
            continue
        r = resumer(x, y, cas)
        r['decalage_dt'] = a.octave
        # sensibilité à l'équilibre
        for w, lib in (((0.33, 0.67), '1_2'), ((0.67, 0.33), '2_1')):
            c2 = paire(M[x], M[y], w, a.octave)
            nf = sum(1 for z in c2 if z['classe'] == 'fusion')
            r['pct_fusion_' + lib] = int(round(100 * nf / len(c2))) if c2 else None
        res.append(r)
        for z in cas:
            detail.append(dict(A=x, B=y, **z))

    res.sort(key=lambda r: -((r['pct_fusion'] or 0) + (r['pct_homogene'] or 0)))
    print('\n{}'.format('B joue {:+d} demi-tons au-dessus de A — paires ORIENTÉES'
                        .format(a.octave) if a.octave else 'À l\'unisson'))
    print('\n{:22s} {:22s} {:>4s} {:>6s} {:>6s} {:>6s} {:>6s} {:>6s} {:>6s}  {}'.format(
        'A', 'B', 'n', 'fusion', 'chimèr', 'homog', 'renfor', 'juxta', 'écart', 'dominante'))
    for r in res[:40]:
        print('{:22s} {:22s} {:4d} {:5d}% {:5d}% {:5d}% {:5d}% {:5d}% {:>6} {}'.format(
            r['A'][:22], r['B'][:22], r['n_hauteurs'], r['pct_fusion'], r['pct_chimere'] or 0,
            r['pct_homogene'], r['pct_renforcement'], r['pct_juxtaposition'],
            r['ecart_parents_bark'] if r['ecart_parents_bark'] is not None else '—',
            r['classe_dominante']))

    with open(os.path.join(a.sortie, 'melanges.csv'), 'w', encoding='utf-8', newline='') as f:
        if res:
            w = csv.DictWriter(f, fieldnames=list(res[0]))
            w.writeheader(); w.writerows(res)
    with open(os.path.join(a.sortie, 'melanges-detail.json'), 'w', encoding='utf-8') as f:
        json.dump(detail, f, ensure_ascii=False)
    print('\n{} paires -> {}'.format(len(res), a.sortie))

if __name__ == '__main__':
    main()
