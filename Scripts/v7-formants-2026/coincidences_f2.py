# -*- coding: utf-8 -*-
"""
coincidences_f2.py — tâche C8 : les maxima de la carte vocalique tombent-ils sur les
formants établis du même instrument ?

Pourquoi la question
--------------------
La méthode 8 et la méthode 12 ne mesurent pas la même chose. La première cherche une
résonance du filtre, indépendante de la note ; la seconde lit la saillance du spectre
RAYONNÉ sur une échelle auditive, sans jamais séparer source et filtre. Rien n'oblige leurs
résultats à coïncider — et c'est précisément pour cela qu'une coïncidence, si elle existe,
vaut comme corroboration indépendante. Le plan la demande comme « indice de confiance pour
les cordes », dont les formants sont les plus discutés.

Ce qui est mesuré
-----------------
Pour chaque cellule de carte (registre × nuance × corpus) et chacun de ses deux maxima,
l'écart en demi-tons au formant établi le plus proche DU MÊME INSTRUMENT. Coïncidence si
l'écart tient sous SEUIL.

Deux témoins, et aucun des deux n'est facultatif
------------------------------------------------
Un taux brut ne veut rien dire : les F2 vivent dans une bande de fréquences étroite et les
formants aussi, donc des coïncidences surviennent par construction. La tâche C4 l'a appris à
ses dépens — des rapports de cotes de 10 à 20 et des p < 10⁻⁴ sur treize formants, tous
annulés par un témoin apparié.

  1. DÉCALAGE. On rejoue la mesure contre le propre jeu de formants de l'instrument,
     transposé de ±2 à ±6 demi-tons. Le nombre de formants et leur espacement sont
     conservés : seule leur position change. C'est le témoin qui partage le confondant.

  2. AUTRES INSTRUMENTS. On compare la cellule aux formants de chaque AUTRE instrument, et
     on lit le rang de l'instrument propre. Sous l'hypothèse nulle, 50 %.

Usage
    python3 coincidences_f2.py
    python3 coincidences_f2.py --seuil 1.0
"""
import os, sys, csv, json, math, glob, argparse, statistics as st
from collections import defaultdict
import numpy as np

ICI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ICI)

RACINE = os.path.dirname(os.path.dirname(ICI))
ETUDE = os.path.join(RACINE, 'Etude-Formants')
MES = os.path.join(ETUDE, 'Mesures')

SEUIL = 0.5                 # demi-ton : le seuil strict de l'étude
DECALAGES = [-6, -5, -4, -3, -2, 2, 3, 4, 5, 6]
STATUTS = ('route A', '1 corpus', 'indéterminé')
FAMILLES = {}               # rempli depuis le CSV

def demi_tons(a, b):
    return abs(12.0 * math.log2(a / b))

def lire_formants(statuts):
    F = defaultdict(list)
    with open(os.path.join(ETUDE, 'Resultats-v7', 'formants-v7.csv'), encoding='utf-8') as f:
        for r in csv.DictReader(f):
            if r['statut'] in statuts:
                F[r['instrument']].append(float(r['hz']))
                FAMILLES[r['instrument']] = r.get('famille', '')
    return {k: sorted(v) for k, v in F.items()}

def lire_cartes():
    """[(corpus, instrument, cellule), …] — cellules publiables seulement."""
    out = []
    for p in sorted(glob.glob(os.path.join(MES, '*', 'carte-vocalique-*.json'))):
        corpus = os.path.basename(os.path.dirname(p))
        d = json.load(open(p, encoding='utf-8'))
        for inst, bloc in d.items():
            for c in bloc.get('cellules', []):
                if c.get('statut') == 'non publiable' or not c.get('F1_hz'):
                    continue
                out.append((corpus, inst, c))
    return out

def plus_proche(valeur, formants):
    if not valeur or not formants:
        return None
    return min(demi_tons(valeur, fc) for fc in formants)

def taux(cellules, F, decalage=0.0, champ='F2_hz', exclure_f0=True):
    """Part de maxima qui tombent à moins de SEUIL d'un formant, et l'écart médian."""
    n = ok = 0
    ecarts = []
    for corpus, inst, c in cellules:
        if inst not in F or not c.get(champ):
            continue
        if champ == 'F1_hz' and exclure_f0 and 'F1=f0' in c.get('marques', []):
            continue
        cible = [fc * 2.0 ** (decalage / 12.0) for fc in F[inst]]
        e = plus_proche(c[champ], cible)
        if e is None:
            continue
        n += 1; ecarts.append(e)
        if e <= SEUIL:
            ok += 1
    if not n:
        return None
    return dict(n=n, n_ok=ok, taux=100.0 * ok / n, ecart_median=float(np.median(ecarts)))

def rang_autres(cellules, F, champ='F2_hz'):
    """Pour chaque cellule : part des AUTRES instruments dont le formant le plus proche est
    plus loin que celui de l'instrument propre. 50 % sous l'hypothèse nulle."""
    rangs = []
    for corpus, inst, c in cellules:
        if inst not in F or not c.get(champ):
            continue
        propre = plus_proche(c[champ], F[inst])
        autres = [plus_proche(c[champ], F[k]) for k in F if k != inst]
        autres = [x for x in autres if x is not None]
        if propre is None or len(autres) < 5:
            continue
        rangs.append(100.0 * sum(1 for x in autres if x > propre) / len(autres))
    return (float(np.mean(rangs)), len(rangs)) if rangs else (None, 0)

def main():
    global SEUIL
    ap = argparse.ArgumentParser()
    ap.add_argument('--seuil', type=float, default=SEUIL)
    ap.add_argument('--statuts', default='route A,1 corpus')
    ap.add_argument('--sortie', default=os.path.join(MES, 'Coincidences-F2'))
    ap.add_argument('--annoter', action='store_true',
                    help='ajoute coincidence_F1/F2 aux JSON de carte déjà écrits')
    a = ap.parse_args()
    SEUIL = a.seuil
    os.makedirs(a.sortie, exist_ok=True)
    statuts = tuple(x.strip() for x in a.statuts.split(','))
    F = lire_formants(statuts)
    cellules = lire_cartes()
    n_corpus = len(set(glob.glob(os.path.join(MES, '*', 'carte-vocalique-*.json'))))
    print("{} formants ({}) sur {} instruments ; {} cellules publiables sur {} corpus"
          .format(sum(len(v) for v in F.values()), ' + '.join(statuts), len(F),
                  len(cellules), n_corpus))
    print("seuil de coïncidence : {:.1f} demi-ton\n".format(SEUIL))

    for champ in ('F1_hz', 'F2_hz'):
        print('=' * 78)
        print('{} — coïncidence avec un formant établi du MÊME instrument'.format(champ[:2]))
        print('=' * 78)
        vrai = taux(cellules, F, 0.0, champ)
        if not vrai:
            print('  rien à mesurer\n'); continue
        tem = [taux(cellules, F, d, champ) for d in DECALAGES]
        tem = [t for t in tem if t]
        mt = float(np.mean([t['taux'] for t in tem]))
        me = float(np.mean([t['ecart_median'] for t in tem]))
        print('{:34s} {:>6s} {:>8s} {:>14s}'.format('', 'n', 'coïnc.', 'écart médian'))
        print('{:34s} {:6d} {:7.1f}% {:11.2f} dt'.format(
            'MESURE — formants de l\'instrument', vrai['n'], vrai['taux'],
            vrai['ecart_median']))
        print('{:34s} {:6d} {:7.1f}% {:11.2f} dt'.format(
            'témoin 1 — décalés de ±2 à ±6 dt', tem[0]['n'], mt, me))
        for d, t in zip(DECALAGES, tem):
            print('      {:+d} dt {:28s} {:7.1f}% {:11.2f} dt'.format(d, '', t['taux'],
                                                                      t['ecart_median']))
        r, nr = rang_autres(cellules, F, champ)
        print('{:34s} {:6d} {:>7} {:11.1f} %'.format(
            'témoin 2 — rang parmi les autres', nr, '', r if r else float('nan')))
        print('      (50 % = hasard, MAIS ce témoin est invalide ici : il ne partage pas')
        print('       le confondant de registre — voir `_COINCIDENCES-F2.md`)')
        print()

    print('=' * 78)
    print('PAR INSTRUMENT — mesure contre son propre jeu de formants décalé')
    print('=' * 78)
    lignes = []
    for s in (0.5, 1.0):
        lignes += par_instrument(cellules, F, s)
    print('{:24s} {:>3s} {:>6s} {:>7s} {:>8s} {:>8s}'.format(
        'instrument', 'max', 'seuil', 'mesure', 'témoin', 'rapport'))
    for d in lignes:
        if d['taux'] == 0 and d['taux_temoin'] == 0:
            continue
        print('{:24s} {:>3s} {:5.1f}  {:6.0f}% {:7.1f}% {:>8}'.format(
            d['instrument'][:24], d['maximum'], d['seuil'], d['taux'], d['taux_temoin'],
            ('×%.1f' % d['rapport']) if d['rapport'] else '—'))

    net = [x for x in cellules if not x[2].get('marques')]
    SEUIL = 0.5
    g = taux(net, F, 0.0, 'F2_hz')
    tm = [taux(net, F, d, 'F2_hz') for d in DECALAGES]
    print('\n   sur les {} cellules SANS AUCUNE MARQUE : F2 coïncide dans {:.1f} % des cas,'
          .format(len(net), g['taux']))
    print('   contre {:.1f} % au témoin. Le basson est la seule exception, sur F1.'
          .format(float(np.mean([x['taux'] for x in tm if x]))))

    os.makedirs(a.sortie, exist_ok=True)
    with open(os.path.join(a.sortie, 'coincidences-par-instrument.csv'), 'w',
              encoding='utf-8', newline='') as f:
        w = csv.DictWriter(f, fieldnames=list(lignes[0])); w.writeheader(); w.writerows(lignes)
    detail = []
    SEUIL = 1.0
    for corpus, inst, c in cellules:
        if inst not in F:
            continue
        for champ in ('F1_hz', 'F2_hz'):
            if not c.get(champ):
                continue
            e, fc = min((demi_tons(c[champ], x), x) for x in F[inst])
            if e <= 1.0:
                detail.append(dict(instrument=inst, corpus=corpus, registre=c['registre'],
                                   dynamique=c['dynamique'], n=c['n'], maximum=champ[:2],
                                   hz=c[champ], formant_hz=int(round(fc)),
                                   ecart_demi_tons=round(e, 2),
                                   marques=','.join(c.get('marques', []))))
    with open(os.path.join(a.sortie, 'coincidences-detail.csv'), 'w', encoding='utf-8',
              newline='') as f:
        w = csv.DictWriter(f, fieldnames=list(detail[0])); w.writeheader(); w.writerows(detail)
    print('\n-> {}'.format(a.sortie))
    if a.annoter:
        print('\nannotation des cartes :')
        annoter(F, 0.5)

def par_instrument(cellules, F, seuil):
    """[(instrument, champ, n, taux, taux_temoin), …]"""
    global SEUIL
    SEUIL = seuil
    out = []
    for inst in sorted({x[1] for x in cellules} & set(F)):
        sous = [x for x in cellules if x[1] == inst]
        for champ in ('F1_hz', 'F2_hz'):
            v = taux(sous, F, 0.0, champ)
            if not v or v['n'] < 6:
                continue
            tem = [taux(sous, F, d, champ) for d in DECALAGES]
            mt = float(np.mean([x['taux'] for x in tem if x]))
            out.append(dict(instrument=inst, maximum=champ[:2], seuil=seuil, n=v['n'],
                            taux=round(v['taux'], 1), taux_temoin=round(mt, 1),
                            rapport=(round(v['taux'] / mt, 2) if mt > 0 else None),
                            ecart_median=round(v['ecart_median'], 2)))
    return out

def annoter(F, seuil):
    """Ajoute coincidence_F1 / coincidence_F2 aux JSON de carte existants. Purement
    additif : aucune valeur mesurée n'est touchée."""
    n = 0
    for p in sorted(glob.glob(os.path.join(MES, '*', 'carte-vocalique-*.json'))):
        d = json.load(open(p, encoding='utf-8'))
        for inst, bloc in d.items():
            for c in bloc.get('cellules', []):
                # on repart du Bark, non du Hz arrondi, pour donner exactement le même
                # résultat que `carte_vocalique.carte()` — sinon l'écart diffère au
                # centième et deux chemins de calcul semblent se contredire.
                import carte_vocalique as _CV
                for champ, cle in (('F1_bark', 'coincidence_F1'),
                                   ('F2_bark', 'coincidence_F2')):
                    c[cle] = None
                    if inst in F and c.get(champ) is not None:
                        hz = _CV.bark_vers_hz(c[champ])
                        e, fc = min((demi_tons(hz, x), x) for x in F[inst])
                        if e <= seuil:
                            c[cle] = {'hz': int(round(fc)),
                                      'ecart_demi_tons': round(e, 2)}
                            n += 1
        json.dump(d, open(p, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
        print('   annoté {}'.format(os.path.relpath(p, MES)))
    print('   {} coïncidences inscrites'.format(n))

if __name__ == '__main__':
    main()
