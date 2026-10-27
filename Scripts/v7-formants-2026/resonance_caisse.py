# -*- coding: utf-8 -*-
"""
resonance_caisse.py — tâche C1 : la méthode 12 voit-elle une résonance de caisse au
violoncelle et à la contrebasse, comme elle en a vu une à l'alto ?

Pourquoi un script de plus
--------------------------
La carte vocalique agrège par cellule : elle rend une médiane de F1 par registre et par
dynamique. Sur l'alto, cette médiane se tient entre 347 et 498 Hz sur trois corpus, et on
en a conclu à la résonance principale de caisse. L'argument repose sur **un seul contrôle**,
la marque `F1=f0` : le premier maximum n'est pas la fondamentale.

Ce contrôle ne suffit pas aux cordes graves. Au violoncelle, la zone de caisse attendue par
la littérature (A0 vers 100 Hz, B1 vers 180–250 Hz) tombe exactement sur les harmoniques 2 à
4 des notes du grave : 2 x 98 = 196, 3 x 65 = 195. Un « F1 stable à 200 Hz » peut donc n'être
que l'harmonique le plus fort, qui suit la note au lieu de rester en place. La marque
`F1=f0` ne le verrait pas, puisqu'elle ne teste que l'harmonique 1.

Le test qui tranche
-------------------
Une résonance ne bouge pas quand la note bouge ; un harmonique, si. On mesure donc, note par
note, la **pente de log2(F1) contre log2(f0)** :

    pente ~ 0   ->  F1 reste en place pendant que f0 se déplace : RÉSONANCE
    pente ~ 1   ->  F1 suit la note : HARMONIQUE, pas une propriété de l'instrument

Le test n'a de force que si f0 se déplace assez : on exige une étendue d'au moins six
demi-tons dans le registre, sans quoi le verdict est « étendue insuffisante ».

On rend en plus, par note, le rang harmonique F1/f0. S'il se groupe sur un entier alors que
la pente est nulle, c'est que la résonance est excitée par des harmoniques différents selon
la note — ce qui est précisément la signature attendue d'une résonance de caisse.

L'alto est traité avec les mêmes règles, comme témoin : c'est le test de l'affirmation déjà
publiée dans `_CARTE-VOCALIQUE.md`.

Usage
    python3 resonance_caisse.py --sortie <dossier>
    python3 resonance_caisse.py --instruments Violoncelle,Contrebasse --registres grave,médium
"""
import os, sys, json, math, argparse
from collections import defaultdict

import numpy as np

ICI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ICI)
import formants_lib as FL
import carte_vocalique as CV

RACINE = os.path.dirname(os.path.dirname(ICI))
ETUDE = os.path.join(RACINE, 'Etude-Formants')
REGISTRES = os.path.join(RACINE, 'Data', 'registres.md')
SOL_OUT = os.path.expanduser('~/work/out')

# instrument -> (clé dans registres.md, fichier SOL, nom dans les corpus v7)
CORDES = {
    'Violoncelle': ('violoncelle', 'Violoncello', 'Violoncelle'),
    'Contrebasse': ('contrebasse', 'Contrabass', 'Contrebasse'),
    'Alto':        ('alto',        'Viola',       'Alto'),
    'Violon':      ('violon',      'Violin',      'Violon'),
}

# Résonances de caisse attendues, d'après la littérature acoustique (Jansson, Woodhouse,
# Bissinger). Elles ne servent qu'à la lecture : aucun verdict ne s'y réfère.
ATTENDU = {
    'Violoncelle': 'A0 ~ 100 Hz, B1 ~ 180–250 Hz',
    'Contrebasse': 'A0 ~ 55–65 Hz, corps ~ 70–110 Hz',
    'Alto':        'A0 ~ 220–250 Hz, B1+ ~ 350–450 Hz',
    'Violon':      'A0 ~ 275 Hz, B1− ~ 460, B1+ ~ 530–550 Hz',
}

SEUIL_ETENDUE = 6.0      # demi-tons de f0 à couvrir pour que le test ait un sens
SEUIL_PENTE_RES = 0.25   # |pente| en dessous : F1 ne suit pas la note
SEUIL_PENTE_HAR = 0.75   # |pente| au-dessus : F1 suit la note
SEUIL_IQR_BARK = 1.5     # dispersion de F1 tolérée pour parler de résonance
MIN_NOTES = 6

# ──────────────────── diagnostic : F1 est-il un centroïde ? ────────────────
# L'échelle de Bark est très comprimée dans le grave : 3,5 Bark, seuil de fusion de
# Chistovich, vaut environ 0–350 Hz en bas de l'échelle. Les harmoniques d'une note de
# contrebasse à 41 Hz (41, 82, 123, 164, 205…) sont donc tous à moins de 3,5 Bark de leur
# voisin et se fusionnent en chaîne : le « F1 » rendu n'est alors pas un maximum mais le
# centre de gravité énergétique des premiers harmoniques, c'est-à-dire un descripteur de
# pente spectrale. On mesure combien de maxima ont été fusionnés, et on compare F1 au
# centroïde pur. Si les deux coïncident, F1 ne dit rien d'une résonance.
def maxima_avec_taille(exc, seuil_rel=25.0):
    """Comme CV.maxima_fusionnes, mais rend aussi la taille de chaque groupe fusionné."""
    pics = [i for i in range(1, len(exc) - 1)
            if exc[i] >= exc[i - 1] and exc[i] > exc[i + 1]]
    if not pics:
        return []
    plancher = exc.max() - seuil_rel
    pics = [i for i in pics if exc[i] >= plancher]
    if not pics:
        return []
    groupes, courant = [], [pics[0]]
    for i in pics[1:]:
        if CV._ZG[i] - CV._ZG[courant[-1]] < CV.FUSION_BARK:
            courant.append(i)
        else:
            groupes.append(courant); courant = [i]
    groupes.append(courant)
    out = []
    for g in groupes:
        ga = np.array(g)
        lin = 10.0 ** (exc[ga] / 10.0)
        z = float((CV._ZG[ga] * lin).sum() / lin.sum())
        out.append((z, len(g), float(CV._ZG[ga[0]]), float(CV._ZG[ga[-1]])))
    return out

def centroide_bark(exc, z_max=8.0):
    """Centre de gravité énergétique du motif d'excitation sous z_max Bark (~900 Hz).
    Aucune notion de maximum : c'est le témoin négatif."""
    lin = 10.0 ** (exc / 10.0)
    m = CV._ZG <= z_max
    return float((CV._ZG[m] * lin[m]).sum() / lin[m].sum())

# ───────────────────────────── collecte ────────────────────────────────────
def resoudre_registres(regs, cle):
    """Même tolérance que `carte()` : « violoncelle » trouve la ligne
    « violoncelle et ensemble de violoncelles » de registres.md."""
    if cle in regs:
        return regs[cle]
    cands = [k for k in regs if k.startswith(cle + ' ') or cle in k.split(' et ')]
    if len(cands) != 1:
        raise SystemExit('registre inconnu ou ambigu : %s (candidats %s)' % (cle, cands))
    return regs[cands[0]]

def chemins(nom):
    """Tous les fichiers de spectres disponibles pour cet instrument, par corpus."""
    cle, fich, nomv7 = CORDES[nom]
    out = {}
    p = os.path.join(SOL_OUT, fich + '.txt')
    if os.path.exists(p):
        out['SOL'] = p
    for c in ('ZSamples', 'Xsamples', 'VO', 'VN'):
        for cand in (nomv7, nomv7.replace('e ', 'e ')):
            p = os.path.join(ETUDE, 'Mesures', c, cand + '.spectres.txt')
            if os.path.exists(p):
                out[c] = p
                break
    return out

def f1_par_note(chemin):
    """Un point par échantillon :
    (f0, midi, dyn, F1_hz, rang, marque, n_fusionnes, etendue_groupe_bark, centroide_hz)"""
    pts = []
    for f0, midi, dyn, spectre in FL.lire_spectres(chemin):
        if not f0 or f0 <= 0:
            continue
        exc = CV.motif_excitation(spectre)
        m = maxima_avec_taille(exc)
        if not m:
            continue
        z, taille, zlo, zhi = m[0]
        h1 = CV.bark_vers_hz(z)
        marque = 'F1=f0' if abs(12 * math.log2(max(h1, 1e-6) / f0)) < 1.0 else ''
        if len(m) < 2:
            marque = ('1 max ' + marque).strip()
        pts.append((f0, midi, dyn, h1, h1 / f0, marque, taille, zhi - zlo,
                    CV.bark_vers_hz(centroide_bark(exc))))
    return pts

# ───────────────────────────── verdict ─────────────────────────────────────
def pente(xs, ys):
    """Moindres carrés sur log2 : pente de log2(F1) contre log2(f0), et r²."""
    x = np.log2(np.asarray(xs, float)); y = np.log2(np.asarray(ys, float))
    x = x - x.mean(); yc = y - y.mean()
    den = float((x * x).sum())
    if den <= 0:
        return None, None
    a = float((x * yc).sum() / den)
    res = yc - a * x
    sst = float((yc * yc).sum())
    r2 = 1.0 - float((res * res).sum()) / sst if sst > 0 else None
    return a, r2

def verdict(etendue, a, iqr_bark, n):
    if n < MIN_NOTES:
        return 'trop peu de notes'
    if etendue < SEUIL_ETENDUE:
        return 'étendue de f0 insuffisante'
    if a is None:
        return '—'
    if abs(a) <= SEUIL_PENTE_RES and iqr_bark <= SEUIL_IQR_BARK:
        return 'RÉSONANCE'
    if abs(a) <= SEUIL_PENTE_RES:
        return 'en place mais dispersé'
    if abs(a) >= SEUIL_PENTE_HAR:
        return 'harmonique (F1 suit la note)'
    return 'intermédiaire'

def analyser(nom, corpus, chemin, regs, registres_voulus):
    pts = f1_par_note(chemin)
    par_reg = defaultdict(list)
    for f0, midi, dyn, h1, rang, marque, taille, span, cen in pts:
        r = CV.registre_de(regs, midi)
        if r:
            par_reg[r].append((f0, h1, rang, marque, midi, taille, span, cen))
    lignes = []
    for r, L in par_reg.items():
        if registres_voulus and r not in registres_voulus:
            continue
        f0s = [x[0] for x in L]; h1s = [x[1] for x in L]
        etendue = 12 * math.log2(max(f0s) / min(f0s)) if min(f0s) > 0 else 0.0
        a, r2 = pente(f0s, h1s)
        zb = CV.hz_vers_bark(np.asarray(h1s))
        iqr_b = float(np.percentile(zb, 75) - np.percentile(zb, 25))
        rangs = np.asarray([x[2] for x in L])
        # un rang est « entier » s'il est à moins de 6 % d'un entier >= 1
        ent = np.abs(rangs - np.round(rangs)) / np.maximum(np.round(rangs), 1)
        lignes.append({
            'instrument': nom, 'corpus': corpus, 'registre': r, 'n': len(L),
            'f0_min': round(min(f0s), 1), 'f0_max': round(max(f0s), 1),
            'etendue_f0_dt': round(etendue, 1),
            'F1_median_hz': int(round(float(np.median(h1s)))),
            'F1_q1_hz': int(round(float(np.percentile(h1s, 25)))),
            'F1_q3_hz': int(round(float(np.percentile(h1s, 75)))),
            'F1_iqr_bark': round(iqr_b, 2),
            'pente': None if a is None else round(a, 2),
            'r2': None if r2 is None else round(r2, 2),
            'rang_median': round(float(np.median(rangs)), 2),
            'pct_rang_entier': int(round(100 * float(np.mean(ent < 0.06)))),
            'pct_F1_egal_f0': int(round(100 * float(np.mean(
                [1.0 if 'F1=f0' in x[3] else 0.0 for x in L])))),
            'pct_1max': int(round(100 * float(np.mean(
                [1.0 if '1 max' in x[3] else 0.0 for x in L])))),
            'maxima_fusionnes_med': round(float(np.median([x[5] for x in L])), 1),
            'groupe_bark_med': round(float(np.median([x[6] for x in L])), 1),
            'centroide_hz': int(round(float(np.median([x[7] for x in L])))),
            'ecart_F1_centroide_dt': round(12 * math.log2(
                float(np.median(h1s)) / max(float(np.median([x[7] for x in L])), 1e-6)), 1),
            'verdict': verdict(etendue, a, iqr_b, len(L)),
        })
    ordre = ['extrême grave', 'grave', 'médium', 'bas médium', 'haut médium',
             'aigu', 'suraigu']
    lignes.sort(key=lambda d: ordre.index(d['registre'])
                if d['registre'] in ordre else 99)
    return lignes

# ──────────────────────────────── figure ───────────────────────────────────
def figure(noms, regs_tous, voulus, chemin):
    """F1 contre f0, note par note. Une résonance donne un nuage horizontal ; un
    harmonique donne un nuage sur la diagonale. C'est tout l'argument, en une image."""
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    from matplotlib.lines import Line2D
    COUL = {'SOL': '#1f5c8b', 'ZSamples': '#b3522c', 'Xsamples': '#2e7d4f',
            'VO': '#7a4b9c', 'VN': '#9c7a1f'}
    plt.rcParams.update({'font.size': 9, 'axes.spines.top': False,
                         'axes.spines.right': False, 'axes.grid': True,
                         'grid.alpha': .25, 'grid.linewidth': .5,
                         'figure.dpi': 130, 'savefig.dpi': 130,
                         'savefig.bbox': 'tight', 'legend.frameon': False})
    # Une ligne par registre : mélanger les registres masque le résultat, puisque
    # c'est précisément le registre qui sépare la résonance de l'harmonique.
    lignes = [r for r in ('grave', 'médium', 'aigu', 'suraigu')
              if not voulus or r in voulus] or ['grave']
    fig, axes = plt.subplots(len(lignes), len(noms),
                             figsize=(3.1 * len(noms), 2.9 * len(lignes)),
                             squeeze=False)
    vus = set()
    # première passe : collecter, pour partager l'axe vertical par instrument
    donnees = {}
    for nom in noms:
        regs = resoudre_registres(regs_tous, CORDES[nom][0])
        for corpus, p in chemins(nom).items():
            for f0, midi, dyn, h1, rang, marque, taille, span, cen in f1_par_note(p):
                r = CV.registre_de(regs, midi)
                if r in lignes:
                    donnees.setdefault((nom, r), {}).setdefault(corpus, ([], []))
                    donnees[(nom, r)][corpus][0].append(f0)
                    donnees[(nom, r)][corpus][1].append(h1)
                    vus.add(corpus)
    for j, nom in enumerate(noms):
        tous = [v for r in lignes for d in [donnees.get((nom, r), {})]
                for xs, ys in d.values() for v in ys]
        ylim = (min(tous) * 0.85, max(tous) * 1.18) if tous else None
        for i, r in enumerate(lignes):
            ax = axes[i][j]
            d = donnees.get((nom, r), {})
            f1s = [v for xs, ys in d.values() for v in ys]
            if not f1s:
                ax.text(.5, .5, 'aucune note', transform=ax.transAxes, ha='center',
                        color='#999', fontsize=8); ax.set_xticks([]); ax.set_yticks([])
                continue
            for corpus, (xs, ys) in d.items():
                ax.scatter(xs, ys, s=12, color=COUL[corpus], alpha=.55,
                           edgecolor='none', zorder=3)
            med = float(np.median(f1s))
            # La pente sur les corpus réunis peut être plate alors qu'aucun corpus pris
            # seul ne l'est : on affiche donc aussi l'écart des pentes par corpus, et
            # c'est le tableau par corpus qui tranche, jamais cette ligne.
            parc = [pente(xs, ys)[0] for xs, ys in d.values() if len(xs) >= MIN_NOTES]
            parc = [p for p in parc if p is not None]
            ax.axhline(med, color='#333', lw=1.1, zorder=4)
            txt = '{:.0f} Hz · pente {:+.2f}'.format(med, pente(
                [x for xs, ys in d.values() for x in xs], f1s)[0] or 0)
            if len(parc) > 1:
                txt += '\npar corpus {:+.2f} à {:+.2f}'.format(min(parc), max(parc))
            ax.annotate(txt, (0.035, med), xycoords=('axes fraction', 'data'),
                        xytext=(0, 4), textcoords='offset points',
                        fontsize=7.5, color='#222', fontweight='bold')
            ax.set_xscale('log'); ax.set_yscale('log')
            if ylim:
                ax.set_ylim(*ylim)
            lo, hi = ax.get_xlim()
            dd = np.array([max(lo, 20.0), hi])
            for k in (1, 2, 3, 4):
                ax.plot(dd, k * dd, color='#ccc', lw=0.7, ls='--', zorder=1)
                if ylim and ylim[0] <= k * dd[1] <= ylim[1]:
                    ax.annotate('×{}'.format(k), (dd[1], k * dd[1]), fontsize=7,
                                color='#aaa', va='center', xytext=(2, 0),
                                textcoords='offset points')
            # sur une échelle log étroite, matplotlib empile les étiquettes mineures
            PALIERS = [20, 25, 30, 40, 50, 60, 70, 80, 100, 125, 150, 200, 250, 300,
                       400, 500, 600, 800, 1000, 1250, 1500, 2000, 3000, 4000]
            for axe, (p0, p1) in ((ax.xaxis, ax.get_xlim()), (ax.yaxis, ax.get_ylim())):
                t = [c for c in PALIERS if p0 <= c <= p1]
                axe.set_minor_locator(matplotlib.ticker.NullLocator())
                if len(t) >= 2:
                    axe.set_major_locator(matplotlib.ticker.FixedLocator(t))
                axe.set_major_formatter(
                    matplotlib.ticker.FuncFormatter(lambda v, _: '%g' % v))
            if i == 0:
                ax.set_title(nom, loc='left', fontweight='bold', fontsize=10)
            if i == len(lignes) - 1:
                ax.set_xlabel('f0 de la note (Hz)')
            if j == 0:
                ax.set_ylabel('{}\nF1 (Hz)'.format(r), fontsize=9)
    fig.legend(handles=[Line2D([], [], marker='o', ls='', color=COUL[c], label=c)
                        for c in ('SOL', 'ZSamples', 'VO', 'VN') if c in vus],
               loc='upper center', bbox_to_anchor=(0.5, 0.0),
               bbox_transform=fig.transFigure, ncol=4, fontsize=8)
    fig.suptitle('F1 contre la note jouée, registre par registre — un nuage horizontal '
                 'est une résonance,\nun nuage le long d\'une diagonale ×k est un '
                 'harmonique qui suit la note',
                 fontsize=9.5, color='#444', y=1.0)
    fig.tight_layout(rect=(0, 0, 1, 0.95))
    fig.savefig(chemin)
    plt.close(fig)
    return True

# ──────────────────────────────── pilote ───────────────────────────────────
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--instruments', default='Violoncelle,Contrebasse,Alto,Violon')
    ap.add_argument('--registres', default='grave,médium')
    ap.add_argument('--sortie', default=os.path.join(ETUDE, 'Mesures', 'Resonance-caisse'))
    ap.add_argument('--figure', metavar='PNG',
                    help='écrit en plus la figure F1 contre f0')
    a = ap.parse_args()
    voulus = {x.strip() for x in a.registres.split(',') if x.strip()} \
        if a.registres not in ('', 'tous') else set()
    os.makedirs(a.sortie, exist_ok=True)

    regs_tous = CV.lire_registres(REGISTRES)
    tout = []
    for nom in [x.strip() for x in a.instruments.split(',') if x.strip()]:
        regs = resoudre_registres(regs_tous, CORDES[nom][0])
        ch = chemins(nom)
        print('\n' + '=' * 78)
        print('{}   corpus : {}'.format(nom, ', '.join(ch) or 'aucun'))
        print('attendu par la littérature : {}'.format(ATTENDU[nom]))
        print('=' * 78)
        print('{:10s} {:9s} {:>4s} {:>13s} {:>5s} {:>16s} {:>5s} {:>6s} {:>5s} '
              '{:>5s} {:>5s} {:>7s}  {}'
              .format('corpus', 'registre', 'n', 'f0', 'ét.dt', 'F1 (q1–méd–q3)',
                      'IQRb', 'pente', 'rang', 'fus.', 'centr', 'F1−ctr', 'verdict'))
        for corpus, p in ch.items():
            for d in analyser(nom, corpus, p, regs, voulus):
                tout.append(d)
                print('{:10s} {:9s} {:4d} {:5.0f}–{:<7.0f} {:5.1f} {:6d}–{:4d}–{:<4d} '
                      '{:5.2f} {:>6} {:5.2f} {:5.1f} {:5d} {:6.1f}dt  {}'.format(
                          d['corpus'], d['registre'], d['n'], d['f0_min'], d['f0_max'],
                          d['etendue_f0_dt'], d['F1_q1_hz'], d['F1_median_hz'],
                          d['F1_q3_hz'], d['F1_iqr_bark'],
                          '—' if d['pente'] is None else d['pente'],
                          d['rang_median'], d['maxima_fusionnes_med'],
                          d['centroide_hz'], d['ecart_F1_centroide_dt'], d['verdict']))

    with open(os.path.join(a.sortie, 'resonance-caisse.json'), 'w', encoding='utf-8') as f:
        json.dump(tout, f, ensure_ascii=False, indent=1)
    import csv as _csv
    if tout:
        with open(os.path.join(a.sortie, 'resonance-caisse.csv'), 'w',
                  encoding='utf-8', newline='') as f:
            w = _csv.DictWriter(f, fieldnames=list(tout[0]))
            w.writeheader(); w.writerows(tout)
    print('\n{} lignes -> {}'.format(len(tout), a.sortie))
    if a.figure:
        noms = [x.strip() for x in a.instruments.split(',') if x.strip()]
        figure(noms, regs_tous, voulus, a.figure)
        print('figure -> {}'.format(a.figure))

if __name__ == '__main__':
    main()
