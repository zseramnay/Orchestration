# -*- coding: utf-8 -*-
"""Figure du doublage à l'octave : la loi se retourne, et la paire devient orientée."""
import os, csv, statistics as st
from collections import defaultdict
import numpy as np, matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D

ICI = os.path.dirname(os.path.abspath(__file__))
RACINE = os.path.dirname(os.path.dirname(ICI))
MES = os.path.join(RACINE, 'Etude-Formants', 'Mesures')
SORTIE = os.path.join(RACINE, 'Etude-Formants', 'Site-v2', 'figures',
                      'synthese-melanges-octave.png')
COUL = {'fusion': '#b3522c', 'homogène': '#1f5c8b',
        'renforcement': '#2e7d4f', 'juxtaposition': '#7a4b9c'}
plt.rcParams.update({'font.size': 9, 'axes.spines.top': False, 'axes.spines.right': False,
                     'axes.grid': True, 'grid.alpha': .25, 'grid.linewidth': .5,
                     'figure.dpi': 130, 'savefig.dpi': 130, 'savefig.bbox': 'tight',
                     'legend.frameon': False})

def lire(p):
    R = list(csv.DictReader(open(p, encoding='utf-8')))
    for r in R:
        for k in ('pct_fusion', 'pct_homogene', 'pct_renforcement', 'pct_juxtaposition',
                  'n_hauteurs'):
            r[k] = int(r[k]) if r[k] else 0
        r['ecart'] = float(r['ecart_parents_bark']) if r['ecart_parents_bark'] else None
    return [r for r in R if r['ecart'] is not None]

U = lire(os.path.join(MES, 'Melanges', 'melanges.csv'))
O = lire(os.path.join(MES, 'Melanges-octave', 'melanges.csv'))
fig, (g, d) = plt.subplots(1, 2, figsize=(11.2, 4.2))

# ── gauche : la loi, unisson contre octave ─────────────────────────────────
B = [(0, 1), (1, 2), (2, 3), (3, 4), (4, 9)]
for L, ls, lib in ((U, '--', 'unisson'), (O, '-', 'octave')):
    xs, ys = [], []
    for lo, hi in B:
        gg = [r for r in L if lo <= r['ecart'] < hi]
        if len(gg) < 5:
            continue
        xs.append((lo + min(hi, 5)) / 2)
        ys.append(float(np.mean([r['pct_fusion'] for r in gg])))
    g.plot(xs, ys, ls, marker='o', ms=5, lw=2, color=COUL['fusion'],
           alpha=1 if ls == '-' else .45)
    xs, ys = [], []
    for lo, hi in B:
        gg = [r for r in L if lo <= r['ecart'] < hi]
        if len(gg) < 5:
            continue
        xs.append((lo + min(hi, 5)) / 2)
        ys.append(float(np.mean([r['pct_juxtaposition'] for r in gg])))
    g.plot(xs, ys, ls, marker='s', ms=4, lw=1.6, color=COUL['juxtaposition'],
           alpha=1 if ls == '-' else .45)
g.axvspan(3, 4, color='#e8b64c', alpha=.22, lw=0, zorder=0)
g.annotate('maximum\nde fusion', (3.5, 6), ha='center', fontsize=8, color='#8a6a1a')
g.set_xlabel('écart de couleur entre les deux instruments (Bark)')
g.set_ylabel('part des cas (%)')
g.set_title('La loi a un maximum', loc='left', fontweight='bold', pad=19)
g.text(0, 1.015, "l'octave peuple le haut de l'échelle et le fait apparaître",
       transform=g.transAxes, fontsize=8.5, color='#666', va='bottom')
g.legend(handles=[Line2D([], [], color=COUL['fusion'], marker='o', label='fusion'),
                  Line2D([], [], color=COUL['juxtaposition'], marker='s',
                         label='juxtaposition'),
                  Line2D([], [], color='#888', ls='-', label='à l\'octave'),
                  Line2D([], [], color='#888', ls='--', alpha=.5, label='à l\'unisson')],
         loc='upper left', ncol=2, fontsize=8)

# ── droite : l'asymétrie des deux sens ─────────────────────────────────────
par = defaultdict(dict)
for r in O:
    if r['n_hauteurs'] >= 20:
        par[frozenset((r['A'], r['B']))][(r['A'], r['B'])] = r
pts = []
for v in par.values():
    if len(v) == 2:
        (k1, r1), (k2, r2) = list(v.items())
        pts.append((min(r1['pct_fusion'], r2['pct_fusion']),
                    max(r1['pct_fusion'], r2['pct_fusion'])))
d.scatter([p[0] for p in pts], [p[1] for p in pts], s=22, color=COUL['fusion'],
          alpha=.5, edgecolor='none', zorder=3)
m = max(max(p) for p in pts) + 5
d.plot([0, m], [0, m], color='#bbb', lw=1, ls='--', zorder=1)
d.annotate('si le sens ne changeait rien,\ntout serait sur cette ligne', (m * .62, m * .62),
           xytext=(6, -26), textcoords='offset points', fontsize=8, color='#888')
d.set_xlabel('fusion dans le sens le moins favorable (%)')
d.set_ylabel('fusion dans le sens le plus favorable (%)')
d.set_title('Une paire à l\'octave est orientée', loc='left', fontweight='bold', pad=19)
d.text(0, 1.015, '{} paires calculées dans les deux sens · écart médian {:.0f} points'
       .format(len(pts), st.median([b - a for a, b in pts])),
       transform=d.transAxes, fontsize=8.5, color='#666', va='bottom')
fig.savefig(SORTIE)
print('figure ->', SORTIE)
