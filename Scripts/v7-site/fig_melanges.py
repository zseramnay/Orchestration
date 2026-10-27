# -*- coding: utf-8 -*-
"""Figure de la tâche C9 : la loi des mélanges, et la matrice des paires."""
import os, sys, csv, math
from collections import defaultdict
import numpy as np, matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D

ICI = os.path.dirname(os.path.abspath(__file__))
RACINE = os.path.dirname(os.path.dirname(ICI))
ETUDE = os.path.join(RACINE, 'Etude-Formants')
SRC = os.path.join(ETUDE, 'Mesures', 'Melanges', 'melanges.csv')
SORTIE = os.path.join(ETUDE, 'Site-v2', 'figures', 'synthese-melanges.png')

COUL = {'fusion': '#b3522c', 'homogène': '#1f5c8b',
        'renforcement': '#2e7d4f', 'juxtaposition': '#7a4b9c'}
plt.rcParams.update({'font.size': 9, 'axes.spines.top': False, 'axes.spines.right': False,
                     'axes.grid': True, 'grid.alpha': .25, 'grid.linewidth': .5,
                     'figure.dpi': 130, 'savefig.dpi': 130, 'savefig.bbox': 'tight',
                     'legend.frameon': False})

R = list(csv.DictReader(open(SRC, encoding='utf-8')))
for r in R:
    for k in ('pct_fusion', 'pct_homogene', 'pct_renforcement', 'pct_juxtaposition',
              'n_hauteurs'):
        r[k] = int(r[k]) if r[k] else 0
    r['ecart'] = float(r['ecart_parents_bark']) if r['ecart_parents_bark'] else None
R = [r for r in R if r['ecart'] is not None]

fig, (g, d) = plt.subplots(1, 2, figsize=(11.2, 4.2),
                           gridspec_kw={'width_ratios': [1.15, 1]})

# ── gauche : la loi, par bandes d'écart ────────────────────────────────────
B = [(0, .5), (.5, 1), (1, 1.5), (1.5, 2), (2, 2.5), (2.5, 3.5), (3.5, 9)]
xs, series = [], defaultdict(list)
for lo, hi in B:
    gg = [r for r in R if lo <= r['ecart'] < hi]
    if not gg:
        continue
    xs.append((lo + min(hi, 4.5)) / 2)
    for k, c in (('pct_fusion', 'fusion'), ('pct_homogene', 'homogène'),
                 ('pct_renforcement', 'renforcement'), ('pct_juxtaposition', 'juxtaposition')):
        series[c].append(float(np.mean([r[k] for r in gg])))
for c, v in series.items():
    g.plot(xs, v, marker='o', ms=5, lw=1.8, color=COUL[c], label=c)
g.set_xlabel('écart de couleur entre les deux instruments (Bark)')
g.set_ylabel('part des cas (%)')
g.set_ylim(0, 72)
g.set_title('La loi des mélanges', loc='left', fontweight='bold', pad=19)
g.text(0, 1.015, "une seule grandeur ordonne presque tout : l'écart de couleur",
       transform=g.transAxes, fontsize=8.5, color='#666', va='bottom')
g.legend(loc='upper center', ncol=2, fontsize=8)

# ── droite : chaque paire, fusion contre écart ─────────────────────────────
for r in R:
    dom = max(('fusion', r['pct_fusion']), ('homogène', r['pct_homogene']),
              ('renforcement', r['pct_renforcement']),
              ('juxtaposition', r['pct_juxtaposition']), key=lambda t: t[1])[0]
    d.scatter(r['ecart'], r['pct_fusion'], s=9 + r['n_hauteurs'] * 0.35,
              color=COUL[dom], alpha=.55, edgecolor='none', zorder=3)
NOTE = [('Trompette en Ut', 'Tuba basse', 'trompette + tuba'),
        ('Flûte', 'Petite flûte', 'flûte + petite flûte'),
        ('Clarinette en Sib', 'Cor en Fa', 'clarinette + cor'),
        ('Hautbois', 'Trombone ténor', 'hautbois + trombone')]
for a, b, lib in NOTE:
    m = [r for r in R if {r['A'], r['B']} == {a, b}]
    if m:
        d.annotate(lib, (m[0]['ecart'], m[0]['pct_fusion']), xytext=(6, 5),
                   textcoords='offset points', fontsize=7.5, color='#333')
d.set_xlabel('écart de couleur (Bark)')
d.set_ylabel('cas de fusion (%)')
d.set_title('Les 215 paires', loc='left', fontweight='bold', pad=19)
d.text(0, 1.015, "couleur : l'issue majoritaire · taille : hauteurs communes",
       transform=d.transAxes, fontsize=8.5, color='#666', va='bottom')
d.legend(handles=[Line2D([], [], marker='o', ls='', color=COUL[c], label=c) for c in COUL],
         loc='upper left', ncol=1, fontsize=7.5)
fig.savefig(SORTIE)
print('figure ->', SORTIE)
