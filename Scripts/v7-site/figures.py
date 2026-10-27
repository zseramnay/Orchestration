# -*- coding: utf-8 -*-
"""
figures.py — regénère toutes les figures du site v2 (tâche B3 du plan de reprise).

Trois familles de figures :

  composite-<id>.png    par instrument : le résidu local médian de la méthode 8,
                        une courbe par corpus, avec les formants établis marqués.
                        C'est la figure qui MONTRE la route A : deux corpus de
                        provenances distinctes qui culminent au même point.

  vocalique-<id>.png    par instrument : les cellules registre x dynamique dans le
                        plan (F1, F2) en Bark, contre les voyelles du français
                        (Calliope, voix d'homme — et non les voyelles allemandes
                        de Meyer : c'est la tâche B4, réglée ici).

  synthese-*.png        par famille / propriété mesurée : convergences, rho par
                        corpus, pentes, balance impair-pair, largeurs confrontées.

Les 47 figures `specenv_*` de l'ancienne étude ne sont pas reprises : elles
montraient des pics de bins de FFT voisins, c'est-à-dire l'artefact lui-même.

Usage :
    python3 figures.py                 # tout
    python3 figures.py --only composite:basson,vocalique:alto,synthese
    python3 figures.py --liste         # audit des noms, n'écrit rien
"""
import os, sys, json, csv, math, re, unicodedata, glob, argparse
from collections import defaultdict

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib.patches import Rectangle

RACINE = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
ETUDE = os.path.join(RACINE, 'Etude-Formants')
MESURES = os.path.join(ETUDE, 'Mesures')
RES = os.path.join(ETUDE, 'Resultats-v7')
LARG = os.path.join(ETUDE, 'Largeurs')
SORTIE = os.path.join(ETUDE, 'Site-v2', 'figures')

CORPUS = ['SOL', 'ZSamples', 'Xsamples', 'VO', 'VN']
COULEUR_CORPUS = {'SOL': '#1f5c8b', 'ZSamples': '#b3522c', 'Xsamples': '#2e7d4f',
                  'VO': '#7a4b9c', 'VN': '#9c7a1f'}
COULEUR_FAMILLE = {'bois': '#2e7d4f', 'cuivres': '#b3522c', 'cordes': '#1f5c8b'}

def legende_bas(fig, handles, ncol):
    """Légende sous la figure. Dans l'axe, elle recouvrait la dernière ligne
    (Tuba contrebasse) ou une barre d'erreur."""
    fig.legend(handles=handles, loc='upper center', bbox_to_anchor=(0.5, 0.0),
               bbox_transform=fig.transFigure, ncol=ncol, fontsize=7.5)

def titre(ax, gauche, droite=None):
    """Titre en deux lignes. Un titre à gauche et un à droite sur la même ligne
    se chevauchent dès que le nom est long — ce qui arrivait sur cinq figures."""
    ax.set_title(gauche, loc='left', fontweight='bold', pad=19)
    if droite:
        ax.text(0, 1.015, droite, transform=ax.transAxes, fontsize=8.5,
                color='#666', va='bottom', ha='left')

plt.rcParams.update({
    'font.size': 9, 'axes.titlesize': 11, 'axes.labelsize': 9,
    'axes.spines.top': False, 'axes.spines.right': False,
    'axes.grid': True, 'grid.alpha': 0.25, 'grid.linewidth': 0.5,
    'figure.dpi': 130, 'savefig.dpi': 130, 'savefig.bbox': 'tight',
    'legend.frameon': False, 'legend.fontsize': 8,
})

# ───────────────────────────── noms ────────────────────────────────────────
def norm(s):
    s = unicodedata.normalize('NFD', s or '')
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return re.sub(r'\s+', ' ', s).strip().lower()

def charger_instruments():
    with open(os.path.join(RES, 'instruments-v7.csv'), encoding='utf-8') as f:
        L = list(csv.DictReader(f))
    for r in L:
        for k in ('ordre', 'n_echantillons', 'n_formants_retenus', 'n_route_A',
                  'harmoniques_par_note'):
            r[k] = int(r[k]) if r.get(k, '').strip() else None
        for k in ('pente_db_octave', 'pente_sd', 'balance_impair_pair_db',
                  'rho_etendue', 'enveloppe_vraie_hz'):
            r[k] = float(r[k]) if r.get(k, '').strip() else None
    return L

def charger_formants():
    with open(os.path.join(RES, 'formants-v7.csv'), encoding='utf-8') as f:
        L = list(csv.DictReader(f))
    for r in L:
        r['hz'] = float(r['hz'])
        for k in ('residu_db', 'largeur_env_hz', 'largeur_lpc_hz', 'Q_enveloppe'):
            r[k] = float(r[k]) if r.get(k, '').strip() else None
    return L

# Les corpus ne nomment pas tous l'instrument de la même façon. SOL écrit « Cor »
# là où la fiche écrit « Cor en Fa ». Table explicite : un appariement par préfixe
# confondrait « Clarinette contrebasse » et « Clarinette contrebasse en Sib » avec
# « Clarinette basse ».
ALIAS = {
    'cor': 'Cor en Fa',
    'trompette': 'Trompette en Ut',
    'clarinette basse': 'Clarinette basse en Sib',
    'clarinette contrebasse': 'Clarinette contrebasse en Sib',
}

def apparier(cible, candidats):
    """cible = nom de fiche ; candidats = noms trouvés dans les données.
    Égalité normalisée, puis table d'alias. Aucun appariement par préfixe :
    il confondrait les clarinettes."""
    n = norm(cible)
    for c in candidats:
        if norm(c) == n:
            return c
    for c in candidats:
        if ALIAS.get(norm(c)) == cible:
            return c
    return None

def resoudre(nom, noms_fiches):
    """nom tel qu'écrit dans les données -> nom de fiche, ou None."""
    n = norm(nom)
    for f in noms_fiches:
        if norm(f) == n:
            return f
    a = ALIAS.get(n)
    return a if a in noms_fiches else None

# ─────────────────────── enveloppes composites ─────────────────────────────
def charger_enveloppes(noms):
    """{nom_de_fiche: {corpus: (centres, mediane)}} — absorbe les deux formats.
    La résolution vers le nom de fiche a lieu ici : sans quoi « Cor » (SOL) et
    « Cor en Fa » (VN, Xsamples) resteraient deux instruments."""
    env = defaultdict(dict)
    perdus = defaultdict(list)
    # SOL : trois fichiers de lot, clés sous 'results', axe commun sous 'centers'
    for lot in ('mesures-2026-10-23', 'mesures-cordes-2026-10-23',
                'mesures-cuivres-2026-10-23'):
        p = os.path.join(MESURES, 'SOL', lot + '.json')
        if not os.path.exists(p):
            continue
        j = json.load(open(p, encoding='utf-8'))
        axe = j['centers']
        for nom, v in j['results'].items():
            f = resoudre(nom, noms)
            if f:
                env[f]['SOL'] = (axe, v['med'])
            else:
                perdus['SOL'].append(nom)
    # les quatre autres : un JSON par instrument, champ 'enveloppe'
    for c in CORPUS[1:]:
        for p in sorted(glob.glob(os.path.join(MESURES, c, '*.json'))):
            nom = os.path.splitext(os.path.basename(p))[0]
            if '+' in nom or ' corde ' in nom:      # sourdines, cordes à vide : suspendus
                continue
            try:
                j = json.load(open(p, encoding='utf-8'))
            except Exception:
                continue
            e = j.get('enveloppe')
            if isinstance(e, dict) and e.get('centres') and e.get('mediane'):
                f = resoudre(nom, noms)
                if f:
                    env[f][c] = (e['centres'], e['mediane'])
                else:
                    perdus[c].append(nom)
    return env, perdus

# ───────────────────────── cartes vocaliques ───────────────────────────────
def charger_cartes(noms):
    """{corpus: {nom_de_fiche: [cellules]}}. Deux formes dans les fichiers :
    la carte SOL du 24/10 porte la liste de cellules directement, les cartes des
    nouveaux corpus la portent sous la clé 'cellules'."""
    out = {}
    for p in sorted(glob.glob(os.path.join(MESURES, '*', 'carte-vocalique-*.json'))):
        corpus = os.path.basename(os.path.dirname(p))
        j = json.load(open(p, encoding='utf-8'))
        d = {}
        for nom, v in j.items():
            cell = v if isinstance(v, list) else (
                v.get('cellules') if isinstance(v, dict) else None)
            f = resoudre(nom, noms)
            if f and isinstance(cell, list):
                d[f] = cell
        out.setdefault(corpus, {}).update(d)
    return out

VOYELLES = [("i", 308, 2064), ("y", 300, 1750), ("u", 315, 764),
            ("e", 365, 1961), ("ø", 381, 1417), ("o", 383, 793),
            ("ɛ", 530, 1718), ("œ", 517, 1391), ("ɔ", 531, 998),
            ("a", 684, 1256)]

def bark(f):
    return 13.0 * math.atan(0.00076 * f) + 3.5 * math.atan((f / 7500.0) ** 2)

VOY_BARK = [(v, bark(a), bark(b)) for v, a, b in VOYELLES]

# Les registres ne sont pas quatre mais douze : les clarinettes ont chalumeau,
# gorge et clarine, les cuivres ont les sons pédale. Classés du grave à l'aigu,
# la couleur suit ce classement — deux instruments de familles différentes
# gardent ainsi des couleurs comparables.
REGISTRES = ['sons pédale', 'extrême grave', 'grave', 'chalumeau', 'bas médium',
             'gorge', 'médium', 'clarine', 'haut médium', 'aigu', 'suraigu']
COUL_REG = {r: plt.cm.viridis(0.06 + 0.84 * k / (len(REGISTRES) - 1))
            for k, r in enumerate(REGISTRES)}
COUL_REG['hors registre'] = '#999999'
DYN = ['doux', 'moyen', 'fort', 'toutes']
MARQ_DYN = {'doux': 'o', 'moyen': 's', 'fort': '^', 'toutes': 'D'}

# ═══════════════════════════ figure 1 : composite ══════════════════════════
def fig_composite(inst, env_par_corpus, formants, chemin):
    fig, ax = plt.subplots(figsize=(7.4, 3.5))
    trace = False
    for c in CORPUS:
        if c not in env_par_corpus:
            continue
        axe, med = env_par_corpus[c]
        xs = [f for f, m in zip(axe, med) if m is not None]
        ys = [m for m in med if m is not None]
        if len(xs) < 4:
            continue
        ax.plot(xs, ys, lw=1.4, color=COULEUR_CORPUS[c], label=c, alpha=.9,
                solid_capstyle='round')
        trace = True
    if not trace:
        plt.close(fig)
        return False

    ax.axhline(0, color='#999', lw=0.7, ls=':')
    ymin, ymax = ax.get_ylim()
    # Les étiquettes vont EN BAS : en haut elles heurtaient la légende et le titre.
    # Deux niveaux en alternance, pour que deux formants distants d'un demi-ton
    # (Si4 et Do♯5 du basson, par exemple) ne s'écrivent pas l'un sur l'autre.
    bas = ymin - (ymax - ymin) * 0.07
    for k, r in enumerate(sorted(formants, key=lambda x: x['hz'])):
        routeA = r['statut'] == 'route A'
        ax.axvline(r['hz'], color='#333' if routeA else '#aaa',
                   lw=1.1 if routeA else 0.8, ls='-' if routeA else '--',
                   alpha=.85, zorder=0)
        ax.annotate(r['note'], (r['hz'], bas - (ymax - ymin) * 0.075 * (k % 2)),
                    ha='center', va='top', fontsize=7.5,
                    color='#222' if routeA else '#888',
                    fontweight='bold' if routeA else 'normal',
                    bbox=dict(facecolor='white', edgecolor='none', pad=0.8))
    if formants:
        ax.set_ylim(bas - (ymax - ymin) * 0.20, ymax + (ymax - ymin) * 0.05)

    ax.set_xscale('log')
    ax.set_xlim(75, 9000)
    ax.set_xticks([100, 200, 500, 1000, 2000, 5000])
    ax.set_xticklabels(['100', '200', '500', '1 k', '2 k', '5 k'])
    ax.set_xlabel('fréquence (Hz, échelle logarithmique)')
    ax.set_ylabel('résidu local médian (dB)')
    n = len(formants)
    nA = sum(1 for r in formants if r['statut'] == 'route A')
    sous = ('aucun formant retenu' if n == 0 else
            '{} formant{} retenu{} · {} par route A'.format(
                n, 's' if n > 1 else '', 's' if n > 1 else '', nA))
    titre(ax, inst['instrument'], 'classe {} · {}'.format(inst['classe'] or '—', sous))
    ax.legend(loc='upper left', ncol=5, fontsize=7.5)
    fig.savefig(chemin)
    plt.close(fig)
    return True

# ════════════════════════ figure 2 : carte vocalique ═══════════════════════
def raison_absence(cells):
    """Pourquoi un instrument n'a pas de carte vocalique. Ce n'est pas un défaut
    de mesure : chez les cuivres graves et le contrebasson, un seul maximum
    survit à la fusion de Chistovich dans toutes les cellules exploitables, donc
    il n'y a pas de couple (F1, F2) à lire."""
    if not cells:
        return 'aucune cellule mesurée'
    m = [x for k in cells for x in (k.get('marques') or [])]
    n1 = sum(1 for x in m if x == '1 max')
    npub = sum(1 for k in cells if k.get('statut') == 'non publiable')
    if n1:
        return ('un seul maximum survit dans {} cellules sur {} : il n\'y a pas de '
                'couple (F1, F2) à lire'.format(n1, len(cells)))
    if npub == len(cells):
        return 'toutes les cellules sont trop peu peuplées'
    return 'aucun couple (F1, F2) exploitable'

def fig_vocalique(inst, cartes_par_corpus, chemin):
    cells = [(c, k) for c in CORPUS for k in cartes_par_corpus.get(c, [])]
    util = [(c, k) for c, k in cells if k.get('F1_bark') and k.get('F2_bark')]
    if not util:
        return False
    fig, ax = plt.subplots(figsize=(5.4, 4.6))

    for v, z1, z2 in VOY_BARK:
        ax.plot(z2, z1, marker='+', ms=7, color='#bbb', mew=1.2, zorder=1)
        ax.annotate('/{}/'.format(v), (z2, z1), xytext=(4, 4),
                    textcoords='offset points', fontsize=10, color='#999', zorder=1)

    vus = set()
    for c, k in util:
        reg = k.get('registre', '')
        dyn = k.get('dynamique', '') or 'moyen'
        ok = k.get('statut') == 'assignable'
        col = COUL_REG.get(reg, '#555')
        # bord sombre sur les marqueurs pleins : sans lui les registres aigus,
        # clairs dans le dégradé, disparaissaient sur le fond blanc
        ax.scatter(k['F2_bark'], k['F1_bark'],
                   marker=MARQ_DYN.get(dyn, 'o'), s=54 if ok else 40,
                   facecolor=col if ok else 'none',
                   edgecolor='#333' if ok else col,
                   linewidth=0.7 if ok else 1.3,
                   alpha=.95 if ok else .55, zorder=3)
        if ok:
            ax.errorbar(k['F2_bark'], k['F1_bark'],
                        xerr=(k.get('IQR2') or 0) / 2, yerr=(k.get('IQR1') or 0) / 2,
                        color=col, lw=0.7, alpha=.28, capsize=0, zorder=2)
        vus.add((reg, dyn, ok))
    if len(cartes_par_corpus) > 1:
        ax.text(0.015, 0.015, 'corpus : ' + ', '.join(
            c for c in CORPUS if cartes_par_corpus.get(c)),
            transform=ax.transAxes, fontsize=7.5, color='#888')

    ax.invert_xaxis()
    ax.invert_yaxis()
    ax.set_xlabel('F2 (Bark) →  plus grave')
    ax.set_ylabel('F1 (Bark) →  plus grave')
    nass = sum(1 for c, k in util if k.get('statut') == 'assignable')
    titre(ax, inst['instrument'],
          '{} cellules, dont {} assignables'.format(len(util), nass))
    leg = [Line2D([], [], marker='o', ls='', color=COUL_REG.get(r, '#555'),
                  mec='#333', mew=0.7, label=r)
           for r in REGISTRES + ['hors registre'] if any(v[0] == r for v in vus)]
    leg += [Line2D([], [], marker=MARQ_DYN[d], ls='', color='#555', label=d)
            for d in DYN if any(v[1] == d for v in vus)]
    leg += [Line2D([], [], marker='o', ls='', mfc='none', color='#555',
                   label='non assignable')]
    ax.legend(handles=leg, loc='center left', bbox_to_anchor=(1.01, 0.5), fontsize=7.5)
    fig.savefig(chemin)
    plt.close(fig)
    return True

# ════════════════════ figure 3 : convergences (synthèse) ═══════════════════
def zones_convergence(FORM):
    """Reprend mot pour mot le regroupement de `page_convergences()` dans
    construire_site.py : la figure et la page ne peuvent donc pas se contredire.
    Unisson = même point de grille, plusieurs instruments. Région = chaîne de
    points dont l'écart consécutif ≤ 1,05 demi-ton ET l'étendue totale ≤ 2."""
    par_hz = {}
    for f in FORM:
        if f['statut'] == 'route A':
            par_hz.setdefault(int(f['hz']), set()).add(f['instrument'])
    hzs = sorted(par_hz)
    unissons = [(h, h, sorted(par_hz[h])) for h in hzs if len(par_hz[h]) > 1]
    groupes = []
    for h in hzs:
        if groupes and (12 * math.log2(h / groupes[-1][-1]) <= 1.05
                        and 12 * math.log2(h / groupes[-1][0]) <= 2.0):
            groupes[-1].append(h)
        else:
            groupes.append([h])
    regions = []
    for g in groupes:
        ins = set()
        for h in g:
            ins |= par_hz[h]
        if len(ins) > 1 and len(g) > 1:
            regions.append((g[0], g[-1], sorted(ins)))
    return unissons, regions

def fig_convergences(INST, FORM, chemin):
    porteurs = [i for i in INST if any(f['id'] == i['id'] and f['statut']
                                       in ('route A', '1 corpus') for f in FORM)]
    # bois en haut : l'ordre de la partition, lu de haut en bas
    porteurs.sort(key=lambda i: (['bois', 'cuivres', 'cordes'].index(i['famille']),
                                 i['ordre']), reverse=True)
    unissons, regions = zones_convergence(FORM)
    note_de = {int(f['hz']): f['note'] for f in FORM}
    fig, ax = plt.subplots(figsize=(8.0, 0.33 * len(porteurs) + 2.1))
    haut = len(porteurs) - 0.45

    for a, b, ins in regions:
        ax.axvspan(a / 2 ** (0.5 / 12), b * 2 ** (0.5 / 12), color='#e8b64c',
                   alpha=.18, lw=0, zorder=0)
    for k, (a, b, ins) in enumerate(unissons):
        ax.axvspan(a / 2 ** (0.5 / 12), a * 2 ** (0.5 / 12), color='#e8b64c',
                   alpha=.45, lw=0, zorder=0)
        ax.annotate(note_de.get(a, str(a)), (a, haut + 0.1 + 0.55 * (k % 2)),
                    ha='center', va='bottom', fontsize=7.5, color='#8a6a1a')

    for y, i in enumerate(porteurs):
        fs = [f for f in FORM if f['id'] == i['id']
              and f['statut'] in ('route A', '1 corpus')]
        col = COULEUR_FAMILLE[i['famille']]
        ax.axhline(y, color='#f0f0f0', lw=0.6, zorder=0)
        for f in fs:
            A = f['statut'] == 'route A'
            ax.scatter(f['hz'], y, s=68 if A else 40, color=col if A else 'none',
                       edgecolor=col, linewidth=1.3, zorder=3)
    # séparateurs de famille
    for y in range(1, len(porteurs)):
        if porteurs[y]['famille'] != porteurs[y - 1]['famille']:
            ax.axhline(y - 0.5, color='#bbb', lw=0.8, ls=':', zorder=1)
    ax.set_yticks(range(len(porteurs)))
    ax.set_yticklabels([i['instrument'] for i in porteurs], fontsize=8)
    ax.set_ylim(-0.8, haut + 1.05)
    ax.set_xscale('log')
    ax.set_xlim(380, 5600)
    ax.set_xticks([400, 600, 900, 1400, 2000, 3000, 4500])
    ax.set_xticklabels(['400', '600', '900', '1 400', '2 000', '3 000', '4 500'])
    ax.set_xlabel('fréquence (Hz, échelle logarithmique)')
    ax.grid(axis='y', visible=False)
    nA = sum(1 for f in FORM if f['statut'] == 'route A')
    nI = len({f['instrument'] for f in FORM if f['statut'] == 'route A'})
    titre(ax, 'Les formants retenus et leurs convergences',
          '{} formants par route A sur {} instruments · {} retenus sur un seul corpus'
          .format(nA, nI, sum(1 for f in FORM if f['statut'] == '1 corpus')))
    ax.legend(handles=[
        Line2D([], [], marker='o', ls='', color='#555', label='route A (deux provenances)'),
        Line2D([], [], marker='o', ls='', mfc='none', color='#555', label='un seul corpus'),
        Rectangle((0, 0), 1, 1, color='#e8b64c', alpha=.45, label='unisson formantique'),
        Rectangle((0, 0), 1, 1, color='#e8b64c', alpha=.18, label='région partagée')],
        loc='upper center', bbox_to_anchor=(0.5, 0.0),
        bbox_transform=fig.transFigure, ncol=4, fontsize=7.5)
    fig.savefig(chemin)
    plt.close(fig)
    return True

# ═════════════════════ figure 4 : rho par corpus ═══════════════════════════
def fig_rho(chemin):
    p = os.path.join(LARG, 'Comparaison', 'rho_par_corpus.csv')
    L = list(csv.DictReader(open(p, encoding='utf-8')))
    for r in L:
        r['_e'] = float(r['etendue']) if r['etendue'].strip() else 0.0
    L.sort(key=lambda r: r['_e'])
    fig, ax = plt.subplots(figsize=(6.6, 0.30 * len(L) + 1.4))
    for y, r in enumerate(L):
        vs = [(c, float(r['rho_' + c])) for c in CORPUS if r.get('rho_' + c, '').strip()]
        if len(vs) > 1:
            ax.plot([min(v for _, v in vs), max(v for _, v in vs)], [y, y],
                    color='#ccc', lw=2.2, solid_capstyle='round', zorder=1)
        for c, v in vs:
            ax.scatter(v, y, s=34, color=COULEUR_CORPUS[c], zorder=3)
    ax.set_yticks(range(len(L)))
    ax.set_yticklabels(['{}  {}'.format(r['instrument'],
                                        ('[' + r['classe_fiche'] + ']') if r['classe_fiche'].strip() else '')
                        for r in L], fontsize=8)
    ax.set_ylim(-0.8, len(L) - 0.2)
    ax.set_xlim(0, 1.02)
    ax.set_xlabel('ρ — corrélation du spectre moyen entre notes voisines')
    ax.grid(axis='y', visible=False)
    titre(ax, 'Stabilité du spectre selon le corpus',
          'le plus instable en haut · entre crochets, la classe de la fiche')
    legende_bas(fig, [Line2D([], [], marker='o', ls='', color=COULEUR_CORPUS[c], label=c)
                      for c in CORPUS], 5)
    fig.savefig(chemin)
    plt.close(fig)
    return True

# ═════════════════ figure 5 : pentes spectrales ════════════════════════════
def fig_pentes(INST, chemin):
    L = [i for i in INST if i['pente_db_octave'] is not None]
    L.sort(key=lambda i: (['bois', 'cuivres', 'cordes'].index(i['famille']), i['ordre']))
    fig, ax = plt.subplots(figsize=(7.4, 3.3))
    xs = range(len(L))
    for x, i in zip(xs, L):
        ax.errorbar(x, i['pente_db_octave'], yerr=i['pente_sd'] or 0,
                    marker='o', ms=5, color=COULEUR_FAMILLE[i['famille']],
                    lw=1, capsize=2.5, alpha=.9)
    ax.set_xticks(list(xs))
    ax.set_xticklabels([i['instrument'] for i in L], rotation=55, ha='right', fontsize=7.5)
    ax.set_ylabel('pente (dB par octave)')
    titre(ax, 'Décroissance spectrale par instrument',
          'barres : écart-type entre notes — les cordes varient moins que les vents')
    ax.grid(axis='x', visible=False)
    legende_bas(fig, [Line2D([], [], marker='o', ls='', color=COULEUR_FAMILLE[f], label=f)
                      for f in ('bois', 'cuivres', 'cordes')], 3)
    fig.savefig(chemin)
    plt.close(fig)
    return True

# ═════════════ figure 6 : balance impair-pair (tuyau cylindrique) ══════════
def fig_balance(INST, chemin):
    L = [i for i in INST if i['balance_impair_pair_db'] is not None]
    L.sort(key=lambda i: i['balance_impair_pair_db'])
    cyl = lambda i: 'larinette' in i['instrument']
    fig, ax = plt.subplots(figsize=(6.6, 0.28 * len(L) + 1.3))
    for y, i in enumerate(L):
        ax.barh(y, i['balance_impair_pair_db'], height=.62,
                color='#b3522c' if cyl(i) else '#c9c9c9',
                edgecolor='none', zorder=2)
    ax.axvline(0, color='#555', lw=0.8)
    ax.set_yticks(range(len(L)))
    ax.set_yticklabels([i['instrument'] for i in L], fontsize=8)
    ax.set_ylim(-0.8, len(L) - 0.2)
    ax.set_xlabel('balance impair − pair (dB)')
    ax.grid(axis='y', visible=False)
    titre(ax, 'Rangs impairs contre rangs pairs',
          'en rouge : la famille de la clarinette — les seules valeurs positives')
    fig.savefig(chemin)
    plt.close(fig)
    return True

# ═══════════════ figure 7 : largeurs confrontées (LPC / enveloppe) ═════════
def fig_largeurs(chemin):
    p = os.path.join(LARG, 'Comparaison', 'confrontation_lpc_enveloppe.csv')
    L = list(csv.DictReader(open(p, encoding='utf-8')))
    for r in L:
        for k in ('formant_hz', 'lpc_v7', 'env_v7', 'Q_env_v7'):
            r[k] = float(r[k]) if r.get(k, '').strip() else None
    L = [r for r in L if r['lpc_v7'] and r['env_v7']]
    L.sort(key=lambda r: r['Q_env_v7'] or 0)        # Q élevé (étroit) en haut
    fig, ax = plt.subplots(figsize=(6.8, 0.30 * len(L) + 1.5))
    for y, r in enumerate(L):
        ax.plot([r['lpc_v7'], r['env_v7']], [y, y], color='#ddd', lw=2.2,
                solid_capstyle='round', zorder=1)
        ax.scatter(r['lpc_v7'], y, s=36, color='#7a4b9c', zorder=3)
        ax.scatter(r['env_v7'], y, s=36, color='#2e7d4f', zorder=3)
        if r['Q_env_v7']:
            ax.annotate('Q {:.1f}'.format(r['Q_env_v7']),
                        (max(r['lpc_v7'], r['env_v7']), y), xytext=(7, -3),
                        textcoords='offset points', fontsize=7, color='#777')
    ax.set_yticks(range(len(L)))
    ax.set_yticklabels(['{} · {:.0f} Hz'.format(r['instrument'], r['formant_hz'])
                        for r in L], fontsize=8)
    ax.set_ylim(-0.8, len(L) - 0.2)
    ax.set_xscale('log')
    ax.set_xlabel('largeur à −3 dB (Hz, échelle logarithmique)')
    ax.grid(axis='y', visible=False)
    titre(ax, 'Deux mesures de la même largeur',
          'les plus étroites en haut · là où elles divergent, c\'est le modèle qui parle')
    legende_bas(fig, [Line2D([], [], marker='o', ls='', color='#7a4b9c', label='pôle LPC'),
                      Line2D([], [], marker='o', ls='', color='#2e7d4f',
                             label='enveloppe vraie, −3 dB')], 2)
    fig.savefig(chemin)
    plt.close(fig)
    return True

# ════════════════════════════════ pilote ═══════════════════════════════════
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--only', default='')
    ap.add_argument('--liste', action='store_true')
    a = ap.parse_args()
    os.makedirs(SORTIE, exist_ok=True)

    INST = charger_instruments()
    FORM = charger_formants()
    NOMS = [i['instrument'] for i in INST]
    ENV, PERDUS = charger_enveloppes(NOMS)
    CARTES = charger_cartes(NOMS)

    par_inst = defaultdict(list)
    for f in FORM:
        if f['statut'] in ('route A', '1 corpus'):
            par_inst[f['id']].append(f)

    if a.liste:
        print('{:26s} {:34s} {:18s} {}'.format('fiche', 'corpus (composite)',
                                               'corpus (voyelles)', 'formants'))
        for i in INST:
            cs = [c for c in CORPUS if i['instrument'] in CARTES.get(c, {})]
            print('{:26s} {:34s} {:18s} {}'.format(
                i['instrument'][:25],
                ','.join(sorted(ENV.get(i['instrument'], {}),
                                key=CORPUS.index)) or '— AUCUN',
                ','.join(cs) or '—',
                len(par_inst.get(i['id'], []))))
        print('\nnoms de données non appariés :',
              dict(PERDUS) if PERDUS else 'aucun')
        return

    sel = [s.strip() for s in a.only.split(',') if s.strip()]
    def veut(genre, ident=None):
        if not sel:
            return True
        return any(s == genre or s == '{}:{}'.format(genre, ident) for s in sel)

    faits = []
    manifeste = {}
    for i in INST:
        fiche = manifeste.setdefault(i['id'], {})
        if veut('composite', i['id']):
            e = ENV.get(i['instrument'])
            if e and fig_composite(i, e, par_inst.get(i['id'], []),
                                   os.path.join(SORTIE, 'composite-%s.png' % i['id'])):
                faits.append('composite-%s.png' % i['id'])
                fiche['composite'] = {'fichier': 'composite-%s.png' % i['id'],
                                      'corpus': sorted(e, key=CORPUS.index)}
            else:
                fiche['composite'] = {'absent': 'aucune enveloppe composite'}
        if veut('vocalique', i['id']):
            cpc = {c: CARTES[c].get(i['instrument'], []) for c in CORPUS if c in CARTES}
            if fig_vocalique(i, cpc, os.path.join(SORTIE, 'vocalique-%s.png' % i['id'])):
                faits.append('vocalique-%s.png' % i['id'])
                fiche['vocalique'] = {'fichier': 'vocalique-%s.png' % i['id'],
                                      'corpus': [c for c in CORPUS if cpc.get(c)]}
            else:
                fiche['vocalique'] = {'absent': raison_absence(
                    [k for c in CORPUS for k in cpc.get(c, [])])}

    if veut('synthese'):
        for nom, fn in [('convergences', lambda p: fig_convergences(INST, FORM, p)),
                        ('rho-corpus', fig_rho),
                        ('pentes', lambda p: fig_pentes(INST, p)),
                        ('balance-impair-pair', lambda p: fig_balance(INST, p)),
                        ('largeurs', fig_largeurs)]:
            p = os.path.join(SORTIE, 'synthese-%s.png' % nom)
            try:
                if fn(p):
                    faits.append('synthese-%s.png' % nom)
            except Exception as ex:
                print('  !! synthese-%s : %s: %s' % (nom, type(ex).__name__, ex))

    if not sel:
        # Le manifeste dit au constructeur du site quelle figure existe, et quand
        # elle n'existe pas, POURQUOI — pour qu'il écrive la raison au lieu de
        # laisser une image manquante.
        with open(os.path.join(SORTIE, 'figures.json'), 'w', encoding='utf-8') as f:
            json.dump(manifeste, f, ensure_ascii=False, indent=1, sort_keys=True)

    print('%d figures écrites dans %s' % (len(faits), SORTIE))
    manquantes = [(k, v['vocalique']['absent']) for k, v in sorted(manifeste.items())
                  if 'absent' in v.get('vocalique', {})]
    if manquantes:
        print('\nsans carte vocalique, et pourquoi :')
        for k, r in manquantes:
            print('   {:26s} {}'.format(k, r))

if __name__ == '__main__':
    main()
