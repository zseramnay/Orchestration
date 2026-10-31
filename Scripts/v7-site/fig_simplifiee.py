# -*- coding: utf-8 -*-
"""
fig_simplifiee.py — cartes vocaliques simplifiées, pour l'enseignement.

Même lecture que les cartes spectrales vocaliques de la V1 (et que Meyer, Reuter,
McCarty) : un axe des fréquences en Hz, les six zones vocaliques de la V1 en fond,
et pour chaque registre une zone principale et une zone secondaire.

AUCUNE MESURE NOUVELLE. Les valeurs sont celles de la carte vocalique (méthode 12,
SOL2020 + Yan_Adds, `Mesures/SOL/carte-vocalique-SOL-*.json`) : F1 et F2 sont les
deux premiers maxima du motif d'excitation auditive, la largeur des barres est
l'écart interquartile à l'intérieur de la cellule, reconverti de Bark en Hz.
Seule la présentation change.

Choix de présentation
---------------------
  · bois et cordes : les dynamiques sont réunies (médiane pondérée par n) —
    leur effet y est faible ;
  · cuivres : deux lignes par registre, doux et fort — la dynamique y déplace
    F1 de près d'une octave, une moyenne n'aurait pas de sens ;
  · barre pleine   : zone stable — elle tient d'un registre au registre voisin
    (même nuance) à 3 demi-tons près, alors que les fondamentales montent
    d'environ une octave ; ou résonance établie par le test de pente ;
  · barre hachurée : la zone suit la note (marque `F1=f0` de la carte ; pour
    les cellules « 1 max », où la carte ne teste pas f0, F1 dans la tessiture
    du registre à un demi-ton près ; ou harmonique établi par le test de pente) ;
  · contour seul   : non tranché ;
  · VERDICTS : les conclusions de _RESONANCE-DE-CAISSE.md priment sur la règle ;
  · cuivres à l'aigu et au suraigu : toujours « suit la note » (borne de C7) ;
  · losange / tireté : formant établi par route A (`formants-v7.csv`) ;
  · largeur affichée plafonnée à 3 Bark ; au-delà le registre est dit instable.

Les noms de zone (u, o, å, a, e, i) sont ceux de la V1 : une voyelle par zone
d'énergie, à la manière de Meyer. Ils ne coïncident pas toujours avec les
voyelles françaises de la carte du site, qui se lisent sur le couple (F1, F2).

Sortie : Site-v2/figures/simple-<id>.png et simple-synthese-<famille>.png

Usage : python3 fig_simplifiee.py
"""
import os, re, csv, json, glob, math
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, Patch
from matplotlib.lines import Line2D
from matplotlib.transforms import blended_transform_factory

RACINE = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
ETUDE = os.path.join(RACINE, 'Etude-Formants')
SORTIE = os.path.join(ETUDE, 'Site-v2', 'figures')

# ─── zones vocaliques de la V1 (Scripts/v6-html-docx/common.py) ─────────────
ZONES = [(100, 400, '#DCEEFB', 'u (oo)', 'Profondeur'),
         (400, 600, '#D5ECD5', 'o (oh)', 'Plénitude'),
         (600, 800, '#FDE8CE', 'å (aw)', 'Transition'),
         (800, 1250, '#F8D5D5', 'a (ah)', 'Puissance'),
         (1250, 2600, '#E8D5F0', 'e (eh)', 'Clarté'),
         (2600, 6000, '#FFF8D0', 'i (ee)', 'Brillance')]

def zone(f):
    for a, b, c, l, d in ZONES:
        if f < b:
            return l.split()[0]
    return 'i'

COUL = {'bois': ('#2E7D32', '#81C784'), 'cuivres': ('#B3522C', '#E8A07E'),
        'cordes': ('#1F5C8B', '#84B4DA')}

# ─── Bark ───────────────────────────────────────────────────────────────────
def bark(f):
    return 13 * np.arctan(0.00076 * f) + 3.5 * np.arctan((f / 7500.0) ** 2)
_FG = np.logspace(math.log10(20), math.log10(16000), 4000)
_ZF = bark(_FG)
def hz(z):
    return float(np.interp(z, _ZF, _FG))

# ─── instruments : (nom dans la carte, clé registres.md, id, famille) ──────
INSTR = [
    ('Petite flûte', 'piccolo', 'petite-flute', 'bois'),
    ('Flûte', 'flûte', 'flute', 'bois'),
    ('Flûte basse', 'flûte basse', 'flute-basse', 'bois'),
    ('Flûte contrebasse', 'flûte contrebasse', 'flute-contrebasse', 'bois'),
    ('Hautbois', 'hautbois', 'hautbois', 'bois'),
    ('Cor anglais', 'cor anglais', 'cor-anglais', 'bois'),
    ('Basson', 'basson', 'basson', 'bois'),
    ('Contrebasson', 'contrebasson', 'contrebasson', 'bois'),
    ('Clarinette en Mib', 'clarinette en mib', 'clarinette-mib', 'bois'),
    ('Clarinette en Sib', 'clarinette', 'clarinette-sib', 'bois'),
    ('Clarinette basse', 'clarinette basse', 'clarinette-basse', 'bois'),
    ('Clarinette contrebasse', 'clarinette contrebasse', 'clarinette-contrebasse', 'bois'),
    ('Cor', 'cor', 'cor', 'cuivres'),
    ('Trompette', 'trompette', 'trompette', 'cuivres'),
    ('Trombone ténor', 'trombones', 'trombone-tenor', 'cuivres'),
    ('Trombone basse', 'trombones', 'trombone-basse', 'cuivres'),
    ('Tuba basse', 'tubas', 'tuba-basse', 'cuivres'),
    ('Tuba contrebasse', 'tubas', 'tuba-contrebasse', 'cuivres'),
    ('Violon', 'violon', 'violon', 'cordes'),
    ('Alto', 'alto', 'alto', 'cordes'),
    ('Violoncelle', 'violoncelle', 'violoncelle', 'cordes'),
    ('Contrebasse', 'contrebasse', 'contrebasse', 'cordes'),
]
NOM_AFFICHE = {'Cor': 'Cor en Fa', 'Trompette': 'Trompette en Ut',
               'Clarinette basse': 'Clarinette basse en Sib',
               'Clarinette contrebasse': 'Clarinette contrebasse en Sib'}

# ─── tessitures par registre (Data/registres.md) ───────────────────────────
_SEMI = {'C': 0, 'D': 2, 'E': 4, 'F': 5, 'G': 7, 'A': 9, 'B': 11}
def note_hz(s):
    m = re.match(r'([A-G])(#?)(-?\d)', s)
    midi = 12 * (int(m[3]) + 1) + _SEMI[m[1]] + (1 if m[2] else 0)
    return 440.0 * 2 ** ((midi - 69) / 12)

def lire_registres():
    out = {}
    for ligne in open(os.path.join(RACINE, 'Data', 'registres.md'), encoding='utf-8'):
        if not ligne.startswith('* ') or ':' not in ligne:
            continue
        nom, reste = ligne[2:].split(':', 1)
        d = {}
        for p in [x.strip() for x in reste.split(',')]:
            p = re.sub(r'^registre\s+', '', p)
            m = re.match(r'([a-zàâçéèêëîïôûù ]+?)\s+en dessous de\s+([A-G]#?\d)', p)
            if m:
                f = note_hz(m[2]); d[m[1].strip()] = (f / 2, f); continue
            m = re.match(r'([a-zàâçéèêëîïôûù ]+?)\s+au-dessus de\s+([A-G]#?\d)', p)
            if m:
                f = note_hz(m[2]); d[m[1].strip()] = (f, 2 * f); continue
            m = re.match(r'([a-zàâçéèêëîïôûù ]+?)\s+([A-G]#?\d)-([A-G]#?\d)', p)
            if m:
                d[m[1].strip()] = (note_hz(m[2]), note_hz(m[3])); continue
            m = re.match(r'([a-zàâçéèêëîïôûù ]+?)\s+(?:à partir de\s+)?([A-G]#?\d)(?:\s+et)?\s*(?:au-dessus)?$', p)
            if m:
                f = note_hz(m[2]); d[m[1].strip()] = (f, 2 * f); continue
            m = re.match(r'([a-zàâçéèêëîïôûù ]+?)\s+au-dessus', p)
            if m and d:
                f = max(v[1] for v in d.values()); d[m[1].strip()] = (f, 2 * f)
        out[nom.strip()] = d
    return out

def tessitures(R, cle):
    if cle in R:
        return R[cle]
    for k in R:
        if k.startswith(cle + ' '):
            return R[k]
    return {}

# ─── données ───────────────────────────────────────────────────────────────
def charger_carte():
    ps = sorted(glob.glob(os.path.join(ETUDE, 'Mesures', 'SOL', 'carte-vocalique-SOL-*.json')))
    p = ps[-1]                       # la plus récente (dates ISO dans le nom)
    return json.load(open(p, encoding='utf-8')), os.path.basename(p)

def charger_etablis():
    E = {}
    for r in csv.DictReader(open(os.path.join(ETUDE, 'Resultats-v7', 'formants-v7.csv'),
                                 encoding='utf-8')):
        if r['statut'] == 'route A':
            E.setdefault(r['id'], []).append(float(r['hz']))
    return E

PLAFOND_IQR = 3.0
DEMI_TON = 2 ** (1 / 12)

def mediane_ponderee(paires):
    paires = sorted(p for p in paires if p[0] is not None)
    if not paires:
        return None
    tot = sum(w for _, w in paires); acc = 0
    for x, w in paires:
        acc += w
        if acc >= tot / 2:
            return x

def lignes(cells, famille, tess):
    """Une ligne par registre (bois, cordes) ou par registre × {doux, fort} (cuivres)."""
    ordre = []
    for c in cells:
        if c['registre'] not in ordre and c['registre'] != 'hors registre':
            ordre.append(c['registre'])
    groupes = []
    for r in ordre:
        cs = [c for c in cells if c['registre'] == r and c['F1_bark'] is not None]
        if famille == 'cuivres':
            for dy, ab in (('doux', 'p'), ('fort', 'f')):
                g = [c for c in cs if c['dynamique'] == dy]
                if not g:
                    g = [c for c in cs if c['dynamique'] == 'toutes']
                    ab = ''
                    if not g or any(x[0] == r for x in groupes):
                        continue
                groupes.append((r, ab, g))
        elif cs:
            groupes.append((r, '', cs))
    out = []
    for r, ab, cs in groupes:
        n = sum(c['n'] for c in cs)
        z1 = mediane_ponderee([(c['F1_bark'], c['n']) for c in cs])
        i1 = mediane_ponderee([(c['IQR1'], c['n']) for c in cs])
        a2 = [c for c in cs if c['F2_bark'] is not None]
        if sum(c['n'] for c in a2) > n / 2:
            z2 = mediane_ponderee([(c['F2_bark'], c['n']) for c in a2])
            i2 = mediane_ponderee([(c['IQR2'], c['n']) for c in a2])
        else:
            z2 = i2 = None
        tr = tess.get(r)
        f1 = hz(z1)
        def suit(c):
            if 'F1=f0' in c['marques']:
                return True
            # « 1 max » : la carte n'a pas testé f0 — on compare à la tessiture
            if '1 max' in c['marques'] and tr:
                return tr[0] / DEMI_TON <= c['F1_hz'] <= tr[1] * DEMI_TON
            return False
        nf0 = sum(c['n'] for c in cs if suit(c))
        nflag = sum(c['n'] for c in cs if 'F1=f0' in c['marques'])
        instable = (i1 or 0) > 2.0 or (i2 or 0) > PLAFOND_IQR
        out.append(dict(reg=r, dyn=ab, tr=tr, z1=z1, i1=min(i1 or 0, PLAFOND_IQR),
                        z2=z2, i2=None if i2 is None else min(i2, PLAFOND_IQR),
                        f1=f1, f2=None if z2 is None else hz(z2),
                        note=nflag > n / 2, note_proxy=nf0 > n / 2,
                        instable=instable))
    return out

# Verdicts établis par l'étude, qui priment sur la règle générale.
# Source : _RESONANCE-DE-CAISSE.md (test de pente note par note, quatre corpus).
VERDICTS = {
    ('Alto', 'grave'): 'stable',          # 374–395 Hz, pente médiane +0,01
    ('Alto', 'médium'): 'note',           # pentes +0,79 à +1,11 : harmonique
    ('Violoncelle', 'médium'): 'stable',  # 209–223 Hz, pente médiane −0,04
    ('Violoncelle', 'grave'): 'indet',    # résonance vue au médium seulement
    ('Violon', 'grave'): 'indet',         # corpus dispersés sur 6,6 demi-tons
    ('Contrebasse', 'grave'): 'indet',    # 34 % des notes sous le plancher de 50 Hz
}
TOL_STABLE = 3.0   # demi-tons

def famille_de(nom):
    return next(f for n, c, i, f in INSTR if n == nom)

def etats(nom, L):
    """'note' : la zone suit la hauteur ; 'stable' : elle tient d'un registre au
    registre voisin (même nuance) à TOL_STABLE demi-tons près alors que les
    fondamentales montent d'environ une octave, ou une résonance est établie ;
    'indet' : ni l'un ni l'autre n'est montré."""
    for k, r in enumerate(L):
        v = VERDICTS.get((nom, r['reg']))
        if v:
            r['etat'] = v; continue
        # Borne de lecture des cuivres (tâche C7, _REGISTRES-PEDALE.md) : à l'aigu
        # et au suraigu, F1/f0 vaut environ 1 — la mesure y est la note.
        if famille_de(nom) == 'cuivres' and r['reg'] in ('aigu', 'suraigu'):
            r['etat'] = 'note'; continue
        if r['note']:                 # marque F1=f0 de la carte : décisive
            r['etat'] = 'note'; continue
        regs = []
        for x in L:
            if x['reg'] not in regs:
                regs.append(x['reg'])
        i = regs.index(r['reg'])
        voisins = [x for x in L if x['dyn'] == r['dyn']
                   and x['reg'] in regs[max(0, i - 1):i] + regs[i + 1:i + 2]]
        ok = any(abs(12 * math.log2(r['f1'] / x['f1'])) <= TOL_STABLE for x in voisins)
        # le repli par la tessiture (cellules « 1 max ») ne passe qu'après le
        # test de stabilité : une zone qui tient d'un registre à l'autre peut
        # tomber par hasard dans la tessiture sans la suivre
        r['etat'] = 'stable' if ok else ('note' if r['note_proxy'] else 'indet')
    return L

# ─── dessin ────────────────────────────────────────────────────────────────
XMIN, XMAX = 80, 6000
TICKS = [100, 150, 200, 300, 400, 500, 600, 800, 1000, 1500, 2000, 3000, 4000, 5000]

def fond(ax, fig):
    for a, b, c, l, d in ZONES:
        ax.axvspan(a, b, color=c, alpha=.8, zorder=0, lw=0)
    tr = blended_transform_factory(ax.transData, ax.transAxes)
    for a, b, c, l, d in ZONES:
        # étiquettes au-dessus du cadre, avec un écart fixe en points :
        # elles ne touchent plus la bordure supérieure
        ax.annotate(l + '\n' + d, (math.sqrt(a * b), 1.0), xycoords=tr,
                    xytext=(0, 5), textcoords='offset points',
                    ha='center', va='bottom', fontsize=8.5, color='#555',
                    fontweight='bold', linespacing=1.15)
    ax.set_xscale('log'); ax.set_xlim(XMIN, XMAX)
    ax.set_xticks(TICKS); ax.set_xticklabels([str(t) for t in TICKS], fontsize=8)
    ax.minorticks_off()
    ax.set_xlabel('Fréquence (Hz) — échelle logarithmique', fontsize=9.5)
    ax.grid(axis='x', color='white', lw=.6, alpha=.7, zorder=0)

def barre(ax, y, z, iq, h, plein, col, alpha=1.0):
    lo, hi = hz(z - iq / 2), hz(z + iq / 2)
    if plein == 'indet':
        ax.add_patch(Rectangle((lo, y - h / 2), hi - lo, h, fc='white', ec=col, lw=1.6,
                               alpha=alpha, zorder=3))
        return lo, hi
    if plein in (True, 'stable'):
        ax.add_patch(Rectangle((lo, y - h / 2), hi - lo, h, fc=col, ec=col, lw=1,
                               alpha=alpha, zorder=3))
    else:
        ax.add_patch(Rectangle((lo, y - h / 2), hi - lo, h, fc='white', ec=col, lw=1.2,
                               hatch='////', alpha=alpha, zorder=3))
    return lo, hi

def legende(fig, fonce, clair, cuivres, ncol=3):
    h = [Patch(fc=fonce, ec=fonce, label='zone principale stable (indépendante de la note)'),
         Patch(fc='white', ec=fonce, hatch='////', label='zone principale qui suit la note jouée'),
         Patch(fc='white', ec=fonce, lw=1.6, label='zone principale non tranchée'),
         Patch(fc=clair, label='zone secondaire'),
         Line2D([], [], color='#777', lw=3, alpha=.45, label='tessiture du registre (fondamentales)'),
         Line2D([], [], marker='D', ls='', mfc='#222', mec='white', ms=6,
                label='formant fixe établi (deux corpus)')]
    if cuivres:
        h.append(Line2D([], [], ls='', label='p = doux · f = fort'))
    fig.legend(handles=h, loc='lower center', ncol=ncol, fontsize=8, frameon=False)

def titre_fig(ax, texte, col):
    ax.set_title(texte, fontsize=12.5, color=col, fontweight='bold', pad=40)

def etiquette(r):
    s = r['reg'] + (' · ' + r['dyn'] if r['dyn'] else '')
    return s

def synthese(famille, donnees, etablis, chemin, source):
    fonce, clair = COUL[famille]
    n = sum(len(L) + 1 for *_, L in donnees)
    fig, ax = plt.subplots(figsize=(11, 0.25 * n + 2.0))
    y = 0; yt = []; yl = []
    for nom, ident, L in donnees:
        ax.text(XMIN * 1.02, y, NOM_AFFICHE.get(nom, nom), fontsize=9.5, fontweight='bold',
                va='center', ha='left', zorder=5)
        for f in etablis.get(ident, []):
            ax.plot(f, y, marker='D', ms=6, mfc='#222', mec='white', zorder=6)
        y += 1
        for r in L:
            if r['tr']:
                ax.plot(r['tr'], [y, y], color='#777', lw=3, alpha=.45,
                        solid_capstyle='butt', zorder=2)
            barre(ax, y, r['z1'], r['i1'], .62, r['etat'], fonce)
            if r['z2'] is not None:
                barre(ax, y, r['z2'], r['i2'], .28, True, clair, .95)
            yt.append(y); yl.append(etiquette(r) + (' *' if r['instable'] else ''))
            y += 1
        y += .35
    ax.set_ylim(y - .3, -1.0)
    ax.set_yticks(yt); ax.set_yticklabels(yl, fontsize=7.5, color='#444')
    ax.yaxis.tick_right(); ax.tick_params(axis='y', length=0)
    fond(ax, fig)
    titre_fig(ax, 'Les {} dans l\'espace vocalique — carte simplifiée par registre'.format(famille),
              fonce)
    legende(fig, fonce, clair, famille == 'cuivres')
    fig.text(0.995, 0.997, '* registre instable (dispersion > 2 Bark) · source : {}'.format(source),
             ha='right', va='top', fontsize=6.5, color='#888')
    fig.tight_layout(rect=(0, 0.055, 1, 1))
    fig.savefig(chemin, dpi=150); plt.close(fig)

def fiche(nom, ident, famille, L, etablis, chemin, source):
    fonce, clair = COUL[famille]
    pas = 1.15
    fig, ax = plt.subplots(figsize=(11, max(3.6, pas * len(L) + 2.4)))
    y = 0; yt = []; yl = []
    for r in L:
        if r['tr']:
            ax.plot(r['tr'], [y + .37, y + .37], color='#777', lw=4, alpha=.45,
                    solid_capstyle='butt', zorder=2)
        lo, hi = barre(ax, y, r['z1'], r['i1'], .55, r['etat'], fonce)
        t = '{} Hz · {}'.format(int(round(r['f1'], -1)), zone(r['f1']))
        if r['etat'] == 'note':
            t += '  (suit la note)'
        elif r['etat'] == 'indet':
            t += '  (?)'
        x2 = None
        if r['z2'] is not None:
            lo2, hi2 = barre(ax, y, r['z2'], r['i2'], .25, True, clair)
            x2 = (lo2, hi2)
        # l'étiquette principale ne doit pas recouvrir la zone secondaire
        if x2 and x2[0] < hi * 4.2 and x2[0] > hi:
            ax.text(lo / 1.04, y, t, va='center', ha='right', fontsize=9, fontweight='bold',
                    color=fonce, zorder=6)
        else:
            ax.text(hi * 1.04, y, t, va='center', fontsize=9, fontweight='bold',
                    color=fonce, zorder=6)
        if x2:
            ax.text(x2[1] * 1.03, y, '{} Hz · {}'.format(int(round(r['f2'], -1)), zone(r['f2'])),
                    ha='left', va='center', fontsize=7.8, color=fonce, zorder=6)
        lab = etiquette(r)
        if r['instable']:
            lab += '\n(instable)'
        if r['tr']:
            lab += '\n{:.0f}–{:.0f} Hz'.format(*r['tr'])
        yt.append(y); yl.append(lab)
        y += pas
    for f in etablis.get(ident, []):
        ax.axvline(f, color='#222', lw=1.1, ls='--', zorder=1)
        ax.text(f, -.62, 'formant fixe {} Hz'.format(int(f)), ha='center', va='center',
                fontsize=7.5, color='#222', bbox=dict(fc='white', ec='none', pad=1), zorder=6)
    ax.set_ylim(y - .3, -1.0)
    ax.set_yticks(yt); ax.set_yticklabels(yl, fontsize=9)
    fond(ax, fig)
    titre_fig(ax, '{} — carte vocalique simplifiée par registre'.format(NOM_AFFICHE.get(nom, nom)),
              fonce)
    h = [Patch(fc=fonce, ec=fonce, label='zone principale stable'),
         Patch(fc='white', ec=fonce, hatch='////', label='suit la note'),
         Patch(fc='white', ec=fonce, lw=1.6, label='non tranchée'),
         Patch(fc=clair, label='zone secondaire'),
         Line2D([], [], color='#777', lw=4, alpha=.45, label='tessiture du registre')]
    if famille == 'cuivres':
        h.append(Line2D([], [], ls='', label='p = doux · f = fort'))
    fig.legend(handles=h, loc='lower center', ncol=len(h), fontsize=8.5, frameon=False)
    fig.text(0.995, 0.997, 'source : ' + source, ha='right', va='top', fontsize=6.5, color='#888')
    fig.tight_layout(rect=(0, 0.07, 1, 1))
    fig.savefig(chemin, dpi=150); plt.close(fig)

def main():
    os.makedirs(SORTIE, exist_ok=True)
    carte, source = charger_carte()
    etablis = charger_etablis()
    R = lire_registres()
    par_famille = {}
    faits = []
    for nom, cle, ident, fam in INSTR:
        if nom not in carte:
            print('  !! absent de la carte :', nom); continue
        v = carte[nom]
        cells = v['cellules'] if isinstance(v, dict) else v
        L = etats(nom, lignes(cells, fam, tessitures(R, cle)))
        if not L:
            print('  !! aucune cellule publiable :', nom); continue
        par_famille.setdefault(fam, []).append((nom, ident, L))
        p = os.path.join(SORTIE, 'simple-%s.png' % ident)
        fiche(nom, ident, fam, L, etablis, p, source)
        faits.append(os.path.basename(p))
    for fam, d in par_famille.items():
        p = os.path.join(SORTIE, 'simple-synthese-%s.png' % fam)
        synthese(fam, d, etablis, p, source)
        faits.append(os.path.basename(p))
    print('%d figures simplifiées écrites dans %s' % (len(faits), SORTIE))

if __name__ == '__main__':
    main()
