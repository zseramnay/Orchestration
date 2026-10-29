# -*- coding: utf-8 -*-
"""
coupure_benade.py — tâche C10 : la coupure du réseau de trous latéraux, descripteur des
instruments sans formant.

Pourquoi ce descripteur
-----------------------
`00-METHODE.md` l'annonce depuis le 23/10 sans l'avoir jamais mesuré : « pour la classe B,
l'invariant pertinent n'est pas un formant mais la fréquence de coupure du réseau de trous
latéraux (Benade) et la pente spectrale ». Les flûtes n'ont aucun formant retenu sur trois
corpus ; il leur faut un chiffre positif à publier, et non seulement une absence.

Le réseau de trous latéraux d'un bois se comporte en filtre : en dessous de la coupure il
réfléchit, au-dessus il laisse passer. Les harmoniques situés sous la coupure sont donc portés
par des résonances d'air fortes ; ceux situés au-dessus ne le sont plus et l'enveloppe y tombe
plus vite. **La signature observable est un coude dans l'enveloppe spectrale** : pente douce en
dessous, pente raide au-dessus.

La mesure
---------
Note par note, sur les amplitudes harmoniques en dB contre le log de la fréquence, on ajuste un
**bâton brisé** — deux droites qui se rejoignent en un point libre. Pour un point de brisure
fixé, le modèle est linéaire ; on balaie donc la brisure sur une grille et on garde le minimum
de résidu. Le coude est retenu seulement si la pente du haut est plus raide que celle du bas
d'au moins SAUT_MIN dB par octave : sans cela, il n'y a pas de coude, seulement une droite.

Le test qui décide
------------------
Une coupure est une propriété du **tuyau**, pas de la note. Elle ne doit donc pas bouger quand
la note bouge. On reprend le test de la tâche C1 : la pente de log2(coude) contre log2(f0)
doit être nulle. Si elle vaut 1, ce qu'on mesure n'est qu'un rang harmonique déguisé.

Et un **témoin négatif** indispensable : les cordes n'ont pas de réseau de trous latéraux. Si
la méthode leur trouve une coupure aussi stable qu'aux bois, c'est qu'elle mesure autre chose.

Usage
    python3 coupure_benade.py
    python3 coupure_benade.py --instruments "Flûte,Petite flûte,Violon"
"""
import os, sys, csv, json, math, glob, argparse, statistics as st
from collections import defaultdict
import numpy as np

ICI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ICI)
import formants_lib as FL
from fenetre_redressement import NOMS

RACINE = os.path.dirname(os.path.dirname(ICI))
ETUDE = os.path.join(RACINE, 'Etude-Formants')
SOL_OUT = os.path.expanduser('~/work/out')
CORPUS = ['SOL', 'ZSamples', 'Xsamples', 'VO', 'VN']

MIN_HARM_COTE = 3        # harmoniques exigés de chaque côté du coude
SAUT_MIN = 4.0           # dB/octave : écart minimal entre les deux pentes
MIN_NOTES = 8
FAMILLE_SANS_TROUS = {'Violon', 'Alto', 'Violoncelle', 'Contrebasse',
                      'Cor en Fa', 'Trompette en Ut', 'Trombone ténor',
                      'Trombone basse', 'Tuba basse', 'Tuba contrebasse'}

def baton_brise(x, y):
    """Ajuste dB = a + b1·(x−c) sous c, a + b2·(x−c) au-dessus, c libre.
    Rend (c, b1, b2, r2) ou None. Pour c fixé le modèle est linéaire : on balaie c."""
    n = x.size
    if n < 2 * MIN_HARM_COTE + 1:
        return None
    best = None
    for i in range(MIN_HARM_COTE, n - MIN_HARM_COTE):
        c = x[i]
        g = x < c
        # colonnes : 1, (x−c)·[x<c], (x−c)·[x≥c]
        A = np.column_stack([np.ones(n), (x - c) * g, (x - c) * (~g)])
        try:
            sol, res, rk, _ = np.linalg.lstsq(A, y, rcond=None)
        except np.linalg.LinAlgError:
            continue
        if rk < 3:
            continue
        sse = float(((A @ sol - y) ** 2).sum())
        if best is None or sse < best[0]:
            best = (sse, c, float(sol[1]), float(sol[2]))
    if best is None:
        return None
    sse, c, b1, b2 = best
    sst = float(((y - y.mean()) ** 2).sum())
    r2 = 1.0 - sse / sst if sst > 0 else None
    return c, b1, b2, r2

def par_note(chemin, f0_max=None):
    """[(f0, coude_hz, pente_bas, pente_haut, r2)] — un point par note qui a un coude.

    f0_max : ne garder que les notes assez graves. Une coupure vers 2 kHz n'est
    mesurable que si **plusieurs harmoniques l'encadrent** ; une note dont la
    fondamentale est déjà au-dessus n'en dit rien, et le bâton brisé s'ajuste alors sur
    un rang harmonique fixe au lieu d'une fréquence fixe. C'est ce qui faisait « suivre
    la note » aux flûtes dans la première version."""
    out = []
    for f0, midi, dyn, sp in FL.lire_spectres(chemin):
        if not f0 or (f0_max and f0 > f0_max):
            continue
        fr, am = FL.harmoniques(sp, f0)
        if fr.size < 2 * MIN_HARM_COTE + 1:
            continue
        m = np.isfinite(am) & (fr > 0)
        r = baton_brise(np.log2(fr[m]), am[m])
        if not r:
            continue
        c, b1, b2, r2 = r
        if b1 - b2 < SAUT_MIN:       # la pente du haut doit être PLUS RAIDE (plus négative)
            continue
        out.append((f0, 2.0 ** c, b1, b2, r2))
    return out

def pente_contre_f0(pts):
    if len(pts) < MIN_NOTES:
        return None, None
    x = np.log2([p[0] for p in pts]); y = np.log2([p[1] for p in pts])
    x = x - x.mean(); yc = y - y.mean()
    d = float((x * x).sum())
    if d <= 0:
        return None, None
    a = float((x * yc).sum() / d)
    sst = float((yc * yc).sum())
    r2 = 1.0 - float(((yc - a * x) ** 2).sum()) / sst if sst > 0 else None
    return round(a, 2), (None if r2 is None else round(r2, 2))

def chemins(nom):
    out = {}
    p = os.path.join(SOL_OUT, (NOMS.get(nom) or '') + '.txt')
    if os.path.exists(p):
        out['SOL'] = p
    for c in CORPUS[1:]:
        p = os.path.join(ETUDE, 'Mesures', c, nom + '.spectres.txt')
        if os.path.exists(p):
            out[c] = p
    return out

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--instruments', default='')
    ap.add_argument('--sortie', default=os.path.join(ETUDE, 'Mesures', 'Coupure-Benade'))
    ap.add_argument('--f0-max', type=float, default=None, metavar='HZ',
                    help='ne garder que les notes dont la fondamentale est sous ce seuil')
    a = ap.parse_args()
    os.makedirs(a.sortie, exist_ok=True)
    noms = [x.strip() for x in a.instruments.split(',') if x.strip()] or sorted(NOMS)

    lignes = []
    print('{:26s} {:9s} {:>4s} {:>5s} {:>9s} {:>7s} {:>7s} {:>7s}  {}'.format(
        'instrument', 'corpus', 'n', '%', 'coude méd.', 'IQR dt', 'pente', 'r²', 'verdict'))
    for nom in noms:
        for corpus, p in chemins(nom).items():
            tot = sum(1 for f0, m, d, s in FL.lire_spectres(p)
                      if f0 and (not a.f0_max or f0 <= a.f0_max))
            pts = par_note(p, a.f0_max)
            if len(pts) < MIN_NOTES:
                continue
            co = [x[1] for x in pts]
            med = float(np.median(co))
            iqr = 12 * math.log2(np.percentile(co, 75) / max(np.percentile(co, 25), 1e-9))
            pe, r2 = pente_contre_f0(pts)
            if pe is None:
                v = 'indéterminé'
            elif abs(pe) <= 0.25 and iqr <= 12:
                v = 'COUPURE STABLE'
            elif abs(pe) >= 0.75:
                v = 'suit la note'
            else:
                v = 'intermédiaire'
            d = dict(instrument=nom, corpus=corpus, n_notes=len(pts), n_total=tot,
                     pct_avec_coude=int(round(100 * len(pts) / max(tot, 1))),
                     coude_median_hz=int(round(med)),
                     coude_q1=int(round(float(np.percentile(co, 25)))),
                     coude_q3=int(round(float(np.percentile(co, 75)))),
                     iqr_demi_tons=round(iqr, 1), pente=pe, r2=r2,
                     pente_bas=round(float(np.median([x[2] for x in pts])), 1),
                     pente_haut=round(float(np.median([x[3] for x in pts])), 1),
                     sans_trous=nom in FAMILLE_SANS_TROUS, f0_max=a.f0_max, verdict=v)
            lignes.append(d)
            print('{:26s} {:9s} {:4d} {:4d}% {:8d}Hz {:6.1f} {:>7} {:>7} {:>5.1f}/{:<5.1f} {}'
                  .format(nom[:26], corpus, len(pts), d['pct_avec_coude'],
                          d['coude_median_hz'], iqr, pe if pe is not None else '—',
                          r2 if r2 is not None else '—',
                          d['pente_bas'], d['pente_haut'], v))

    with open(os.path.join(a.sortie, 'coupure-benade.csv'), 'w', encoding='utf-8',
              newline='') as f:
        if lignes:
            w = csv.DictWriter(f, fieldnames=list(lignes[0]))
            w.writeheader(); w.writerows(lignes)
    print('\n{} mesures -> {}'.format(len(lignes), a.sortie))

if __name__ == '__main__':
    main()
