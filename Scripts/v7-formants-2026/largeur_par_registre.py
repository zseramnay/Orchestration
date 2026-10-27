# -*- coding: utf-8 -*-
"""
largeur_par_registre.py — tâche C5 : la largeur de bande par registre et par dynamique.

L'hypothèse à tester
--------------------
Le plan de reprise la pose ainsi : « chez les cuivres la localisation devrait s'améliorer en jeu
fort ». Le raisonnement physique est connu — un cuivre joué fort raidit le spectre par
non-linéarité de propagation dans la perce, l'énergie monte en fréquence et les résonances du
pavillon se détachent mieux du fond. Si c'est vrai, la largeur à −3 dB d'un formant doit se
resserrer, et la proportion de pics « ouverts » — ceux qui ne redescendent pas de 3 dB des deux
côtés, donc non localisables — doit baisser du doux vers le fort.

Les données existent déjà : `Largeurs/Enveloppe/SOL-v7/*_env.csv` porte, note par note, chaque
pic de l'enveloppe vraie avec sa largeur, sa proéminence, sa dynamique et son MIDI. Rien à
recalculer, seulement à croiser.

Ce qui est mesuré
-----------------
Pour chaque formant établi fc, on retient les pics situés à moins de deux demi-tons de fc — la
même tolérance qu'en C4 — puis on agrège par registre × dynamique :

    · part de pics OUVERTS : la mesure directe de la localisation. Un pic ouvert n'a pas de
      largeur, donc pas de Q : l'écarter en silence biaiserait la médiane vers les pics étroits.
      C'est pourquoi on le compte plutôt que de l'ignorer ;
    · largeur médiane à −3 dB, et Q = fc / largeur, sur les seuls pics fermés.

Le test de l'hypothèse se lit alors sur une seule ligne par instrument : part d'ouverts au doux
contre part d'ouverts au fort.

Garde-fou
---------
La dynamique et le registre ne sont pas indépendants dans un corpus d'échantillons : les notes
extrêmes ne sont pas toujours jouées à toutes les nuances. On rapporte donc l'effectif de chaque
cellule, et l'effet de dynamique est aussi calculé **à registre constant**, ce qui est le seul
calcul qui réponde vraiment à la question.

Usage
    python3 largeur_par_registre.py --familles cuivres
    python3 largeur_par_registre.py --tout
"""
import os, sys, csv, math, json, argparse
from collections import defaultdict
import numpy as np

ICI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ICI)
import carte_vocalique as CV
from fenetre_redressement import NOMS

RACINE = os.path.dirname(os.path.dirname(ICI))
ETUDE = os.path.join(RACINE, 'Etude-Formants')
ENV = os.path.join(ETUDE, 'Largeurs', 'Enveloppe', 'SOL-v7')
TOL_DT = 2.0

DYN = {'ppp': 'doux', 'pp': 'doux', 'p': 'doux', 'mp': 'doux',
       'mf': 'moyen', 'f': 'fort', 'ff': 'fort', 'fff': 'fort'}
ORDRE_DYN = ['doux', 'moyen', 'fort']

FAMILLE = {
    'Alto': 'cordes', 'Violon': 'cordes', 'Violoncelle': 'cordes', 'Contrebasse': 'cordes',
    'Basson': 'anches doubles', 'Hautbois': 'anches doubles',
    'Cor anglais': 'anches doubles', 'Contrebasson': 'anches doubles',
    'Clarinette en Sib': 'clarinettes', 'Clarinette en Mib': 'clarinettes',
    'Clarinette basse en Sib': 'clarinettes', 'Clarinette contrebasse en Sib': 'clarinettes',
    'Cor en Fa': 'cuivres', 'Trompette en Ut': 'cuivres', 'Trombone ténor': 'cuivres',
    'Trombone basse': 'cuivres', 'Tuba basse': 'cuivres', 'Tuba contrebasse': 'cuivres',
    'Flûte': 'flûtes', 'Petite flûte': 'flûtes', 'Flûte basse': 'flûtes',
    'Flûte contrebasse': 'flûtes',
}
CLE_REGISTRE = {
    'Alto': 'alto', 'Violon': 'violon', 'Violoncelle': 'violoncelle',
    'Contrebasse': 'contrebasse', 'Basson': 'basson', 'Hautbois': 'hautbois',
    'Cor anglais': 'cor anglais', 'Contrebasson': 'contrebasson',
    'Clarinette en Sib': 'clarinette en si bémol', 'Clarinette en Mib': 'clarinette en mi bémol',
    'Clarinette basse en Sib': 'clarinette basse',
    'Clarinette contrebasse en Sib': 'clarinette contrebasse',
    # `registres.md` écrit « trombones » au pluriel et ne distingue pas ténor et basse :
    # les deux pointent donc sur la même ligne. Sans cela ils tombaient dans un registre
    # « tous », ce qui effaçait justement la structure par registre qu'on cherche.
    'Cor en Fa': 'cor', 'Trompette en Ut': 'trompette', 'Trombone ténor': 'trombones',
    'Trombone basse': 'trombones', 'Tuba basse': 'tubas', 'Tuba contrebasse': 'tubas',
    'Flûte': 'flûte', 'Petite flûte': 'petite flûte', 'Flûte basse': 'flûte basse',
    'Flûte contrebasse': 'flûte contrebasse',
}

def resoudre(regs, cle):
    if cle in regs:
        return regs[cle]
    c = [k for k in regs if k.startswith(cle + ' ') or cle in k.split(' et ')]
    return regs[c[0]] if len(c) == 1 else None

def pics(nom):
    """[(midi, dyn, freq, bw|None, ouvert)] pour un instrument."""
    p = os.path.join(ENV, (NOMS.get(nom) or '') + '_env.csv')
    if not os.path.exists(p):
        return None
    out = []
    with open(p, encoding='utf-8') as f:
        for r in csv.DictReader(f):
            try:
                midi = int(float(r['midi'])) if r['midi'] else None
                fr = float(r['freq_hz'])
            except ValueError:
                continue
            bw = None
            try:
                bw = float(r['bw_hz']) if r['bw_hz'] else None
            except ValueError:
                bw = None
            out.append((midi, DYN.get((r['dynamic'] or '').lower()), fr, bw,
                        r['ouvert'] == '1'))
    return out

def agreger(P, fc, regs):
    """{(registre, dyn): dict} pour les pics proches de fc."""
    acc = defaultdict(lambda: {'n': 0, 'ouv': 0, 'bw': []})
    lfc = math.log2(fc)
    for midi, dyn, fr, bw, ouv in P:
        if midi is None or dyn is None:
            continue
        if abs(12 * (math.log2(max(fr, 1e-9)) - lfc)) > TOL_DT:
            continue
        reg = CV.registre_de(regs, midi) if regs else 'tous'
        d = acc[(reg, dyn)]
        d['n'] += 1
        if ouv or bw is None:
            d['ouv'] += 1
        else:
            d['bw'].append(bw)
    out = {}
    for k, d in acc.items():
        out[k] = dict(n=d['n'], pct_ouvert=int(round(100 * d['ouv'] / d['n'])),
                      bw_median=None if not d['bw'] else round(float(np.median(d['bw'])), 1),
                      Q_median=None if not d['bw'] else round(fc / float(np.median(d['bw'])), 2),
                      n_ferme=len(d['bw']))
    return out

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--familles', default='')
    ap.add_argument('--tout', action='store_true')
    ap.add_argument('--sortie', default=os.path.join(ETUDE, 'Mesures', 'Largeur-par-registre'))
    a = ap.parse_args()
    os.makedirs(a.sortie, exist_ok=True)
    fams = {x.strip() for x in a.familles.split(',') if x.strip()}
    regs_tous = CV.lire_registres(os.path.join(RACINE, 'Data', 'registres.md'))

    with open(os.path.join(ETUDE, 'Resultats-v7', 'formants-v7.csv'), encoding='utf-8') as f:
        FORM = [r for r in csv.DictReader(f) if r['statut'] in ('route A', '1 corpus')]
    par_inst = defaultdict(list)
    for r in FORM:
        par_inst[r['instrument']].append((float(r['hz']), r['statut']))

    lignes, resume = [], []
    for nom in sorted(par_inst):
        fam = FAMILLE.get(nom, '?')
        if fams and fam not in fams and not a.tout:
            continue
        P = pics(nom)
        if P is None:
            print('{:28s} — pas de largeurs d\'enveloppe'.format(nom)); continue
        regs = resoudre(regs_tous, CLE_REGISTRE.get(nom, ''))
        print('\n=== {}  ({})'.format(nom, fam))
        print('{:8s} {:10s} {:>5s} {:>9s} {:>9s} {:>7s} {:>7s}'.format(
            'formant', 'registre', 'dyn', 'n', 'ouverts', 'bw méd', 'Q méd'))
        glob_dyn = defaultdict(lambda: {'n': 0, 'ouv': 0})
        for fc, statut in sorted(par_inst[nom]):
            A = agreger(P, fc, regs)
            ordre = [l for l, lo, hi in (regs or [])] or ['tous']
            for k in sorted(A, key=lambda x: (ordre.index(x[0]) if x[0] in ordre else 99,
                                              ORDRE_DYN.index(x[1]) if x[1] in ORDRE_DYN else 9)):
                d = A[k]
                lignes.append(dict(instrument=nom, famille=fam, formant_hz=int(fc),
                                   statut=statut, registre=k[0], dynamique=k[1], **d))
                glob_dyn[k[1]]['n'] += d['n']; glob_dyn[k[1]]['ouv'] += d['n'] * d['pct_ouvert'] / 100
                print('{:7d}  {:10s} {:>5s} {:5d} {:8d}% {:>7} {:>7}'.format(
                    int(fc), k[0][:10], k[1], d['n'], d['pct_ouvert'],
                    d['bw_median'] if d['bw_median'] is not None else '—',
                    d['Q_median'] if d['Q_median'] is not None else '—'))
        o = {}
        for dy in ORDRE_DYN:
            g = glob_dyn.get(dy)
            o[dy] = None if not g or not g['n'] else int(round(100 * g['ouv'] / g['n']))
        resume.append(dict(instrument=nom, famille=fam, ouvert_doux=o['doux'],
                           ouvert_moyen=o['moyen'], ouvert_fort=o['fort'],
                           effet=None if None in (o['doux'], o['fort'])
                           else o['doux'] - o['fort']))
        print('  → part de pics ouverts : doux {}% · moyen {}% · fort {}%{}'.format(
            o['doux'], o['moyen'], o['fort'],
            '   effet doux→fort : {:+d} points'.format(o['doux'] - o['fort'])
            if None not in (o['doux'], o['fort']) else ''))

    print('\n' + '=' * 72)
    print("L'HYPOTHÈSE : la localisation s'améliore-t-elle en jeu fort ?")
    print('=' * 72)
    print('{:16s} {:28s} {:>7s} {:>7s} {:>7s} {:>9s}'.format(
        'famille', 'instrument', 'doux', 'moyen', 'fort', 'effet'))
    for r in sorted(resume, key=lambda r: (r['famille'],
                                           -(r['effet'] if r['effet'] is not None else -99))):
        print('{:16s} {:28s} {:>6}% {:>6}% {:>6}% {:>9}'.format(
            r['famille'], r['instrument'][:28], r['ouvert_doux'], r['ouvert_moyen'],
            r['ouvert_fort'],
            '—' if r['effet'] is None else '{:+d} pts'.format(r['effet'])))
    par_fam = defaultdict(list)
    for r in resume:
        if r['effet'] is not None:
            par_fam[r['famille']].append(r['effet'])
    print('\n{:16s} {:>3s} {:>12s}  {}'.format('famille', 'n', 'effet médian', 'détail'))
    for fam, v in sorted(par_fam.items(), key=lambda kv: -float(np.median(kv[1]))):
        print('{:16s} {:3d} {:+11.0f} pts  {}'.format(
            fam, len(v), float(np.median(v)), ' '.join('%+d' % x for x in sorted(v, reverse=True))))

    with open(os.path.join(a.sortie, 'largeur-par-registre.json'), 'w', encoding='utf-8') as f:
        json.dump({'cellules': lignes, 'resume': resume}, f, ensure_ascii=False, indent=1)
    if lignes:
        with open(os.path.join(a.sortie, 'largeur-par-registre.csv'), 'w', encoding='utf-8',
                  newline='') as f:
            w = csv.DictWriter(f, fieldnames=list(lignes[0]))
            w.writeheader(); w.writerows(lignes)
    print('\n{} cellules -> {}'.format(len(lignes), a.sortie))

if __name__ == '__main__':
    main()
