# -*- coding: utf-8 -*-
"""
localisation.py — tâche C6 : répliquer la détection de localisation sur les cinq corpus.

La question
-----------
`statut_formants.py` marque « aucun pic localisé » un formant dont le Q de l'enveloppe vraie
tombe sous 2 : la bosse est si large qu'elle ne sélectionne rien. La trompette et le trombone
basse ont été classés ainsi, sur la foi de **SOL seul**. Le plan de reprise pose la question
qu'il faut : *le Q ≈ 1 de la trompette est-il l'instrument, ou ce corpus ?*

Les largeurs d'enveloppe existent pour les cinq corpus — `Largeurs/Enveloppe/` — et n'avaient
jamais été croisées. Rien à recalculer.

Ce qui est mesuré
-----------------
Pour chaque formant établi et chaque corpus : les pics de l'enveloppe vraie situés à moins de
deux demi-tons du formant, puis

    part de pics OUVERTS   ceux qui ne redescendent pas de 3 dB des deux côtés. Sans largeur,
                           donc sans Q : les compter plutôt que les écarter, sinon la médiane
                           des Q ne porte que sur les pics les plus étroits
    largeur médiane        sur les seuls pics fermés
    Q = fc / largeur       sous 2, le formant ne sélectionne rien

Le verdict se lit alors corpus par corpus, et sa **stabilité** est la réponse à la question.

Usage
    python3 localisation.py
    python3 localisation.py --instruments "Trompette en Ut,Trombone basse"
"""
import os, sys, csv, json, math, glob, argparse, statistics as st
from collections import defaultdict
import numpy as np

ICI = os.path.dirname(os.path.abspath(__file__))
RACINE = os.path.dirname(os.path.dirname(ICI))
ETUDE = os.path.join(RACINE, 'Etude-Formants')
ENV = os.path.join(ETUDE, 'Largeurs', 'Enveloppe')
CORPUS = ['SOL-v7', 'ZSamples', 'Xsamples', 'VO', 'VN']
ETIQ = {'SOL-v7': 'SOL'}
TOL_DT = 2.0
SEUIL_Q = 2.0          # même seuil que `statut_formants.charger_localisation`

# noms SOL (anglais) ; les autres corpus nomment comme les fiches, à quelques variantes près
from fenetre_redressement import NOMS
ALIAS = {'cor': 'Cor en Fa', 'trompette': 'Trompette en Ut',
         'clarinette basse': 'Clarinette basse en Sib',
         'clarinette contrebasse': 'Clarinette contrebasse en Sib',
         'flute': 'Flûte', 'flute basse': 'Flûte basse',
         'flute contrebasse': 'Flûte contrebasse', 'petite flute': 'Petite flûte',
         'trombone tenor': 'Trombone ténor'}

def norm(s):
    import unicodedata, re
    s = unicodedata.normalize('NFD', s or '')
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return re.sub(r'\s+', ' ', s).strip().lower()

def fichier(corpus, nom):
    """Le CSV de largeurs de cet instrument dans ce corpus, ou None."""
    if corpus == 'SOL-v7':
        p = os.path.join(ENV, corpus, (NOMS.get(nom) or '') + '_env.csv')
        return p if os.path.exists(p) else None
    cible = norm(nom)
    for p in glob.glob(os.path.join(ENV, corpus, '*_env.csv')):
        b = os.path.basename(p)[:-len('_env.csv')]
        if any(x in b for x in ('+sordina', '+sourdine', '+piombo', 'Ensemble', ' corde')):
            continue
        if norm(b) == cible or ALIAS.get(norm(b)) == nom:
            return p
    return None

def pics(p):
    """[(freq, bw|None, ouvert)]"""
    out = []
    with open(p, encoding='utf-8') as f:
        for r in csv.DictReader(f):
            try:
                fr = float(r['freq_hz'])
            except (ValueError, KeyError):
                continue
            try:
                bw = float(r['bw_hz']) if r['bw_hz'] else None
            except ValueError:
                bw = None
            out.append((fr, bw, r.get('ouvert') == '1'))
    return out

def mesurer(P, fc):
    lfc = math.log2(fc)
    sel = [x for x in P if abs(12 * (math.log2(max(x[0], 1e-9)) - lfc)) <= TOL_DT]
    if len(sel) < 5:
        return None
    bws = [x[1] for x in sel if x[1] and not x[2]]
    ouv = 100.0 * sum(1 for x in sel if x[2] or x[1] is None) / len(sel)
    if not bws:
        return dict(n=len(sel), pct_ouvert=int(round(ouv)), bw=None, Q=None)
    bw = float(np.median(bws))
    return dict(n=len(sel), pct_ouvert=int(round(ouv)), bw=round(bw, 1),
                Q=round(fc / bw, 2))

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--instruments', default='')
    ap.add_argument('--sortie', default=os.path.join(ETUDE, 'Mesures', 'Localisation'))
    a = ap.parse_args()
    os.makedirs(a.sortie, exist_ok=True)
    voulus = {x.strip() for x in a.instruments.split(',') if x.strip()}

    with open(os.path.join(ETUDE, 'Resultats-v7', 'formants-v7.csv'), encoding='utf-8') as f:
        FORM = [r for r in csv.DictReader(f) if r['statut'] in ('route A', '1 corpus')]
    par = defaultdict(list)
    for r in FORM:
        if not voulus or r['instrument'] in voulus:
            par[r['instrument']].append((float(r['hz']), r['statut']))

    CH = {n: {c: fichier(c, n) for c in CORPUS} for n in par}
    P = {n: {c: pics(p) for c, p in d.items() if p} for n, d in CH.items()}

    lignes = []
    print('{:26s} {:>7s} {:>8s}  {}'.format('instrument', 'formant', 'statut',
                                            '  '.join('%-13s' % ETIQ.get(c, c)
                                                      for c in CORPUS)))
    for nom in sorted(par):
        for fc, statut in sorted(par[nom]):
            cols, Qs, d = [], [], dict(instrument=nom, formant_hz=int(fc), statut=statut)
            for c in CORPUS:
                m = mesurer(P.get(nom, {}).get(c, []), fc) if P.get(nom, {}).get(c) else None
                e = ETIQ.get(c, c)
                if not m:
                    cols.append('—'); d['Q_' + e] = None; continue
                d['Q_' + e] = m['Q']; d['ouvert_' + e] = m['pct_ouvert']; d['n_' + e] = m['n']
                if m['Q'] is None:
                    cols.append('tous ouv.')
                else:
                    Qs.append(m['Q'])
                    cols.append('Q{:5.1f} {:3d}%'.format(m['Q'], m['pct_ouvert']))
            d['n_corpus'] = len(Qs)
            d['Q_median'] = round(float(np.median(Qs)), 2) if Qs else None
            d['Q_min'] = min(Qs) if Qs else None
            d['Q_max'] = max(Qs) if Qs else None
            d['localise_partout'] = (all(q >= SEUIL_Q for q in Qs) if len(Qs) >= 2 else None)
            d['non_localise_partout'] = (all(q < SEUIL_Q for q in Qs) if len(Qs) >= 2 else None)
            lignes.append(d)
            print('{:26s} {:6d}Hz {:>8s}  {}'.format(
                nom[:26], int(fc), statut[:8], '  '.join('%-13s' % x for x in cols)))

    print('\n' + '=' * 78)
    av = [d for d in lignes if d['n_corpus'] >= 2]
    print("RÉPLICATION DU VERDICT, sur les {} formants mesurés dans au moins deux corpus"
          .format(len(av)))
    print('  localisé (Q ≥ 2) dans TOUS les corpus      : {}'.format(
        sum(1 for d in av if d['localise_partout'])))
    print('  NON localisé dans tous les corpus          : {}'.format(
        sum(1 for d in av if d['non_localise_partout'])))
    print('  verdict qui CHANGE selon le corpus         : {}'.format(
        sum(1 for d in av if not d['localise_partout'] and not d['non_localise_partout'])))
    disc = [d for d in av if not d['localise_partout'] and not d['non_localise_partout']]
    if disc:
        print('\n  les discordants :')
        for d in sorted(disc, key=lambda d: d['instrument']):
            print('     {:26s} {:5d} Hz   Q de {:.1f} à {:.1f} sur {} corpus'.format(
                d['instrument'][:26], d['formant_hz'], d['Q_min'], d['Q_max'], d['n_corpus']))
    with open(os.path.join(a.sortie, 'localisation.csv'), 'w', encoding='utf-8',
              newline='') as f:
        ch = sorted({k for d in lignes for k in d})
        w = csv.DictWriter(f, fieldnames=['instrument', 'formant_hz', 'statut', 'n_corpus',
                                          'Q_median', 'Q_min', 'Q_max', 'localise_partout',
                                          'non_localise_partout']
                           + [k for k in ch if k.startswith(('Q_S', 'Q_Z', 'Q_X', 'Q_V',
                                                             'ouvert_', 'n_S', 'n_Z', 'n_X',
                                                             'n_V'))])
        w.writeheader()
        for d in lignes:
            w.writerow({k: v for k, v in d.items() if k in w.fieldnames})
    print('\n{} formants -> {}'.format(len(lignes), a.sortie))

if __name__ == '__main__':
    main()
