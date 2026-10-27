# -*- coding: utf-8 -*-
"""
evaluer_migration.py — avant / après sur le seul chiffre qui décide.

Les listes de formants par corpus bougent beaucoup quand la fenêtre change : c'est attendu,
un maximum pris sur un seul corpus est bruité. Ce qui compte est ce qui **survit à la route
A** : un couple de corpus de provenances distinctes au même point de grille.

Le script recalcule ce comptage à la fenêtre courante, sans écrire ni fiche ni CSV, et le
compare à l'état archivé. Il faut ce chiffre AVANT de toucher aux 22 fiches.

SOL n'a pas de JSON par instrument — ses formants vivent dans le YAML des fiches — donc on le
remesure ici depuis ses fichiers de spectres.
"""
import os, sys, re, json, glob, math, itertools, argparse
import numpy as np

ICI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ICI)
import formants_lib as FL
import statut_formants as SF
from fenetre_redressement import NOMS, chemins

RACINE = os.path.dirname(os.path.dirname(ICI))
ETUDE = os.path.join(RACINE, 'Etude-Formants')
SOL_OUT = os.path.expanduser('~/work/out')
ORDRE = ['SOL', 'ZSamples', 'Xsamples', 'VO', 'VN']

def provenances():
    """{clé: IRCAM|VSL} d'après `source_echantillons` des fiches, comme statut_formants."""
    prov = {}
    for p in glob.glob(os.path.join(ETUDE, 'Instruments', '*', '[0-9][0-9]-*.md')):
        t = open(p, encoding='utf-8').read()
        m = re.match(r'^---\n(.*?)\n---\n', t, re.S)
        if not m:
            continue
        y = m.group(1)
        nom = re.search(r'^nom:\s*(.+)$', y, re.M)
        src = re.search(r'^source_echantillons:\s*(.+)$', y, re.M)
        if nom:
            prov[SF.cle(nom.group(1))] = ('VSL' if src and 'Yan_Adds' in src.group(1)
                                          else 'IRCAM')
    return prov

def retenus_corpus(dossier):
    """{clé: [hz retenus]} pour un corpus à JSON par instrument."""
    out = {}
    for p in sorted(glob.glob(os.path.join(dossier, '*.json'))):
        b = os.path.basename(p)[:-5]
        if any(x in b for x in (' corde', '+sourdine', '+piombo', '+sord')):
            continue
        try:
            j = json.load(open(p, encoding='utf-8'))
        except Exception:
            continue
        if 'formants' not in j:
            continue
        out[SF.cle(b)] = sorted(int(round(float(x['hz']))) for x in j['formants']
                                if str(x.get('verdict', '')).startswith('retenu'))
    return out

def mesurer_sol():
    """{clé: ([hz retenus], res)} remesuré depuis ~/work/out à la fenêtre courante."""
    out = {}
    for nom, fich in sorted(NOMS.items()):
        p = os.path.join(SOL_OUT, fich + '.txt')
        if not os.path.exists(p):
            continue
        r = FL.mesurer(p)
        if 'erreur' in r:
            continue
        out[SF.cle(nom)] = (sorted(int(round(float(x['hz']))) for x in r.get('formants', [])
                                   if str(x.get('verdict', '')).startswith('retenu')), r)
        print('  SOL {:32s} {}'.format(nom, out[SF.cle(nom)][0]))
    return out

def bilan(C, prov, etiquette):
    """Compte les formants route A, par la logique exacte de statut_formants."""
    total, insts, detail = 0, set(), {}
    for k in set().union(*[set(C[c]) for c in ORDRE if C.get(c)]):
        et = SF.centres_etablis(k, C, prov)
        if et:
            total += len(et); insts.add(k); detail[k] = et
    print('\n{} : {} formants route A sur {} instruments'.format(etiquette, total, len(insts)))
    return total, insts, detail

def zones(detail):
    par_hz = {}
    for k, et in detail.items():
        for item in et:
            hz = int(round(item[0] if not isinstance(item, (int, float)) else item))
            par_hz.setdefault(hz, set()).add(k)
    hzs = sorted(par_hz)
    uni = [(h, sorted(par_hz[h])) for h in hzs if len(par_hz[h]) > 1]
    gr = []
    for h in hzs:
        if gr and (12 * math.log2(h / gr[-1][-1]) <= 1.05
                   and 12 * math.log2(h / gr[-1][0]) <= 2.0):
            gr[-1].append(h)
        else:
            gr.append([h])
    reg = []
    for g in gr:
        ins = set()
        for h in g:
            ins |= par_hz[h]
        if len(ins) > 1 and len(g) > 1:
            reg.append((g[0], g[-1], sorted(ins)))
    return uni, reg

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--archives', default=os.path.join(
        ETUDE, 'Archives', 'halfwin-0.75-2026-10-27'))
    a = ap.parse_args()
    prov = provenances()
    print('fenêtre courante : ±{} octave'.format(FL.HALFWIN))

    # ---------- état archivé (±0,75) ----------
    Cav = {c: {} for c in ORDRE}
    for c in ORDRE[1:]:
        Cav[c] = retenus_corpus(os.path.join(a.archives, 'Mesures', c))
    for p in glob.glob(os.path.join(a.archives, 'Instruments', '*', '[0-9][0-9]-*.md')):
        t = open(p, encoding='utf-8').read()
        m = re.match(r'^---\n(.*?)\n---\n', t, re.S)
        if not m:
            continue
        nom = re.search(r'^nom:\s*(.+)$', m.group(1), re.M)
        if nom:
            Cav['SOL'][SF.cle(nom.group(1))] = SF.formants_de_yaml(m.group(1))

    # ---------- état courant (±1,00) ----------
    print('\nremesure de SOL à la fenêtre courante :')
    sol = mesurer_sol()
    Cap = {c: {} for c in ORDRE}
    Cap['SOL'] = {k: v[0] for k, v in sol.items()}
    for c in ORDRE[1:]:
        Cap[c] = retenus_corpus(os.path.join(ETUDE, 'Mesures', c))

    tav, iav, dav = bilan(Cav, prov, 'AVANT  ±0,75')
    tap, iap, dap = bilan(Cap, prov, 'APRÈS  ±1,00')

    uav, rav = zones(dav); uap, rap = zones(dap)
    print('\nconvergences — unissons : {} -> {}   régions : {} -> {}'.format(
        len(uav), len(uap), len(rav), len(rap)))
    print('\nunissons AVANT :')
    for h, ins in uav:
        print('   {:6d} Hz  {}'.format(h, ' · '.join(ins)))
    print('unissons APRÈS :')
    for h, ins in uap:
        print('   {:6d} Hz  {}'.format(h, ' · '.join(ins)))
    print('\ninstruments perdus :', sorted(iav - iap) or 'aucun')
    print('instruments gagnés :', sorted(iap - iav) or 'aucun')

    with open(os.path.join(ETUDE, 'Mesures', 'bilan-migration.json'), 'w',
              encoding='utf-8') as f:
        json.dump({'avant': {'total': tav, 'instruments': sorted(iav),
                             'unissons': [[h, i] for h, i in uav],
                             'regions': [[x, y, i] for x, y, i in rav]},
                   'apres': {'total': tap, 'instruments': sorted(iap),
                             'unissons': [[h, i] for h, i in uap],
                             'regions': [[x, y, i] for x, y, i in rap]},
                   'sol_apres': {k: v[0] for k, v in sol.items()}},
                  f, ensure_ascii=False, indent=1)

if __name__ == '__main__':
    main()
