#!/usr/bin/env python3
"""Largeurs de bande des formants par pôles LPC — reprise de mai26.

Pourquoi ce script
------------------
`Analyses_formants-mai26/formant_analysis.py` calcule les racines du polynôme
LPC (ligne 175) puis ne garde que **l'angle** de chaque pôle, converti en
fréquence. Le **module** du pôle, qui porte la largeur de bande, était dans la
même variable et n'a jamais été sauvegardé :

    BW = -(sr / pi) * ln|r|        r = racine complexe du polynôme LPC

Les 45 CSV de mai26 n'ont donc que des fréquences centrales. Ce script reprend
EXACTEMENT le même ensemble de fichiers — via la colonne `filename` de ces CSV —
et produit les largeurs, par note et par dynamique.

Deux configurations
-------------------
**fidele**  44 100 Hz, ordre 18, attaque coupée à 100 ms, fenêtre 500 ms,
            préaccentuation 0,97, fenêtre de Hann. Reproduit mai26 à
            l'identique, et y ajoute la largeur.

**corrige** décimation à 12 000 Hz puis ordres 12, 16, 20, 24.
            Motif : à 44 100 Hz, l'ordre 18 ne laisse que **neuf paires de
            pôles pour couvrir 22 kHz**, soit une résonance tous les 2,4 kHz.
            Le modèle est sous-dimensionné et force des pôles larges — c'est
            l'artefact qui fait que le LPC « trouve toujours quelque chose » et
            qu'on le dessine en ellipses larges. À 12 kHz, l'ordre 16 donne huit
            paires sur 0–6 kHz, ce qui est le dimensionnement usuel.
            Les quatre ordres servent à tester la **stabilité** : une largeur
            qui change avec l'ordre est une propriété du modèle, pas de
            l'instrument.

Usage
-----
    python3 largeurs_lpc.py --mai26 DOSSIER --corpus RACINE --sortie DOSSIER
    python3 largeurs_lpc.py ... --seulement Bassoon,Violin
    python3 largeurs_lpc.py ... --lister
"""

import argparse, csv, glob, math, os, sys, collections
import numpy as np
import soundfile as sf

# scipy n'est pas présent sur la machine de travail : les deux seules fonctions
# qu'il fournissait sont réécrites en numpy ci-dessous.


def _toeplitz(r, n):
    """Matrice d'autocorrélation R[i,j] = r[|i-j|]."""
    i = np.abs(np.subtract.outer(np.arange(n), np.arange(n)))
    return r[i]


def _reechantillonner(x, sr_in, sr_out):
    """Rééchantillonnage par FFT. La troncature du spectre EST le filtre
    anti-repliement, donc rien à ajouter."""
    n_in = len(x)
    n_out = int(round(n_in * sr_out / sr_in))
    if n_out < 16:
        return x
    X = np.fft.rfft(x)
    Y = np.zeros(n_out // 2 + 1, dtype=complex)
    m = min(len(X), len(Y))
    Y[:m] = X[:m]
    return np.fft.irfft(Y, n_out) * (n_out / n_in)

PRE = 0.97
SKIP_MS = 100          # comme mai26
DUREE_MS = 500         # comme mai26
ORDRE_FIDELE = 18
SR_CORRIGE = 12000
ORDRES_CORRIGE = (12, 16, 20, 24)
FMAX = 6000.0
BW_MAX = 4000.0        # au-delà, le pôle ne décrit plus une résonance


def index_corpus(racine):
    """Nom de fichier -> chemin. OrchideaSOL2020/ et FullSOL2020/ sont deux
    copies identiques (vérifié par taille et empreinte) : on préfère FullSOL2020,
    qui est la source citée dans les fiches."""
    idx = {}
    for p in glob.glob(os.path.join(racine, '**', '*.wav'), recursive=True):
        b = os.path.basename(p)
        if b not in idx or ('FullSOL2020' in p and 'FullSOL2020' not in idx[b]):
            idx[b] = p
    return idx


def _onset(x, sr):
    """Même détection que FL.spectre_moyen : première trame dépassant 20 % du
    RMS maximal, sur une moyenne glissante de 10 ms."""
    k = max(1, int(sr * 0.01))
    rms = np.convolve(np.abs(x), np.ones(k) / k, mode='same')
    idx = np.nonzero(rms > 0.2 * rms.max())[0]
    return int(idx[0]) if idx.size else 0


def segment(chemin, sr_cible=None, mode='mai26'):
    """mode='mai26' : 100 ms fixes depuis le DÉBUT DU FICHIER, fenêtre 500 ms.
       mode='v7'    : onset + 200 ms, fenêtre 1 500 ms — convention du projet.

    L'écart n'est pas cosmétique. Sur SOL, l'onset médian est à 25 ms mais
    **16,7 % des fichiers ont leur onset après 100 ms** : le segment de mai26 y
    commence avant le son et contient l'attaque entière, et 0,7 % ne contiennent
    pas la note du tout."""
    x, sr = sf.read(chemin, always_2d=True)
    x = x.mean(axis=1).astype(np.float64)
    if mode == 'v7':
        o = _onset(x, sr)
        d = o + int(sr * 200 / 1000)
        f = d + int(sr * 1500 / 1000)
    else:
        d = int(sr * SKIP_MS / 1000)
        f = d + int(sr * DUREE_MS / 1000)
    seg = x[d:min(f, len(x))]
    if len(seg) < 512:
        seg = x
    if sr_cible and sr != sr_cible:
        seg = _reechantillonner(seg, sr, sr_cible)
        sr = sr_cible
    return seg, sr


def poles_lpc(seg, sr, ordre):
    """Rend [(freq, largeur, rayon), …] triés par fréquence.

    C'est ici que se joue l'ajout : `np.roots` donne des racines complexes ;
    mai26 n'en lisait que l'angle. Le module donne la largeur."""
    if len(seg) < ordre * 4:
        return []
    pre = np.append(seg[0], seg[1:] - PRE * seg[:-1])
    fr = pre * np.hanning(len(pre))
    r = np.array([np.dot(fr[k:], fr[:len(fr) - k]) for k in range(ordre + 1)])
    if r[0] <= 0:
        return []
    try:
        a = np.linalg.solve(_toeplitz(r, ordre), -r[1:ordre + 1])
    except np.linalg.LinAlgError:
        return []
    racines = np.roots(np.concatenate([[1.0], a]))
    racines = racines[np.imag(racines) > 1e-9]        # une par paire conjuguée
    out = []
    for z in racines:
        mod = abs(z)
        if not (0 < mod < 1):                          # pôle instable : écarté
            continue
        f = math.atan2(z.imag, z.real) * sr / (2 * math.pi)
        bw = -(sr / math.pi) * math.log(mod)
        if 50 < f < FMAX and 0 < bw < BW_MAX:
            out.append((f, bw, mod))
    return sorted(out)


def traiter(csv_mai26, idx, sortie, fidele=True, corrige=True, mode='mai26'):
    nom = os.path.basename(csv_mai26).replace('_formant_results.csv', '')
    lignes = list(csv.DictReader(open(csv_mai26, encoding='utf-8')))
    res, absents, echecs = [], 0, 0
    for row in lignes:
        p = idx.get(row['filename'])
        if not p:
            absents += 1
            continue
        base = dict(filename=row['filename'], instrument=row.get('instrument', ''),
                    technique=row.get('technique', ''), note=row.get('note', ''),
                    dynamic=row.get('dynamic', ''), midi=row.get('midi', ''))
        try:
            if fidele:
                seg, sr = segment(p, mode=mode)
                for k, (f, bw, mod) in enumerate(poles_lpc(seg, sr, ORDRE_FIDELE), 1):
                    res.append({**base, 'config': 'fidele', 'segment': mode, 'sr': sr,
                                'ordre': ORDRE_FIDELE, 'k': k,
                                'freq_hz': round(f, 1), 'bw_hz': round(bw, 1),
                                'rayon': round(mod, 5)})
            if corrige:
                seg, sr = segment(p, SR_CORRIGE, mode=mode)
                for o in ORDRES_CORRIGE:
                    for k, (f, bw, mod) in enumerate(poles_lpc(seg, sr, o), 1):
                        res.append({**base, 'config': 'corrige', 'segment': mode, 'sr': sr,
                                    'ordre': o, 'k': k,
                                    'freq_hz': round(f, 1), 'bw_hz': round(bw, 1),
                                    'rayon': round(mod, 5)})
        except Exception:
            echecs += 1
    if not res:
        return nom, 0, len(lignes), absents, echecs
    champs = ['filename', 'instrument', 'technique', 'note', 'dynamic', 'midi',
              'config', 'segment', 'sr', 'ordre', 'k', 'freq_hz', 'bw_hz', 'rayon']
    dest = os.path.join(sortie, nom + '_largeurs.csv')
    with open(dest, 'w', newline='') as fh:
        w = csv.DictWriter(fh, fieldnames=champs)
        w.writeheader()
        w.writerows(res)
    return nom, len(res), len(lignes), absents, echecs


def traiter_dossier(dossier, nom, sortie, fidele=True, corrige=True, mode='v7'):
    """Même traitement que `traiter()`, mais la liste des notes vient du DOSSIER
    d'échantillons et non d'un CSV de mai 2026.

    Nécessaire parce que l'analyse de mai 2026 a omis des instruments : le cor
    ordinaire n'y figure pas — seul `Horn+sordina` a été traité. Ce n'est pas un
    trou de la reprise mais de l'étude d'origine, et il se voyait d'autant moins
    que tous les scripts aval prenaient leur liste d'instruments de ce dossier-là.
    (Ajouté le 27/10/2026 pour la tâche C4.)"""
    import formants_lib as _FL
    res, echecs = [], 0
    fichiers = sorted(glob.glob(os.path.join(dossier, '*.wav')))
    for p in fichiers:
        b = os.path.basename(p)
        parties = b[:-4].split('-')
        f0, midi, dyn = _FL.analyser_nom(p)
        base = dict(filename=b,
                    instrument=parties[0] if parties else '',
                    technique=parties[1] if len(parties) > 1 else '',
                    note=parties[2] if len(parties) > 2 else '',
                    dynamic=dyn or (parties[3] if len(parties) > 3 else ''),
                    midi='' if midi is None else midi)
        try:
            if fidele:
                seg, sr = segment(p, mode=mode)
                for k, (f, bw, mod) in enumerate(poles_lpc(seg, sr, ORDRE_FIDELE), 1):
                    res.append({**base, 'config': 'fidele', 'segment': mode, 'sr': sr,
                                'ordre': ORDRE_FIDELE, 'k': k,
                                'freq_hz': round(f, 1), 'bw_hz': round(bw, 1),
                                'rayon': round(mod, 5)})
            if corrige:
                seg, sr = segment(p, SR_CORRIGE, mode=mode)
                for o in ORDRES_CORRIGE:
                    for k, (f, bw, mod) in enumerate(poles_lpc(seg, sr, o), 1):
                        res.append({**base, 'config': 'corrige', 'segment': mode, 'sr': sr,
                                    'ordre': o, 'k': k,
                                    'freq_hz': round(f, 1), 'bw_hz': round(bw, 1),
                                    'rayon': round(mod, 5)})
        except Exception:
            echecs += 1
    if not res:
        return nom, 0, len(fichiers), 0, echecs
    champs = ['filename', 'instrument', 'technique', 'note', 'dynamic', 'midi',
              'config', 'segment', 'sr', 'ordre', 'k', 'freq_hz', 'bw_hz', 'rayon']
    os.makedirs(sortie, exist_ok=True)
    with open(os.path.join(sortie, nom + '_largeurs.csv'), 'w', newline='') as fh:
        w = csv.DictWriter(fh, fieldnames=champs)
        w.writeheader(); w.writerows(res)
    return nom, len(res), len(fichiers), 0, echecs


def main():
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('--mai26')
    p.add_argument('--depuis-dossier', metavar='DOSSIER',
                   help="liste les notes depuis un dossier d'échantillons au lieu "
                        "d'un CSV de mai 2026 — pour les instruments que mai 2026 a omis")
    p.add_argument('--nom', help='nom de sortie avec --depuis-dossier')
    p.add_argument('--corpus', required=True)
    p.add_argument('--sortie', required=True)
    p.add_argument('--seulement')
    p.add_argument('--lister', action='store_true')
    p.add_argument('--sans-fidele', action='store_true')
    p.add_argument('--sans-corrige', action='store_true')
    p.add_argument('--segment', choices=('mai26', 'v7'), default='mai26',
                   help="mai26 : 100 ms fixes (comparable a mai26) ; v7 : onset+200 ms")
    a = p.parse_args()

    if a.depuis_dossier:
        nom = a.nom or os.path.basename(a.depuis_dossier.rstrip('/'))
        r = traiter_dossier(a.depuis_dossier, nom, a.sortie,
                            fidele=not a.sans_fidele, corrige=not a.sans_corrige,
                            mode=a.segment)
        print('{:32s} {:5d} notes  {:7d} pôles'.format(r[0], r[2], r[1])
              + (f'  {r[4]} échecs' if r[4] else ''))
        return
    if not a.mai26:
        p.error('--mai26 ou --depuis-dossier est requis')

    fichiers = sorted(glob.glob(os.path.join(a.mai26, '*_formant_results.csv')))
    noms = [os.path.basename(f).replace('_formant_results.csv', '') for f in fichiers]
    if a.lister:
        for n, f in zip(noms, fichiers):
            print('{:32s} {:5d} lignes'.format(n, sum(1 for _ in open(f)) - 1))
        return
    garde = {x.strip() for x in a.seulement.split(',')} if a.seulement else None
    os.makedirs(a.sortie, exist_ok=True)
    idx = index_corpus(a.corpus)
    print(f'corpus indexé : {len(idx)} noms distincts', file=sys.stderr)
    for f, n in zip(fichiers, noms):
        if garde and n not in garde:
            continue
        nom, nres, nl, abs_, ech = traiter(f, idx, a.sortie,
                                           fidele=not a.sans_fidele,
                                           corrige=not a.sans_corrige, mode=a.segment)
        print('{:32s} {:5d} notes  {:7d} pôles'.format(nom, nl, nres)
              + (f'  {abs_} absents' if abs_ else '')
              + (f'  {ech} échecs' if ech else ''))


if __name__ == '__main__':
    main()
