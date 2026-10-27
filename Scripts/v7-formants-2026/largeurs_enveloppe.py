#!/usr/bin/env python3
"""Largeurs de bande à −3 dB sur l'enveloppe vraie (cepstre à ordre adapté).

Pourquoi, face au LPC
---------------------
Le LPC impose un modèle **tout-pôle source-filtre**, c'est-à-dire un modèle du
conduit vocal : une source indépendante excitant un résonateur passif. Chez un
instrument à vent, l'anche et la colonne d'air sont **couplées** — le résonateur
impose l'oscillation au lieu de filtrer une source autonome. Et le LPC rend
toujours `ordre/2` résonances, qu'elles existent ou non : il ne peut pas dire
« il n'y a pas de formant ici », ce que nous devons pouvoir dire.

Le cepstre ne suppose rien. Il sépare la variation lente du log-spectre de la
variation rapide du peigne harmonique.

Mais un ordre FIXE porte le défaut d'Orchidea : l'ordre N ne retire le peigne
que si f0 < fe/N. Ordre 160 → valide sous 276 Hz. **Ordre 80 → sous 551 Hz.**
Mesuré sur la trompette : 44 % des pics tombent sur un harmonique avec l'ordre
fixe 80, **0 %** avec l'ordre adapté. D'où l'enveloppe vraie de la méthode 6,
dont l'ordre est réglé sur f0, note par note.

La largeur se lit alors directement à −3 dB sur la courbe, sans modèle de pôles
et sans nombre de résonances imposé.

Honnêteté de la mesure
----------------------
Une largeur n'est retenue que si le niveau redescend de 3 dB **des deux côtés**
à l'intérieur de la plage. Sinon le pic est marqué `ouvert` et sa largeur n'est
pas publiée — c'est exactement le défaut de `bande_db3` dans les fiches, où plus
de la moitié des valeurs sont des butées de plage prises pour des largeurs.

Deux segmentations
------------------
`--segment mai26` : 100 ms coupés, fenêtre de 500 ms — identique au calcul LPC,
                    donc comparable note à note.
`--segment v7`    : onset + 200 ms, fenêtre de 1 500 ms — convention du projet.

Usage
-----
    python3 largeurs_enveloppe.py --mai26 DOSSIER --corpus RACINE --sortie DOSSIER
    python3 largeurs_enveloppe.py --spectres DOSSIER --sortie DOSSIER   # corpus déjà mesuré
"""

import argparse, csv, glob, math, os, sys
import numpy as np
import soundfile as sf

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import formants_lib as FL

DB = 20.0 / math.log(10.0)     # népers -> dB
F_BAS, F_HAUT = 150.0, 6000.0
CHUTE = 3.0                    # dB
PROEMINENCE_MIN = 1.0          # dB


def _onset(x, sr):
    k = max(1, int(sr * 0.01))
    rms = np.convolve(np.abs(x), np.ones(k) / k, mode='same')
    idx = np.nonzero(rms > 0.2 * rms.max())[0]
    return int(idx[0]) if idx.size else 0


def segment(x, sr, mode='mai26'):
    """mode='mai26' : 100 ms fixes depuis le DEBUT DU FICHIER, fenetre 500 ms.
       mode='v7'    : onset + 200 ms, fenetre 1 500 ms — convention du projet.

    Sur SOL, l'onset median est a 25 ms mais 16,7 % des fichiers l'ont apres
    100 ms : le segment mai26 y commence avant le son et contient l'attaque."""
    if mode == 'v7':
        d = _onset(x, sr) + int(sr * 200 / 1000)
        f = d + int(sr * 1500 / 1000)
    else:
        d = int(sr * 100 / 1000)
        f = d + int(sr * 500 / 1000)
    seg = x[d:min(f, len(x))]
    return seg if len(seg) >= 512 else x


def spectre_moyen_segment(seg):
    """Mêmes paramètres de FFT que FL.spectre_moyen, sur un segment imposé."""
    W = np.hanning(FL.NFFT)
    tr = [np.abs(np.fft.rfft(seg[i:i + FL.NFFT] * W, n=FL.NFFT))
          for i in range(0, len(seg) - FL.NFFT + 1, FL.HOP)]
    if not tr:
        pad = np.zeros(FL.NFFT)
        n = min(len(seg), FL.NFFT)
        pad[:n] = seg[:n] * W[:n]
        tr = [np.abs(np.fft.rfft(pad, n=FL.NFFT))]
    m = np.mean(tr, axis=0)[:FL.NBINS]
    s = m.sum()
    return m / s if s > 0 else m


def largeurs(spectre, f0):
    """Rend [(freq, largeur ou None, proéminence, ouvert), …] sur l'enveloppe vraie."""
    V, C = FL.enveloppe_vraie(spectre, f0)
    D = V * DB
    i0, i1 = int(F_BAS / FL.BW), min(int(F_HAUT / FL.BW), FL.NBINS - 2)
    out = []
    for i in range(i0 + 1, i1):
        if not (D[i] >= D[i - 1] and D[i] > D[i + 1]):
            continue
        cible = D[i] - CHUTE
        g = i
        while g > i0 and D[g] > cible:
            g -= 1
        d = i
        while d < i1 and D[d] > cible:
            d += 1
        ferme_g, ferme_d = D[g] <= cible, D[d] <= cible
        # proéminence : écart au plus haut des deux minima latéraux
        mg = D[g:i + 1].min() if g < i else D[i]
        md = D[i:d + 1].min() if d > i else D[i]
        pro = D[i] - max(mg, md)
        if pro < PROEMINENCE_MIN:
            continue
        ouvert = not (ferme_g and ferme_d)
        bw = None if ouvert else (d - g) * FL.BW
        out.append((i * FL.BW, bw, pro, ouvert, C))
    return out


def ecrire(lignes, dest):
    champs = ['source', 'instrument', 'note', 'dynamic', 'midi', 'f0_hz',
              'segment', 'ordre_cepstre', 'freq_hz', 'bw_hz', 'proeminence_db', 'ouvert']
    with open(dest, 'w', newline='') as fh:
        w = csv.DictWriter(fh, fieldnames=champs)
        w.writeheader()
        w.writerows(lignes)


def depuis_mai26(dossier_mai26, racine, sortie, seulement=None, mode='mai26'):
    idx = {}
    for p in glob.glob(os.path.join(racine, '**', '*.wav'), recursive=True):
        b = os.path.basename(p)
        if b not in idx or ('FullSOL2020' in p and 'FullSOL2020' not in idx[b]):
            idx[b] = p
    print(f'corpus indexé : {len(idx)} noms', file=sys.stderr)
    os.makedirs(sortie, exist_ok=True)
    for f in sorted(glob.glob(os.path.join(dossier_mai26, '*_formant_results.csv'))):
        nom = os.path.basename(f).replace('_formant_results.csv', '')
        if seulement and nom not in seulement:
            continue
        lignes, nn, ko = [], 0, 0
        for r in csv.DictReader(open(f, encoding='utf-8')):
            p = idx.get(r['filename'])
            if not p:
                ko += 1
                continue
            try:
                x, sr = sf.read(p, always_2d=True)
                x = x.mean(axis=1).astype(np.float64)
                if sr != FL.FS:
                    ko += 1
                    continue
                sp = spectre_moyen_segment(segment(x, sr, mode))
                nt = FL.note_vers_f0(r.get('note', ''))
                if nt is None:
                    ko += 1
                    continue
                f0 = nt[0]
                nn += 1
                for fr, bw, pro, ouv, C in largeurs(sp, f0):
                    lignes.append(dict(source='mai26', segment=mode, instrument=nom,
                                       note=r.get('note', ''), dynamic=r.get('dynamic', ''),
                                       midi=r.get('midi', ''), f0_hz=round(f0, 1),
                                       ordre_cepstre=C, freq_hz=round(fr, 1),
                                       bw_hz='' if bw is None else round(bw, 1),
                                       proeminence_db=round(pro, 2),
                                       ouvert=int(ouv)))
            except Exception:
                ko += 1
        if lignes:
            ecrire(lignes, os.path.join(sortie, nom + '_env.csv'))
        print('{:30s} {:5d} notes  {:6d} pics'.format(nom, nn, len(lignes))
              + (f'  {ko} ignorés' if ko else ''))


def depuis_dossier(dossier, nom, sortie, mode='v7'):
    """Liste les notes depuis un DOSSIER d'échantillons, et non depuis un CSV de
    mai 2026. Nécessaire parce que l'analyse de mai 2026 a omis des instruments —
    le cor ordinaire n'y figure pas, seul `Horn+sordina` — et que tous les scripts
    aval tiraient leur liste d'instruments de ce dossier, si bien que l'absence ne
    se voyait nulle part. (Ajouté le 27/10/2026, tâche C5.)"""
    os.makedirs(sortie, exist_ok=True)
    lignes, nn, ko = [], 0, 0
    for p in sorted(glob.glob(os.path.join(dossier, '*.wav'))):
        b_ = os.path.basename(p)
        parties = b_[:-4].split('-')
        f0, midi, dyn = FL.analyser_nom(p)
        if f0 is None:
            ko += 1
            continue
        try:
            x, sr = sf.read(p, always_2d=True)
            x = x.mean(axis=1).astype(np.float64)
            if sr != FL.FS:
                ko += 1
                continue
            sp = spectre_moyen_segment(segment(x, sr, mode))
            nn += 1
            for fr, bw, pro, ouv, C in largeurs(sp, f0):
                lignes.append(dict(source='dossier', segment=mode, instrument=nom,
                                   note=parties[2] if len(parties) > 2 else '',
                                   dynamic=dyn or '',
                                   midi='' if midi is None else midi,
                                   f0_hz=round(f0, 1), ordre_cepstre=C,
                                   freq_hz=round(fr, 1),
                                   bw_hz='' if bw is None else round(bw, 1),
                                   proeminence_db=round(pro, 2), ouvert=int(ouv)))
        except Exception:
            ko += 1
    if lignes:
        ecrire(lignes, os.path.join(sortie, nom + '_env.csv'))
    print('{:30s} {:5d} notes  {:6d} pics'.format(nom, nn, len(lignes))
          + (f'  {ko} ignorés' if ko else ''))


def depuis_spectres(dossier, sortie, etiquette):
    os.makedirs(sortie, exist_ok=True)
    for p in sorted(glob.glob(os.path.join(dossier, '*.spectres.txt'))):
        nom = os.path.basename(p)[:-13]
        ech = FL.lire_spectres(p)
        lignes = []
        for f0, midi, dyn, sp in ech:
            for fr, bw, pro, ouv, C in largeurs(sp, f0):
                lignes.append(dict(source=etiquette, segment='v7', instrument=nom, note='',
                                   dynamic=dyn or '', midi=midi, f0_hz=round(f0, 1),
                                   ordre_cepstre=C, freq_hz=round(fr, 1),
                                   bw_hz='' if bw is None else round(bw, 1),
                                   proeminence_db=round(pro, 2), ouvert=int(ouv)))
        if lignes:
            ecrire(lignes, os.path.join(sortie, nom + '_env.csv'))
        print('{:30s} {:5d} notes  {:6d} pics'.format(nom, len(ech), len(lignes)))


def main():
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('--mai26')

    p.add_argument('--depuis-dossier', metavar='DOSSIER')

    p.add_argument('--nom')
    p.add_argument('--corpus')
    p.add_argument('--spectres')
    p.add_argument('--etiquette', default='corpus')
    p.add_argument('--sortie', required=True)
    p.add_argument('--seulement')
    p.add_argument('--segment', choices=('mai26', 'v7'), default='mai26')
    a = p.parse_args()

    if a.depuis_dossier:

        depuis_dossier(a.depuis_dossier,

                       a.nom or os.path.basename(a.depuis_dossier.rstrip('/')),

                       a.sortie, a.segment)

        return
    s = {x.strip() for x in a.seulement.split(',')} if a.seulement else None
    if a.mai26:
        if not a.corpus:
            sys.exit('--corpus requis avec --mai26')
        depuis_mai26(a.mai26, a.corpus, a.sortie, s, a.segment)
    elif a.spectres:
        depuis_spectres(a.spectres, a.sortie, a.etiquette)
    else:
        p.print_help()


if __name__ == '__main__':
    main()
