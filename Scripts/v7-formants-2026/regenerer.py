# -*- coding: utf-8 -*-
"""
regenerer.py — régénère les mesures de méthode 8 après un changement de HALFWIN.

Ne relit PAS l'audio : les fichiers `.spectres.txt` existent déjà et sont indépendants
de la fenêtre de redressement (ils ne portent que des spectres moyens). On rappelle donc
`FL.mesurer()` sur eux, ce qui évite aussi de dépendre des volumes d'échantillons montés.

Deux précautions reprises de `corpus.py` :
  · VO corrige l'accord. Le fichier de spectres est déjà recalé, mais les fréquences des
    formants étaient ensuite remultipliées par le rapport médian. On relit ce rapport dans
    l'ancien JSON (`accord_rapport`) et on le réapplique, sinon VO sortirait décalé.
  · Les configurations dérivées (cordes à vide, sourdines, piombo) sont régénérées comme
    les autres, mais elles restent suspendues pour la publication.

Usage
    python3 regenerer.py --corpus ZSamples,Xsamples,VO,VN
    python3 regenerer.py --sol                 # les 22 instruments de SOL, dans ~/work/out
"""
import os, sys, json, glob, argparse, shutil

ICI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ICI)
import formants_lib as FL

RACINE = os.path.dirname(os.path.dirname(ICI))
ETUDE = os.path.join(RACINE, 'Etude-Formants')
SOL_OUT = os.path.expanduser('~/work/out')

def _hz(x):
    return int(round(float(x)))

def regenerer_corpus(corpus):
    d = os.path.join(ETUDE, 'Mesures', corpus)
    sp = sorted(glob.glob(os.path.join(d, '*.spectres.txt')))
    print('\n=== {} : {} configurations ==='.format(corpus, len(sp)))
    bilan = []
    for p in sp:
        nom = os.path.basename(p)[:-len('.spectres.txt')]
        pj = os.path.join(d, nom + '.json')
        ancien = {}
        if os.path.exists(pj):
            try:
                ancien = json.load(open(pj, encoding='utf-8'))
            except Exception:
                pass
        avant = sorted(_hz(x['hz']) for x in ancien.get('formants', [])
                       if str(x.get('verdict', '')).startswith('retenu'))
        res = FL.mesurer(p)
        if 'erreur' in res:
            print('  ✗ {:34s} {}'.format(nom, res['erreur']))
            continue
        # correction d'accord : on reprend le rapport mesuré lors de la lecture de l'audio
        r_med = float(ancien.get('accord_rapport') or 1.0)
        corr = bool(ancien.get('accord_corrige'))
        res['accord_corrige'] = corr
        res['accord_rapport'] = round(r_med, 4)
        res['accord_cents'] = ancien.get('accord_cents', 0)
        if corr and r_med != 1.0:
            for x in res.get('formants', []):
                if x.get('hz'):
                    x['hz'] = float(x['hz']) * r_med
        res['fenetre_redressement'] = FL.HALFWIN
        with open(pj, 'w', encoding='utf-8') as fh:
            json.dump(res, fh, ensure_ascii=False, indent=1, default=float)
        apres = sorted(_hz(x['hz']) for x in res.get('formants', [])
                       if str(x.get('verdict', '')).startswith('retenu'))
        chg = '' if avant == apres else '   <<< changé'
        print('  {:34s} {} -> {}{}'.format(nom, avant or '[]', apres or '[]', chg))
        bilan.append((corpus, nom, avant, apres))
    return bilan

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--corpus', default='ZSamples,Xsamples,VO,VN')
    ap.add_argument('--bilan', default=os.path.join(ETUDE, 'Mesures',
                                                    'bilan-regeneration.json'))
    a = ap.parse_args()
    tout = []
    for c in [x.strip() for x in a.corpus.split(',') if x.strip()]:
        tout += regenerer_corpus(c)
    chg = [x for x in tout if x[2] != x[3]]
    print('\n{} configurations régénérées, {} changées ({:.0f} %)'.format(
        len(tout), len(chg), 100 * len(chg) / max(len(tout), 1)))
    with open(a.bilan, 'w', encoding='utf-8') as f:
        json.dump([{'corpus': c, 'config': n, 'avant': a_, 'apres': b}
                   for c, n, a_, b in tout], f, ensure_ascii=False, indent=1)

if __name__ == '__main__':
    main()
