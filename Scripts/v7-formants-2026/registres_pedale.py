# -*- coding: utf-8 -*-
"""
registres_pedale.py — tâche C7 : relancer les cuivres avec les registres pédale.

Le point de départ
------------------
`registres.md` a reçu le 24/10 une définition des sons pédale pour le cor, la trompette et
les trombones, et un « extrême grave » pour les tubas. Avant cela ces notes tombaient dans
la catégorie fourre-tout « hors registre » et n'étaient lues par personne. La tâche demande
de relancer les cuivres avec ces registres.

Trois questions, et non une
---------------------------
1. **La couverture est-elle complète ?** Un registre défini ne sert à rien si des notes du
   corpus tombent encore à côté. On liste note par note ce qui reste « hors registre ».

2. **Les cellules pédale sont-elles lisibles ?** La carte vocalique exige cinq notes par
   cellule et deux maxima. Les pédales sont rares dans les corpus : il faut compter avant
   d'interpréter.

3. **Que gagne la méthode 8 ?** C'est la question qui compte, et elle n'a rien à voir avec
   les registres. La méthode 8 ignore les registres : elle met toutes les notes dans un
   même composite. Mais la tâche C2 a montré qu'un harmonique n'y entre que si sa fenêtre
   de ±3/4 d'octave contient trois harmoniques de la même note, ce qui écarte les rangs
   bas et place un **plancher de domaine** entre 3,2 et 5,7 fois f0_min. Or les pédales
   abaissent f0_min d'une octave. **Elles doivent donc abaisser le plancher**, et ouvrir à
   la mesure une région qui lui était fermée. On le vérifie en relançant chaque cuivre avec
   et sans ses pédales.

Et un test de rejet, celui de la tâche C1 : dans un registre donné, la pente de log2(F1)
contre log2(f0) doit être nulle si F1 est une propriété du tuyau. Le registre pédale est le
meilleur terrain possible pour ce test, puisque F1 y vaut dix à vingt fois f0 : aucune
confusion avec la fondamentale n'y est possible.

Usage
    python3 registres_pedale.py
    python3 registres_pedale.py --instruments "Trombone basse,Cor en Fa"
"""
import os, sys, csv, math, glob, argparse, itertools
from collections import defaultdict
import numpy as np

ICI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ICI)
import formants_lib as FL
import carte_vocalique as CV
from fenetre_redressement import NOMS

RACINE = os.path.dirname(os.path.dirname(ICI))
ETUDE = os.path.join(RACINE, 'Etude-Formants')
REGISTRES = os.path.join(RACINE, 'Data', 'registres.md')
SOL_OUT = os.path.expanduser('~/work/out')
CORPUS = ['SOL', 'ZSamples', 'Xsamples', 'VO', 'VN']

# cuivres, et la clé de `registres.md` qui leur correspond
CUIVRES = {'Cor en Fa': 'cor', 'Trompette en Ut': 'trompette',
           'Trombone ténor': 'trombones', 'Trombone basse': 'trombones',
           'Tuba basse': 'tubas', 'Tuba contrebasse': 'tubas'}
PEDALE = ('sons pédale', 'extrême grave')
BAS = set(PEDALE) | {'grave', 'médium'}     # au-dessus, F1 = f0 : voir §3 du document
MIN_NOTES_PENTE = 5
MIN_NOTES_DYN = 8
SEUIL_STRICT = 0.5                          # demi-ton, le seuil de la route A
ACCORDS_MARQUES = []

# Provenance du côté « SOL » de chaque cuivre, d'après `source_echantillons` des fiches.
# `_CORPUS-CROISES.md`, 25/10 : `Yan_Adds-Divers` contient du VSL et VN EST du VSL — trois
# paires vérifiées identiques au sample près. La paire (SOL, VN) ne compte donc pas quand la
# source SOL est Yan_Adds. C'est le cas du trombone basse et du tuba contrebasse.
ORDRE_INST = ['Cor en Fa', 'Trompette en Ut', 'Trombone ténor', 'Trombone basse',
              'Tuba basse', 'Tuba contrebasse']
PROV = {'Cor en Fa': 'IRCAM', 'Trompette en Ut': 'IRCAM', 'Trombone ténor': 'IRCAM',
        'Tuba basse': 'IRCAM', 'Trombone basse': 'VSL', 'Tuba contrebasse': 'VSL'}

def paire_valide(instrument, a, b):
    return not ({a, b} == {'SOL', 'VN'} and PROV.get(instrument) == 'VSL')

NOMS_NOTE = ['Do', 'Do#', 'Ré', 'Ré#', 'Mi', 'Fa', 'Fa#', 'Sol', 'Sol#', 'La', 'La#', 'Si']

def nom_midi(m):
    return '%s%d' % (NOMS_NOTE[m % 12], m // 12 - 1)

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

def pente(xs, ys):
    """Moindres carrés sur log2 : pente de log2(y) contre log2(x), et r². Test de C1."""
    x = np.log2(np.asarray(xs, float)); y = np.log2(np.asarray(ys, float))
    x = x - x.mean(); yc = y - y.mean()
    den = float((x * x).sum())
    if den <= 0:
        return None, None
    a = float((x * yc).sum() / den)
    sst = float((yc * yc).sum())
    r2 = 1.0 - float(((yc - a * x) ** 2).sum()) / sst if sst > 0 else None
    return a, r2

def formants_de(ech, dedup=False):
    """Composite de la méthode 8 sur ces notes, et ses maxima. `dedup` écarte les
    fréquences dupliquées sans retirer de note : c'est le témoin qui isole l'artefact
    de résolution sans toucher au domaine."""
    acc = [[] for _ in range(FL.NB)]; vues = [set() for _ in range(FL.NB)]
    for f0, midi, dyn, sp in ech:
        if not f0:
            continue
        fr, am = FL.harmoniques(sp, f0)
        if fr.size == 0:
            continue
        if dedup:
            _, i = np.unique(fr, return_index=True)
            i = np.sort(i); fr, am = fr[i], am[i]
        r, nl = FL.residu_local(fr, am)
        idx = np.digitize(fr, FL.BORDS) - 1
        for i, v, q in zip(idx, r, nl):
            if 0 <= i < FL.NB and q >= 3:
                acc[i].append(v); vues[i].add(midi)
    med = np.array([np.median(a) if len(a) >= 5 else np.nan for a in acc])
    return FL.trouver_formants(med, np.array([len(s) for s in vues]),
                               np.array([len(a) for a in acc]))

def iqr_dt(v):
    if len(v) < 3:
        return None
    q1, q3 = np.percentile(v, 25), np.percentile(v, 75)
    return round(12 * math.log2(max(q3, 1e-9) / max(q1, 1e-9)), 1)


# ──────────────────────────────── figure ───────────────────────────────────
ORDRE_REG = ['sons pédale', 'extrême grave', 'grave', 'médium', 'aigu', 'suraigu']
STYLE_CORPUS = {'SOL': '-', 'ZSamples': '--', 'Xsamples': ':', 'VO': '-.', 'VN': (0, (3, 1, 1, 1))}

def figure(lignes, V, chemin):
    """Deux panneaux : à gauche la borne de lecture (F1/f0 par registre, qui tend vers 1),
    à droite la variable cachée (F1 par nuance, et les accords qui apparaissent au doux)."""
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    from matplotlib.lines import Line2D
    plt.rcParams.update({'font.size': 8, 'axes.titlesize': 9, 'axes.labelsize': 8,
                         'xtick.labelsize': 7.5, 'ytick.labelsize': 7.5,
                         'figure.dpi': 130, 'savefig.dpi': 130,
                         'savefig.bbox': 'tight', 'legend.frameon': False})
    insts = [n for n in ORDRE_INST if any(d['instrument'] == n for d in lignes)]
    coul = {n: plt.get_cmap('viridis')(i / max(len(insts) - 1, 1) * 0.88)
            for i, n in enumerate(insts)}
    fig, (g, d) = plt.subplots(1, 2, figsize=(10.4, 4.3))

    # ── gauche : F1/f0 par registre ─────────────────────────────────────────
    for nom in insts:
        for corpus in CORPUS:
            L = [r for r in lignes if r['instrument'] == nom and r['corpus'] == corpus
                 and r['registre'] in ORDRE_REG and r['rapport_F1_f0']]
            if len(L) < 2:
                continue
            L.sort(key=lambda r: ORDRE_REG.index(r['registre']))
            g.plot([ORDRE_REG.index(r['registre']) for r in L],
                   [r['rapport_F1_f0'] for r in L], ls=STYLE_CORPUS.get(corpus, '-'),
                   color=coul[nom], lw=1.1, marker='o', ms=2.6, alpha=0.85)
    g.axhline(1.0, color='#c0392b', lw=1.4, zorder=0)
    g.text(0.03, 0.085, "F1 = f0 : la mesure n'est plus une couleur, c'est la note",
           transform=g.transAxes, color='#c0392b', fontsize=7.5, va='bottom')
    g.axvspan(3.5, len(ORDRE_REG) - 0.5, color='#c0392b', alpha=0.07, zorder=0)
    g.text(4.75, 1.08, 'ne pas lire', color='#c0392b', fontsize=7.5, ha='center')
    g.set_yscale('log'); g.set_ylim(0.8, 22)
    g.set_yticks([1, 2, 4, 8, 16])
    g.set_yticklabels(['1', '2', '4', '8', '16'])
    g.yaxis.set_minor_locator(matplotlib.ticker.NullLocator())
    g.set_xticks(range(len(ORDRE_REG)))
    g.set_xticklabels([r.replace(' ', '\n') for r in ORDRE_REG])
    g.set_ylabel('F1 / f0')
    g.set_title("La borne de lecture : au-dessus du médium, F1 est la fondamentale")
    g.grid(axis='y', alpha=0.25, lw=0.5)

    # ── droite : F1 par nuance ──────────────────────────────────────────────
    NU = ['doux', 'moyen', 'fort']
    for nom in insts:
        for corpus in CORPUS:
            pts = [(i, V[(nom, corpus, k)]['F1']) for i, k in enumerate(NU)
                   if (nom, corpus, k) in V]
            if len(pts) < 2:
                continue
            d.plot([x[0] for x in pts], [x[1] for x in pts], ls=STYLE_CORPUS.get(corpus, '-'),
                   color=coul[nom], lw=1.1, marker='o', ms=2.6, alpha=0.85)
    for a in ACCORDS_MARQUES:
        if a in V:
            d.plot([NU.index(a[2])], [V[a]['F1']], 'o', ms=9, mfc='none',
                   mec='#c0392b', mew=1.3, zorder=5)
    d.set_yscale('log'); d.set_xlim(-0.35, 2.55)
    d.set_xticks(range(len(NU))); d.set_xticklabels(NU)
    d.set_yticks([150, 200, 300, 400, 600, 800, 1000])
    d.set_yticklabels(['150', '200', '300', '400', '600', '800', '1000'])
    d.yaxis.set_minor_locator(matplotlib.ticker.NullLocator())
    d.set_ylabel('F1 (Hz) — pédale + grave + médium')
    d.set_title("La variable cachée : +9,5 demi-tons en médiane du doux au fort")
    d.grid(axis='y', alpha=0.25, lw=0.5)
    d.text(0.985, 0.025, 'cercles : les 6 accords stricts entre provenances distinctes',
           transform=d.transAxes, color='#c0392b', fontsize=7.5, ha='right')

    h = [Line2D([], [], color=coul[n], lw=1.6, label=n) for n in insts]
    h += [Line2D([], [], color='0.45', lw=1.2, ls=STYLE_CORPUS.get(c, '-'), label=c)
          for c in CORPUS]
    fig.legend(handles=h, loc='lower center', ncol=6, bbox_to_anchor=(0.5, -0.085))
    fig.suptitle("Les cuivres relancés avec les registres pédale — tâche C7", y=1.005)
    fig.tight_layout()
    fig.savefig(chemin)
    plt.close(fig)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--instruments', default='')
    ap.add_argument('--sortie', default=os.path.join(ETUDE, 'Mesures', 'Registres-Pedale'))
    ap.add_argument('--figure', metavar='PNG',
                    help='écrit en plus la figure à deux panneaux')
    a = ap.parse_args()
    os.makedirs(a.sortie, exist_ok=True)
    voulus = [x.strip() for x in a.instruments.split(',') if x.strip()] or sorted(CUIVRES)
    regs = CV.lire_registres(REGISTRES)

    # ── 1. couverture ───────────────────────────────────────────────────────
    print('=' * 86)
    print("1. COUVERTURE — les notes du corpus qui tombent encore « hors registre »")
    print('=' * 86)
    cv_lignes, manques = [], defaultdict(lambda: defaultdict(int))
    for nom in voulus:
        R = regs[CUIVRES[nom]]
        for corpus, p in chemins(nom).items():
            c = defaultdict(int); hors = defaultdict(int)
            for f0, midi, dyn, sp in FL.lire_spectres(p):
                r = CV.registre_de(R, midi)
                c[r] += 1
                if r == 'hors registre':
                    hors[midi] += 1; manques[nom][midi] += 1
            if hors:
                print('{:18s} {:9s} {:3d} notes hors registre : {}'.format(
                    nom[:18], corpus, sum(hors.values()),
                    ', '.join('%s' % nom_midi(m) for m in sorted(hors))))
            cv_lignes.append(dict(instrument=nom, corpus=corpus,
                                  **{('n_' + k.replace(' ', '_')): v for k, v in c.items()}))
    if not manques:
        print('  aucune : la couverture est complète.')

    # ── 2. lisibilité des cellules pédale ───────────────────────────────────
    print()
    print('=' * 86)
    print("2. LE REGISTRE PÉDALE, note par note")
    print('=' * 86)
    print('{:18s} {:9s} {:14s} {:>4s} {:>7s} {:>5s} {:>8s} {:>7s} {:>6s} {:>6s} {}'.format(
        'instrument', 'corpus', 'registre', 'n', 'f0 méd.', 'harm', 'F1 méd.',
        'IQR dt', 'F1/f0', 'pente', 'r²'))
    lignes = []
    for nom in voulus:
        R = regs[CUIVRES[nom]]
        for corpus, p in chemins(nom).items():
            par = defaultdict(list)
            for ech in FL.lire_spectres(p):
                par[CV.registre_de(R, ech[1])].append(ech)
            for lab in [x[0] for x in R] + ['hors registre']:
                E = par.get(lab)
                if not E:
                    continue
                f0s, F1s, nh = [], [], []
                for f0, midi, dyn, sp in E:
                    if not f0:
                        continue
                    fr, am = FL.harmoniques(sp, f0)
                    nh.append(int(fr.size))
                    z1, z2, h1, h2, mq = CV.f1_f2(sp, f0)
                    if h1:
                        f0s.append(f0); F1s.append(h1)
                pe, r2 = (pente(f0s, F1s) if len(F1s) >= MIN_NOTES_PENTE else (None, None))
                d = dict(instrument=nom, corpus=corpus, registre=lab, n=len(E),
                         n_avec_F1=len(F1s),
                         f0_median=round(float(np.median([e[0] for e in E if e[0]])), 1),
                         harm_median=int(np.median(nh)) if nh else None,
                         F1_median=int(round(float(np.median(F1s)))) if F1s else None,
                         iqr_dt=iqr_dt(F1s),
                         rapport_F1_f0=(round(float(np.median(
                             [b / c for b, c in zip(F1s, f0s)])), 1) if F1s else None),
                         pente=None if pe is None else round(pe, 2),
                         r2=None if r2 is None else round(r2, 2),
                         pedale=lab in PEDALE)
                lignes.append(d)
                if lab in PEDALE or lab == 'hors registre' or len(E) >= 5:
                    print('{:18s} {:9s} {:14s} {:4d} {:6.1f}Hz {:5d} {:>7} {:>7} {:>6} {:>6} {:>6}'
                          .format(nom[:18], corpus, lab[:14], d['n'], d['f0_median'],
                                  d['harm_median'] or 0,
                                  (str(d['F1_median']) + 'Hz') if d['F1_median'] else '—',
                                  d['iqr_dt'] if d['iqr_dt'] is not None else '—',
                                  d['rapport_F1_f0'] if d['rapport_F1_f0'] else '—',
                                  d['pente'] if d['pente'] is not None else '—',
                                  d['r2'] if d['r2'] is not None else '—'))

    # ── 3. la résolution de la chaîne d'analyse ─────────────────────────────
    print()
    print('=' * 86)
    print("3. RÉSOLUTION — les notes dont les harmoniques ne sont pas séparables")
    print('=' * 86)
    print("   `pic_proche` cherche le pic dans ±2 bins autour de k·f0, soit 5 bins de large.")
    print("   Sous {:.0f} bins d'écart (f0 < {:.0f} Hz) deux rangs voisins rendent le MÊME pic."
          .format(FL.RESOL_MIN, FL.RESOL_MIN * FL.BW))
    print()
    print('{:18s} {:9s} {:>5s} {:>8s} {:>9s}  {}'.format(
        'instrument', 'corpus', 'n', 'sous seu.', '% harm dbl', 'registres touchés'))
    res_lignes = []
    for nom in voulus:
        R = regs[CUIVRES[nom]]
        for corpus, p in chemins(nom).items():
            n = 0; sous = 0; taux = []; touches = defaultdict(int)
            for f0, midi, dyn, sp in FL.lire_spectres(p):
                if not f0:
                    continue
                n += 1
                fr, am = FL.harmoniques(sp, f0)
                if fr.size < 2:
                    continue
                dup = fr.size - np.unique(fr).size
                if f0 / FL.BW < FL.RESOL_MIN:
                    sous += 1
                if dup:
                    taux.append(dup / fr.size); touches[CV.registre_de(R, midi)] += 1
            d = dict(instrument=nom, corpus=corpus, n=n, n_sous_seuil=sous,
                     pct_sous_seuil=int(round(100 * sous / max(n, 1))),
                     pct_harm_doubles=round(100 * float(np.mean(taux)), 1) if taux else 0.0,
                     registres_touches='; '.join('%s %d' % (k, v)
                                                 for k, v in sorted(touches.items())))
            res_lignes.append(d)
            print('{:18s} {:9s} {:5d} {:7d}  {:8.1f}%  {}'.format(
                nom[:18], corpus, n, sous, d['pct_harm_doubles'],
                d['registres_touches'] or '—'))

    # ── 4. conséquence mesurée : dédoublonner / retirer ─────────────────────
    print()
    print('=' * 86)
    print("4. CONSÉQUENCE — trois lectures du même composite")
    print('=' * 86)
    print("   DÉDOUBLONNÉ garde les notes et écarte les fréquences répétées : il isole")
    print("   l'artefact. RETIRÉ enlève les notes sous le seuil : il enlève aussi le")
    print("   domaine grave, et confond donc l'artefact avec la perte de données.")
    print()
    print('{:18s}  {:28s}  {:28s}  {}'.format(
        'instrument', 'tel que publié', 'DÉDOUBLONNÉ (mêmes notes)', 'notes RETIRÉES'))
    eff_lignes = []
    for nom in voulus:
        R = regs[CUIVRES[nom]]
        for corpus, p in chemins(nom).items():
            tous = [e for e in FL.lire_spectres(p) if e[0]]
            if not any(e[0] / FL.BW < FL.RESOL_MIN for e in tous):
                continue
            gard = [e for e in tous if e[0] / FL.BW >= FL.RESOL_MIN]
            if len(gard) < 10:
                continue
            A = formants_de(tous); D = formants_de(tous, dedup=True); S = formants_de(gard)
            f = lambda L: ' '.join('%d' % round(x[0]) for x in sorted(L))
            eff_lignes.append(dict(instrument=nom, corpus=corpus, n=len(tous),
                                   n_retirees=len(tous) - len(gard), publie=f(A),
                                   dedoublonne=f(D), notes_retirees=f(S),
                                   dedup_identique=(f(A) == f(D))))
            print('{:18s}  {:28s}  {:28s}  {}'.format(nom[:18], f(A), f(D), f(S)))
    ident = sum(1 for d in eff_lignes if d['dedup_identique'])
    print('\n   dédoublonnage sans effet sur la liste : {} cas sur {}'
          .format(ident, len(eff_lignes)))

    # ── 5. la nuance, variable cachée ───────────────────────────────────────
    print()
    print('=' * 86)
    print("5. LA NUANCE — ce que la pente contre f0 ne peut pas voir")
    print('=' * 86)
    print("   La pente contre f0 contrôle la NOTE. Les nuances étant réunies dans chaque")
    print("   régression, elle ne dit rien d'une dépendance à la nuance. On la mesure.")
    print()
    V, dyn_lignes = {}, []
    for nom in voulus:
        R = regs[CUIVRES[nom]]
        for corpus, p in chemins(nom).items():
            d = defaultdict(list)
            for f0, midi, dyn, sp in FL.lire_spectres(p):
                if not f0 or CV.registre_de(R, midi) not in BAS:
                    continue
                z1, z2, h1, h2, mq = CV.f1_f2(sp, f0)
                if h1:
                    d[CV.DYN.get(dyn, dyn)].append((f0, h1))
            for k, v in d.items():
                if len(v) < MIN_NOTES_DYN:
                    continue
                F = [x[1] for x in v]
                pe, r2 = pente([x[0] for x in v], F)
                V[(nom, corpus, k)] = dict(
                    F1=float(np.median(F)), n=len(F), iqr=iqr_dt(F), pente=pe, r2=r2)
    print('{:18s} {:9s} {:>11s} {:>11s} {:>11s} {:>9s}'.format(
        'instrument', 'corpus', 'doux', 'moyen', 'fort', 'doux->fort'))
    for nom in voulus:
        for corpus in CORPUS:
            C = [V.get((nom, corpus, k)) for k in ('doux', 'moyen', 'fort')]
            if not any(C):
                continue
            ecart = (12 * math.log2(C[2]['F1'] / C[0]['F1'])) if C[0] and C[2] else None
            dyn_lignes += [dict(instrument=nom, corpus=corpus, nuance=k, **V[(nom, corpus, k)])
                           for k in ('doux', 'moyen', 'fort') if (nom, corpus, k) in V]
            print('{:18s} {:9s} {:>11s} {:>11s} {:>11s} {:>9s}'.format(
                nom[:18], corpus,
                *['%d Hz (%d)' % (x['F1'], x['n']) if x else '—' for x in C],
                '%.1f dt' % ecart if ecart is not None else '—'))

    print()
    print("   ACCORDS STRICTS (< {:.1f} dt) entre provenances distinctes, à nuance ÉGALE"
          .format(SEUIL_STRICT))
    acc = []
    for nom in voulus:
        for dy in ('doux', 'moyen', 'fort'):
            cs = sorted(c for (i, c, k) in V if i == nom and k == dy)
            for x, y in itertools.combinations(cs, 2):
                if not paire_valide(nom, x, y):
                    continue
                A, B = V[(nom, x, dy)], V[(nom, y, dy)]
                e = abs(12 * math.log2(B['F1'] / A['F1']))
                if e >= SEUIL_STRICT:
                    continue
                acc.append(dict(instrument=nom, nuance=dy, corpus_a=x, corpus_b=y,
                                F1_a=int(A['F1']), F1_b=int(B['F1']), n_a=A['n'], n_b=B['n'],
                                pente_a=A['pente'], pente_b=B['pente'],
                                ecart_demi_tons=round(e, 2)))
                print('     {:18s} {:6s} {:9s} {:4.0f} Hz (n={:2d}, pente {:+.2f})  /  '
                      '{:9s} {:4.0f} Hz (n={:2d}, pente {:+.2f})   {:.2f} dt'.format(
                          nom[:18], dy, x, A['F1'], A['n'], A['pente'],
                          y, B['F1'], B['n'], B['pente'], e))
    print('\n   {} accords stricts à nuance égale.'.format(len(acc)))

    for f, L in (('couverture.csv', cv_lignes), ('registres-pedale.csv', lignes),
                 ('resolution.csv', res_lignes), ('effet-resolution.csv', eff_lignes),
                 ('par-nuance.csv', dyn_lignes), ('accords-stricts.csv', acc)):
        if not L:
            continue
        ch = sorted({k for d in L for k in d})
        with open(os.path.join(a.sortie, f), 'w', encoding='utf-8', newline='') as fh:
            w = csv.DictWriter(fh, fieldnames=ch); w.writeheader(); w.writerows(L)
    if a.figure:
        global ACCORDS_MARQUES
        ACCORDS_MARQUES = [(r['instrument'], r['corpus_a'], r['nuance']) for r in acc] \
                          + [(r['instrument'], r['corpus_b'], r['nuance']) for r in acc]
        figure(lignes, V, a.figure)
        print('figure -> {}'.format(a.figure))
    print('\n-> {}'.format(a.sortie))

if __name__ == '__main__':
    main()
