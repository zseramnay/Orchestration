# -*- coding: utf-8 -*-
"""
fenetre_redressement.py — tâche C2 : élargir la fenêtre de `residu_local()`.

Le diagnostic publié, et pourquoi il est faux
---------------------------------------------
`residu_local()` porte dans sa docstring : « aveugle aux résonances plus larges que
~1,5 octave ». On en a déduit que la résonance de caisse de l'alto (375–395 Hz, établie en
C1 sur quatre corpus) échappait à la méthode 8 parce qu'elle est trop large, et que l'élargir
la ferait apparaître.

La vraie raison est ailleurs, et elle est arithmétique. Un harmonique n'entre dans
l'enveloppe composite que si sa fenêtre de ±w octave contient **au moins trois harmoniques de
la même note** (`q >= 3` dans `composite()`). Or le nombre d'harmoniques dans
[k/2^w, k·2^w] croît avec le rang k : aux rangs les plus bas la fenêtre est trop étroite.

    ±0,50 octave  ->  rang le plus bas admis : 4  ->  plancher = 4 x f0_min
    ±0,75 octave  ->  rang le plus bas admis : 3  ->  plancher = 3 x f0_min
    ±1,00 à ±1,50 ->  rang le plus bas admis : 2  ->  plancher = 2 x f0_min
    ±2,00 octave  ->  rang le plus bas admis : 1  ->  plancher = 1 x f0_min

Pour l'alto, f0_min = 131 Hz : le plancher vaut 393 Hz à ±3/4 d'octave. La résonance mesurée
est à 374–395 Hz. **Elle n'est pas redressée, elle est hors du domaine calculé.** L'enveloppe
composite ne porte aucune valeur en dessous — ni pour l'alto, ni pour le violoncelle.

Élargir la fenêtre répond donc bien au problème, mais par un autre mécanisme que celui
annoncé : cela **abaisse le plancher du composite**, au lieu de « voir plus large ».

Ce que le script mesure
-----------------------
1. Le plancher réel du composite à chaque largeur — vérification de l'arithmétique ci-dessus.
2. Ce que `trouver_formants()` trouve à chaque largeur, et en particulier dans la zone que la
   méthode 12 a établie en C1.
3. **Le coût.** Une fenêtre large cesse de redresser : la médiane locale finit par couvrir
   toute la tessiture harmonique de la note, et le résidu ne mesure plus qu'une pente
   spectrale. On le quantifie par la corrélation entre résidu et log-fréquence — nulle pour un
   bon redressement, franchement négative pour une fenêtre trop large.
4. **Les dégâts.** Combien des formants déjà établis par route A survivent à chaque largeur.

Le test de rejet, adapté
------------------------
`robustesse_fenetre()` compare à ±0,5 et ±1,0 octave. Appliqué tel quel à un pic découvert à
±1,5, il le rejetterait d'office : à ±0,5 le plancher est à 4 x f0_min, donc le pic ne PEUT
pas y exister. On déplace donc les fenêtres de comparaison avec la référence : pour une
référence w, on teste à w − 0,25 et w + 0,25. Le test 1, l'indépendance au rang harmonique,
reste inchangé — c'est lui qui porte le contenu.

Usage
    python3 fenetre_redressement.py --instruments Alto,Violoncelle --cibles 384,219
    python3 fenetre_redressement.py --degats        # survie des formants établis
"""
import os, sys, json, math, csv, argparse
import numpy as np

ICI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ICI)
import formants_lib as FL

RACINE = os.path.dirname(os.path.dirname(ICI))
ETUDE = os.path.join(RACINE, 'Etude-Formants')
SOL_OUT = os.path.expanduser('~/work/out')
LARGEURS = (0.75, 1.0, 1.25, 1.5, 2.0)

# nom de fiche -> (fichier SOL sans extension, nom dans Mesures/<corpus>/)
NOMS = {
    'Alto': 'Viola', 'Violoncelle': 'Violoncello', 'Violon': 'Violin',
    'Contrebasse': 'Contrabass', 'Hautbois': 'Oboe', 'Basson': 'Bassoon',
    'Cor anglais': 'EnglishHorn', 'Clarinette en Sib': 'Clarinet_Bb',
    'Flûte': 'Flute', 'Petite flûte': 'Piccolo', 'Flûte basse': 'Bass-Flute',
    'Flûte contrebasse': 'Contrabass-Flute', 'Contrebasson': 'Contrebasson',
    'Clarinette en Mib': 'Clarinet-Eb', 'Clarinette basse en Sib': 'Bass-Clarinet-Bb',
    'Clarinette contrebasse en Sib': 'Contrabass-Clarinet-Bb',
    'Cor en Fa': 'Horn', 'Trompette en Ut': 'Trumpet_C', 'Trombone ténor': 'Trombone',
    'Trombone basse': 'Bass-Trombone', 'Tuba basse': 'Bass_Tuba',
    'Tuba contrebasse': 'Contrabass-Tuba',
}

def chemins(nom):
    """{corpus: fichier de spectres} pour cet instrument."""
    out = {}
    p = os.path.join(SOL_OUT, NOMS.get(nom, nom) + '.txt')
    if os.path.exists(p):
        out['SOL'] = p
    for c in ('ZSamples', 'Xsamples', 'VO', 'VN'):
        p = os.path.join(ETUDE, 'Mesures', c, nom + '.spectres.txt')
        if os.path.exists(p):
            out[c] = p
    return out

# ─────────────────────────── mesures par largeur ───────────────────────────
def pente_residuelle(med):
    """Corrélation entre résidu médian et log-fréquence. Un redressement correct la
    laisse proche de zéro ; une fenêtre trop large la rend franchement négative, signe
    que le résidu ne mesure plus qu'une pente spectrale."""
    m = ~np.isnan(med)
    if m.sum() < 8:
        return None
    x = np.log2(FL.CENTRES[m]); y = med[m]
    x = x - x.mean(); y = y - y.mean()
    d = math.sqrt(float((x * x).sum()) * float((y * y).sum()))
    return None if d <= 0 else round(float((x * y).sum()) / d, 2)

def robustesse_adaptee(ech, fc, w, delta=0.25):
    """Test 2, mais avec des fenêtres de comparaison centrées sur la référence w."""
    for hw in (w - delta, w + delta):
        if hw <= 0:
            continue
        m, n, p = FL.composite(ech, hw)
        pics = [a for a, b, c, d, e in FL.trouver_formants(m, n, p)]
        if not any(abs(math.log2(x / fc)) <= 1 / 3 for x in pics):
            return False
    return True

def analyser(nom, corpus, chemin, cibles, rng):
    ech = FL.lire_spectres(chemin)
    if not ech:
        return []
    f0min = min(f0 for f0, m, d, s in ech if f0)
    notes = None
    lignes = []
    for w in LARGEURS:
        med, nn, npt = FL.composite(ech, w)
        ok = ~np.isnan(med)
        fmin = float(FL.CENTRES[ok][0]) if ok.any() else None
        pics = FL.trouver_formants(med, nn, npt)
        d = {'instrument': nom, 'corpus': corpus, 'fenetre': w, 'n_ech': len(ech),
             'f0_min': round(f0min, 1),
             'plancher_hz': None if fmin is None else int(round(fmin)),
             'plancher_sur_f0': None if fmin is None else round(fmin / f0min, 2),
             'bandes_peuplees': int(ok.sum()),
             'pente_residuelle': pente_residuelle(med),
             'formants': [int(round(a)) for a, b, c, e, f in pics],
             'residus': [round(b, 2) for a, b, c, e, f in pics]}
        # résidu dans la zone cible, et verdict si un pic y tombe
        for cible in cibles:
            i = int(np.argmin(np.abs(FL.CENTRES - cible)))
            v = med[i]
            d['residu_%d' % cible] = None if np.isnan(v) else round(float(v), 2)
            proche = [a for a, b, c, e, f in pics if abs(12 * math.log2(a / cible)) <= 2]
            d['pic_%d' % cible] = int(round(proche[0])) if proche else None
            if proche:
                if notes is None:
                    notes = FL.par_note(ech, w)
                km = FL._rang_median(notes, proche[0])
                G = ([FL._bootstrap_groupe(notes, proche[0], 1, km, rng=rng),
                      FL._bootstrap_groupe(notes, proche[0], km + 1, 99, rng=rng)]
                     if km else [])
                d['rang_ok_%d' % cible] = sum(1 for g in G if g.get('passe'))
                d['rang_testable_%d' % cible] = sum(1 for g in G if g.get('testable'))
                d['robuste_%d' % cible] = robustesse_adaptee(ech, proche[0], w)
        lignes.append(d)
        notes = None    # par_note dépend de la fenêtre
    return lignes

# ──────────────────────────── dégâts collatéraux ───────────────────────────
def degats(sortie):
    """Pour chaque formant déjà établi, dit s'il se retrouve encore à moins de 1/3
    d'octave quand on élargit la fenêtre. C'est le prix à payer, et c'est lui qui
    décide si la fenêtre de référence peut changer."""
    with open(os.path.join(ETUDE, 'Resultats-v7', 'formants-v7.csv'),
              encoding='utf-8') as f:
        R = [r for r in csv.DictReader(f) if r['statut'] in ('route A', '1 corpus')]
    par_inst = {}
    for r in R:
        par_inst.setdefault(r['instrument'], []).append((float(r['hz']), r['statut']))

    survie = {w: {'route A': [0, 0], '1 corpus': [0, 0]} for w in LARGEURS}
    detail = []
    print('{:30s} {:9s} {:>7s}  {}'.format('instrument', 'corpus', 'fenêtre',
                                           'formants établis retrouvés'))
    for nom, fs in sorted(par_inst.items()):
        ch = chemins(nom)
        if not ch:
            print('{:30s} — aucun fichier de spectres'.format(nom)); continue
        for corpus, p in ch.items():
            ech = FL.lire_spectres(p)
            if not ech:
                continue
            for w in LARGEURS:
                m, n, q = FL.composite(ech, w)
                pics = [a for a, b, c, d, e in FL.trouver_formants(m, n, q)]
                trouve = []
                for hz, st in fs:
                    ok = any(abs(math.log2(x / hz)) <= 1 / 3 for x in pics)
                    trouve.append((hz, st, ok))
                    if corpus == 'SOL' or True:
                        survie[w][st][1] += 1
                        survie[w][st][0] += 1 if ok else 0
                detail.append({'instrument': nom, 'corpus': corpus, 'fenetre': w,
                               'retrouves': sum(1 for h, s, o in trouve if o),
                               'total': len(trouve),
                               'perdus': [int(h) for h, s, o in trouve if not o]})
                print('{:30s} {:9s} {:7.2f}  {}/{}{}'.format(
                    nom, corpus, w, sum(1 for h, s, o in trouve if o), len(trouve),
                    '   perdus : ' + ' · '.join(str(int(h)) for h, s, o in trouve if not o)
                    if any(not o for h, s, o in trouve) else ''))
    print('\n{:>7s}  {:>18s}  {:>18s}'.format('fenêtre', 'route A retrouvés',
                                              '1 corpus retrouvés'))
    for w in LARGEURS:
        a1, a2 = survie[w]['route A']; b1, b2 = survie[w]['1 corpus']
        print('{:7.2f}  {:>9s} {:5.0f}%  {:>9s} {:5.0f}%'.format(
            w, '%d/%d' % (a1, a2), 100 * a1 / max(a2, 1),
            '%d/%d' % (b1, b2), 100 * b1 / max(b2, 1)))
    with open(os.path.join(sortie, 'degats-fenetre.json'), 'w', encoding='utf-8') as f:
        json.dump({'detail': detail,
                   'survie': {str(w): survie[w] for w in LARGEURS}}, f,
                  ensure_ascii=False, indent=1)

# ───────────────────────────────── pilote ──────────────────────────────────
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--instruments', default='Alto,Violoncelle,Violon,Contrebasse,Hautbois')
    ap.add_argument('--cibles', default='384,219')
    ap.add_argument('--sortie', default=os.path.join(ETUDE, 'Mesures', 'Fenetre-redressement'))
    ap.add_argument('--degats', action='store_true',
                    help='survie des formants déjà établis, au lieu de l\'exploration')
    a = ap.parse_args()
    cibles = [int(x) for x in a.cibles.split(',') if x.strip()]
    os.makedirs(a.sortie, exist_ok=True)
    rng = np.random.default_rng(20261027)
    if a.degats:
        degats(a.sortie)
        return

    tout = []
    for nom in [x.strip() for x in a.instruments.split(',') if x.strip()]:
        ch = chemins(nom)
        print('\n' + '=' * 96)
        print('{}   corpus : {}'.format(nom, ', '.join(ch) or 'aucun'))
        print('=' * 96)
        print('{:9s} {:>6s} {:>5s} {:>9s} {:>7s} {:>7s}  {:28s} {}'.format(
            'corpus', 'fenêtre', 'f0min', 'plancher', 'bandes', 'pente r',
            'formants trouvés (Hz)', 'zone cible'))
        for corpus, p in ch.items():
            for d in analyser(nom, corpus, p, cibles, rng):
                tout.append(d)
                zc = []
                for cible in cibles:
                    r = d.get('residu_%d' % cible)
                    pic = d.get('pic_%d' % cible)
                    if r is None and pic is None:
                        continue
                    s = '{}:{}'.format(cible, '—' if r is None else '%.1f dB' % r)
                    if pic:
                        s += ' PIC {} rang {}/{} rob {}'.format(
                            pic, d.get('rang_ok_%d' % cible),
                            d.get('rang_testable_%d' % cible),
                            'oui' if d.get('robuste_%d' % cible) else 'non')
                    zc.append(s)
                print('{:9s} {:6.2f} {:5.0f} {:>6}Hz{:>3} {:7d} {:>7}  {:28s} {}'.format(
                    d['corpus'], d['fenetre'], d['f0_min'],
                    d['plancher_hz'] if d['plancher_hz'] else '—',
                    '×%.1f' % d['plancher_sur_f0'] if d['plancher_sur_f0'] else '',
                    d['bandes_peuplees'],
                    '—' if d['pente_residuelle'] is None else d['pente_residuelle'],
                    ' · '.join(str(x) for x in d['formants'][:5]) or '(aucun)',
                    ' | '.join(zc)))

    with open(os.path.join(a.sortie, 'fenetre-redressement.json'), 'w',
              encoding='utf-8') as f:
        json.dump(tout, f, ensure_ascii=False, indent=1)
    print('\n{} lignes -> {}'.format(len(tout), a.sortie))

if __name__ == '__main__':
    main()
