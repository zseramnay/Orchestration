#!/usr/bin/env python3
"""Met les fiches d'instrument à jour avec le statut de chaque formant.

Ce que fait ce script, pour chaque formant retenu d'une fiche :

1. **Calcule sa route** en confrontant les cinq corpus.
   - *route A* : un couple de corpus au MÊME point de grille (< 0,5 demi-ton),
     de provenances distinctes.
   - *indéterminé* : concordance à un demi-ton sans couple strict. Le test de
     permutation de `_REPLICATION.md` montre qu'environ une concordance sur
     quatre à ce seuil est fortuite : on ne peut pas l'affirmer.
   - *1 corpus* : rien ailleurs.
2. **Applique la règle de provenance.** `Yan_Adds-Divers` contient du VSL et le
   corpus VN est du VSL : trois paires de fichiers vérifiées identiques au
   sample près, corrélation jusqu'à 1,00000. Le couple (SOL, VN) ne compte donc
   pas pour les dix instruments issus de Yan_Adds.
3. **Reporte la localisation** mesurée dans `Largeurs/Comparaison/`. Un formant
   dont l'enveloppe vraie donne Q < 2 n'est pas un pic : c'est une région
   d'énergie, et la valeur LPC qu'on lirait ailleurs est une fabrication du
   modèle.
4. **Renomme `bande_db3` en `etendue_residu`.** Ce champ mesure où le résidu de
   la méthode 8 redescend de 3 dB ; dans plus de la moitié des cas c'est la
   butée de la plage de recherche, et il serait lu comme une largeur de bande.
   Les vraies largeurs sont dans `_LARGEURS-DE-BANDE.md`.

Usage
-----
    python3 statut_formants.py --racine ETUDE --lister
    python3 statut_formants.py --racine ETUDE --seulement 14-trompette,13-cor
    python3 statut_formants.py --racine ETUDE --tout --ecrire
Sans `--ecrire`, rien n'est modifié : le script affiche ce qu'il ferait.
"""

import argparse, csv, glob, json, math, os, re, sys, unicodedata, itertools

ORDRE = ['SOL', 'ZSamples', 'Xsamples', 'VO', 'VN']
SEUIL_STRICT = 0.5
SEUIL_PROCHE = 1.05
Q_NON_LOCALISE = 2.0


def slug(s):
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn').lower()
    s = re.sub(r'\b(en)\s+(sib|mib|fa|ut)\b', '', s)
    return re.sub(r'[^a-z]+', '', s)


ALIAS = {'corenfa': 'cor', 'trompetteenut': 'trompette', 'trombone': 'trombonetenor'}


def cle(s):
    k = slug(s)
    return ALIAS.get(k, k)


def lire_fiche(p):
    t = open(p, encoding='utf-8').read()
    m = re.match(r'^---\n(.*?)\n---\n(.*)$', t, re.S)
    if not m:
        return None
    return dict(chemin=p, yaml=m.group(1), corps=m.group(2))


def formants_de_yaml(y, verdicts=('retenu',)):
    out = []
    for ln in y.split('\n'):
        if 'hz:' not in ln:
            continue
        mv = re.search(r'verdict:\s*([\w-]+)', ln)
        mh = re.search(r'hz:\s*(\d+)', ln)
        if mh and mv and any(mv.group(1).startswith(v) for v in verdicts):
            out.append(int(mh.group(1)))
    return sorted(out)


def charger_corpus(racine):
    """Formants retenus par corpus. SOL vient des fiches, le reste des JSON."""
    C = {c: {} for c in ORDRE}
    prov = {}
    for p in glob.glob(os.path.join(racine, 'Instruments', '*', '[0-9][0-9]-*.md')):
        f = lire_fiche(p)
        if not f:
            continue
        nom = re.search(r'^nom:\s*(.+)$', f['yaml'], re.M)
        src = re.search(r'^source_echantillons:\s*(.+)$', f['yaml'], re.M)
        if not nom:
            continue
        k = cle(nom.group(1))
        C['SOL'][k] = formants_de_yaml(f['yaml'])
        prov[k] = 'VSL' if (src and 'Yan_Adds' in src.group(1)) else 'IRCAM'
    for c in ('ZSamples', 'Xsamples', 'VO', 'VN'):
        d = os.path.join(racine, 'Mesures', c)
        for p in sorted(glob.glob(os.path.join(d, '*.json'))):
            b = os.path.basename(p)[:-5]
            if any(x in b for x in (' corde', '+sourdine', '+piombo', '+sord')):
                continue
            try:
                j = json.load(open(p))
            except Exception:
                continue
            C[c][cle(b)] = sorted(int(round(float(x['hz']))) for x in j.get('formants', [])
                                  if str(x.get('verdict', '')).startswith('retenu'))
    return C, prov


def charger_localisation(racine):
    """Q de l'enveloppe vraie par (instrument, formant). Q < 2 : pas un pic."""
    p = os.path.join(racine, 'Largeurs', 'Comparaison', 'confrontation_lpc_enveloppe.csv')
    loc = {}
    if not os.path.exists(p):
        return loc
    for r in csv.DictReader(open(p)):
        try:
            q = float(r['Q_env_v7'])
        except (ValueError, KeyError):
            continue
        loc[(cle(r['instrument']), int(r['formant_hz']))] = q
    return loc


def centres_etablis(k, C, prov):
    """Couples stricts de provenances distinctes, cherchés SANS regroupement
    préalable. Rend [(centre_hz, {corpus}), …].

    Pourquoi sans regroupement : un regroupement séquentiel chaîne depuis le
    premier élément d'une zone, si bien que deux valeurs qui s'apparient
    peuvent tomber dans des zones voisines. C'est ce qui avait fait manquer
    trois formants le 25/10, dont celui de l'alto."""
    pts = [(f, c) for c in ORDRE for g in [C[c].get(k)] if g for f in g]
    etab = []
    for (a, ca), (b, cb) in itertools.combinations(pts, 2):
        if ca == cb or abs(12 * math.log2(a / b)) >= SEUIL_STRICT:
            continue
        if {ca, cb} == {'SOL', 'VN'} and prov.get(k) == 'VSL':
            continue
        etab.append((math.sqrt(a * b), ca, cb))
    etab.sort()
    fus = []
    for h, ca, cb in etab:
        if fus and abs(12 * math.log2(h / fus[-1][0])) < 1.0:
            fus[-1][1] |= {ca, cb}
        else:
            fus.append([h, {ca, cb}])
    return [(h, cs) for h, cs in fus]


def statut(k, f0, C, prov):
    """Rend (route, texte). La paire porte la preuve : on la nomme."""
    pts = [(f, c) for c in ORDRE for g in [C[c].get(k)] if g for f in g]
    zone = [(f, c) for f, c in pts if abs(12 * math.log2(f / f0)) <= SEUIL_PROCHE]
    corpus = sorted({c for _, c in zone})
    for h, cs in centres_etablis(k, C, prov):
        if abs(12 * math.log2(f0 / h)) <= SEUIL_PROCHE:
            e = 12 * math.log2(f0 / h)
            txt = ('**route A** — {} au même point de grille, {:.0f} Hz · zone vue par {}'
                   .format(' + '.join(sorted(cs)), h, ', '.join(sorted(set(corpus) | cs))))
            if abs(e) > SEUIL_STRICT:
                txt += ('. ⚠ La valeur de cette fiche, {} Hz, est à {:+.1f} demi-ton du '
                        'couple : c\'est **{:.0f} Hz** qui est établi'.format(f0, e, h))
            return 'route A', txt
    if len(corpus) < 2:
        return '1 corpus', 'présent dans le seul corpus SOL'
    autres = sorted(c for c in corpus if c != 'SOL')
    bloque = any({ca, cb} == {'SOL', 'VN'} for (a, ca), (b, cb)
                 in itertools.combinations(zone, 2)
                 if ca != cb and abs(12 * math.log2(a / b)) < SEUIL_STRICT)
    if bloque and prov.get(k) == 'VSL':
        return 'indéterminé', ('**indéterminé** — le seul couple strict est SOL + VN, '
                               'or les deux sont du VSL (voir `_CORPUS-CROISES.md`)')
    return 'indéterminé', ('**indéterminé** — concordance à un demi-ton avec {} mais aucun '
                           'couple au même point de grille. Environ une concordance sur '
                           'quatre à ce seuil est fortuite.'.format(', '.join(autres)))


def bloc_markdown(nom, k, fs, C, prov, loc):
    L = ['> ## Statut des formants — révision du 26/10/2026',
         '>',
         '> Confrontation des **cinq corpus**. Provenance de ce corpus : **{}**.'
         .format('VSL / ConTimbre / perso (Yan_Adds)' if prov.get(k) == 'VSL'
                 else 'IRCAM (FullSOL2020)'),
         '>']
    if not fs:
        L.append('> Aucun formant retenu. Si le résultat négatif est répliqué, '
                 'c\'est un acquis — voir `_CORPUS-CROISES.md`.')
        return '\n'.join(L) + '\n'
    L += ['> | Formant | Statut |', '> |---|---|']
    for f0 in fs:
        r, txt = statut(k, f0, C, prov)
        q = loc.get((k, f0))
        if q is not None and q < Q_NON_LOCALISE:
            txt += (' · ⚠ **aucun pic localisé** : Q = {:.1f} sur l\'enveloppe vraie, '
                    'c\'est une région d\'énergie et non une résonance'.format(q))
        elif q is not None:
            txt += ' · pic localisé, Q = {:.1f}'.format(q)
        L.append('> | **{} Hz** | {} |'.format(f0, txt))
    L += ['>',
          '> `etendue_residu` dans l\'en-tête **n\'est pas une largeur de bande** : c\'est '
          'l\'étendue du résidu de la méthode 8, souvent une butée de plage. Les largeurs '
          'mesurées sont dans `_LARGEURS-DE-BANDE.md`.']
    return '\n'.join(L) + '\n'


def traiter(p, C, prov, loc, ecrire=False):
    f = lire_fiche(p)
    if not f:
        return None
    nom = re.search(r'^nom:\s*(.+)$', f['yaml'], re.M)
    if not nom:
        return None
    nom = nom.group(1).strip()
    k = cle(nom)
    fs = formants_de_yaml(f['yaml'])

    y = f['yaml']
    y = y.replace('bande_db3:', 'etendue_residu:')
    if 'provenance:' not in y:
        y = re.sub(r'^(source_echantillons:.*)$',
                   r'\1\nprovenance: ' + ('VSL' if prov.get(k) == 'VSL' else 'IRCAM'),
                   y, count=1, flags=re.M)
    # statut par formant, dans l'en-tête
    for f0 in fs:
        r, _ = statut(k, f0, C, prov)
        y = re.sub(r'(\{\s*hz:\s*%d,)(?![^}]*statut:)' % f0,
                   r'\1 statut: "%s",' % r, y, count=1)

    corps = f['corps']
    # remplace une bannière existante (bloc de lignes '>' juste après le titre)
    m = re.search(r'^(# [^\n]+\n\n)((?:>[^\n]*\n)+)', corps, re.M)
    bloc = bloc_markdown(nom, k, fs, C, prov, loc)
    if m:
        corps = corps[:m.start(2)] + bloc + corps[m.end(2):]
        remplace = True
    else:
        m2 = re.search(r'^(# [^\n]+\n\n)', corps, re.M)
        if m2:
            corps = corps[:m2.end(1)] + bloc + '\n' + corps[m2.end(1):]
        remplace = False

    if ecrire:
        open(p, 'w').write('---\n' + y + '\n---\n' + corps)
    return dict(fiche=os.path.basename(p), nom=nom, provenance=prov.get(k),
                formants=fs, banniere_remplacee=remplace,
                statuts=[statut(k, f0, C, prov)[0] for f0 in fs])


def main():
    a = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    a.add_argument('--racine', required=True, help="dossier Etude-Formants")
    a.add_argument('--seulement')
    a.add_argument('--tout', action='store_true')
    a.add_argument('--lister', action='store_true')
    a.add_argument('--ecrire', action='store_true')
    o = a.parse_args()
    C, prov = charger_corpus(o.racine)
    loc = charger_localisation(o.racine)
    fiches = sorted(glob.glob(os.path.join(o.racine, 'Instruments', '*', '[0-9][0-9]-*.md')))
    if o.lister:
        for p in fiches:
            print(os.path.basename(p)[:-3])
        return
    garde = {x.strip() for x in o.seulement.split(',')} if o.seulement else None
    if not garde and not o.tout:
        sys.exit('préciser --seulement ou --tout')
    n = 0
    for p in fiches:
        b = os.path.basename(p)[:-3]
        if garde and b not in garde:
            continue
        r = traiter(p, C, prov, loc, o.ecrire)
        if not r:
            continue
        n += 1
        print('{:28s} {:6s} {:2d} formants  {}  {}'.format(
            r['fiche'][:28], r['provenance'] or '?', len(r['formants']),
            ' · '.join('{} {}'.format(f, s) for f, s in zip(r['formants'], r['statuts'])) or '—',
            '' if r['banniere_remplacee'] else '(bannière ajoutée)'))
    print('\n{} fiche(s) {}'.format(n, 'écrite(s)' if o.ecrire else
                                    'analysée(s) — rien écrit, ajouter --ecrire'))


if __name__ == '__main__':
    main()
