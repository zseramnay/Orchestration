#!/usr/bin/env python3
"""Construit le site v2 de l'étude formantique.

Principe
--------
La **fiche `.md` est la source**. Chaque page d'instrument est le rendu de sa
fiche ; les chiffres de l'en-tête et des sommaires viennent de
`Resultats-v7/formants-v7.csv` et `instruments-v7.csv`, eux-mêmes produits par
`exporter_csv.py` depuis les mêmes fiches. Rien n'est retranscrit à la main :
c'est la retranscription qui avait fait annoncer 20 formants là où il y en a 23.

Pourquoi un constructeur neuf plutôt qu'une adaptation de v6
------------------------------------------------------------
La chaîne v6 (~7 000 lignes) est bâtie sur `get_f()`, qui rend six formants
fixes avec quartiles, écarts-types et largeurs. Ce modèle suppose que tout
instrument possède six formants ordonnés, ce que l'étude a réfuté : le nombre
va de zéro à quatre, et zéro est un résultat. Adapter v6 aurait voulu dire
réécrire sa couche de données et tous ses consommateurs. Le CSS de v6 est en
revanche repris tel quel, pour la continuité visuelle.

Arborescence produite
---------------------
    Site-v2/
      index.html                 sommaire : 22 instruments par famille, avec statut
      methode.html
      bois/ cuivres/ cordes/     une page par instrument
      syntheses/                 une page par sujet
      style.css

Usage
-----
    python3 construire_site.py --racine ETUDE --sortie ETUDE/Site-v2 [--css V6/common.py]
"""

import argparse, csv, glob, html, json, math, os, re, sys
import markdown as md

FAMILLES = [('bois', 'Les bois', range(1, 13)),
            ('cuivres', 'Les cuivres', range(13, 19)),
            ('cordes', 'Les cordes', range(19, 23))]

SYNTHESES = [
    ('convergences', 'Synthèse — le chapitre', None),
    ('formants-etablis', 'Les formants établis', '_CORPUS-CROISES.md'),
    ('replication', 'Le critère de réplication', '_REPLICATION.md'),
    ('carte-vocalique', 'Carte vocalique', '_CARTE-VOCALIQUE.md'),
    ('largeurs', 'Largeurs de bande', '_LARGEURS-DE-BANDE.md'),
    ('resonance-de-caisse', 'Résonances de caisse des cordes', '_RESONANCE-DE-CAISSE.md'),
    ('fenetre-de-redressement', 'La fenêtre de redressement', '_FENETRE-DE-REDRESSEMENT.md'),
    ('suraigu', 'Le plancher de la fondamentale', '_SURAIGU.md'),
    ('appariement-lpc', 'Corroboration par le LPC', '_APPARIEMENT-LPC.md'),
    ('largeur-par-registre', 'Largeur par registre et dynamique', '_LARGEUR-PAR-REGISTRE.md'),
    ('melanges', 'Fusion, renforcement, homogénéité', '_MELANGES.md'),
    ('melanges-octave', "Le doublage à l'octave", '_MELANGES-OCTAVE.md'),
    ('hors-espace', "Les bords de l'espace vocalique", '_HORS-ESPACE.md'),
    ('localisation', 'La localisation, répliquée', '_LOCALISATION.md'),
    ('coupure-benade', 'La coupure du réseau de trous', '_COUPURE-BENADE.md'),
    ('registres-pedale', 'Les registres pédale des cuivres', '_REGISTRES-PEDALE.md'),
    ('coincidences-f2', 'Les deux méthodes se corroborent-elles ?', '_COINCIDENCES-F2.md'),
    ('reference-feminine', 'Voix d\'homme ou voix de femme ?', '_REFERENCE-FEMININE.md'),
    ('verification-vocalique', 'La carte, corpus contre corpus', '_VERIFICATION-VOCALIQUE.md'),
    ('heritage', "L'ancienne étude, et pourquoi", '_VERIFICATION-GRAVES.md'),
]

# Les figures sont produites par figures.py, qui écrit aussi figures.json :
# le manifeste dit quelle figure existe et, quand elle manque, POURQUOI — on
# écrit alors la raison au lieu de laisser une image absente.
FIGURES_SYNTHESE = {
    'convergences': [('synthese-convergences.png',
                      'Les formants retenus, par famille. Bande foncée : un unisson '
                      'formantique, plusieurs instruments au même point de grille. '
                      'Bande claire : une région partagée, large d\'au plus deux '
                      'demi-tons.')],
    'formants-etablis': [('synthese-convergences.png',
                          'Les formants retenus et leurs convergences.')],
    'largeurs': [('synthese-largeurs.png',
                  'La même largeur mesurée deux fois : pôle LPC et enveloppe vraie '
                  'à −3 dB. L\'écart n\'est pas du bruit — c\'est le modèle '
                  'tout-pôle qui répond là où le signal ne dit rien.'),
                 ('synthese-balance-impair-pair.png',
                  'La balance impair − pair sépare la famille de la clarinette du '
                  'reste de l\'orchestre : ce sont les seules valeurs positives.'),
                 ('synthese-pentes.png',
                  'La décroissance spectrale par instrument.')],
    'carte-vocalique': [
        ('vocalique-alto.png',
         'L\'alto, mesuré sur trois corpus de provenances distinctes. C\'est le seul '
         'instrument où la carte vocalique réplique là où la méthode 8 ne retient rien : '
         'la résonance de caisse est trop large pour le redressement à ±3/4 d\'octave.'),
        ('vocalique-hautbois.png',
         'Le hautbois, sur trois corpus. Classé en B : quatre corpus donnent quatre '
         'réponses formantiques incompatibles, alors que la carte vocalique, elle, tient.')],
    'resonance-de-caisse': [('synthese-resonance-caisse.png',
        'F1 contre la note jouée, registre par registre, quatre corpus superposés. '
        'Un nuage horizontal est une résonance ; un nuage le long d\'une diagonale ×k est '
        'un harmonique qui suit la note. La droite porte la pente sur les corpus réunis et '
        'l\'écart des pentes par corpus — c\'est le tableau par corpus qui tranche, jamais '
        'cette droite.')],
    'melanges': [('synthese-melanges.png',
        'À gauche, la loi : l\'écart de couleur entre deux instruments ordonne les quatre '
        'issues du mélange. La fusion croît, l\'homogénéité décroît, le renforcement culmine '
        'au milieu. À droite, les 215 paires, colorées par leur issue majoritaire et '
        'dimensionnées par le nombre de hauteurs communes.')],
    'melanges-octave': [('synthese-melanges-octave.png',
        'À gauche, la loi à l\'unisson (pointillé) et à l\'octave (trait plein). La fusion '
        'culmine entre 3 et 4 Bark puis redescend, tandis que la juxtaposition continue de '
        'monter : c\'est le seuil de fusion de Chistovich. Le pointillé de l\'unisson monte '
        'encore au-delà de 4 Bark, mais sur treize paires seulement — c\'est le manque de '
        'paires éloignées qui cachait le maximum, et l\'octave qui le révèle en peuplant le '
        'haut de l\'échelle. À droite, chaque paire portée par ses deux sens : si le sens ne '
        'changeait rien, tous les points seraient sur la diagonale.')],
    'registres-pedale': [('synthese-registres-pedale.png',
        'À gauche, la borne de lecture : le rapport F1/f0 descend de 10,8 au registre pédale '
        'à 1,0 à l\'aigu, où la mesure cesse d\'être une couleur pour devenir la note. Le '
        'registre pédale, longtemps laissé « hors registre », est au contraire le plus propre. '
        'À droite, la variable que la pente contre f0 ne peut pas voir : la couleur monte de '
        '9,5 demi-tons en médiane du doux au fort, pour les six instruments et les quatre '
        'corpus. Les cercles marquent les six accords stricts entre provenances distinctes, '
        'qui n\'apparaissent qu\'à nuance égale — quatre d\'entre eux au doux.')],
    'replication': [('synthese-rho-corpus.png',
                     'ρ par corpus. L\'étendue entre corpus mesure ce que le choix '
                     'de la banque, à lui seul, fait varier.')],
}

def figure(nom, legende, profondeur=1):
    racine = '../' * profondeur
    return ('<figure><img src="%sfigures/%s" alt="%s" loading="lazy">'
            '<figcaption>%s</figcaption></figure>\n'
            % (racine, nom, html.escape(legende[:120]), html.escape(legende)))

CSS_V7 = """
.bandeau { background:#1a237e; color:#fff; padding:22px 26px; border-radius:8px; margin-bottom:26px; }
.bandeau h1 { color:#fff; border:0; margin:0 0 6px; }
.bandeau p { margin:2px 0; color:#c5cae9; }
figure { margin:26px 0; }
figure img { width:100%; max-width:860px; height:auto; display:block;
              border:1px solid #e0e0e0; border-radius:6px; background:#fff; }
figcaption { font-size:.88em; color:#555; margin-top:8px; max-width:860px;
             line-height:1.45; }
.sans-figure { font-size:.9em; color:#666; border-left:3px solid #ddd;
               padding:8px 14px; margin:22px 0; background:#fafafa; }
.fil { font-size:.9em; color:#666; margin-bottom:18px; }
.fil a { color:#1a237e; }
table.sommaire { width:100%; border-collapse:collapse; background:#fff; margin:14px 0 30px; }
table.sommaire th { background:#1a237e; color:#fff; padding:8px 10px; text-align:left; font-size:.9em; }
table.sommaire td { padding:7px 10px; border-bottom:1px solid #e0e0e0; font-size:.92em; }
table.sommaire tr:hover td { background:#f0f3ff; }
.pastille { display:inline-block; padding:1px 8px; border-radius:10px; font-size:.82em; font-weight:600; }
.route-a   { background:#1b5e20; color:#fff; }
.indet     { background:#f9a825; color:#222; }
.uncorpus  { background:#eceff1; color:#555; }
.aucun     { background:#b71c1c; color:#fff; }
.cartes { display:grid; grid-template-columns:repeat(auto-fit,minmax(280px,1fr)); gap:16px; }
.carte { background:#fff; border:1px solid #ddd; border-left:5px solid #2e7d32; border-radius:6px; padding:16px 18px; }
.carte h3 { margin:0 0 6px; }
.carte p { margin:4px 0; font-size:.92em; color:#555; }
blockquote { background:#fffde7; border-left:4px solid #f9a825; margin:16px 0; padding:10px 16px; }
blockquote h2, blockquote h3 { margin-top:4px; }
code { background:#eceff1; padding:1px 5px; border-radius:3px; font-size:.9em; }
footer { margin-top:50px; padding-top:14px; border-top:1px solid #ddd; font-size:.82em; color:#777; }
"""


def css_de_v6(chemin):
    if not chemin or not os.path.exists(chemin):
        return ''
    s = open(chemin, encoding='utf-8').read()
    m = re.search(r'CSS_GLOBAL\s*=\s*("""|\'\'\')(.*?)\1', s, re.S)
    return m.group(2) if m else ''


def lire_fiche(p):
    t = open(p, encoding='utf-8').read()
    m = re.match(r'^---\n(.*?)\n---\n(.*)$', t, re.S)
    if not m:
        return None
    return m.group(1), m.group(2)


def champ(y, nom, defaut=''):
    m = re.search(r'^%s:\s*(.+?)\s*(?:#.*)?$' % re.escape(nom), y, re.M)
    return m.group(1).strip() if m else defaut


def rendre(corps):
    return md.markdown(corps, extensions=['tables', 'attr_list', 'sane_lists'])


def page(titre, corps, profondeur=1, sous_titre='', fil=''):
    racine = '../' * profondeur
    return ('<!DOCTYPE html>\n<html lang="fr">\n<head>\n<meta charset="UTF-8">\n'
            '<meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
            f'<title>{html.escape(titre)}</title>\n'
            f'<link rel="stylesheet" href="{racine}style.css">\n</head>\n<body>\n'
            f'<div class="bandeau"><h1>{html.escape(titre)}</h1>'
            + (f'<p>{html.escape(sous_titre)}</p>' if sous_titre else '')
            + '</div>\n'
            + (f'<p class="fil">{fil}</p>\n' if fil else '')
            + corps
            + '\n<footer>Référence formantique de l\'orchestre — Yan Maresz · '
              'page produite par <code>construire_site.py</code> depuis les fiches '
              'et <code>Resultats-v7/</code>.</footer>\n</body>\n</html>\n')


def pastille(n, classe, libelle):
    return f'<span class="pastille {classe}">{n} {libelle}</span>' if int(n or 0) else ''


def construire(racine, sortie, css_v6=''):
    os.makedirs(sortie, exist_ok=True)
    open(os.path.join(sortie, 'style.css'), 'w').write(css_v6 + CSS_V7)

    INST = {r['id']: r for r in csv.DictReader(
        open(os.path.join(racine, 'Resultats-v7', 'instruments-v7.csv')))}
    FORM = {}
    for r in csv.DictReader(open(os.path.join(racine, 'Resultats-v7', 'formants-v7.csv'))):
        FORM.setdefault(r['id'], []).append(r)

    MANIF = {}
    pm = os.path.join(sortie, 'figures', 'figures.json')
    if os.path.exists(pm):
        MANIF = json.load(open(pm, encoding='utf-8'))
    else:
        print('  figures.json absent : lancez figures.py avant construire_site.py')

    fiches = {}
    for p in glob.glob(os.path.join(racine, 'Instruments', '*', '[0-9][0-9]-*.md')):
        n = int(os.path.basename(p)[:2])
        fiches[n] = p

    # ---------- pages d'instrument ----------
    liens = {}
    for n, p in sorted(fiches.items()):
        y, corps = lire_fiche(p)
        ident = champ(y, 'id')
        dossier = next(d for d, _, rg in FAMILLES if n in rg)
        os.makedirs(os.path.join(sortie, dossier), exist_ok=True)
        base = os.path.basename(p)[:-3] + '.html'
        liens[n] = (dossier + '/' + base, champ(y, 'nom'), ident)

        fs = FORM.get(ident, [])
        ra = [f for f in fs if f['statut'] == 'route A']
        entete = ['<table class="sommaire"><tr><th>Classe</th><th>Provenance</th>'
                  '<th>Échantillons</th><th>Tessiture</th><th>Formants établis</th></tr>'
                  f'<tr><td>{html.escape(champ(y, "classe"))}</td>'
                  f'<td>{html.escape(champ(y, "provenance"))}</td>'
                  f'<td>{html.escape(champ(y, "n_echantillons"))}</td>'
                  f'<td>{html.escape(champ(y, "tessiture_mesuree"))}</td>'
                  f'<td>{len(ra)}</td></tr></table>']
        if ra:
            entete.append('<table class="sommaire"><tr><th>Formant</th><th>Note</th>'
                          '<th>Corpus</th><th>Localisation</th></tr>')
            for f in ra:
                entete.append(f'<tr><td><b>{f["hz"]} Hz</b></td><td>{f["note"]}</td>'
                              f'<td>{html.escape(f["corpus"])}</td>'
                              f'<td>{html.escape(f["localisation"])}</td></tr>')
            entete.append('</table>')
        m = MANIF.get(ident, {})
        fg = m.get('composite', {})
        if fg.get('fichier'):
            entete.append(figure(fg['fichier'],
                'Enveloppe composite : le résidu local médian, une courbe par corpus '
                '(%s). Un formant retenu est un maximum que deux corpus de provenances '
                'distinctes placent au même point de grille.' % ', '.join(fg['corpus'])))
        fg = m.get('vocalique', {})
        if fg.get('fichier'):
            entete.append(figure(fg['fichier'],
                'Carte vocalique : chaque cellule est un registre à une dynamique, '
                'placée dans le plan (F1, F2) en Bark contre les voyelles du français '
                '(Calliope, voix d\'homme). Marqueur creux : cellule non assignable. '
                'Le nom de la voyelle dépend de la référence : en voix de femme, '
                '41 %% des cellules en changent et 11 instruments sur 20 changent de '
                'voyelle dominante. Le plancher de l\'espace, lui, ne bouge pas. '
                'Corpus : %s.' % ', '.join(fg['corpus'])))
        elif fg.get('absent'):
            entete.append('<p class="sans-figure"><b>Pas de carte vocalique.</b> %s. '
                          'Ce n\'est pas un défaut de mesure : l\'instrument sort du '
                          'triangle vocalique par le bas.</p>\n'
                          % html.escape(fg['absent']))

        fil = ('<a href="../index.html">Sommaire</a> › '
               + next(lib for d, lib, rg in FAMILLES if n in rg))
        open(os.path.join(sortie, dossier, base), 'w').write(
            page(champ(y, 'nom'), ''.join(entete) + rendre(corps), 1,
                 champ(y, 'nom_en'), fil))

    # ---------- sommaire ----------
    c = ['<p>Vingt-deux instruments, cinq corpus indépendants. '
         'La pastille dit par quelle route chaque formant est établi.</p>']
    for d, libelle, rg in FAMILLES:
        c.append(f'<h2>{libelle}</h2><table class="sommaire">'
                 '<tr><th>Instrument</th><th>Classe</th><th>Provenance</th>'
                 '<th>Formants</th><th>ρ (SOL → autres corpus)</th></tr>')
        for n in rg:
            if n not in liens:
                continue
            lien, nom, ident = liens[n]
            i = INST.get(ident, {})
            fs = FORM.get(ident, [])
            p = (pastille(sum(1 for f in fs if f['statut'] == 'route A'), 'route-a', 'route A')
                 + ' ' + pastille(i.get('n_indetermine', 0), 'indet', 'indéterminé')
                 + ' ' + pastille(i.get('n_un_corpus', 0), 'uncorpus', '1 corpus'))
            if not p.strip():
                p = '<span class="pastille aucun">aucun formant</span>'
            rhos = [i.get(k, '') for k in ('rho_SOL', 'rho_ZSamples', 'rho_Xsamples',
                                           'rho_VO', 'rho_VN')]
            rhos = ' · '.join(x for x in rhos if x)
            c.append(f'<tr><td><a href="{lien}">{html.escape(nom)}</a></td>'
                     f'<td>{html.escape(i.get("classe", ""))}</td>'
                     f'<td>{html.escape(i.get("provenance", ""))}</td>'
                     f'<td>{p}</td><td style="font-size:.85em">{rhos}</td></tr>')
        c.append('</table>')
    c.append('<h2>Synthèses</h2><div class="cartes">')
    for slug, titre, src in SYNTHESES:
        c.append(f'<div class="carte"><h3><a href="syntheses/{slug}.html">'
                 f'{html.escape(titre)}</a></h3>'
                 + (f'<p>d\'après <code>{src}</code></p>' if src else
                    '<p>généré depuis <code>formants-v7.csv</code></p>') + '</div>')
    c.append('<div class="carte"><h3><a href="methode.html">Méthode</a></h3>'
             '<p>critères de rejet, typologie, limites</p></div></div>')
    open(os.path.join(sortie, 'index.html'), 'w').write(
        page("Référence formantique de l'orchestre", ''.join(c), 0,
             '22 instruments · 5 corpus · mesures de 2026'))

    # ---------- méthode ----------
    y, corps = lire_fiche(os.path.join(racine, 'Instruments', '00-METHODE.md')) or ('', '')
    if corps:
        open(os.path.join(sortie, 'methode.html'), 'w').write(
            page('Méthode', rendre(corps), 0, '',
                 '<a href="index.html">Sommaire</a> › Méthode'))

    # ---------- synthèses ----------
    os.makedirs(os.path.join(sortie, 'syntheses'), exist_ok=True)
    for slug, titre, src in SYNTHESES:
        fil = '<a href="../index.html">Sommaire</a> › Synthèses'
        if src is None:
            chap = os.path.join(racine, 'Instruments', '_CHAPITRE-SYNTHESE.md')
            tableaux = page_convergences(FORM, INST)
            if os.path.exists(chap):
                lu = lire_fiche(chap)
                texte = lu[1] if lu else open(chap, encoding='utf-8').read()
                # le marqueur reçoit les tableaux, qui restent produits depuis le CSV
                if '<!-- TABLEAUX-CONVERGENCES -->' in texte:
                    av, ap = texte.split('<!-- TABLEAUX-CONVERGENCES -->', 1)
                    corps = rendre(av) + tableaux + rendre(ap)
                else:
                    corps = rendre(texte) + tableaux
            else:
                corps = tableaux
        else:
            p = os.path.join(racine, 'Instruments', src)
            if not os.path.exists(p):
                continue
            lu = lire_fiche(p)
            corps = rendre(lu[1] if lu else open(p, encoding='utf-8').read())
        figs = ''.join(figure(f, l) for f, l in FIGURES_SYNTHESE.get(slug, []))
        open(os.path.join(sortie, 'syntheses', slug + '.html'), 'w').write(
            page(titre, figs + corps, 1, '', fil))
    return liens


def page_convergences(FORM, INST):
    """Le chapitre : les convergences, construites depuis le CSV.

    Deux tableaux, et la distinction compte. **Au même point de grille**, c'est
    un unisson formantique. **À un demi-ton**, c'est une région partagée : la
    résolution de la méthode étant d'un demi-ton, confondre les deux reviendrait
    à refaire l'erreur des « convergences à Δ = 11 Hz » de l'ancienne étude.

    Les régions sont obtenues en regroupant les zones établies — écart entre
    valeurs consécutives ≤ 1,05 demi-ton ET étendue totale ≤ 2 demi-tons, pour
    qu'une chaîne ne s'étende pas indéfiniment. Sans ce regroupement, un
    appariement deux à deux produit des lignes transitives : 2 216–2 225,
    2 225–2 348 et 2 216–2 348 décrivent une seule région.
    """
    par_hz = {}
    for fs in FORM.values():
        for f in fs:
            if f['statut'] == 'route A':
                par_hz.setdefault(int(f['hz']), set()).add(f['instrument'])
    hzs = sorted(par_hz)

    exact = [(h, sorted(par_hz[h])) for h in hzs if len(par_hz[h]) > 1]

    groupes = []
    for h in hzs:
        if groupes and (12 * math.log2(h / groupes[-1][-1]) <= 1.05
                        and 12 * math.log2(h / groupes[-1][0]) <= 2.0):
            groupes[-1].append(h)
        else:
            groupes.append([h])
    regions = []
    for g in groupes:
        ins = set()
        for h in g:
            ins |= par_hz[h]
        if len(ins) > 1 and not (len(g) == 1):
            regions.append((g[0], g[-1], sorted(ins)))

    note_de = {}
    for fs in FORM.values():
        for f in fs:
            note_de[int(f['hz'])] = f['note']

    c = ['<p>Chaque membre de ces zones est validé <b>indépendamment</b>, par un couple de '
         'corpus de provenances distinctes au même point de grille. À ne pas confondre avec '
         'les « convergences » de l\'ancienne étude, qui comparaient des médianes de '
         'tessiture tombant dans des bins de FFT voisins.</p>']
    c.append('<h2>Au même point de grille</h2>'
             '<p>Les seules à présenter comme des <b>unissons formantiques</b>.</p>'
             '<table class="sommaire"><tr><th>Zone</th><th>Note</th><th>Instruments</th></tr>')
    for h, ins in exact:
        c.append(f'<tr><td><b>{h} Hz</b></td><td>{note_de.get(h, "")}</td>'
                 f'<td>{html.escape(" · ".join(ins))}</td></tr>')
    c.append('</table>')
    c.append('<h2>Régions partagées, à un demi-ton près</h2>'
             '<p>Réelles, mais à présenter comme des <b>régions</b> et non comme des '
             'unissons : la résolution de la méthode est d\'un demi-ton.</p>'
             '<table class="sommaire"><tr><th>Région</th><th>Étendue</th>'
             '<th>Instruments</th></tr>')
    for a, b, ins in regions:
        et = '0' if a == b else '{:.1f} demi-ton'.format(12 * math.log2(b / a))
        c.append(f'<tr><td>{a}–{b} Hz</td><td>{et}</td>'
                 f'<td>{html.escape(" · ".join(ins))}</td></tr>')
    c.append('</table>')
    return ''.join(c)


def main():
    a = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    a.add_argument('--racine', required=True)
    a.add_argument('--sortie', required=True)
    a.add_argument('--css')
    o = a.parse_args()
    liens = construire(o.racine, o.sortie, css_de_v6(o.css))
    n = sum(len(fs) for _, _, fs in os.walk(o.sortie) if True)
    print('{} pages d\'instrument + sommaire + méthode + {} synthèses'
          .format(len(liens), len(SYNTHESES)))
    print('→', o.sortie)


if __name__ == '__main__':
    main()
