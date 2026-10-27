# -*- coding: utf-8 -*-
"""Branche les figures de figures.py dans construire_site.py (tâche B3)."""
import io, os, re, sys

P = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'construire_site.py')
t = io.open(P, encoding='utf-8').read()
orig = t
fait = []

def sub(ancre, nouveau, nom):
    global t
    if nouveau.split('\n')[1][:40] in t:
        print('  = %s : déjà présent' % nom); return
    if ancre not in t:
        print('  !! %s : ancre introuvable' % nom); sys.exit(1)
    if t.count(ancre) != 1:
        print('  !! %s : ancre en %d exemplaires' % (nom, t.count(ancre))); sys.exit(1)
    t = t.replace(ancre, nouveau, 1)
    fait.append(nom)

# ── 1. figures de synthèse, par page ───────────────────────────────────────
sub("""CSS_V7 = \"\"\"""",
    """# Les figures sont produites par figures.py, qui écrit aussi figures.json :
# le manifeste dit quelle figure existe et, quand elle manque, POURQUOI — on
# écrit alors la raison au lieu de laisser une image absente.
FIGURES_SYNTHESE = {
    'convergences': [('synthese-convergences.png',
                      'Les formants retenus, par famille. Bande foncée : un unisson '
                      'formantique, plusieurs instruments au même point de grille. '
                      'Bande claire : une région partagée, large d\\'au plus deux '
                      'demi-tons.')],
    'formants-etablis': [('synthese-convergences.png',
                          'Les formants retenus et leurs convergences.')],
    'largeurs': [('synthese-largeurs.png',
                  'La même largeur mesurée deux fois : pôle LPC et enveloppe vraie '
                  'à −3 dB. L\\'écart n\\'est pas du bruit — c\\'est le modèle '
                  'tout-pôle qui répond là où le signal ne dit rien.'),
                 ('synthese-balance-impair-pair.png',
                  'La balance impair − pair sépare la famille de la clarinette du '
                  'reste de l\\'orchestre : ce sont les seules valeurs positives.'),
                 ('synthese-pentes.png',
                  'La décroissance spectrale par instrument.')],
    'replication': [('synthese-rho-corpus.png',
                     'ρ par corpus. L\\'étendue entre corpus mesure ce que le choix '
                     'de la banque, à lui seul, fait varier.')],
}

def figure(nom, legende, profondeur=1):
    racine = '../' * profondeur
    return ('<figure><img src="%sfigures/%s" alt="%s" loading="lazy">'
            '<figcaption>%s</figcaption></figure>\\n'
            % (racine, nom, html.escape(legende[:120]), html.escape(legende)))

CSS_V7 = \"\"\"""", 'FIGURES_SYNTHESE + figure()')

# ── 2. CSS ─────────────────────────────────────────────────────────────────
sub(""".fil { font-size:.9em; color:#666; margin-bottom:18px; }""",
    """figure { margin:26px 0; }
figure img { width:100%; max-width:860px; height:auto; display:block;
              border:1px solid #e0e0e0; border-radius:6px; background:#fff; }
figcaption { font-size:.88em; color:#555; margin-top:8px; max-width:860px;
             line-height:1.45; }
.sans-figure { font-size:.9em; color:#666; border-left:3px solid #ddd;
               padding:8px 14px; margin:22px 0; background:#fafafa; }
.fil { font-size:.9em; color:#666; margin-bottom:18px; }""", 'CSS des figures')

# ── 3. chargement du manifeste ─────────────────────────────────────────────
sub("""    fiches = {}""",
    """    MANIF = {}
    pm = os.path.join(sortie, 'figures', 'figures.json')
    if os.path.exists(pm):
        MANIF = json.load(open(pm, encoding='utf-8'))
    else:
        print('  figures.json absent : lancez figures.py avant construire_site.py')

    fiches = {}""", 'chargement de figures.json')

# ── 4. figures sur la page d'instrument ────────────────────────────────────
sub("""        fil = ('<a href="../index.html">Sommaire</a> › '
               + next(lib for d, lib, rg in FAMILLES if n in rg))""",
    """        m = MANIF.get(ident, {})
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
                '(Calliope, voix d\\'homme). Marqueur creux : cellule non assignable. '
                'Corpus : %s.' % ', '.join(fg['corpus'])))
        elif fg.get('absent'):
            entete.append('<p class="sans-figure"><b>Pas de carte vocalique.</b> %s. '
                          'Ce n\\'est pas un défaut de mesure : l\\'instrument sort du '
                          'triangle vocalique par le bas.</p>\\n'
                          % html.escape(fg['absent']))

        fil = ('<a href="../index.html">Sommaire</a> › '
               + next(lib for d, lib, rg in FAMILLES if n in rg))""",
    "figures sur les pages d'instrument")

# ── 5. figures sur les pages de synthèse ───────────────────────────────────
sub("""        open(os.path.join(sortie, 'syntheses', slug + '.html'), 'w').write(
            page(titre, corps, 1, '', fil))""",
    """        figs = ''.join(figure(f, l) for f, l in FIGURES_SYNTHESE.get(slug, []))
        open(os.path.join(sortie, 'syntheses', slug + '.html'), 'w').write(
            page(titre, figs + corps, 1, '', fil))""",
    'figures sur les pages de synthèse')

if 'import json' not in t:
    t = t.replace('import os', 'import json, os', 1)
    fait.append('import json')

io.open(P, 'w', encoding='utf-8').write(t)
print('construire_site.py : %d modifications' % len(fait))
for f in fait:
    print('   +', f)
