# -*- coding: utf-8 -*-
"""Branche les cartes simplifiées (fig_simplifiee.py) dans construire_site.py."""
import io, os, sys
P = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'construire_site.py')
t = io.open(P, encoding='utf-8').read()
if 'carte-simplifiee' in t:
    print('déjà branché'); sys.exit(0)

def sub(a, b):
    global t
    assert t.count(a) == 1, a
    t = t.replace(a, b)

# 1. page de synthèse, avant la carte vocalique
sub("""    ('carte-vocalique', 'Carte vocalique', '_CARTE-VOCALIQUE.md'),""",
    """    ('carte-simplifiee', 'Carte vocalique simplifiée — pour l\\'enseignement',
     '_CARTE-SIMPLIFIEE.md'),
    ('carte-vocalique', 'Carte vocalique', '_CARTE-VOCALIQUE.md'),""")

# 2. ses trois figures
sub("""    'carte-vocalique': [""",
    """    'carte-simplifiee': [
        ('simple-synthese-bois.png',
         'Les bois, registre par registre, devant les six zones vocaliques de la première '
         'étude. Barre pleine : zone stable. Hachurée : la zone suit la note. Contour : '
         'non tranché. Barre claire : zone secondaire. Trait gris : tessiture du registre.'),
        ('simple-synthese-cuivres.png',
         'Les cuivres, avec deux lignes par registre — p (doux) et f (fort) : la nuance y '
         'déplace la zone principale de près d\\'une octave.'),
        ('simple-synthese-cordes.png',
         'Les cordes. Résonance de caisse stable à l\\'alto au grave et au violoncelle au '
         'médium ; au-dessus, la zone suit la note.')],
    'carte-vocalique': [""")

# 3. sur chaque page d'instrument, en tête des figures
sub("""        m = MANIF.get(ident, {})
        fg = m.get('composite', {})""",
    """        m = MANIF.get(ident, {})
        if os.path.exists(os.path.join(sortie, 'figures', 'simple-%s.png' % ident)):
            entete.append(figure('simple-%s.png' % ident,
                'Carte vocalique simplifiée : pour chaque registre, la zone principale '
                '(barre pleine : stable ; hachurée : suit la note ; contour : non tranché) '
                'et la zone secondaire, devant les six zones vocaliques de la première '
                'étude. Mêmes mesures que la carte en Bark ci-dessous, lecture en hertz. '
                'Voir la page « Carte vocalique simplifiée ».'))
        fg = m.get('composite', {})""")

io.open(P, 'w', encoding='utf-8').write(t)
print('branché')
