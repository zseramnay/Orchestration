# Graphique d'origine `voyelles_IPA.svg` : erreurs relevées

Fichier conservé tel quel. Sous-titre d'origine : « F1 et F2 sur axe fréquentiel linéaire d'après Meyer (2009), Giesler (1985) ».

## Ce qui est correct

Les points d'ancrage sur l'axe (extrémités des traits pointillés) tombent tous à leur valeur nominale à ±1 px près (échelle 0–3000 Hz sur 985 px).

## Erreurs

1. **« F1 : hauteur → » est inversé.** F1 augmente avec l'ouverture : plus la voyelle est haute (fermée), plus F1 est bas. Il faudrait « aperture → » ou « ouverture → ». « F2 : antériorité → » est correct.
2. **Le « cluster 420–550 Hz » ne correspond pas aux voyelles mises en évidence.** /ø/ (400 Hz) est hors de la bande mais coloré « cluster », alors que /e/ (400 Hz aussi) ne l'est pas. /ɔ/ (560 Hz) est à la limite et n'est pas inclus.
3. **La bande jaune traverse aussi la rangée F2**, où aucune valeur n'est entre 420 et 550 Hz.

## Choix trompeurs

- **Rangée F1 : cercles très décalés de leurs ancrages.** Exemples : /a/ ancré à 800 Hz, cercle dessiné vers 1120 Hz ; /ɑ/ ancré à 750 Hz, cercle vers 1040 Hz. L'œil lit la position du cercle.
- **/ɑ/ en pointillés sur F1**, sans explication dans la légende.
- **Rectangles colorés de fond** (environ 600–2340 Hz sur F2, 215–780 Hz sur F1) sans signification. Le second exclut /a/ (800 Hz).
- **/ə/ classé « mi-fermée »** : plutôt une voyelle centrale moyenne, souvent réalisée [ø] ou [œ] en français.

## Valeurs discutables

- **Les sources ne correspondent pas aux exemples.** Meyer (2009) et Giesler sont des données d'acoustique musicale issues principalement de locuteurs germanophones, alors que la légende annonce des « exemples en français standard ».
- **/o/ F1 = 460 Hz** est plus ouvert que /e/ et /ø/ (400 Hz), alors que les trois sont mi-fermées. Calliope, homme : 383 Hz.
- **/ɔ/ (F2 900) et /ɑ/ (F2 950)** sont assez bas pour le français. Calliope, homme : /ɔ/ F2 = 998 Hz ; femme : 1180 Hz.
- **Type de voix non indiqué.**
- **Opposition /a/–/ɑ/** pratiquement disparue en français standard contemporain.

Version corrigée : `../graphiques/plan_F1_F2.html` et `../graphiques/F1_seul.html`.
