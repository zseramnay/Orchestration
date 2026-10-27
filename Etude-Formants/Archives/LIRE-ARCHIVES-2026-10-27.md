# Ce que contiennent ces deux dossiers

## `scripts-morts/tests_de_rejet.py`

Archivé le 27/10/2026. Le script pointait encore vers un répertoire de travail temporaire
d'une session de mai 2026 (`/tmp/claude-0/…/scratchpad/composite.py`) et ne tournait plus
depuis. Il était trompeur : on pouvait croire que les tests de rejet vivaient là.

**Ses deux tests vivent désormais dans `Scripts/v7-formants-2026/formants_lib.py`**, dans
`verdict()` :

- **test 1**, indépendance au rang harmonique — `_rang_median()` et `_bootstrap_groupe()`.
  Le découpage se fait au rang médian et non à un seuil fixe, ce qui préserve le test pour les
  instruments très graves dont tout le signal est à rang élevé ;
- **test 2**, robustesse de fenêtre — `robustesse_fenetre()`.

## `halfwin-0.75-2026-10-27/`

L'état complet de l'étude **avant** le passage de la fenêtre de redressement de ±3/4 à
±1 octave : les mesures des quatre corpus à fichiers par instrument, les deux CSV de
`Resultats-v7/`, et les 22 fiches.

La raison du changement, les gains et les coûts mesurés sont dans
`Instruments/_FENETRE-DE-REDRESSEMENT.md`. En deux lignes : à ±3/4 d'octave, l'enveloppe
composite ne mesurait rien en dessous de 3,2 à 5,7 fois la fondamentale la plus grave du
corpus, parce qu'un harmonique n'y entre que si sa fenêtre contient trois harmoniques de la
même note. ±1 octave abaisse ce plancher de cinq demi-tons en médiane **et** retrouve plus de
formants déjà établis que ±3/4 (96 % contre 95 % pour la route A).

Ces fichiers sont conservés pour qu'on puisse comparer chiffre à chiffre, et pour qu'on puisse
revenir en arrière si la migration se révélait regrettable.
