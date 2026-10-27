# Formants des voyelles orales du français

Synthèse d'une séance de travail (23 octobre 2026) : vérification d'un graphique formantique IPA existant, puis construction de graphiques corrects pour le français, en Hz et en Bark, pour voix d'homme, de femme et d'enfant (estimée).

## Contenu du dossier

```
Formants-Francais/
├── README.md                     ce document
├── donnees/
│   ├── formants_calliope.csv     F1, F2 (Hz et Bark), Δ F2−F1, note de F1 : homme, femme, enfant estimé
│   └── sources.md                références et liens
├── graphiques/
│   ├── plan_F1_F2.html           plan F1–F2 interactif (Hz/Bark, seuil 3,5 Bark, étiquettes sans masquage)
│   └── F1_seul.html              F1 seul, classé par aperture, trois voix sur couloirs séparés
└── original/
    ├── voyelles_IPA.svg          graphique d'origine (non modifié)
    └── NOTE_erreurs.md           liste des erreurs relevées
```

Les deux fichiers HTML s'ouvrent dans n'importe quel navigateur, hors ligne, et suivent le mode clair/sombre du système. Les données sont intégrées à chaque page.

## Données retenues

Source principale : **Calliope** (Tubach, *La parole et son traitement automatique*, Masson, 1989), référence classique pour le français.

| Voyelle | Ex. | API | Aperture | Homme F1 | Homme F2 | Femme F1 | Femme F2 |
|---|---|---|---|---|---|---|---|
| i | vie | [vi] | fermée | 308 | 2064 | 306 | 2456 |
| y | lune | [lyn] | fermée | 300 | 1750 | 305 | 2046 |
| u | tout | [tu] | fermée | 315 | 764 | 311 | 804 |
| e | été | [ete] | mi-fermée | 365 | 1961 | 417 | 2351 |
| ø | feu | [fø] | mi-fermée | 381 | 1417 | 469 | 1605 |
| o | eau | [o] | mi-fermée | 383 | 793 | 461 | 855 |
| ɛ | fête | [fɛt] | mi-ouverte | 530 | 1718 | 660 | 2080 |
| œ | peur | [pœʁ] | mi-ouverte | 517 | 1391 | 647 | 1690 |
| ɔ | or | [ɔʁ] | mi-ouverte | 531 | 998 | 634 | 1180 |
| a | patte | [pat] | ouverte | 684 | 1256 | 788 | 1503 |

Valeurs en Hz. Notation API : voyelles (phonèmes) entre barres obliques, mots-exemples en transcription phonétique entre crochets (R uvulaire noté [ʁ]). Le CSV contient en plus la transcription, les valeurs en Bark, l'écart F2 − F1 et la note la plus proche de F1.

### Limites à garder en tête

- **Enfant : valeurs non mesurées.** Aucune donnée publiée couvrant les dix voyelles chez l'enfant francophone n'a été trouvée. Les valeurs « enfant » sont extrapolées (femme × 1,15, enfant d'environ 10 ans). C'est un ordre de grandeur, moins fiable encore pour /i y u/ (voir l'effet Helmholtz plus bas).
- **Contexte de mesure.** Calliope mesure les voyelles en contexte /pV/ (e o u y ø) et /pVʀ/ (i ɛ a ɔ œ). Le /ʀ/ final tend à relever F1 et à abaisser F2. Les auteurs précisent eux-mêmes que ces valeurs ne doivent pas être prises comme la norme du français. En voyelles isolées et hyperarticulées (Georgeton et al., 2012, 40 locutrices), l'espace vocalique est plus étendu.
- **/ɑ/ et /ə/ absents.** Calliope ne mesure que dix voyelles. L'opposition /a/–/ɑ/ (« patte »/« pâte ») a pratiquement disparu en français standard, et /ə/ se réalise le plus souvent [ø] ou [œ].
- **Variabilité.** Écarts-types typiques de 30 à 60 Hz sur F1 et de 50 à 120 Hz sur F2 (Calliope, femmes). En parole continue, les voyelles se centralisent nettement.

## Observations

### Homme / femme : pas un simple facteur d'échelle

Les formants sont en général plus élevés chez les femmes, **sauf ceux qui relèvent surtout d'une résonance de Helmholtz : F1 de /i y u/ et F2 de /u/**. Les voyelles fermées restent presque au même endroit (environ 300–315 Hz pour F1). L'écart grandit ensuite avec l'ouverture : environ +50 à +90 Hz pour les mi-fermées, +120 à +130 Hz pour les mi-ouvertes, +100 Hz pour /a/.

### F1 code l'aperture et presque rien d'autre

Les quatre degrés d'aperture forment quatre paliers nets, sans recouvrement. À l'intérieur d'un même degré, les trois voyelles (étirée, arrondie, postérieure) ont un F1 presque identique. En Bark, ces paliers sont à peu près équidistants (0,7 à 1,6 Bark entre degrés successifs).

Hauteurs approximatives de F1 chez l'homme (La3 = 440 Hz) : fermées vers Ré–Ré♯3, mi-fermées vers Fa♯–Sol3, mi-ouvertes vers Do4, /a/ vers Fa4.

## L'échelle Bark

Le Bark découpe l'axe des fréquences en **bandes critiques**. Une bande critique est la largeur de bande à l'intérieur de laquelle l'oreille intègre l'énergie : deux composantes situées dans la même bande se masquent et fusionnent au lieu d'être entendues séparément. On compte environ 24 Bark entre 20 Hz et 15,5 kHz, et un Bark correspond à peu près à 1,3 mm de membrane basilaire.

- **Sous 500 Hz environ**, une bande critique fait à peu près 100 Hz, quelle que soit la fréquence : l'échelle y est presque linéaire.
- **Au-dessus**, la largeur croît en proportion de la fréquence, autour de 15 à 20 % (grossièrement une tierce mineure) : l'échelle y devient quasi logarithmique.

C'est pourquoi le Bark convient mieux au plan vocalique que les Hz ou un pur logarithme. F1 se situe surtout dans la zone quasi linéaire, F2 dans la zone quasi logarithmique, et le Bark traite les deux comme l'oreille.

Conversion utilisée (Traunmüller, 1990) :

```
z = 26,81 · f / (1960 + f) − 0,53        f = 1960 · (z + 0,53) / (26,28 − z)
```

Ses petites corrections aux extrémités (sous 2 Bark et au-dessus de 20,1 Bark) n'ont pas été appliquées, elles ne touchent pas cette plage.

### Ce qui change visuellement par rapport aux Hz

- **La région antérieure se tasse.** Les 700 Hz qui séparent F2 de /i/ et de /ɛ/ chez la femme ne représentent qu'environ 1,7 Bark. En Hz, la partie gauche du trapèze exagère les différences.
- **La région postérieure et les F1 élevés se dilatent.** Les écarts entre /u/, /o/ et /ɔ/, minuscules en Hz, prennent leur poids perceptif réel.
- **L'homme et la femme se rapprochent.** Une bonne part de leur différence est un décalage global en Bark.

### Le seuil de 3,5 Bark

Hypothèse de l'**intégration spectrale à grande échelle** (Chistovich et collègues, fin des années 1970). Quand deux formants sont à moins de 3 à 3,5 Bark l'un de l'autre, l'oreille ne les perçoit plus séparément. Elle les intègre en un seul « centre de gravité » spectral, le formant effectif F2′.

Les postérieures françaises /u o ɔ/ sont dans ce cas : écart F2 − F1 de 3,3 à 3,8 Bark chez l'homme, et /a/ est juste à la limite (3,5). Cela rejoint la notion de **voyelles focales** : chez les postérieures françaises, F1 et F2 sont très proches (400 à 550 Hz d'écart selon Georgeton et al.).

La valeur exacte et la portée de ce seuil restent discutées. Certaines études perceptives ne retrouvent pas d'organisation binaire nette autour de 3 à 3,5 Bark. À lire comme un repère, pas comme une frontière.

## Pistes pour l'orchestration

- **/u o ɔ/** : il est souvent inutile de répartir F1 et F2 sur deux groupes instrumentaux distincts. Un seul pic d'énergie au centre de gravité F1–F2 peut suffire à évoquer la voyelle.
- **/i/ et /y/** : la perception dépend aussi du rapprochement **F2–F3 pour /y/** (environ 500 Hz) et **F3–F4 pour /i/** (environ 700 Hz), caractéristique du français. Le plan F1–F2 ne le montre pas. Un plan F2–F3 serait le complément naturel.
- **/ø œ/** : formants à peu près équidistants (voyelles « acoustiquement centrales »).
- **Voix** : pour les voyelles fermées, la même cible F1 (environ 300 Hz) vaut pour l'homme et la femme. Pour les voyelles ouvertes, il faut adapter.

## Pistes non traitées

- Plan F2–F3 (et F3–F4 pour /i/).
- Remplacer les valeurs enfant extrapolées par des mesures publiées si on en trouve.
- Comparaison Calliope / Georgeton et al. (voyelles isolées) / Gendrot et Adda-Decker (parole radiophonique) sur le même plan.
- Voyelles nasales.
