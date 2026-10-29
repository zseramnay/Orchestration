---
id: localisation
titre: La localisation répliquée — le Q ≈ 1 de la trompette est bien l'instrument
date: 2026-10-28
tache: C6
script: Scripts/v7-formants-2026/localisation.py
donnees: Etude-Formants/Mesures/Localisation/
---

# La localisation, répliquée sur cinq corpus

## La question

`statut_formants.py` marque « aucun pic localisé » un formant dont le Q de l'enveloppe vraie
tombe sous 2 : la bosse est si large qu'elle ne sélectionne rien. La trompette et le trombone
basse avaient été classés ainsi — **sur la foi de SOL seul**. Le plan de reprise posait la bonne
question : *le Q ≈ 1 de la trompette est-il l'instrument, ou ce corpus ?*

Les largeurs d'enveloppe existaient déjà pour les cinq corpus, dans `Largeurs/Enveloppe/`, et
n'avaient jamais été croisées. Quarante-deux formants, mesurés dans deux à cinq corpus chacun.

## Premier résultat, décourageant

| Verdict sur les cinq corpus | Formants |
|---|---|
| localisé (Q ≥ 2) partout | 32 |
| **non localisé partout** | **0** |
| verdict qui change selon le corpus | 10 |

**Aucun formant n'est « non localisé » dans tous les corpus.** Pris au pied de la lettre, le
verdict publié ne réplique nulle part.

Et la trompette, au cœur de la question :

| | SOL | ZSamples | Xsamples | VN |
|---|---|---|---|---|
| 2 792 Hz | Q 0,97 | Q 1,16 | **Q 6,92** | Q 1,11 |
| 4 695 Hz | Q 1,10 | Q 1,13 | **Q 9,28** | Q 1,03 |

Trois corpus serrés autour de 1, un seul très au-dessus. Avant de conclure, il fallait demander
ce que fait Xsamples.

## Un corpus est biaisé, et cela se démontre

Comparaison **appariée** — pour chaque formant présent dans SOL et dans l'autre corpus, le
rapport des deux Q :

| Corpus | Formants comparés | Rapport médian des Q | Plus net que SOL |
|---|---|---|---|
| ZSamples | 39 | 1,06 | 22 / 39 |
| **Xsamples** | 18 | **1,37** | **15 / 18** |
| VO | 31 | 1,03 | 16 / 31 |
| VN | 26 | 0,96 | 11 / 26 |

**Xsamples rend les pics 37 % plus étroits que SOL, et plus étroits dans quinze formants sur
dix-huit.** Sa part de pics ouverts est de 4 %, contre 10 à 16 % partout ailleurs, et **son Q ne
descend jamais en dessous de 3**, sur aucun instrument. Ce corpus ne peut structurellement pas
produire un verdict de non-localisation.

Les trois autres sont sans biais mesurable par rapport à SOL — 1,06, 1,03 et 0,96.

## Second résultat, en écartant le corpus biaisé

| Verdict sur SOL, ZSamples, VO et VN | Formants |
|---|---|
| localisé partout | 32 |
| **non localisé partout** | **2** |
| discordant | 8 |

Et les deux non localisés sont **les deux formants de la trompette** :

| Formant | SOL | ZSamples | VN | Médian |
|---|---|---|---|---|
| **2 792 Hz** (route A) | 1,0 | 1,2 | 1,1 | **1,1** |
| **4 695 Hz** (route A) | 1,1 | 1,1 | 1,0 | **1,1** |

**Le Q ≈ 1 de la trompette est l'instrument, pas le corpus.** Trois corpus de provenances
distinctes, sans biais démontré, s'accordent à 0,2 près sur les deux formants. L'affirmation
publiée sur la foi de SOL seul est donc confirmée, et elle l'est maintenant par réplication.

C'est aussi le seul cas de l'étude où la mesure dit **que la notion de pic ne s'applique pas**,
plutôt que de donner une valeur. Chez la trompette, la largeur égale la fréquence centrale : la
« résonance » occupe toute la région et ne sélectionne rien.

## Ce qui ne réplique pas : le trombone basse

`_LARGEURS-DE-BANDE.md` écrit : « la trompette **et le trombone basse** n'ont aucun pic
localisé », en visant le trombone basse **vers 1 670 Hz**, et ajoute « et, de peu, le trombone
ténor à 879 Hz (Q ≈ 1,4) ». Ces deux-là ne répliquent pas.

| Formant | SOL | ZSamples | VO | Verdict |
|---|---|---|---|---|
| Trombone basse **1 759 Hz** *(le 1 670 Hz du document)* | 1,1 | 1,1 | 8,4 | **discordant** |
| Trombone ténor **830 Hz** *(le 879 Hz du document)* | 1,5 | 2,4 | 3,9 | **discordant** |
| *Trombone basse 879 Hz, pour mémoire* | 2,9 | 2,3 | 4,8 | localisé dans les trois |

Dans les deux cas, SOL donne Q ≈ 1,1 à 1,5 et les autres corpus montent jusqu'à 8,4 et 3,9 : rien
n'est tranché. **Seule la trompette survit à la réplication.** Et le formant du trombone basse
que le CSV publie — 879 Hz, celui qui porte l'unisson avec le trombone ténor — est localisé dans
les trois corpus, ce que le CSV disait déjà.

## Les huit discordants

| Formant | Q selon les corpus |
|---|---|
| Trombone ténor 3 134 Hz | 9,9 · 1,0 · 20,8 |
| Trombone basse 1 759 Hz | 1,1 · 1,1 · 8,4 |
| Tuba basse 659 Hz | 1,2 · 2,8 · 6,1 · 5,1 |
| Trombone ténor 830 Hz | 1,5 · 2,4 · 3,9 |
| Cor 739 Hz | 1,3 · 2,0 · 3,6 |
| Violoncelle 554 Hz | 4,7 · 4,0 · 1,6 · 4,0 |
| Alto 1 318 Hz | 3,7 · 3,5 · 3,5 · 1,4 |
| Tuba basse 498 Hz | 2,0 · 3,9 · 3,6 · 3,9 |

Sept des huit sont des cuivres, et c'est cohérent avec la tâche C5 : chez eux la largeur dépend
fortement de la nuance, et les corpus n'ont pas la même répartition de nuances. Les deux
exceptions, violoncelle 554 Hz et alto 1 318 Hz, sont deux valeurs déjà signalées comme fragiles
ailleurs dans l'étude.

## Conséquence pour la méthode

**Un verdict de localisation ne doit pas être prononcé sur un seul corpus.** Sur les
42 formants, 10 changent d'avis selon le corpus, soit près d'un quart. La règle à suivre :

- prononcer sur la **médiane des corpus non biaisés**, et non sur un corpus ;
- exclure Xsamples de ce calcul, ou le pondérer, tant que son biais de +37 % n'est pas expliqué ;
- écrire « discordant » quand les corpus ne s'accordent pas, plutôt que de retenir celui qui
  arrange.

Le biais de Xsamples mériterait d'être compris plutôt que seulement écarté : il est régulier, il
va toujours dans le même sens, et il n'affecte pas la **position** des formants — seulement leur
largeur. Une normalisation ou un traitement dynamique en amont est l'hypothèse la plus simple.

## Journal

- **2026-10-28** — Tâche C6. La question est tranchée : **le Q ≈ 1 de la trompette est bien
  l'instrument**, confirmé sur trois corpus indépendants. Mais la route pour y arriver a demandé
  d'identifier un biais de corpus — Xsamples rend les pics 37 % plus étroits et ne descend jamais
  sous Q = 3 — sans quoi la réplication paraissait échouer entièrement. **Et la moitié de
  l'affirmation publiée tombe** : le trombone basse à 879 Hz est localisé dans les trois corpus.
