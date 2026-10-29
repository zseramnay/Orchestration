---
id: coincidences-f2
titre: Les deux méthodes ne se corroborent pas — sauf pour le basson
date: 2026-10-29
maj: 2026-10-29 — chiffres recalculés sur les cinq corpus après la tâche C13
tache: C8
instruments: les 17 instruments ayant au moins un formant établi
corpus: SOL, Xsamples, VO, VN
script: Scripts/v7-formants-2026/coincidences_f2.py
donnees: Etude-Formants/Mesures/Coincidences-F2/
---

# Les coïncidences entre la carte vocalique et les formants établis

## La question, et pourquoi elle est bonne

La méthode 8 et la méthode 12 ne mesurent pas la même chose. La première cherche une
résonance du **filtre**, indépendante de la note ; la seconde lit la saillance du spectre
**rayonné** sur une échelle auditive, sans jamais séparer source et filtre. Rien n'oblige
leurs résultats à coïncider — et c'est exactement pour cela qu'une coïncidence, si elle
existe, vaudrait comme corroboration indépendante.

Le plan la demandait comme « **indice de confiance pour les cordes** », dont les formants
sont les plus discutés de l'étude.

**La réponse est non, et elle est particulièrement non pour les cordes.** Une seule
exception, forte et nette : le basson.

## La mesure, et ses deux témoins

632 cellules de carte — registre × nuance × corpus — sur **cinq** corpus, la tâche C13 ayant ajouté les 35 cartes qu'un échec d'appariement de noms écartait. Pour chaque cellule
et chacun de ses deux maxima, l'écart en demi-tons au formant établi le plus proche **du même
instrument**. Coïncidence si l'écart tient sous un demi-ton, le seuil strict de l'étude.

Un taux brut ne veut rien dire : les F2 vivent dans une bande étroite et les formants aussi,
donc des coïncidences surviennent par construction. La tâche C4 l'a appris à ses dépens.

| Témoin | Ce qu'il contrôle |
|---|---|
| **1. décalage** | les propres formants de l'instrument, transposés de ±2 à ±6 demi-tons. Le **nombre** de formants et leur **espacement** sont conservés : seule la position change |
| 2. rang parmi les autres | l'écart au formant le plus proche de chaque **autre** instrument. 50 % sous l'hypothèse nulle |

## Le second témoin est invalide, et c'est le premier résultat

Il faut le dire avant les chiffres, sinon ils se lisent à l'envers.

| | Coïncidence mesurée | Rang parmi les autres |
|---|---|---|
| **F1 des cuivres** | **1,2 %** — contre 4,2 % au témoin | **83 %** |

Une coïncidence trois fois plus rare que le hasard, et pourtant l'instrument propre paraît
« plus proche » que les autres dans 83 % des comparaisons. Les deux ne se contredisent pas : **le rang est confondu par le
registre.** Le F1 d'un tuba est voisin de ses propres formants parce que les deux sont graves,
tandis que les formants d'un violon sont deux octaves au-dessus de tout F1. Le témoin
récompense les instruments dont les formants habitent la même octave que leur F1 — ce qui
n'est pas la question posée.

**Seul le témoin par décalage partage le confondant**, parce qu'il garde l'instrument. C'est
lui qui décide dans tout ce qui suit. Encore la même leçon : un témoin doit partager le
confondant de ce qu'il contrôle.

## Le résultat : la coïncidence n'existe pas

| Sous-ensemble | n | Mesure | Témoin décalé | |
|---|---|---|---|---|
| F2, tous instruments | 309 | 5,8 % | **6,1 %** | le témoin fait mieux |
| **F2, cordes** | 149 | 6,0 % | **8,2 %** | le témoin fait mieux |
| F2, bois | 136 | 5,9 % | 4,2 % | rien |
| F2, cuivres | 24 | 4,2 % | 3,8 % | rien |
| **F2, cellules sans aucune marque** | 55 | **3,6 %** | 4,0 % | rien |
| F1, cordes | 79 | 0,0 % | 0,5 % | rien |
| F1, cuivres | 86 | 1,2 % | 4,2 % | rien |

Et le chiffre qui résume la tâche : **hors le basson, il n'existe pas une seule coïncidence
sous un demi-ton dans une cellule non marquée**, sur 632 cellules, 17 instruments et cinq
corpus.

Instrument par instrument, le rapport mesure / témoin :

| Instrument | Maximum | Mesure | Témoin | Rapport |
|---|---|---|---|---|
| **Basson** | **F1** | **44 %** | 7,8 % | **×5,7** |
| **Basson** | **F2** | **23 %** | 3,1 % | **×7,5** |
| Violon | F2 | 7 % | 5,4 % | ×1,4 |
| Violoncelle | F2 | 10 % | 9,5 % | ×1,1 |
| **Alto** | F2 | 5 % | 15,1 % | **×0,3** |
| Flûte · hautbois · trombones · tuba basse · violoncelle (F1) | — | 0 % | 1,6 à 10 % | — |

Les trois cordes visées par la tâche donnent ×1,4, ×1,1 et ×0,3. **L'indice de confiance
demandé n'aurait rien mesuré là où on le voulait.**

> À tolérance relâchée d'un demi-ton à un ton entier, le violon passe à ×2,4 et le violoncelle
> à ×1,9. Il serait tentant de s'y arrêter. Mais ces coïncidences-là sont presque toutes dans
> des cellules déjà marquées « F1=f0 » ou « dispersé » — l'aigu et le suraigu des cordes, que
> l'étude ne publie pas. Sur les cellules non marquées, elles tombent à zéro.

## L'exception : le basson, et elle est forte

| | Mesure | Témoin | Rapport |
|---|---|---|---|
| F1, seuil 0,5 dt | 44 % | 7,8 % | ×5,7 |
| F1, seuil 1,0 dt | 89 % | 14,8 % | ×6,0 |
| F2, seuil 0,5 dt | 23 % | 3,1 % | ×7,5 |
| **F1, cellules non marquées, 1,0 dt** | **94 %** — 16 cellules sur 17 | 13,5 % | **×7,0** |

Et la coïncidence porte sur **ses trois formants publiés**, dans **deux corpus de provenances
distinctes** — SOL (`FullSOL2020`, IRCAM) et Xsamples :

| Formant | Cellules sous 1,0 dt | dont sous 0,5 dt | Corpus |
|---|---|---|---|
| **494 Hz** | 7 | 3 | SOL · Xsamples |
| **554 Hz** | 11 | 6 | SOL · Xsamples |
| **1 108 Hz** | 8 | 3 | SOL · Xsamples |

Quinze de ces vingt-six cellules ne portent aucune marque. Les registres concernés sont le
grave, le bas médium et le haut médium, aux trois nuances.

**Deux méthodes qui ne partagent ni hypothèse ni chaîne de calcul trouvent le même
instrument au même endroit.** C'est la corroboration la plus forte que l'étude possède pour
un instrument isolé.

### Et cela complète exactement ce que le LPC refusait

Le formant de 1 108 Hz est l'un des trois unissons formantiques, basson · violoncelle. La
tâche C4 avait trouvé son membre **basson faible** — 29,4ᵉ centile à l'appariement LPC, contre
67,6 pour le violoncelle — et l'avait expliqué : le modèle tout-pôle du LPC échoue sur les
anches doubles, où l'anche couplée à la perce n'est pas une source blanche.

**La carte vocalique corrobore précisément là où le LPC ne pouvait pas.** Le basson à
1 108 Hz n'est donc pas un membre faible de l'unisson : c'est un membre que l'un des deux
instruments de mesure ne pouvait pas voir.

## Pourquoi la coïncidence est l'exception et non la règle

La réponse est dans la tâche C2, et elle est arithmétique.

**L'écart médian de F1 au formant établi le plus proche est de 15,5 demi-tons** — plus d'une
octave. Celui de F2, de 4,1 demi-tons. Chez les cordes l'écart de F1 monte à 22,6 demi-tons,
chez les bois à 25,1 ; chez les cuivres il tombe à 7,8, parce que leurs formants sont bas. Les deux méthodes ne regardent tout
simplement pas la même région du spectre : la méthode 8 ne mesure rien sous **3,2 à 5,7 fois
f0_min**, à cause du plancher de domaine, tandis que F1 vit par construction juste au-dessus
de la fondamentale.

Une coïncidence n'est donc possible que là où un instrument possède un formant **bas**, dans
la région où F1 se trouve. Le basson en a deux, à 494 et 554 Hz. C'est tout l'écart entre
lui et les autres.

> Ce n'est pas un défaut mais une répartition des rôles, et elle était déjà écrite dans le
> chapitre : « les cordes ont une résonance de caisse, et la méthode 8 ne la voit pas ». La
> tâche C8 en donne la réciproque chiffrée — **là où la méthode 8 voit, la méthode 12 ne
> regarde pas.**

## Ce qui a été implémenté

Le signalement demandé existe, mais il est **nommé pour ce qu'il est et non pour ce qu'on
espérait** :

- `carte_vocalique.coincidence(hz, instrument)` rend le formant établi le plus proche s'il est
  à moins de 0,5 demi-ton, et `carte()` inscrit `coincidence_F1` et `coincidence_F2` dans
  chaque cellule ;
- sa docstring porte **les taux de témoin mesurés ci-dessus**, pour que personne ne puisse le
  lire comme un indice de confiance sans rencontrer le démenti ;
- `coincidences_f2.py --annoter` inscrit le champ dans les cartes déjà écrites —
  **34 coïncidences sur les cinq corpus**, purement additif, aucune valeur mesurée touchée. Il
  repart du Bark et non du Hz arrondi, pour donner exactement le même résultat que
  `carte()`.

Le champ reste utile : il rend le cas du basson lisible d'un coup d'œil, et il rendra
visible tout cas semblable qui apparaîtrait sur un corpus futur.

## Ce qui reste ouvert

| Question | Ce qui la trancherait |
|---|---|
| **Le contrebasson a-t-il la même propriété que le basson ?** | c'est l'autre anche double grave, et ses candidats 329 et 698 Hz sont bas. La tâche C13 lui a donné **quatre corpus de carte**, dont Xsamples et ZSamples qui échappent à la règle de provenance : le côté vocalique est donc désormais testable. Reste à vérifier si ses formants le sont aussi — `formants-v7.csv` lui donne n_corpus = 2 et 3, ce que la phrase « aucun second corpus indépendant » du chapitre ne laisse pas attendre |
| Le seuil de 0,5 demi-ton est-il le bon ? | une cellule est une médiane sur 7 à 49 notes, avec un IQR de 1 à 2 Bark : la tolérance devrait dépendre de la dispersion de la cellule, pas être fixe. À faire si le champ sert un jour à autre chose qu'au basson |
| **La réciproque** : un formant établi bas prédit-il une coïncidence ? | trois instruments seulement ont un formant sous 600 Hz et une carte lisible. L'échantillon est trop petit pour en faire une règle — c'est une hypothèse, pas un résultat |

## Journal

- **2026-10-29** — Tâche C8. **La coïncidence demandée n'existe pas** : 5,8 % contre 6,1 % au
  témoin qui partage le confondant, et **3,6 % contre 4,0 % sur les cellules non marquées**.
  Elle n'existe surtout pas chez les cordes, pour qui le plan la voulait : ×1,4, ×1,1 et ×0,3.
  *(Chiffres recalculés le jour même sur les cinq corpus, après que la tâche C13 eut ajouté
  35 cartes : la conclusion est inchangée et le basson en sort renforcé, de ×4,7 à ×5,7.)*
  **Une exception, forte : le basson**, dont F1 tombe sur ses formants de 494 et 554 Hz dans
  16 cellules non marquées sur 17 (×7,0), et dont F2 corrobore le 1 108 Hz de l'unisson avec
  le violoncelle — précisément là où le LPC de la tâche C4 échouait par construction sur les
  anches doubles. **Un témoin a dû être déclaré invalide** : le rang parmi les autres
  instruments est confondu par le registre, ce que les cuivres démontrent en donnant 84 % de
  rang pour 1,2 % de coïncidence. Et la raison structurelle est dans C2 : F1 vit à 15,5
  demi-tons en médiane du formant le plus proche, parce que le plancher de domaine interdit à
  la méthode 8 de mesurer là où F1 se trouve.
