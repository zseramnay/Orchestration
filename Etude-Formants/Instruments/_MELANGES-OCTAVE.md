---
id: melanges-octave
titre: Le doublage à l'octave — la loi des mélanges se retourne
date: 2026-10-28
tache: suite de C9
script: Scripts/v7-formants-2026/melanges.py --octave 12
donnees: Etude-Formants/Mesures/Melanges-octave/
corpus: SOL · 402 paires orientées
---

# Le doublage à l'octave

La tâche C9 avait mesuré 215 paires **à l'unisson**, en signalant que c'était le cas rare : on
double à l'octave bien plus souvent qu'à l'unisson. Voici l'octave.

## Une paire à l'octave est orientée

À l'unisson, mélanger A et B ou B et A revient au même. **Dès que les deux parties ne sont plus
à la même hauteur, ce n'est plus vrai** : flûte au-dessus du hautbois n'est pas hautbois au-dessus
de la flûte. Le calcul énumère donc 402 paires **orientées** et non 215 paires.

Et l'asymétrie n'est pas une subtilité : sur les 133 paires calculées dans les deux sens avec au
moins vingt hauteurs communes, **l'écart médian de fusion entre les deux sens est de 14 points**,
et il monte à plus de 60. Hautbois et cor, par exemple : **54 % de fusion quand le cor est au
dessus, 20 % quand c'est le hautbois.**

## Ce que l'octave change

Moyennes sur toutes les paires.

| | Unisson | Octave |
|---|---|---|
| Fusion | 23 % | **29 %** |
| Homogène | 33 % | **16 %** |
| Renforcement | 27 % | 29 % |
| Juxtaposition | 7 % | **18 %** |
| Écart de couleur médian | 1,65 Bark | **2,71 Bark** |

Deux mouvements, et un seul mécanisme derrière les deux. Transposer un parent d'une octave
**écarte sa couleur de celle de l'autre** — l'écart médian passe de 1,65 à 2,71 Bark. Par la loi
établie en C9, cet écart agrandi fait monter la fusion et effondrer l'homogénéité. C'est bien ce
qu'on observe : l'homogénéité est divisée par deux.

Mais la juxtaposition, elle, est **multipliée par 2,5**. Et c'est là que la loi se retourne.

## La loi a un maximum, et l'octave le révèle

| Écart (Bark) | Paires | Fusion | Homogène | Renforcement | Juxtaposition |
|---|---|---|---|---|---|
| 0 – 1 | 22 | 4 % | **56 %** | 19 % | 8 % |
| 1 – 2 | 91 | 16 % | 30 % | 30 % | 13 % |
| 2 – 3 | 121 | 32 % | 12 % | 32 % | 15 % |
| **3 – 4** | 82 | **41 %** | 7 % | 28 % | 19 % |
| 4 et plus | 86 | 35 % ↓ | 5 % | 29 % | **28 %** ↑ |

**La fusion culmine entre 3 et 4 Bark, puis redescend** — tandis que la juxtaposition continue de
monter. À l'unisson, ce plafond n'était pas visible : trop peu de paires atteignaient cette
distance, et la fusion semblait croître indéfiniment. L'octave, en écartant les couleurs, peuple
le haut de l'échelle et fait apparaître le maximum.

**C'est le seuil de Chistovich qui remonte à la surface.** La fusion de deux maxima auditifs
n'opère que s'ils sont distants de moins de 3,5 Bark ; au-delà, ils ne se fondent pas en un
centre de gravité commun, ils restent deux. Le mécanisme psychoacoustique qui est au cœur de la
méthode 12 se lit donc directement dans une table d'orchestration.

**Conséquence pratique** : il existe une distance de couleur **optimale** pour obtenir une
couleur nouvelle, autour de 3 à 4 Bark. En dessous on fond sans créer ; au-delà on juxtapose.
Choisir deux instruments « le plus différents possible » n'est pas la bonne stratégie.

## Dans quel sens écrire le doublage

Sur les 133 paires calculées dans les deux sens, **le sens le plus fusionnant met le timbre le
plus clair au-dessus dans 71 % des cas**.

C'est cohérent avec la loi, et c'est même sa conséquence directe. Transposer un parent d'une
octave pousse sa couleur vers le haut. Mettre au-dessus **le plus clair** l'éloigne encore de
l'autre : l'écart grandit, la fusion monte. Mettre au-dessus **le plus sombre** le rapproche de
son partenaire : l'écart se referme, et l'on retombe vers l'homogénéité ou le renforcement.

> **Réserve.** 71 % est une tendance, pas une règle. Et les deux sens ne sont pas comparables
> toutes choses égales : pour partager des hauteurs à l'octave, l'instrument du haut joue
> nécessairement dans une autre partie de sa tessiture que celui du bas. L'effet de registre et
> l'effet de sens sont donc partiellement confondus. C'est la limite principale de cette page.

Pour mémoire, la couleur propre de chaque instrument — F1 médian, qui est la grandeur qui décide :

| Les plus sombres | | Les plus clairs | |
|---|---|---|---|
| Contrebasse | 1,86 Bark | Petite flûte | 11,10 Bark |
| Clarinette contrebasse | 2,20 | Hautbois | 8,51 |
| Tuba basse | 2,78 | Flûte | 7,09 |
| Contrebasson | 3,00 | Trompette | 7,04 |
| Violoncelle | 3,10 | Clarinette en Mi♭ | 6,15 |

## Les doublages que l'on écrit réellement

« X ⇑ Y » se lit : **Y double X à l'octave supérieure.**

| Doublage | n | Fusion | Homogène | Renforcement | Juxtaposition | *rappel unisson* |
|---|---|---|---|---|---|---|
| **Violoncelle ⇑ Violon** | 126 | **56 %** | 8 % | 26 % | 10 % | *24 %* |
| **Alto ⇑ Violon** | 123 | **44 %** | 8 % | 30 % | 17 % | *19 %* |
| Trombone ténor ⇑ Trompette | 92 | 42 % | 25 % | 14 % | 16 % | *32 %* |
| Tuba basse ⇑ Trombone ténor | 93 | 38 % | 4 % | 38 % | 10 % | *36 %* |
| Clarinette ⇑ Flûte | 109 | 36 % | 4 % | 31 % | 24 % | *9 %* |
| Contrebasse ⇑ Violoncelle | 135 | 33 % | 19 % | 29 % | 10 % | *7 %* |
| Violoncelle ⇑ Alto | 147 | 31 % | 5 % | **50 %** | 8 % | *12 %* |
| Clarinette basse ⇑ Clarinette | 126 | 28 % | 7 % | 37 % | 13 % | *5 %* |
| **Hautbois ⇑ Flûte** | 85 | 24 % | 2 % | **61 %** | 12 % | *18 %* |
| Cor ⇑ Hautbois | 96 | 20 % | 5 % | 32 % | **41 %** | *40 %* |
| Basson ⇑ Hautbois | 90 | 18 % | 24 % | 18 % | **38 %** | *35 %* |
| **Flûte ⇑ Petite flûte** | 101 | 15 % | 3 % | 13 % | **67 %** | *0 %* |
| Basson ⇑ Cor anglais | 96 | 10 % | 30 % | 25 % | 28 % | *12 %* |

Trois lectures.

**Les cordes doublées à l'octave fusionnent, et c'est le seul groupe qui le fasse autant.**
Violoncelle ⇑ violon 56 %, alto ⇑ violon 44 %, contrebasse ⇑ violoncelle 33 % — contre 24, 19 et
7 % à l'unisson. Le doublage à l'octave de la famille des cordes **crée** une couleur, là où
l'unisson la laissait inchangée. C'est probablement le résultat le plus directement utilisable de
cette page.

**La petite flûte ne se fond pas, elle se superpose.** 67 % de juxtaposition — de loin le
maximum de la table — contre 75 % d'homogénéité à l'unisson. Les deux chiffres disent la même
chose : à l'unisson, les deux instruments ont la même couleur et l'on n'entend qu'un timbre ; à
l'octave, la couleur de la petite flûte part si haut qu'elle ne fusionne plus du tout. C'est
exactement ce qu'on entend d'un pupitre de petites flûtes doublant les flûtes.

**La flûte au-dessus du hautbois renforce (61 %), elle ne fusionne pas.** Le doublage le plus
courant des bois aigus n'apporte pas une couleur nouvelle : il ajoute de l'énergie à la couleur
du hautbois. Et dans l'autre sens — hautbois au-dessus de la flûte — on passe à 41 % de
juxtaposition : on entend deux instruments.

## Ce que cette page ne dit pas

Les limites de la page sur l'unisson valent toutes ici — un seul corpus, deux instruments à la
fois, et le spectral qui n'est pas le perceptif : ni phase, ni espace, ni musiciens qui
s'ajustent. S'y ajoutent deux limites propres :

1. **L'effet de registre et l'effet de sens sont confondus**, comme dit plus haut ;
2. **seule l'octave juste est calculée.** La quinte, la double octave et l'octave + tierce sont à
   portée immédiate — c'est le même paramètre `--octave` — et diraient si le maximum de fusion se
   déplace avec l'intervalle.

## Ce qui reste ouvert

| Question | Ce qui la trancherait |
|---|---|
| **Le maximum de fusion se déplace-t-il avec l'intervalle ?** | relancer à 7, 19 et 24 demi-tons. Si l'optimum reste à 3–4 Bark quel que soit l'intervalle, c'est bien le seuil de Chistovich et non une propriété de l'octave |
| Démêler registre et sens | comparer les deux sens **à registre égal** pour l'instrument du haut, en restreignant aux hauteurs où les deux jouent dans le même registre nommé |
| La règle « le plus clair au-dessus » tient-elle par famille ? | les 133 paires sont dans le CSV, non ventilées |

## Journal

- **2026-10-28** — Suite de C9. L'octave n'ajoute pas une colonne à la table de l'unisson : elle
  en change la forme. Deux résultats de fond : **une paire à l'octave est orientée**, avec
  14 points d'écart médian entre les deux sens ; et **la loi des mélanges a un maximum**, vers
  3–4 Bark, que l'unisson ne laissait pas voir et qui est le seuil de fusion de Chistovich lu au
  niveau de l'orchestre.
