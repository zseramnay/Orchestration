---
id: reference-feminine
titre: La référence féminine — le plancher tient par la physique, les noms de voyelles ne tiennent pas
date: 2026-10-29
maj: 2026-10-29 — chiffres recalculés sur les cinq corpus après la tâche C13
tache: C12
corpus: SOL, Xsamples, VO, VN
script: Scripts/v7-formants-2026/reference_feminine.py
donnees: Etude-Formants/Mesures/Reference-Feminine/
---

# La référence vocalique : voix d'homme ou voix de femme ?

## Ce que la référence décide

Toute la carte vocalique est lue contre les dix voyelles orales du français mesurées par
Calliope (Tubach 1989) **en voix d'homme**. Ce choix n'a jamais été justifié : c'était le
tableau sous la main. Il décide pourtant de trois choses.

| | Ce qui en dépend | Verdict de cette tâche |
|---|---|---|
| **Le plancher** — le plus bas F1 des dix voyelles | le critère « hors espace par le bas » de la tâche C11 | **robuste : +0,29 demi-ton** |
| **Les noms de voyelles** | « le basson sonne /ɔ/ », les 22 fiches, le chapitre | **non robuste : 10 instruments sur 20 changent** |
| **Le plafond** — le plus haut F1 | « la frontière tombe vers 700 Hz » (chapitre, partie II) | **non robuste : +2,45 demi-tons** |

Les valeurs des deux voix sont dans le dépôt depuis le 23 octobre —
`Formants-Francais/donnees/formants_calliope.csv` — et n'avaient jamais été croisées.

## Les deux espaces ne sont pas homothétiques, et c'est la clé

| Voyelle | Homme F1 | Femme F1 | Écart F1 | Écart F2 |
|---|---|---|---|---|
| **/i/** | 308 Hz | 306 Hz | **−0,11 dt** | +3,01 dt |
| **/u/** | 315 Hz | 311 Hz | **−0,22 dt** | +0,88 dt |
| **/y/** | 300 Hz | 305 Hz | **+0,29 dt** | +2,71 dt |
| /e/ | 365 Hz | 417 Hz | +2,31 dt | +3,14 dt |
| /o/ | 383 Hz | 461 Hz | +3,21 dt | +1,30 dt |
| /ø/ | 381 Hz | 469 Hz | +3,60 dt | +2,16 dt |
| /ɔ/ | 531 Hz | 634 Hz | +3,07 dt | +2,90 dt |
| /ɛ/ | 530 Hz | 660 Hz | +3,80 dt | +3,31 dt |
| /œ/ | 517 Hz | 647 Hz | +3,88 dt | +3,37 dt |
| /a/ | 684 Hz | 788 Hz | +2,45 dt | +3,11 dt |

**Les trois voyelles fermées ne bougent pas. Toutes les autres montent de 2,3 à 3,9
demi-tons.** Le `README.md` de `Formants-Francais/` en donne déjà la raison : « les formants
sont en général plus élevés chez les femmes, **sauf ceux qui relèvent surtout d'une résonance
de Helmholtz : F1 de /i y u/ et F2 de /u/** ». Le F1 d'une voyelle fermée est fixé par un
conduit resserré, dont la résonance dépend peu de la longueur du conduit ; le F1 d'une
voyelle ouverte, lui, suit cette longueur.

## 1. Le plancher tient — et il tient par la physique, pas par chance

Le plancher de l'espace vocalique est le plus bas F1 des dix voyelles : **/y/, 300 Hz chez
l'homme, 305 Hz chez la femme**. Soit **0,29 demi-ton**, quand le seuil strict de l'étude est
d'un demi-ton.

C'est le résultat le plus utile de la tâche, et il n'est pas accidentel : **le plancher est,
par construction, le F1 d'une voyelle fermée, c'est-à-dire la seule grandeur de la référence
qui ne dépende presque pas de la voix.** Le critère de la tâche C11 était donc bien choisi,
pour une raison que C11 ignorait.

Sur les 632 cellules, le changement de référence déplace **sept cellules** sous le plancher,
et un seul instrument change de classe :

| Instrument | Sous le plancher, homme | Sous le plancher, femme | |
|---|---|---|---|
| Clarinette contrebasse · contrebasse | 75 % | 75 % | inchangé |
| Tuba contrebasse | 67 % | 71 % | inchangé |
| Tuba basse | 53 % | 56 % | inchangé |
| Clarinette contrebasse | 73 % | 82 % | inchangé |
| Violoncelle · clarinette basse · flûtes graves · clarinettes · cor · cor anglais | — | — | **tous inchangés** |
| **Contrebasson** | **40 %** | **51 %** | **franchit le seuil de 50 %** |

**Le contrebasson passe de « hors espace dans le grave seulement » à « hors espace par le
bas ».** C'est le seul reclassement, et il était prévisible : 40 % et 51 % encadrent le seuil.
Il faut le noter dans sa fiche comme un cas limite, pas trancher.

**La limite n° 1 de `_HORS-ESPACE.md` peut être levée.** Elle disait : « une référence
féminine déplacerait le plancher vers le haut et reclasserait sans doute quelques cas
limites ». Mesuré : elle le déplace de 0,29 demi-ton et reclasse un instrument sur treize.

## 2. Les noms de voyelles ne tiennent pas

| | Cellules |
|---|---|
| ayant un couple (F1, F2) | 435 |
| **dont la voyelle change** | **181 — soit 42 %** |
| dont le statut change (assignable / approximative / hors espace) | 88 — soit 20 % |

Les reclassements les plus fréquents : /ɛ/ → /œ/ (29 cellules), /ɛ/ → /ø/ (28), /œ/ → /ɔ/
(22), /ø/ → /u/ (22), /a/ → /ɔ/ (22).

Et instrument par instrument, sur la voyelle **dominante** :

| | Homme | Femme | |
|---|---|---|---|
| **Basson** | **/ɔ/** | **/ɔ/** | tient |
| Hautbois · violon · petite flûte · clarinettes Mi♭ et Si♭ | /ɛ/ | /ɛ/ | tiennent |
| Violoncelle | /u/ | /u/ | tient |
| Trombone ténor · flûte basse | /ɔ/ | /ɔ/ | tiennent |
| **Alto** | /ɛ/ | **/ø/** | change |
| **Contrebasse** | /ø/ | **/u/** | change |
| **Trompette** | /ɛ/ | **/a/** | change |
| **Cor · trombone basse** | /ɛ/ · /œ/ | **/ø/** · **/ø/** | changent |
| **Flûtes basse et contrebasse · trombone ténor** | /a/ · /œ/ · /a/ | **/ɔ/** | changent |
| **Clarinettes basse et contrebasse** | /ø/ | **/u/** | changent |

**Dix instruments sur vingt.** La question du plan — « un instrument classé /ɔ/ en voix
d'homme peut se classer autrement » — est donc largement confirmée, et le phénomène est plus
vaste qu'annoncé.

Deux choses survivent, et ce sont les deux affirmations publiées qui comptent :

- **le basson garde /ɔ/**, donc « le basson, seul instrument à voyelle unique et stable sur
  toute sa tessiture » (chapitre, partie II section 5) tient dans les deux références ;
- les instruments les plus graves changent **tous vers /u/**, c'est-à-dire vers la voyelle la
  plus fermée et la plus postérieure. Ce n'est pas un désordre : c'est le même fait vu
  autrement — leur couleur est sombre et fermée, quel que soit le tableau.

## 3. Le plafond ne tient pas non plus

Le plafond est le F1 de /a/ : **684 Hz chez l'homme, 788 Hz chez la femme**, soit
2,45 demi-tons. Le chapitre en tire une frontière : « la frontière tombe vers 700 Hz de F1,
qui est exactement le plafond de /a/ chez l'homme — 684 Hz. Au-delà, aucune voyelle française
n'est atteignable ».

**176 cellules sont au-dessus du plafond masculin ; 51 d'entre elles repassent sous le plafond
féminin** — clarinette basse, cor anglais, flûte basse, hautbois, petite flûte, trombone
basse, trombone ténor, trompette, violoncelle.

L'affirmation reste vraie, mais elle doit dire **de quelle voix** : la frontière est à 684 Hz
en voix d'homme et à 788 Hz en voix de femme.

## 4. « Mieux ajusté » n'est pas un critère — démonstration

Il était tentant de choisir la référence par la qualité de l'ajustement. La voix de femme
gagne nettement :

| | Distance médiane à la voyelle la plus proche | Cellules assignables | Cellules hors espace |
|---|---|---|---|
| voix d'homme | 1,32 Bark | 292 | 85 |
| **voix de femme** | **1,19 Bark** | **307** | **60** |
| voix d'enfant *(extrapolée)* | **1,19 Bark** | **322** | **46** |

**C'est un piège, et le témoin le montre.** On dilate l'espace masculin d'un facteur commun
sur F1 et F2, et on relit :

| Facteur | ×1,00 | ×1,05 | ×1,10 | **×1,15** | ×1,20 | ×1,30 | ×1,50 | ×2,00 |
|---|---|---|---|---|---|---|---|---|
| Distance médiane | 1,32 | 1,23 | 1,16 | **1,10** | 1,12 | 1,16 | 1,33 | 1,61 |

La voix de femme a un F1 moyen de 500 Hz contre 431 pour l'homme, soit un facteur 1,16. **À ce
facteur, l'espace masculin dilaté donne 1,10 Bark — c'est-à-dire mieux que la voix de femme
elle-même (1,19), et mieux que les trois voix de Calliope.** Le gain n'est donc pas dans la
*forme* de l'espace féminin mais dans sa *taille* : il dit seulement que **l'orchestre se situe au-dessus de l'espace vocalique
masculin**, ce que l'étude savait déjà.

Pire : le critère ne s'arrête nulle part de raisonnable. Il préfère la voix d'enfant, dont le
`README.md` du dossier précise qu'elle est **extrapolée et non mesurée** — « femme × 1,15 »,
« c'est un ordre de grandeur ». Choisir la référence à l'ajustement, c'est choisir le tableau
le moins fiable.

## Décision

**La référence reste la voix d'homme.** Les motifs ne sont pas ceux de l'ajustement :

1. elle est **mesurée**, là où la voix d'enfant est extrapolée ;
2. elle est celle que publient les 22 fiches, le chapitre et les 20 cartes — en changer
   reclasserait 42 % des cellules pour un gain nul ;
3. **rien de substantiel n'en dépend** : le seul critère de l'étude qui s'appuie sur la
   référence est le plancher de C11, et il bouge de 0,29 demi-ton ;
4. le critère d'ajustement qui plaiderait pour en changer vient d'être invalidé.

**Trois conséquences à inscrire** :

- tout nom de voyelle publié doit porter la mention « voix d'homme » — il n'est pas une
  propriété de l'instrument mais d'un couple instrument × référence ;
- l'affirmation du chapitre sur la frontière à 700 Hz doit nommer la voix ;
- le **contrebasson** est un cas limite du critère « hors espace par le bas » : 37 % en voix
  d'homme, 52 % en voix de femme.

## Ce qui reste ouvert

| Question | Ce qui la trancherait |
|---|---|
| **Calliope mesure en contexte /pV/ et /pVʀ/**, et ses auteurs écrivent eux-mêmes que ces valeurs ne sont pas la norme du français | Georgeton et al. 2012 donnent 40 locutrices en voyelles **isolées**, avec un espace plus étendu. Le dossier les cite déjà comme source des valeurs femme : les reprendre en entier serait une troisième référence, mesurée celle-là |
| **La centralisation en parole continue** | les écarts-types de Calliope sont de 30 à 60 Hz sur F1 : à peu près un demi-ton. Une cellule de carte a un IQR de 1 à 2 Bark, donc bien davantage — l'incertitude de la référence n'est pas le facteur limitant |
| Faut-il publier les deux lectures côte à côte ? | `cellules.csv` les contient déjà toutes les deux. C'est une décision de présentation : deux noms de voyelle par cellule informent, ou bien embrouillent |

## Journal

- **2026-10-29** — Tâche C12. **Le plancher de C11 tient, et il tient par la physique** : il
  est le F1 de /y/, une voyelle fermée, donc une résonance de Helmholtz qui dépend peu de la
  longueur du conduit — 300 Hz chez l'homme, 305 chez la femme, soit 0,29 demi-ton. Un seul
  instrument change de classe, le contrebasson, qui encadrait le seuil. **En revanche les noms
  de voyelles ne tiennent pas** : 42 % des cellules et **10 instruments sur 20** changent de
  voyelle dominante, et le plafond monte de 2,45 demi-tons. **Et le critère qui aurait plaidé
  pour changer de référence est invalide** : la voix de femme ajuste mieux (1,19 contre
  1,32 Bark), mais l'espace masculin dilaté du même facteur fait mieux (1,10), et la voix
  d'enfant — extrapolée, non mesurée — l'égale (1,19). « Mieux ajusté » ne dit que ceci :
  l'orchestre est au-dessus de l'espace vocalique masculin.
