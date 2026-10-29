---
id: hors-espace
titre: « Hors espace par le bas » — un critère qui tient enfin d'un corpus à l'autre
date: 2026-10-28
tache: C11
donnees: Etude-Formants/Mesures/*/carte-vocalique-*.json
---

# « Hors espace par le bas »

## Le problème

La carte vocalique classait un instrument « hors espace par le bas » quand **un seul maximum
survivait** à la fusion de Chistovich : sans second maximum, pas de couple (F1, F2), donc pas de
voyelle. `_CARTE-VOCALIQUE.md` signalait déjà que ce critère ne tient pas :

> La présence d'un second maximum dépend du corpus, pas seulement de l'instrument. Le
> contrebasson n'en a aucun chez SOL et deux chez ZSamples. C'est une différence de prise de
> son — plus de contenu aigu dans un cas. Le critère « 1 max » n'est donc **pas** une propriété
> de l'instrument.

Mesuré sur les treize instruments disposant d'au moins deux corpus, **l'étendue médiane du
critère entre corpus est de 17 points**, et elle monte à 50 :

| Instrument | Part de cellules « 1 max » | Étendue |
|---|---|---|
| **Cor anglais** | SOL 0 % · VN 22 % · Xsamples 50 % | **50 pts** |
| **Contrebasson** | SOL 75 % · VN 75 % · Xsamples 33 % | **42 pts** |
| Tuba basse | SOL 75 % · VN 67 % · VO 100 % | 33 pts |
| Clarinette en Si♭ | SOL 8 % · VO 25 % · Xsamples 0 % | 25 pts |

Un critère qui dit « jamais » dans un corpus et « une fois sur deux » dans un autre ne décrit
pas l'instrument.

## Le critère de remplacement

Il était sous nos yeux. « Hors espace par le bas » veut dire que la couleur de l'instrument se
trouve **sous** l'espace vocalique français. Or cet espace a un plancher mesurable : le plus bas
F1 de ses dix voyelles, celui de /y/, à **300 Hz, soit 2,92 Bark**.

On compare donc **F1 au plancher**, cellule par cellule, au lieu de compter les maxima. Deux
avantages, et le second compte autant que le premier :

1. F1 est la grandeur la plus stable que l'étude possède — la tâche C1 l'a vue s'accorder à un
   demi-ton près sur quatre corpus ;
2. **F1 existe même dans les cellules que l'ancien critère jetait.** Une cellule « 1 max » a bien
   un F1 ; c'est son F2 qui manque. L'ancien critère refusait de lire une donnée qu'il avait.

> **Note de méthode, et correction d'une première version de cette page.** J'avais annoncé
> « 17 points contre 0 ». Les deux chiffres étaient calculés sur des ensembles de cellules
> différents — toutes les cellules d'un côté, seulement celles ayant un F1 de l'autre — et la
> comparaison ne valait rien. À dénominateur identique, et en prenant **la moyenne et non la
> médiane**, qui vaut zéro des deux côtés parce que la majorité des instruments sont à 0 % sur
> les deux critères :

| Critère | Étendue **moyenne** entre corpus | Pire cas |
|---|---|---|
| « 1 max » | **24,3 points** | 67 |
| **F1 sous le plancher** | **9,5 points** | 40 |

*(Chiffres recalculés le 29/10 sur les 22 instruments et les cinq corpus, après que la tâche
C13 eut ajouté 35 cartes. Le rapport entre les deux critères est inchangé — le nouveau est
deux fois et demie plus stable — mais son pire cas n'est plus nul : il vaut 40 points, sur la
clarinette contrebasse, qui n'a que onze cellules.)*

Le gain est net, et il porte précisément là où il faut — sur les instruments que l'ancien
critère traitait le plus mal :

| Instrument | « 1 max » entre corpus | F1 sous le plancher |
|---|---|---|
| **Trombone basse** | 40 % · 100 % → **60 pts** | 0 % · 0 % → **0 pt** |
| Contrebasson | 100 % · 100 % · 44 % → 56 pts | 33 % · 33 % · 44 % → 11 pts |
| Cor anglais | 22 % · 0 % · 50 % → 50 pts | 0 % · 0 % · 17 % → 17 pts |
| Clarinette en Si♭ | 25 % · 8 % · 0 % → 25 pts | 25 % · 25 % · 25 % → **0 pt** |
| Hautbois | 8 % · 8 % · 25 % → 17 pts | 0 % · 0 % · 0 % → **0 pt** |

Sept instruments sur treize donnent une valeur **identique** d'un corpus à l'autre avec le
nouveau critère. La contrebasse : 75 % et 75 %. La clarinette en Si♭ : 25 % trois fois. Le
critère est devenu une propriété de l'instrument.

## Ce que le changement révèle : l'ancien critère confondait trois choses

C'est le vrai résultat de cette tâche. « Un seul maximum » et « sous l'espace vocalique » ne sont
pas la même chose, et les mélanger produisait des classements faux dans les deux sens.

| | Sous le plancher | « 1 max » | Lecture |
|---|---|---|---|
| **Contrebasse** | **75 %** | 29 % | hors espace, et l'ancien critère la manquait |
| **Trombone basse** | **0 %** | **79 %** | dans l'espace — l'ancien critère la déclarait dehors |
| **Tuba basse** | 54 % | 100 % | les deux à la fois |

### 1. Hors espace par le bas — une propriété, stable

F1 sous le plancher dans au moins la moitié des cellules.

| Instrument | Sous le plancher | F1 médian | Corpus |
|---|---|---|---|
| Contrebasse | 75 % | 205 Hz | 2 |
| Clarinette contrebasse | 75 % | 220 Hz | 1 |
| Tuba contrebasse | 67 % | 294 Hz | 2 |
| Tuba basse | 54 % | 292 Hz | 3 |

Les quatre instruments les plus graves de l'orchestre, et eux seuls. Le résultat est celui qu'on
attend, ce qui est en soi un contrôle de la mesure.

### 2. Hors espace dans le grave seulement — par registre

Entre 20 et 50 % des cellules : violoncelle 45 %, clarinette basse 42 %, **contrebasson 37 %
(52 % en référence féminine : cas limite, voir `_REFERENCE-FEMININE.md`)**,
flûte basse et flûte contrebasse 33 %, clarinette en Mi♭ et en Si♭ 25 %, cor 22 %.

Ce n'est pas un défaut de mesure : **c'est leur registre grave qui sort de l'espace**, pendant
que le reste de leur tessiture y reste. Le critère étant calculé par cellule, il le dit tout seul
— là où le comptage de maxima, agrégé, ne pouvait que trancher en bloc.

### 3. Dans l'espace, mais sans second maximum

Trois instruments ont une couleur parfaitement située — F1 de 529 à 764 Hz, en plein milieu de
l'espace — et pourtant pas de F2 : **trombone ténor, trombone basse, trompette.** Les appeler
« hors espace par le bas » était un contresens.

Et ce troisième cas se dédouble selon la nuance :

| Instrument | doux | moyen | fort | |
|---|---|---|---|---|
| Trombone basse | 100 % | 100 % | 100 % | **sans F2 à toute nuance** |
| Trombone ténor | 75 % | 75 % | 75 % | **sans F2 à toute nuance** |
| Tuba basse | 100 % | 100 % | 100 % | **sans F2 à toute nuance** |
| **Trompette** | 67 % | 67 % | **0 %** | le second maximum **apparaît en jeu fort** |
| **Cor** | 100 % | 67 % | **0 %** | le second maximum **apparaît en jeu fort** |
| *Hautbois (témoin)* | 25 % | 8 % | 8 % | — |
| *Violon (témoin)* | 0 % | 0 % | 0 % | — |

**Pour la trompette et le cor, l'absence de second maximum n'est pas une propriété : c'est un
effet de nuance.** Joués forte, ils en ont un. Les classer d'après un corpus agrégé sur toutes
les nuances revenait à les décrire au piano.

Pour les trombones et le tuba basse, en revanche, c'est constant : leur spectre ne présente pas
de second renflement, quelle que soit la force de jeu.

> J'avais espéré que ces trois-là soient exactement les cuivres à perce cylindrique de la tâche
> C5 — trombones et trompette — et que l'effet soit le même phénomène vu deux fois. **Ce n'est
> pas le cas** : la trompette se comporte comme le cor et non comme les trombones, et le tuba
> basse, à perce conique, rejoint les trombones. La coïncidence partielle ne doit pas être
> publiée comme une loi.

## Ce qu'il faut écrire dans les fiches

| Classement | Instruments |
|---|---|
| **hors espace par le bas** | contrebasse, clarinette contrebasse, tuba contrebasse, tuba basse |
| **hors espace dans le grave seulement** | violoncelle 46 %, **contrebasson¹** 40 %, clarinette basse 35 %, flûte contrebasse 33 %, clarinette en Mi♭ 29 %, clarinette en Si♭ 26 %, flûte basse 24 % |
| **dans l'espace, sans second maximum à toute nuance** | trombone basse |
| **sans second maximum au doux seulement** | cor, trombone ténor, trompette — à ne **pas** classer hors espace |
| dans l'espace | alto, violon, hautbois, basson, cor anglais, flûte, petite flûte |

¹ Le contrebasson est le seul **cas limite** : 40 % des cellules sous le plancher en voix
d'homme, 51 % en voix de femme. Il est donc de part et d'autre du seuil de 50 % selon la
référence, et ne doit pas être présenté comme tranché (tâche C12).

> **Trois instruments ont changé de classe le 29/10**, quand la tâche C13 a porté la
> couverture de 19 à 22 instruments et jusqu'à quatre corpus chacun. Le **cor** quitte « hors
> espace dans le grave seulement » (22 % → 8 % de cellules sous le plancher) ; le **trombone
> ténor** quitte « sans second maximum à toute nuance » pour « au doux seulement » ; et le
> **tuba basse** reste « hors espace par le bas », où il figurait déjà. Les quatre instruments
> réellement sous l'espace et les sept qui y sont pleinement n'ont pas bougé. C'est la
> démonstration, par l'exemple, de la règle que la tâche C6 avait tirée : **un verdict ne se
> prononce pas sur un seul corpus.**

## Limites

1. ~~**Le plancher à 300 Hz vient de Calliope, voix d'homme.** Une référence féminine le
   déplacerait vers le haut et reclasserait sans doute quelques cas limites.~~ **Levée le
   29/10 par la tâche C12** : la référence féminine donne un plancher de **305 Hz, soit
   0,29 demi-ton**, et le critère n'est pas accidentellement robuste — le plancher est par
   construction le F1 d'une voyelle fermée, c'est-à-dire une résonance de Helmholtz, la seule
   grandeur de la référence qui ne dépende presque pas de la voix. **Un seul reclassement** :
   le contrebasson passe de 37 % à 52 % et franchit le seuil, donc il est à traiter comme un
   cas limite. Voir `_REFERENCE-FEMININE.md`.
2. **Neuf instruments n'ont qu'un seul corpus**, donc pas de contrôle de stabilité : clarinette
   contrebasse, flûte basse, flûte contrebasse, cor, trompette, trombone ténor, flûte, petite
   flûte, clarinette basse. Leur classement est à prendre comme provisoire.
3. Le seuil de 50 % pour « hors espace » et celui de 20 % pour « par registre » sont des
   conventions de lecture. Les valeurs continues sont dans les données et doivent primer sur la
   case.

## Journal

- **2026-10-28** — Tâche C11. Le critère de remplacement — F1 contre le plancher de l'espace
  vocalique — fait tomber l'étendue **moyenne** entre corpus de 24,3 à 9,5 points, et le pire
  cas de 60 à 0. La première version de cette page annonçait « 17 contre 0 » : deux
  dénominateurs différents, et une médiane qui vaut zéro des deux côtés. Quatrième fois cette
  semaine qu'un chiffre d'apparence décisive ne résistait pas à la question « comparé à quoi,
  exactement ? ». Mais le résultat de fond
  est ailleurs : l'ancien comptage de maxima **confondait trois phénomènes**, et se trompait dans
  les deux sens. Il manquait la contrebasse, qui est réellement sous l'espace, et déclarait
  dehors les trombones, qui sont en plein milieu.
