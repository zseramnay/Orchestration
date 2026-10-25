---
id: replication
titre: Réplication entre corpus indépendants — SOL2020 / Yan_Adds contre ZSamples
date: 2026-10-24
---

# Réplication entre deux corpus indépendants

Ce fichier est le plus important du dossier. Il dit **ce qui est établi et ce qui ne l'est
pas**, et il doit être lu avant toute fiche d'instrument.

## Pourquoi il existe

Jusqu'au 24 octobre 2026, l'étude reposait sur un seul corpus — SOL2020 pour les instruments
solistes, Yan_Adds pour les auxiliaires. Les deux critères de rejet de `00-METHODE.md`
(indépendance au rang harmonique, robustesse de fenêtre) sont **internes à un corpus** : ils
vérifient la cohérence d'une mesure, pas sa généralité.

C'est exactement la faiblesse qui avait produit l'étude antérieure : des chiffres
internement cohérents, jamais confrontés à autre chose qu'eux-mêmes. La reprise d'octobre
l'a reproduite d'un cran plus haut, avec des tests plus rigoureux mais toujours une seule
source.

L'arrivée du corpus **ZSamples** (`~/Documents/Instrumental_formants/Zsamples`, 3 540 fichiers
retenus, 42 configurations) permet enfin la vérification.

## Règle de lecture

La grille d'analyse est au **douzième d'octave, soit le demi-ton**. Un écart d'un demi-ton
entre deux corpus n'est donc pas un désaccord : c'est la quantification de la grille
elle-même. Seuil retenu : **≤ 1 demi-ton = concordance**, 2 demi-tons = proche,
au-delà = non répliqué.

> ### Révision du 24/10/2026 — recompte et test de permutation
>
> Ce seuil a été mis en cause le 24/10, en découvrant que **nos propres formants tombent tous
> sur des fréquences de notes tempérées** (554 Hz = Do♯5, 2 348 = Ré7 ; voir
> `_VERIFICATION-GRAVES.md`). La résolution de la méthode 8 est donc d'un demi-ton, et le
> critère « ≤ 1 demi-ton » accepte un pas de grille adjacent — la même indulgence que le
> « Δ = 11 Hz » de l'ancienne étude, un niveau plus haut.
>
> Recompte sur les 40 formants retenus des deux corpus :
>
> | Critère | Formants concordants |
> |---|---|
> | ≤ 1 demi-ton (critère initial) | **18 / 40** |
> | **Strict — même point de grille** | **7 / 40** |
>
> **Test de permutation** (20 000 tirages, appariement instrument ↔ corpus mélangé) pour
> savoir si ces concordances dépassent le hasard :
>
> | Critère | Observé | Nul médian | IC 95 % du nul | p |
> |---|---|---|---|---|
> | ≤ 1 demi-ton | **18** | 9 | [4, 14] | **0,0014** |
> | strict (= 0) | **7** | 3 | [0, **7**] | 0,033 |
>
> Trois lectures, qui vont dans des sens différents et doivent toutes être retenues.
>
> **1. Le critère « ≤ 1 demi-ton » est justifié comme résultat d'ensemble.** 18 concordances
> contre 9 attendues par hasard, p = 0,0014, et l'observé est très au-delà de l'IC du nul.
> La réplication entre corpus n'est pas un artefact de grille.
>
> **2. Mais la moitié des concordances individuelles est attendue par hasard.** Le nul médian
> est de 9 sur 40. On ne peut donc **pas** affirmer les 18 un par un : l'excès est d'environ
> neuf formants, sans qu'on sache lesquels. *L'ensemble est établi, les éléments ne le sont
> pas tous.* C'est la correction la plus importante à apporter à ce fichier.
>
> **3. Le critère strict ne sauve pas la situation.** 7 observés contre 3 attendus, mais
> l'observé tombe **exactement sur la borne haute de l'IC du nul** (p = 0,033). Le critère
> strict est plus exigeant mais statistiquement plus faible, parce qu'il travaille sur de
> trop petits nombres.
>
> Conclusion méthodologique : **la réplication entre deux corpus ne suffit pas à établir un
> formant individuel.** Il faut une seconde voie — voir « Deux routes » ci-dessous.

## Instruments nus — distance moyenne entre corpus

| Instrument | Harm./note | Formants SOL retenus | Distances (demi-tons) | Moyenne |
|---|---|---|---|---|
| **Basson** | 57 | 1 974 · 554 · 1 108 | 0,0 · 2,0 · 0,0 | **0,7** |
| **Trombone basse** | 90 | 1 759 · 3 320 · 879 | 1,0 · 1,0 · 0,0 | **0,7** |
| Trompette | 23 | 4 695 · 2 792 | 2,0 · 0,0 | **1,0** |
| Violon | 15 | 4 432 · 2 348 | 2,0 · 1,0 | **1,5** |
| Tuba basse | 85 | 659 · 1 479 · 466 · 987 | 1,0 · 1,0 · 1,0 · 4,0 | 1,7 |
| Contrebasse | 83 | 1 396 · 466 · 2 635 · 3 948 | 0,0 · 1,0 · 0,0 · 7,0 | 2,0 |
| Trombone ténor | 62 | 3 134 · 1 567 · 830 · 2 091 | 2,0 · 1,0 · 1,0 · 4,0 | 2,0 |
| Clarinette contrebasse | 90 | 1 759 | 2,0 | 2,0 |
| Violoncelle | 47 | 2 216 · 1 108 · 3 948 | 4,0 · 3,0 · 2,0 | 3,0 |
| Alto | 22 | 3 517 · 4 974 · 2 091 · 1 318 | 1,0 · 2,0 · 2,0 · 10,0 | 3,8 |
| Cor | 56 | 739 · 2 216 · 4 183 | 7,0 · 4,0 · 4,0 | 5,0 |
| Hautbois | 17 | 4 183 | 5,0 | 5,0 |
| Contrebasson | 90 | 698 · 329 · 1 174 | 2,0 · 15,0 · 1,0 | 6,0 |
| Clarinette en Sib | 22 | 1 567 · 3 517 | 14,0 · 0,0 | 7,0 |
| Flûte | 14 | 3 320 | 8,0 | 8,0 |
| Petite flûte | 7 | aucun | — | **concordance sur le résultat négatif** |
| Flûte basse · flûte contrebasse · cor anglais · clarinette en Mib | — | — | aucun recoupement | — |

## Formants établis — deux routes

> **Liste révisée le 24/10/2026.** La liste de dix-sept formants qui figurait ici reposait sur
> la seule réplication à un demi-ton, dont le test de permutation montre qu'elle laisse passer
> environ neuf concordances fortuites. Elle est remplacée par deux routes distinctes.

Un formant est **établi** s'il satisfait l'une des deux conditions suivantes :

- **Route A — réplication stricte.** Même point de grille dans les deux corpus (écart nul).
- **Route B — corroboration par méthodes indépendantes.** Concordance à ≤ 1 demi-ton avec au
  moins deux familles de méthodes n'ayant aucune étape commune avec la méthode 8 : LPC et
  cepstre d'ordre fixe (`Analyses_formants-mai26`), méthode 12 (carte vocalique), littérature
  acoustique primaire.

### Route A — réplication stricte : 7 formants, 5 instruments

| Instrument | Formants | Corroboration mai26 |
|---|---|---|
| **Contrebasse** | **1 396 Hz** | cep-80 : 1 389 Hz (0,1 dt) |
| **Contrebasse** | **2 635 Hz** | — |
| **Clarinette en Sib** | **3 517 Hz** | LPC F3 : 3 625 Hz (0,5 dt) |
| **Trombone basse** | **879 Hz** | cep-80 : 926 Hz (0,9 dt) |
| **Basson** | **1 108 Hz** | — |
| **Basson** | **1 974 Hz** | LPC F2 : 1 825 Hz (1,4 dt) |
| **Trompette** | **2 792 Hz** | — |

Trois d'entre eux — contrebasse 1 396, clarinette en Sib 3 517, trombone basse 879 — passent
**les deux routes**. Ce sont les trois résultats les plus solides de toute l'étude.

### Route B — corroboration multiple

| Formant | Lignes de preuve indépendantes |
|---|---|
| **Violon 2 348 Hz** — le *bridge hill* | cep-80 : 2 455 Hz (0,8 dt) · LPC mai26 : 2 179 Hz · méthode 12 : F2 = 2 136 Hz à l'aigu fort · littérature acoustique du violon (2–3 kHz) |
| **Violoncelle 1 108 Hz** | cep-160 : 1 109 Hz (**0,0 dt**) · méthode 12 : F2 = 1 079 Hz au médium doux |
| **Alto 2 091 Hz** | cep-160 : 2 089 Hz (**0,0 dt**) · méthode 12 : F2 = 2 161 Hz au médium fort |
| **Basson 554 Hz** | cep-80 : 528 Hz (0,8 dt) · méthode 12 : F1 stable 470–590 Hz sur 12 cellules et sur deux corpus · littérature unanime (Backus 470, Giesler 500, Meyer 500) |
| **Cor anglais 2 216 Hz** | LPC F2 : 2 312 Hz (0,7 dt) · **mais aucun formant dans ZSamples** — corroboré par méthode, non répliqué par corpus |

Le cas du **violon** justifie à lui seul l'existence de la route B : son *bridge hill* échoue
à la réplication stricte (1 demi-ton) mais réunit quatre lignes de preuve indépendantes, dont
une validation externe par la littérature. Un critère à voie unique l'aurait écarté à tort.

Le cas du **cor anglais** montre la limite symétrique : corroboré par le LPC, mais ZSamples
n'y trouve rien. **Les deux routes ne sont pas interchangeables** — il faut dire par laquelle
chaque formant passe.

### Déclassés par la révision

| Formant | Statut |
|---|---|
| Trombone basse 1 759 · 3 320 Hz | 1 demi-ton, non corroborés → **1 corpus** |
| Contrebasse 466 Hz | 1 demi-ton, non corroboré → **1 corpus** |
| Tuba basse 466 · 659 · 1 479 Hz | 1 demi-ton, non corroborés → **1 corpus** |
| Trombone ténor 830 · 1 567 Hz | 1 demi-ton ; 2 091 Hz corroboré par LPC (0,2 dt) mais à 2 dt de réplication |
| Alto 3 517 Hz | 1 demi-ton, non corroboré → **1 corpus** |

Ces dix formants ne sont pas réfutés : ils sont **indéterminés**. Ils concordent à un pas de
grille, ce que le hasard produit dans environ un cas sur quatre. Les trancher demande une
troisième source — un corpus supplémentaire, ou l'appariement note par note avec le LPC de
mai26, qui reste à faire.

### Le résultat négatif, inchangé

La **petite flûte** ne présente aucun formant dans les deux corpus, avec ρ = 0,96 et une pente
de −21 dB/octave de part et d'autre. Un résultat négatif répliqué ne souffre pas du problème
de comparaisons multiples : il n'y a rien à apparier par hasard. Il reste donc l'un des
acquis les plus sûrs de l'étude — et c'est lui qui répond à l'anomalie initiale.

S'y ajoute un résultat **négatif** répliqué, qui a autant de valeur : la petite flûte ne
présente aucun formant dans les deux corpus, avec un ρ de 0,96 et une pente de −21 dB/octave
de part et d'autre.

## Ce qui n'est PAS établi

- **Cor anglais, 2 216 et 4 183 Hz.** Présentés dans la fiche comme « le résultat le plus
  solide des anches doubles aiguës ». ZSamples n'y trouve **aucun** formant. Non établis.
- **La convergence inter-familles à 4 200 Hz** (violon, cor, hautbois, cor anglais,
  trompette), présentée comme la plus large de la reprise. Trois de ses cinq membres ne se
  répliquent pas. **Écartée.** La fiche du violon posait la question d'un artefact de chaîne
  de mesure : la réponse est probablement oui.
- **Hautbois 4 183 Hz**, **cor 739 / 2 216 / 4 183 Hz**, **contrebasson 329 Hz**,
  **clarinette en Sib 1 567 Hz**, **flûte 3 320 Hz** — tous à 4 demi-tons ou plus.
- **La pente spectrale comme descripteur d'instrument.** Elle varie fortement d'un corpus à
  l'autre : clarinette basse −9,4 contre −16,5 dB/octave. Elle dépend de la distance de
  micro et du local. **Comparable à l'intérieur d'un corpus seulement.** Toute la description
  des flûtes repose sur elle et doit être relue dans ce sens.

## Sourdines — la réplication échoue

Dix-huit paires comparables, **écart médian de 4,0 demi-tons**. Cinq paires seulement sous
2 demi-tons : trompette sèche (0,0), violon piombo (1,0), trombone sèche (1,7),
cor avec sourdine (2,0), trombone wah ouvert (2,0). Les pires : tuba basse avec sourdine
(8,3), trompette wah fermé (7,0), violoncelle piombo (5,5), trompette harmon (5,5).

**Ce n'est probablement pas un échec de mesure, mais un fait.** Une sourdine est un objet
physique : matériau, forme, ajustement au pavillon varient d'un exemplaire et d'un
instrumentiste à l'autre. Les formants d'une sourdine ne sont donc pas génériques. Il faut
parler de « la sourdine sèche de cet instrument-là », non de « la sourdine sèche ».

Conséquence pratique : **les fiches de sourdines ne doivent pas être rédigées comme des
résultats établis.** Les 28 fiches prévues le 24/10 au matin sont suspendues.

## Pourquoi les corpus diffèrent — quatre causes légitimes

1. **Ce ne sont pas les mêmes instruments ni les mêmes instrumentistes.** Pour des résonances
   de corps, une variation est attendue, en particulier sur les cordes et les sourdines.
2. **Ni la même chaîne de prise de son.** Micro, distance, local. Effet fort sur la pente,
   modéré sur les positions.
3. **Ni les mêmes notes.** ZSamples échantillonne par pas espacés (environ une tierce
   mineure), SOL chromatiquement. Les bandes ne sont pas éclairées de la même manière.
4. **La fiabilité dépend du nombre d'harmoniques par note.** Les instruments à plus de
   40 harmoniques concordent à 2,0 demi-tons en médiane, ceux à moins de 25 à 4,4. Le lien
   existe mais n'est pas une loi : le contrebasson a 90 harmoniques et concorde mal (6,0),
   la trompette en a 23 et concorde bien (1,0).

## Le test inter-cordes ne conclut pas

ZSamples organise les cordes par corde jouée (`string_I` à `string_V`), ce qui promettait un
test d'indépendance plus fort que celui du rang harmonique : une même hauteur sur deux cordes
différentes, c'est deux sources pour un même corps.

Résultat : les formants mesurés sur l'ensemble des cordes ne se retrouvent que sur une à deux
cordes sur quatre. **Mais le test est structurellement biaisé et ne doit pas être lu comme
une infirmation.** Deux raisons :

- chaque corde ne compte que 21 à 40 échantillons, donc la mesure par corde est sous-dimensionnée ;
- surtout, une corde aiguë **ne peut pas** éclairer une bande à 3 700 Hz par ses harmoniques
  d'ordre 5 à 15, alors qu'une corde grave le peut. Le défaut est le même que celui corrigé
  pour les groupes de rang : il faut ne compter que les cordes capables d'éclairer la bande.

Correction à apporter à `zsamples.py`, fonction `test_cordes()`.

## Anomalie de corpus relevée

Le dossier `Zsamples/Clarinets/bass_clarinet_boehm_system` contient des fichiers nommés
`bb_clarinet_boehm_system-*`, d'ambitus **B3–A#6** — celui de la clarinette en Sib, non d'une
clarinette basse, qui descend à Ré2. Les deux mesures donnaient d'ailleurs des formants
identiques (5 583 et 3 517 Hz). **Il n'y a pas de clarinette basse dans ZSamples.** Même
schéma que le doublon `cor`/`horn` déjà relevé dans les analyses antérieures : à vérifier
systématiquement sur tout nouveau corpus.

## Faux positif du contrôle de nomenclature

Le contrôle a alerté sur la clarinette basse de ZSamples (rapport 1,118 au lieu de 1,00).
C'était un **faux positif** : il cherchait le pic le plus proche de f0, or la clarinette a une
fondamentale faible — c'est sa signature. Le contrôle est désormais fondé sur une estimation
de f0 par **somme harmonique** (`f0_par_somme_harmonique` dans `formants_lib.py`), et rend
1,00 partout, y compris là où il criait à tort.

## Fichiers de données

- `synthese-formants-2026-10-23.csv` — les 54 configurations SOL/Yan_Adds.
- `zsamples-2026-10-24/` et `../Instruments-ZSamples/` — les 42 configurations ZSamples,
  mesures JSON et fiches pré-remplies.
