---
id: suraigu
titre: Le plancher de la fondamentale — correction de l'instabilité au suraigu
date: 2026-10-27
tache: C3
scripts: Scripts/v7-formants-2026/suraigu.py, impact_regle_f0.py
donnees: Etude-Formants/Mesures/Suraigu/
statut: correctif adopté · cartes régénérées
---

# Le plancher de la fondamentale

## Le défaut

`_CARTE-VOCALIQUE.md` le donnait comme **le seul défaut de méthode reproductible identifié** :

> Le violon est instable au suraigu dans les deux corpus. SOL donne F1 = 496 Hz au doux et
> 1 478 Hz au moyen ; ZSamples donne 977 Hz avec un IQR de 8,6 Bark, soit un tiers de l'échelle.
> Ce n'est pas un artefact d'un jeu d'échantillons : c'est un basculement du détecteur entre
> fondamentale et résonance de caisse. **À corriger avant toute publication de la ligne du
> violon.**

Le diagnostic était juste. Il restait à en tirer une règle.

## La règle

`f1_f2()` prenait pour F1 **le maximum survivant le plus bas**, sans borne inférieure. Or un
spectre harmonique **n'a aucun partiel en dessous de f0**. Ce qui se trouve là est du bruit
d'archet, la réponse de la caisse à une excitation large bande, ou la jupe de la fonction
d'étalement de Schroeder — jamais un formant au sens vocalique, puisque la lecture (F1, F2)
suppose une source harmonique filtrée par des résonances.

Tant que f0 reste **sous** la zone de résonance, la question ne se pose pas : à l'alto grave
(f0 131–196 Hz) la caisse est à 375–395 Hz, donc au-dessus de toute fondamentale du registre.
Mais au suraigu du violon (f0 ≥ 1 175 Hz), la fondamentale a dépassé la caisse (~530 Hz) : le
détecteur a alors le choix entre deux maxima sans rapport l'un avec l'autre, et bascule de l'un
à l'autre selon la dynamique.

**Correctif : F1 est cherché parmi les maxima situés à f0 ou au-dessus**, à un demi-ton de
tolérance près (`PLANCHER_F0` dans `carte_vocalique.py`). Quand des maxima sont écartés, la
cellule porte la marque `sous f0 écarté` — l'information n'est pas perdue en silence.

## Combien de notes étaient concernées

C'est le chiffre qui justifie la règle à lui seul : **entre un quart et plus de la moitié des
notes du suraigu avaient un « F1 » situé sous leur propre fondamentale.**

| Instrument | SOL | ZSamples | VO | VN |
|---|---|---|---|---|
| Violon, suraigu | 37 % | 32 % | 21 % | 4 % |
| Alto, suraigu | 15 % | 28 % | **45 %** | 25 % |
| Violoncelle, suraigu | 26 % | **57 %** | 31 % | 13 % |
| tous, registres grave et médium | **0 %** | **0 %** | **0 %** | **0 %** |

Aucune cellule ne devient vide : la règle ne détruit aucune mesure, elle en écarte une partie
impossible.

## Ce que la règle change

### Le violon, cas par cas

| Registre · dynamique | F1 avant | F1 après | IQR avant | IQR après |
|---|---|---|---|---|
| grave, les trois | 451 · 460 · 476 | **identiques** | — | identiques |
| médium, les trois | 507 · 525 · 525 | **identiques** | — | identiques |
| aigu : doux · moyen · fort | 664 · 741 · 788 | **829 · 829 · 829** | 2,57 | 2,05 |
| suraigu : doux · moyen · fort | **496** · 1 478 · 1 569 | **1 665 · 1 665 · 1 610** | 0,50 | 2,25 |

Le cas qui avait motivé la tâche — 496 Hz au doux contre 1 478 Hz au moyen — disparaît : les
trois dynamiques s'accordent. L'IQR par cellule au suraigu tombe de **7,4 à 2,3 Bark chez SOL**
et de **9,0 à 3,4 chez ZSamples**.

On notera que l'IQR de la cellule « suraigu doux » *monte* de 0,50 à 2,25. Ce n'est pas une
dégradation : l'ancien 0,50 était la cohérence d'une réponse constamment fausse, le détecteur se
verrouillant sur la caisse pour toutes les notes de la cellule.

### L'accord entre corpus s'améliore partout où la règle agit

C'est le gain le plus utile, parce qu'il porte sur la réplication et non sur une dispersion
interne.

| | Avant (étendue inter-corpus) | Après |
|---|---|---|
| Violon, aigu | 702–854 Hz · 3,4 dt | **807–854 Hz · 1,0 dt** |
| Violon, suraigu | 1 317–2 094 Hz · 8,0 dt | 1 659–2 094 Hz · 4,0 dt |
| Alto, suraigu | 829–1 108 Hz · 5,0 dt | **1 112–1 173 Hz · 1,0 dt** |
| Violoncelle, suraigu | 311–878 Hz · 18,0 dt | 788–1 245 Hz · 8,0 dt |

Le suraigu de l'alto passe d'un éparpillement sur cinq demi-tons à un accord à un demi-ton sur
quatre corpus. C'est un **résultat neuf**, produit par le correctif.

### Et ailleurs : presque rien

Vérifié cellule par cellule sur les cinq corpus avant adoption.

| | SOL (206 cellules) | ZSamples + Xsamples + VO + VN (≈ 440 cellules) |
|---|---|---|
| F1 déplacé de plus d'un demi-ton | 9 · **4 %** | 23 · **5 %** |
| Voyelle changée | 1 · **0 %** | 9 · **2 %** |
| Dispersion améliorée / dégradée | 8 / 3 | 18 / 14 |

Et la répartition par registre montre que le correctif est **exactement ciblé** :

| Registre | Cellules dont F1 bouge | |
|---|---|---|
| suraigu | 27 % (SOL) · 28 % (autres) | là où était le défaut |
| aigu | 14 % · 10 % | registre de transition |
| médium | 0 % · 1 % | |
| **grave** | **0 %** · **0 %** | |
| chalumeau, gorge, clarine, bas et haut médium, sons pédale, extrême grave | **0 %** | |

**Les acquis de C1 sont intacts**, ce qui était la condition d'adoption : la caisse de l'alto au
grave (371 · 374 · 381 Hz selon la dynamique) et celle du violoncelle au médium
(211 · 225 · 236 Hz) sont inchangées au Hz près. C'est attendu : ces résonances sont au-dessus
de la fondamentale de leur registre, donc la borne ne les touche pas.

## Deux effets de bord, tous deux favorables

**1. Le cor anglais revenait manquant de chaque régénération.** Son entrée avait été retirée de
`carte_vocalique.INSTRUMENTS` avec le commentaire « ABSENT de registres.md — registres à définir
avant mesure ». Ses registres y ont été ajoutés depuis (grave Mi2–Do4, médium Do4–Do5, aigu
Do5–Do6, suraigu au-dessus) mais le commentaire est resté, si bien que la carte retombait à 21
instruments à chaque relance. Réintégré.

**2. Dix-sept cartes par corpus gagnées.** Les cartes des nouvelles banques ne couvraient que le
hautbois et l'alto — les deux instruments pour lesquels elles avaient été lancées explicitement
en octobre. En relançant le script complet : Xsamples passe de 1 à 5 instruments, VO de 1 à 8,
VN de 2 à 8. Aucun n'est perdu. **Cela avance la tâche C13** (étendre la vérification vocalique
aux instruments non comparés) sans qu'on l'ait cherché.

Conséquence directe : **le contrebasson et le trombone basse ont maintenant une carte
vocalique**, alors qu'ils figuraient parmi les quatre « hors espace par le bas » sans couple
(F1, F2). Il n'en reste que deux, le tuba basse et le tuba contrebasse. Cela confirme la
limite n° 1 déjà écrite dans `_CARTE-VOCALIQUE.md` : *la présence d'un second maximum dépend du
corpus, pas seulement de l'instrument*. Le classement « hors espace par le bas » n'est donc
toujours pas une propriété d'instrument — c'est la tâche **C11**.

## Ce qui reste ouvert

| Question | Ce qui la trancherait |
|---|---|
| Le suraigu du violoncelle reste dispersé (8 dt entre corpus, IQR 3,1 à 4,5) | il descend de 18 dt, donc le correctif agit, mais pas assez. À regarder après C11 |
| Faut-il marquer les cellules où la fondamentale a dépassé la zone de résonance ? | elles n'ont plus de « voyelle » au sens propre : F1 y est la hauteur jouée. La marque `F1=f0` le dit déjà en partie ; un marquage explicite « hors espace par le haut » serait plus clair |
| Le plancher doit-il valoir aussi pour F2 ? | F2 est déjà au-dessus de F1, donc au-dessus de f0 par construction. Sans objet |

## Journal

- **2026-10-27** — Tâche C3. Le défaut est corrigé par une borne physique et non par un
  réglage : aucun partiel n'existe sous la fondamentale, donc aucun maximum n'y est un formant.
  Mesuré avant adoption sur les cinq corpus — 4 à 5 % des cellules bougent, 0 à 2 % changent de
  voyelle, rien au grave ni au médium. Cartes des cinq corpus régénérées, anciennes conservées
  dans `Archives/cartes-avant-C3-2026-10-27/`.
