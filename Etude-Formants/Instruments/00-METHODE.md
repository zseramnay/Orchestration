---
id: methode
titre: Méthode de mesure formantique — version 2026
version_mesure: 2026-10-23-wav-v1
---

# Méthode de mesure — reprise d'octobre 2026

> **À lire d'abord : `_REPLICATION.md`.** Il distingue les formants établis — répliqués
> dans deux corpus indépendants — de ceux qui ne reposent que sur une source.

Ce fichier définit une fois pour toutes les méthodes et les critères utilisés dans les
fichiers d'instruments. Chaque fichier d'instrument y renvoie et n'en reprend que les
résultats.

## Pourquoi tout a été refait

Trois défauts cumulés invalidaient les mesures antérieures.

**1. Les enveloppes `specenv` d'Orchidea ne sont pas des enveloppes de formants.**
Le code (`src/fourier.h`, fonction `cepstralEnvelope`, appelée depuis `src/analysis.h`)
applique un lissage cepstral d'ordre fixe `NUM_SMOOTH = 160`, codé en dur. À 44 100 Hz,
l'indice cepstral du peigne harmonique vaut fs / f0 ; l'ordre 160 n'exclut donc ce peigne
que pour f0 < 44 100 / 160 = **276 Hz**, soit environ do# 4. Au-dessus, la périodicité de la
note passe *dans* l'enveloppe, et le « pic le plus proéminent » devient un harmonique.
Ce n'est pas un défaut d'Orchidea : pour la recherche de combinaisons, on veut justement
un descripteur sensible à la hauteur. C'est un défaut d'emploi.

**2. Le Fp par centroïde ne mesure pas une résonance.** Un centroïde calculé dans une
bande [a, b] tombe nécessairement entre a et b : il mesure la pente du spectre à
l'intérieur d'une fenêtre imposée. Sa stabilité apparente est une conséquence de son
bornage, non une qualité timbrale. Changer la bande change la valeur : la flûte passe de
1 535 Hz (bande par instrument) à 1 177 Hz (bande unique 600–1 800 Hz).

**3. Les écarts publiés étaient sous la résolution de mesure.** Les valeurs progressaient
par pas de 10,77 Hz, largeur d'un bin de FFT à 4096 points. Un « Δ = 11 Hz, convergence
quasi-parfaite » signifiait « un bin d'écart », pour des dispersions de 100 à 400 Hz.

## Source des données

Spectres moyens **recalculés depuis les fichiers WAV**, et non repris des bases Orchidea.
Raison : `compute_features` démarre sa boucle à l'échantillon 0 et pondère chaque trame par
son énergie ; l'attaque, la plus énergétique, pèse donc lourd dans la moyenne.

| Paramètre | Valeur |
|---|---|
| FFT | 4 096 points, saut 2 048, fenêtre de Hann |
| Fréquence d'échantillonnage | 44 100 Hz (vérifiée fichier par fichier) |
| Résolution | 10,77 Hz/bin · 1 024 bins · 0–11 015 Hz |
| Segment analysé | onset + 200 ms, durée 1 500 ms |
| Détection d'onset | premier dépassement de 20 % du RMS maximal (fenêtre 10 ms) |
| Normalisation | par l'énergie totale du spectre moyen |
| Technique | `ordinario` (sauf contrebasson : `non-vibrato`, seule disponible) |

Contrôle de validité préalable : le rapport entre le pic réel et le f0 déduit du nom de
fichier vaut **1,00** pour les douze instruments. La nomenclature des bases est donc en
hauteurs réelles, et la chaîne de lecture est juste.

## Méthode 8 — enveloppe composite (socle)

Principe : un formant est une propriété du *filtre*, indépendante de la note. Mais une note
tenue n'éclaire ce filtre qu'aux fréquences de ses harmoniques. On agrège donc les
amplitudes harmoniques de **toutes** les notes de la tessiture sur un axe fréquentiel
commun : des centaines de sources différentes éclairent le même filtre, ce qui est la
définition physique du formant. C'est l'approche de Luce & Clark (1967).

Procédure :

1. Pour chaque note, relever l'amplitude de chaque harmonique k·f0 (pic local, ±2 à 3 bins).
2. **Redressement local** : retirer à chaque harmonique la médiane des amplitudes de la
   même note dans une fenêtre de ±3/4 d'octave. Seules survivent les déviations plus
   étroites que ~1,5 octave.
3. Agréger les résidus par bande de 1/12 d'octave, et prendre la médiane.
4. Un formant apparaît comme une zone de résidu médian positif.

Le redressement **local** est essentiel. Un redressement linéaire en log-fréquence, essayé
d'abord, fabrique une bosse médiane artificielle sur tout spectre concave : il avait produit
des pics à 1 318 et 1 863 Hz pour le hautbois, qui se sont évanouis dès le passage au
redressement local.

**Limite à connaître :** cette méthode est aveugle par construction aux résonances plus
larges que ~1,5 octave — **et, limite distincte corrigée le 27/10, elle ne mesure rien en
dessous de 3,2 à 5,7 fois la fondamentale la plus grave du corpus**, parce qu'un harmonique
n'entre dans le composite que si sa fenêtre contient trois harmoniques de la même note. Les deux
limites ont été confondues : c'est la seconde, et non la première, qui fait manquer la résonance
de caisse de l'alto et du violoncelle. Voir `_FENETRE-DE-REDRESSEMENT.md`. C'est le prix de
l'immunité à l'artefact de pente, et c'est pourquoi
les flûtes ressortent sans formant — leur zone d'énergie est une coupure de rayonnement
large, non une résonance étroite.

## Méthode 6 — enveloppe vraie à ordre adaptatif (lecture note par note)

Cepstre itératif à la Röbel & Rodet : à chaque itération, V ← max(log|X|, V) puis lissage
cepstral. L'ordre n'est pas fixe mais **adapté à la hauteur** : C = 0,7 × q_peigne, où
q_peigne est l'indice cepstral du peigne harmonique. Les ordres employés vont de 3 (piccolo
suraigu) à 529 (contrebasson grave).

Usage principal : fournir la mesure d'indépendance à la hauteur. On corrèle la position du
pic trouvé avec la hauteur jouée. **ρ proche de 0 = formant fixe ; ρ proche de 1 = le
spectre suit la note.** C'est cette mesure qui fonde la typologie en deux classes.

## Méthode 12 — motifs d'excitation auditive (doublures) — à venir

Pour la question des doublures, les descripteurs formantiques ne suffisent pas. Deux
fonctions distinctes doivent être mesurées sur des **mélanges réels** :

- **Fusion** : le centroïde du mélange reste stable, les timbres se confondent.
- **Brillance / renforcement** : le centroïde monte, mais la fusion est maintenue par la
  coïncidence harmonique.

Non encore implémentée.

## Critères de rejet

Aucun chiffre n'est publié sans avoir passé ces tests. Ils auraient suffi à écarter les
trois défauts énumérés plus haut.

**Test 1 — indépendance au rang harmonique.** La bande doit rester élevée quel que soit le
rang k qui vient l'éclairer. Un pic qui n'est qu'un effet de hauteur s'effondre dès que le
rang change. Critère : résidu ≥ +1 dB dans au moins **deux** groupes de rang distincts
(k = 2–3, 4–7, 8–15, ≥ 16), chacun alimenté par ≥ 5 notes différentes. Le rang k = 1 est
exclu, étant confondu avec la fondamentale.

**Test 2 — robustesse à la fenêtre de redressement.** Le pic doit se retrouver à moins de
1/3 d'octave lorsque la fenêtre passe de ±3/4 à ±1/2 puis ±1 octave. Ce test écarte les
grandeurs fabriquées par leur propre fenêtre d'analyse, défaut du Fp.

**Verdicts :**

| Verdict | Signification |
|---|---|
| **retenu** | les deux tests passés |
| **fragile** | un seul des deux tests passé — à confirmer, non publiable seul |
| **rejeté** | aucun des deux tests passé |

## Test 3 — réplication sur un corpus indépendant (24/10/2026)

**Ajouté après coup, et c'est le critère le plus important.** Les deux tests ci-dessus sont
*internes à un corpus* : ils vérifient la cohérence d'une mesure, pas sa généralité. C'est la
faiblesse même qui avait produit l'étude antérieure — des chiffres cohérents entre eux,
jamais confrontés à autre chose qu'eux-mêmes.

Un formant n'est **établi** que s'il est retrouvé, à un demi-ton près, dans un second corpus
enregistré indépendamment. La grille d'analyse étant au douzième d'octave, un écart d'un
demi-ton est la quantification de la grille et non un désaccord.

| Statut | Signification |
|---|---|
| **établi** | répliqué à ≤ 1 demi-ton dans deux corpus indépendants |
| **1 corpus** | passe les tests 1 et 2, mais sur une seule source — **non publiable comme acquis** |
| **non répliqué** | mesuré dans deux corpus, écart ≥ 2 demi-tons |

Au 24/10/2026, la confrontation SOL2020/Yan_Adds contre ZSamples donne **17 formants établis
sur 9 instruments**, tous dans le registre grave ou médium. Le détail, instrument par
instrument, et la liste de ce qui n'est PAS établi, sont dans `_REPLICATION.md`, à lire avant
toute fiche.

Deux conséquences immédiates :

- la **pente spectrale** n'est pas un descripteur d'instrument : elle varie fortement d'un
  corpus à l'autre (clarinette basse −9,4 contre −16,5 dB/octave) car elle dépend de la
  distance de micro et du local. Elle ne vaut qu'à l'intérieur d'un corpus ;
- les **sourdines ne répliquent pas** (écart médian 4,0 demi-tons sur 18 paires). Ce n'est
  probablement pas un échec de mesure mais un fait : une sourdine est un objet physique dont
  le matériau et l'ajustement varient. Il faut parler de « la sourdine de cet instrument-là ».

## Critère assoupli pour corpus réduits (23/10/2026)

Quatre instruments du corpus comptent moins de 50 échantillons et il n'en existe pas
d'autres. Le critère a donc été assoupli — **sans abaisser aucun seuil**. Le résidu minimal
de +1 dB, le test de robustesse de fenêtre et le redressement local à ±3/4 d'octave sont
inchangés. Deux leviers seulement ont bougé.

**Levier 1 — découpage au rang médian.** Les quatre groupes fixes (k = 2–3, 4–7, 8–15, ≥ 16)
sont remplacés par deux groupes séparés au **rang médian des points de la bande**, avec une
tolérance portée de ±1/12 à ±1/6 d'octave. Le test logique est identique : le résidu
survit-il à un changement de rang harmonique ? Mais les deux groupes sont peuplés par
construction.

Un découpage fusionné à seuil fixe (k = 2–5 / k ≥ 6) avait d'abord été essayé : il **perd le
test** pour les instruments très graves, dont tous les harmoniques de la bande sont à rang
élevé, et il classait à tort la clarinette contrebasse « non mesurable ». Le découpage
adaptatif l'évite.

**Levier 2 — bootstrap par note.** Le seuil de comptage « ≥ 5 notes distinctes par groupe »
est remplacé par un intervalle de confiance. On rééchantillonne avec remise les **notes** —
et non les points, les harmoniques d'une même note étant corrélés — et l'on exige, pour
chaque groupe, un résidu médian ≥ +1 dB **et** une borne basse d'IC à 90 % strictement
positive. Un petit nombre de notes produit un intervalle large, donc un échec : l'incertitude
est rendue plutôt que masquée par un seuil arbitraire.

| Verdict assoupli | Signification |
|---|---|
| **retenu-cr** | les deux groupes de rang passent le bootstrap, et la robustesse de fenêtre passe |
| **fragile-cr** | un seul des deux tests passe |
| **rejeté** | aucun |
| **non mesurable** | moins de deux groupes testables (< 3 notes ou < 5 points) — aucun chiffre produit |

Le suffixe `-cr` doit être conservé dans toute publication : il signale que le verdict repose
sur le critère assoupli, afin qu'un formant de corpus réduit ne soit jamais présenté comme
l'équivalent d'un formant établi sur 300 échantillons.

### Contrôle de validité de l'assouplissement

Un critère assoupli n'a de valeur que s'il ne retient pas ce que le critère strict rejetait à
juste titre. Le contrôle a été fait sur les 22 instruments : **les sept rejets significatifs
sont tous préservés**, dont le cas de référence de la clarinette basse à 466 Hz
(résidu −0,3 dB au-delà du rang 5, IC de −2,0), le hautbois à 2 792 Hz et la clarinette en
Sib à 1 108 Hz. Deux formants auparavant retenus sont au contraire **déclassés en fragiles**
(trombone basse 3 320 Hz, clarinette basse 5 270 Hz), leur moitié de rang élevé étant
négative. Le critère assoupli est donc plus discriminant sur certains points qu'il n'est
permissif sur d'autres.

### Méthode écartée : le test de permutation

Un test de permutation global sur la bande avait d'abord été implémenté : permutation,
note par note, de l'affectation des résidus aux fréquences de ses propres harmoniques.
**Il est inutilisable.** Il donnait p < 0,005 à presque tous les candidats, y compris à ceux
que le critère strict rejette correctement, pour deux raisons :

- **biais de sélection** — la bande a été *détectée* parce que son résidu médian était un
  maximum local, puis le test porte sur cette même positivité. Le raisonnement est circulaire.
- il ne teste pas l'indépendance au rang harmonique, qui est le seul point en cause.

C'est consigné ici pour que l'erreur ne soit pas refaite.

## Typologie — trois classes

La typologie a été révisée le 23/10/2026, après mesure des cordes. Deux classes ne
suffisaient pas.

| Classe | Définition | ρ annoncé | ρ mesuré | Formants retenus | Instruments |
|---|---|---|---|---|---|
| **A** | résonance fixe dominante | faible | **0,48 – 0,50** | oui | anches doubles *(sauf le hautbois)*, trombones, tubas |
| **A/B** | intermédiaire | — | **0,57** | oui | **cor, trompette** |
| **B** | spectre suiveur, pas de résonance étroite | élevé | **0,77 – 0,88** | non ou peu | flûtes, clarinettes, **hautbois** |
| **C** | résonances multiples étroites, aucune dominante | **élevé** | **0,84** | **oui, nombreuses** | cordes |

> **Vérifié le 27/10/2026 contre les ρ désormais mesurés sur les cinq corpus (tâche D3).** La
> typologie tient : chaque classe retrouve le ρ qu'elle annonce, et les groupes ne se
> chevauchent pas. Deux ajouts rendus nécessaires par la mesure :
>
> - **le cor et la trompette n'étaient dans aucune ligne.** Leur ρ, 0,55 et 0,58, tombe entre la
>   classe A (0,48–0,50) et la classe B (0,77–0,88) : d'où le **A/B** que portent déjà leurs
>   fiches, qui a maintenant une justification chiffrée et une ligne ;
> - **le hautbois est en classe B** depuis le 26/10 — quatre corpus, quatre réponses
>   formantiques incompatibles — alors que la ligne A citait les anches doubles en bloc.
>   L'exception est désormais écrite.

**Pourquoi une troisième classe.** La mesure ρ compare la position du pic *le plus
proéminent* de l'enveloppe vraie à la hauteur jouée. Elle répond donc à la question : « un
pic unique domine-t-il, indépendamment de la note ? » Un ρ élevé peut avoir **deux causes
opposées** :

- il n'existe aucune résonance étroite, et le pic suit la fondamentale — cas des flûtes ;
- il existe beaucoup de résonances étroites, aucune ne domine assez pour que le
  pic-tracker s'y fixe, et celui-ci retombe sur la fondamentale — cas des cordes.

Les cordes affichent ρ = 0,78 à 0,90, aussi élevé que les flûtes, **tout en** livrant trois
à quatre formants retenus chacune, là où les flûtes n'en livrent aucun. Le ρ seul ne
distingue donc pas les deux situations : il faut le croiser avec le nombre de formants que
l'agrégation retient. C'est le couple (ρ, nombre de formants retenus) qui classe, non ρ seul.

Conséquence pratique : **ρ ne doit jamais être publié seul comme indice d'« absence de
formant »**. C'est une erreur que la version antérieure de ce fichier aurait pu induire.

Le principe d'indépendance de la convergence formantique à la hauteur vaut pour les classes
A et C. Pour la classe B, l'invariant pertinent n'est pas un formant mais la fréquence de
coupure du réseau de trous latéraux (Benade) et la pente spectrale.

> **Mesuré le 28/10/2026, tâche C10 — et la moitié de cette phrase est à retirer.** Le coude
> spectral se mesure bien, et il est stable **chez les clarinettes seulement** : pente de
> log2(coude) contre log2(f0) de 0,07 en médiane, et la clarinette en Si♭ donne 2 094 · 2 078 ·
> 1 970 · 2 013 Hz sur quatre corpus, soit 1,1 demi-ton d'écart.
>
> **Chez les flûtes, il ne se mesure pas** — pente 1,22, c'est-à-dire un rang harmonique et non
> une fréquence fixe. Elles cumulent trop peu d'harmoniques (7 par note pour la petite flûte),
> une tessiture qui place leurs fondamentales au-dessus de la région de coupure, et une
> enveloppe trop lisse. **Le descripteur promis aux instruments sans formant est précisément
> celui qu'ils ne permettent pas de mesurer.** Pour elles, il reste la pente spectrale, et le
> fait mesuré en C9 qu'elles ne fusionnent avec rien.
>
> Enfin, l'attribution au réseau de trous **n'est pas démontrée** : le violon, qui n'en a aucun,
> donne le coude le plus stable de l'étude (pente 0,02). Publier « coude spectral, compatible
> avec une coupure de réseau de trous ». Voir `_COUPURE-BENADE.md`.

Validation indépendante : la balance impair − pair, mesurée sur les harmoniques 2 à 9, est
**positive pour les quatre clarinettes seulement** (+0,7 à +5,1 dB) et croît vers le grave de
la famille. C'est la signature du tuyau cylindrique fermé, mesurée et non postulée.

## Quand la notion de formant n'a pas de sens

Chaque fichier d'instrument porte un champ `pertinence_formant` :

- **pertinente** — résonance fixe établie, mesure fiable ;
- **partielle** — spectre suiveur, mais tessiture et richesse harmonique suffisantes pour
  que l'agrégation détecte des résonances réelles ;
- **non pertinente** — soit la tessiture est entièrement au-dessus du domaine de validité,
  soit le spectre ne présente aucune résonance étroite, soit le corpus est trop petit. Dans
  ce cas, les descripteurs de classe B (pente, coupure) remplacent le formant, et il faut le
  dire plutôt que de produire un chiffre.

## État et limites

- Corpus : 1 166 échantillons, 12 instruments, 0 erreur de lecture.
- Concordance attaque incluse / exclue : les formants ne bougent que d'un à deux douzièmes
  d'octave. Les valeurs sont donc doublement validées.
- **Non encore fait :** vérification des valeurs de littérature aux sources primaires — elles
  sont pour l'instant reprises des tableaux de l'étude antérieure. *(Les trois autres points de
  cette liste sont faits : la mesure par registre par les tâches C5 et C7, la méthode 12 par la
  carte vocalique, et la coupure des flûtes par C10 — négativement, voir ci-dessus.)*

### Résolution en fréquence : rien n'est séparable sous 75 Hz *(tâche C7, 28/10/2026)*

L'analyse travaille à 10,77 Hz par bin, et `pic_proche` cherche le maximum des harmoniques
dans ±2 bins autour de k·f0 — une fenêtre de 5 bins. **Quand l'écart entre deux harmoniques
tombe sous cette largeur, les fenêtres de deux rangs voisins rendent le même pic**, et la
liste des harmoniques contient des fréquences dupliquées.

| Écart f0/bin | 1 | 2 | 3 | 4 | 5 | 6 | **≥ 7** |
|---|---|---|---|---|---|---|---|
| Notes avec doublon | 100 % | 100 % | 100 % | 74 % | 51 % | 28 % | **0 %** |

Soit **f0 ≥ 75 Hz**. Cela touche 4,8 % des notes du corpus SOL, et jusqu'à 40 % des notes d'un
seul instrument — clarinette contrebasse, tuba contrebasse, contrebasson, contrebasses, sons
pédale et extrême grave des cuivres.

**Conséquence mesurée sur les résultats publiés : aucune.** Dédoublonner sans retirer de note
laisse la route A à 23 formants, les trois unissons et les six régions à l'identique :
`residu_local` retranche une médiane locale, et un point compté deux fois la repondère sans
la déplacer. Trois valeurs de statut « indéterminé » bougent d'un ou deux pas de grille
(contrebasse 1 396, contrebasson 698, tuba contrebasse 879 Hz) et **doivent porter cette
mention**. Le drapeau `FORMANTS_DEDOUBLONNER=1` rend la vérification reproductible ; il reste
inactif, faute de gain.

**Et surtout l'inverse est vrai** : *retirer* ces notes fait disparaître les formants graves
(le tuba contrebasse perd 370 Hz, le cor 466, le contrebasson 329), non parce qu'ils étaient
faux mais parce que le plancher de domaine de la méthode 8 vaut 3,2 à 5,7 fois f0_min. **Ce
sont ces notes très graves qui portent les formants graves.** Voir `_REGISTRES-PEDALE.md`.

### Les cuivres ne se lisent pas au-dessus du médium *(tâche C7)*

Sur les six cuivres et tous corpus confondus, le rapport F1/f0 de la méthode 12 descend de
10,8 au registre pédale à 1,0 à l'aigu, et la pente de log2(F1) contre log2(f0) atteint 1,03
avec r² = 0,85 au suraigu : **la carte vocalique y mesure la note, pas une couleur.** Rien ne
doit être lu au-dessus du médium pour cette famille. À l'inverse le registre pédale, longtemps
classé « hors registre », est le terrain le plus propre — F1 y vaut onze fois f0.

Et la couleur d'un cuivre dépend de la nuance au point que **comparer deux corpus sans
apparier la nuance revient à comparer deux répartitions de nuances** : 1 seul accord strict
entre provenances distinctes en nuances réunies, **6 à nuance égale**.
- Quatre instruments reposent sur moins de 50 échantillons (flûte basse 39, flûte
  contrebasse 38, clarinette en Mib 45, clarinette contrebasse 45) : leurs résultats sont
  provisoires.
