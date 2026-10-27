---
id: appariement-lpc
titre: L'appariement note par note avec le LPC — la corroboration devient un test
date: 2026-10-27
tache: C4
script: Scripts/v7-formants-2026/appariement_lpc.py
donnees: Etude-Formants/Mesures/Appariement-LPC/
corpus: SOL (seul corpus pour lequel le LPC note par note existe)
---

# L'appariement note par note avec le LPC

## Ce que la corroboration valait jusqu'ici

La route B disait : « corroboré par deux familles de méthodes indépendantes ». En pratique on
comparait des **listes** — la méthode 8 trouve un formant vers 2 216 Hz, le LPC trouve un pôle
dans la même région, donc les deux se confirment.

C'est un indice, pas un test. Avec six à douze pôles répartis sur 6 kHz, un pôle tombe près de
n'importe quelle fréquence assez souvent pour qu'une coïncidence globale ne prouve rien.

## Le test

Les deux analyses portent sur **les mêmes fichiers**, et le nom de fichier est dans les deux
jeux de données : l'appariement note par note est donc exact (309 notes sur 309 pour l'alto,
291 sur 291 pour le violoncelle). Pour un formant à fc et chaque note n :

    A(n) = 1  si la méthode 8 donne un résidu local POSITIF à l'harmonique le plus proche de fc
    B(n) = 1  si le LPC de cette même note place un pôle à moins de deux demi-tons de fc

La question cesse d'être « les deux méthodes trouvent-elles la même chose en moyenne » pour
devenir « **trouvent-elles la même chose sur les mêmes notes** ». On lit P(B | A=1) contre
P(B | A=0) — ce second terme **est** le niveau de hasard, mesuré et non supposé.

Choix fixés d'avance : ordre LPC 12 (le moins de pôles, donc le hasard le plus bas), tolérance
±2 demi-tons (les largeurs de pôle valant des centaines de hertz, plus serré serait injuste
envers le LPC), segment v7 — le même que la méthode 8, pour ne pas mélanger deux différences.

## Le premier résultat était un faux positif, et le témoin l'a dit

Sur les premiers instruments, l'association paraissait écrasante : alto à 2 091 Hz,
P(B|A=1) = 0,90 contre P(B|A=0) = 0,42, cote 11,8, p de Fisher inférieur à 10⁻⁴. Treize
formants sur treize donnaient le même genre de chiffres.

**Le témoin au triton les donnait aussi, et souvent plus forts.** À 2 091 × √2 ≈ 2 957 Hz, où il
n'y a aucun formant, la cote montait à 22,5 — le double. Au violoncelle à 2 631 Hz : cote 6,9 au
formant, 14,4 au témoin.

L'association est donc **réelle mais non localisée** : les deux méthodes s'accordent sur *quelles
notes* ont de l'énergie dans une large région du spectre, pas sur *l'endroit* de la résonance.
Une note brillante a un résidu positif un peu partout dans l'aigu et des pôles un peu partout
aussi. Publier ces cotes aurait été publier un artefact de plus.

## Le test corrigé : le rang de fc dans son propre spectre

Puisque le facteur de confusion est le contenu spectral de la note, le témoin doit le partager.
On calcule donc la cote **à toutes les fréquences de la grille tempérée du domaine** — 13 à 52
témoins selon l'instrument, bande de garde de deux demi-tons autour de fc et des autres formants
établis — et on lit **le rang de fc dans cette distribution**. Toutes les fréquences témoins
subissent la même variation d'une note à l'autre : le facteur de confusion est neutralisé.

Lecture : un rang de 50 % signifie que fc ne se distingue pas du reste du spectre ; un rang élevé
signifie que l'accord des deux méthodes y est **propre à fc**.

## Résultats

42 formants testés sur 17 instruments, 40 testables (deux cellules n'ont qu'un seul groupe
peuplé, toutes les notes ayant A = 1).

| Famille | n | Rang médian de fc | Au-dessus du 50ᵉ centile |
|---|---|---|---|
| Flûtes | 3 | 82,6 % | 67 % |
| **Cordes** | **13** | **79,2 %** | **85 %** |
| Clarinettes | 6 | 78,3 % | 83 % |
| Cuivres | 12 | 69,3 % | 58 % |
| **Anches doubles** | **6** | **15,4 %** | **0 %** |
| *ensemble* | *40* | *73,8 %* | *62 %* |

**La corroboration tient pour quatre familles sur cinq, et échoue entièrement pour la
cinquième.** Les six formants des anches doubles — basson, hautbois, cor anglais, contrebasson —
sont **tous** sous la médiane, et quatre d'entre eux sous le 30ᵉ centile. Le basson place trois
de ses quatre formants dans les 6 % les plus bas de son propre spectre.

Le détail des cordes, où l'accord est le plus net :

| Instrument | Formant | Statut | Rang |
|---|---|---|---|
| Violoncelle | 554 · 2 631 · 3 948 Hz | 1 corpus · route A · route A | **91 %** chacun |
| Contrebasse | 2 635 Hz | route A | 90 % |
| Violoncelle | 2 216 Hz | 1 corpus | 88 % |
| Contrebasse | 3 948 Hz | 1 corpus | 83 % |
| Violon | 3 959 Hz | route A | 79 % |
| Violon | 2 348 Hz | route A | 71 % |
| Alto | 3 748 Hz | route A | **15 %** |

L'alto fait exception parmi les cordes : son unique formant route A, à 3 748 Hz, est au 15ᵉ
centile. C'est cohérent avec tout ce que l'étude a trouvé sur lui — sa caisse est à 375–395 Hz
(tâche C1), la méthode 8 ne mesure rien si bas (tâche C2), et son 1 318 Hz s'était révélé faux de
dix demi-tons sur ZSamples. L'alto reste l'instrument le plus mal décrit par la méthode 8.

## Les trois unissons, membre par membre

C'est la lecture la plus directement utile, puisque les unissons sont la matière du chapitre de
synthèse. Chaque membre est établi **indépendamment** par la route A ; le rang dit si le LPC le
corrobore en plus, note à note.

| Unisson | Membre | Rang |
|---|---|---|
| **1 244 Hz** (Ré♯6) | Clarinette basse | 75,6 % |
| | Tuba basse | 74,2 % |
| 1 108 Hz (Do♯6) | Violoncelle | 67,6 % |
| | **Basson** | **29,4 %** |
| 2 216 Hz (Do♯7) | Violon (2 348 Hz) | 70,8 % |
| | **Cor** | **30,8 %** |

**L'unisson de 1 244 Hz est le seul dont les deux membres sont corroborés.** Les deux autres sont
asymétriques, et dans les deux cas le membre faible s'explique : le basson est une anche double,
famille où le LPC échoue par construction ; le cor est le seul cuivre du lot dont le formant
route A tombe sous la médiane.

Cela ne retire rien à ces unissons — la route A, deux corpus de provenances distinctes au même
point de grille, est le critère le plus exigeant, et il est satisfait des deux côtés dans les
trois cas. Mais si l'un des trois doit être mis en avant dans le chapitre, c'est celui de
1 244 Hz.

## Pourquoi les anches doubles, et pas les autres

Le LPC suppose un modèle **source–filtre tout-pôle** : une source à spectre plat filtrée par une
cavité. Il convient à une corde — corde excitatrice, caisse filtrante, les deux réellement
séparables. Chez une anche double, l'anche et la perce sont **fortement couplées**, et la
résonance propre de l'anche fait partie de la source : il n'y a pas de filtre séparable, donc les
pôles ne se posent pas sur les résonances de la perce.

Cette limite était déjà écrite dans les notes de méthode — « les vents sont à anche et perce
couplées » — mais elle n'avait jamais été **mesurée**. C'est maintenant le cas, et avec une
conséquence pratique nette.

Que les clarinettes passent (78 %) alors que les anches doubles échouent est à prendre comme une
lecture, pas comme un acquis : la clarinette est à anche simple et joue très en dessous de la
résonance propre de son anche, donc le couplage y est plus faible. L'étude ne le démontre pas.

## Conséquence pour la route B

**La route B peut être chiffrée pour les cordes, les clarinettes, les cuivres et les flûtes ; elle
doit être refusée aux anches doubles.** Concrètement :

- les formants du basson, du hautbois, du cor anglais et du contrebasson restent établis **par la
  route A** — deux corpus de provenances distinctes au même point de grille — et cela ne change
  pas. La route A est d'ailleurs le critère le plus exigeant des deux ;
- mais ils **ne peuvent pas revendiquer de corroboration LPC**, et les fiches ne doivent pas le
  laisser croire ;
- pour les cordes, la corroboration est désormais un chiffre et non une impression : au
  violoncelle, 91ᵉ centile sur trois formants.

## Ce que le test ne couvre pas

| Manque | Pourquoi |
|---|---|
| ~~Le cor~~ | **Comblé le 27/10.** `Largeurs/LPC-v7/` ne contenait que `Horn+sordina` — **l'analyse de mai 2026 avait omis le cor ordinaire**, et tous les scripts aval prenaient leur liste d'instruments de ce dossier, si bien que l'absence ne se voyait pas. `largeurs_lpc.py` reçoit une option `--depuis-dossier` qui construit la liste depuis les échantillons. 134 notes, 4 973 pôles. Le cor place son formant route A au 31ᵉ centile |
| Les autres corpus | le LPC note par note n'existe que pour SOL. ZSamples, Xsamples, VO et VN n'ont pas été passés au LPC |
| Le LPC de mai 2026 | à 44,1 kHz et ordre 18, sur un autre segment. L'apparier à la méthode 8 mélangerait deux différences ; il faudrait d'abord le refaire sur le segment v7 |
| La dépendance à l'ordre | les cotes aux ordres 12, 16, 20 et 24 sont dans le CSV, mais le rang n'a été calculé qu'à l'ordre 12 |

## Journal

- **2026-10-27** — Tâche C4. La corroboration devient un test, et le test dit deux choses : elle
  tient pour quatre familles sur cinq, et elle échoue pour les anches doubles, ce que le modèle
  tout-pôle du LPC prédisait. Le premier test, sans témoin apparié, donnait treize résultats
  écrasants et faux ; c'est le témoin au triton qui l'a révélé. **Même leçon que la tâche C2 sous
  une autre forme : le témoin doit partager le facteur de confusion de ce qu'il contrôle.**
