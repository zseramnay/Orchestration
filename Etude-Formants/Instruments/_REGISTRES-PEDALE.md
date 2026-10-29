---
id: registres-pedale
titre: Les registres pédale — et la découverte que la couleur d'un cuivre est une propriété de la nuance
date: 2026-10-28
tache: C7
instruments: cor, trompette, trombone ténor, trombone basse, tuba basse, tuba contrebasse
corpus: SOL, ZSamples, Xsamples, VO, VN
script: Scripts/v7-formants-2026/registres_pedale.py
donnees: Etude-Formants/Mesures/Registres-Pedale/
figure: Site-v2/figures/synthese-registres-pedale.png
---

# Les cuivres, relancés avec leurs registres pédale

## Ce que la tâche demandait, et ce qu'elle a trouvé

`registres.md` a reçu le 24 octobre une définition des sons pédale pour le cor, la trompette
et les trombones, et un « extrême grave » pour les tubas. Avant cela ces notes tombaient dans
la catégorie fourre-tout « hors registre » et personne ne les lisait.

| | Résultat |
|---|---|
| **Ce qui était demandé** | les pédales sont assignées, mais **81 échantillons de trombone restent hors registre**, et les cellules pédale sont trop minces pour la carte vocalique |
| **Résultat de méthode** | la carte vocalique d'un cuivre **ne se lit que jusqu'au médium** : au-dessus, son F1 *est* la fondamentale |
| **Résultat de fond** | **le désaccord des cuivres entre corpus était un désaccord de nuances.** À nuance égale, 6 accords stricts entre provenances distinctes au lieu de 1 |
| **Grandeur neuve** | **trombone basse, F1 = 403–412 Hz au doux**, trois corpus de provenances distinctes, **0,39 demi-ton** — et elle n'apparaît qu'à nuance égale |
| **Limite de la chaîne** | sous 75 Hz les harmoniques ne sont plus séparables. Conséquence mesurée sur les résultats publiés : **aucune** |

## 1. La couverture est incomplète, et la cause est nommée dans le fichier lui-même

`registres.md` écrit : « trombones : sons pédale E1-A#1, grave E2-E3 (**trou B1-D#2 sans
barillet**), médium E3-E4, aigu au-dessus ». Deux défauts en découlent.

| Instrument | Corpus | Notes hors registre |
|---|---|---|
| Trombone basse | SOL | 25 — La#0, Do#1, Ré#1, **Si1, Do2, Do#2, Ré2, Ré#2** |
| Trombone basse | ZSamples | 16 — Do1, Ré#1, **Do2, Ré#2** |
| Trombone basse | VO | 9 — La#0, Do#1, **Do#2** |
| Trombone ténor | SOL | 15 — **Si1, Do2, Do#2, Ré2, Ré#2** |
| Trombone ténor | ZSamples | 16 — Si0, Do1, Ré#1, **Do2, Ré#2** |

**Le trou B1-D#2 est réel sur le trombone ténor sans barillet, et il ne l'est pas sur le
trombone basse**, qui possède le barillet et joue ces notes. La ligne est écrite pour le ténor
et appliquée aux deux. Second défaut, indépendant : **les pédales des deux trombones descendent
sous E1** — le corpus contient La#0, Si0, Do1, Do#1, Ré#1, toutes sous la borne.

> Ce n'est pas une erreur de `registres.md` mais une ambiguïté de ce qu'il décrit : la
> tessiture **notée** d'un instrument conventionnel, ou l'étendue **échantillonnée** du corpus.
> Les deux ne coïncident pas, et la carte vocalique a besoin de la seconde. Proposition :
> une ligne par trombone, et des bornes de pédale prises sur le corpus.

## 2. Les cellules pédale sont trop minces — sauf une

La carte exige cinq notes par cellule et deux maxima survivants. Dynamiques réunies, le
registre pédale pèse : trombone basse **24 · 8 · 9** notes selon le corpus, trombone ténor
**3 · 12**, cor **3 · 2**, trompette **2 · 9 · 4**. Cellule par cellule — registre × nuance —
presque tout retombe sous le seuil ou ne donne qu'un seul maximum.

**Définir le registre pédale n'a donc pas ouvert de cellules lisibles**, hors le trombone
basse. C'est la réponse honnête à la tâche telle qu'elle était posée. Ce qui suit est ce que
la relance a trouvé en plus.

## 3. Au-dessus du médium, F1 *est* la fondamentale

Test de la tâche C1 appliqué registre par registre aux six cuivres, tous corpus confondus : la
pente de log2(F1) contre log2(f0) doit être nulle si F1 est une propriété du tuyau, et valoir 1
si ce n'est que la hauteur jouée.

| Registre | Cellules | Pente méd. | r² méd. | F1/f0 | r² ≥ 0,5 |
|---|---|---|---|---|---|
| **sons pédale** | 5 | 0,44 | **0,06** | **10,8** | 0 / 5 |
| **extrême grave** | 4 | 0,21 | **0,04** | 7,9 | 0 / 4 |
| **grave** | 16 | −0,07 | **0,03** | 4,2 | 0 / 16 |
| **médium** | 17 | 0,16 | **0,05** | 2,1 | 1 / 17 |
| aigu | 17 | 0,42 | 0,31 | **1,0** | 6 / 17 |
| **suraigu** | 5 | **1,03** | **0,85** | **1,0** | 3 / 5 |

Le tableau se lit en une ligne : **F1/f0 descend de 10,8 à 1,0, et au bout du chemin la pente
vaut 1 avec r² = 0,85.** Dans le suraigu d'un cuivre, la carte vocalique ne mesure pas une
couleur, elle mesure la note. Dans l'aigu, six cellules sur dix-sept y sont déjà.

La carte marquait déjà « F1=f0 » cellule par cellule. Ce qui est neuf est que **c'est une
propriété de registre, pas un accident de cellule**, et qu'elle a une frontière : le haut du
médium. Toute la suite ne porte donc que sur pédale + grave + médium.

Et le renversement : **le registre pédale est le terrain le plus propre de tous.** F1 y vaut
onze fois f0 — aucune confusion avec la fondamentale n'y est possible — et r² y tombe à 0,06.
Le registre qu'on avait laissé « hors registre » est celui où la mesure est la moins ambiguë.

## 4. Mon premier verdict était vrai et trompeur

En réunissant pédale + grave + médium, f0 couvre 1,3 à 3,3 octaves — le levier maximal du test
de pente. Le résultat est net : **quinze mesures sur dix-sept donnent une pente dans ±0,25**,
donc une propriété indépendante de la note. Seule la trompette échoue, avec des pentes de signe
non constant (−0,29 · +0,28 · −0,32).

J'ai d'abord écrit que cela établissait une propriété fixe de l'instrument. **C'est faux, et
l'erreur est exactement celle que cette étude a déjà commise quatre fois : le contrôle ne
partage pas le confondant de ce qu'il contrôle.** La pente contre f0 contrôle la *note*. Elle
ne dit rien de la *nuance*, puisque les nuances sont réunies dans chaque régression.

| Instrument, corpus SOL | doux | moyen | fort | **doux → fort** |
|---|---|---|---|---|
| Tuba basse | 149 Hz | 244 Hz | 364 Hz | **15,5 dt** |
| Trombone ténor | 374 Hz | 539 Hz | 716 Hz | **11,2 dt** |
| Trombone basse | 412 Hz | 663 Hz | 713 Hz | **9,5 dt** |
| Cor | 291 Hz | 350 Hz | 439 Hz | 7,1 dt |
| Trompette | 653 Hz | 861 Hz | 950 Hz | 6,5 dt |
| Tuba contrebasse | 245 Hz | 303 Hz | 308 Hz | 4,0 dt |

**F1 est plat contre la note et monte de 9,5 demi-tons en médiane avec la nuance.** La
progression est monotone pour les six instruments, dans les quatre corpus. Une grandeur peut
donc avoir une pente nulle contre f0 et bouger de plus d'une octave : la platitude contre la
note ne prouve pas qu'on mesure l'instrument.

**Et l'étude le savait déjà.** La partie II du chapitre de synthèse ouvre sa section 2 par
« la dynamique déplace le spectre des cuivres de près d'une octave — c'est le résultat le plus
robuste de la famille », avec le trombone ténor au grave à 355 → 718 Hz, soit ×2,02, répliqué
sur un corpus indépendant à 359 → 707. Mon erreur n'a donc pas été d'ignorer un fait inconnu,
mais de **ne pas voir qu'un résultat déjà publié invalidait ma lecture du test de pente**. Les
chiffres ci-dessus ne sont pas une découverte : ce sont les mêmes, étendus aux six cuivres, à
quatre corpus et aux trois registres bas réunis.

Ce qui est neuf est la **conséquence**, que le chapitre énonçait sans l'exploiter : « toute
mesure de cuivre agrégée sur plusieurs nuances est nécessairement étalée ». Si l'étalement
vient de là, alors le désaccord entre corpus doit disparaître à nuance égale. C'est la section
suivante, et c'est le résultat de cette tâche.

## 5. À nuance égale, la réplication apparaît

Si la nuance est la variable cachée, alors comparer deux corpus revient à comparer deux
répartitions de nuances, et non deux instruments. On refait donc la comparaison **nuance par
nuance**, en appliquant la règle de provenance de `_CORPUS-CROISES.md`.

| | Accords stricts (< 0,5 dt) entre provenances distinctes |
|---|---|
| nuances réunies | **1** — et douteux : tuba basse SOL/VN à 0,26 dt, mais IQR de 14,4 et 7,4 dt |
| **à nuance égale** | **6** |

Les six :

| Instrument | Nuance | Les deux corpus | Écart | Pentes |
|---|---|---|---|---|
| **Trombone basse** | **doux** | SOL 412 Hz (n=12) · **VO** 408 Hz (n=12) | **0,15 dt** | +0,07 · +0,08 |
| **Trombone basse** | **doux** | **VO** 408 Hz · **ZSamples** 403 Hz (n=20) | **0,24 dt** | +0,08 · +0,10 |
| **Trombone basse** | **doux** | SOL 412 Hz · **ZSamples** 403 Hz | **0,39 dt** | +0,07 · +0,10 |
| Trombone basse | fort | SOL 713 Hz (n=63) · ZSamples 695 Hz | 0,43 dt | −0,09 · +0,19 |
| Trompette | doux | SOL 654 Hz (n=19) · VN 658 Hz (n=57) | 0,12 dt | −0,19 · −0,18 |
| Tuba basse | doux | SOL 150 Hz (n=24) · VO 147 Hz (n=13) | 0,30 dt | −0,20 · −0,25 |

**Le trombone basse au doux est le résultat de cette tâche.** Trois corpus de provenances
distinctes — Yan_Adds (VSL), VO et ZSamples — donnent 412, 408 et 403 Hz : **0,39 demi-ton
d'étendue totale**, avec des IQR de 2,6, 2,5 et 6,4 demi-tons et des pentes contre f0 de +0,07,
+0,08 et +0,10. Plat contre la note, serré, et répliqué trois fois.

Et il réplique aussi au fort, à 713 et 695 Hz. **La dépendance à la nuance est elle-même
répliquée** : le même instrument passe de ~408 à ~704 Hz, soit 9,5 demi-tons, dans deux
bibliothèques indépendantes. Ce n'est donc pas un artefact de prise de son.

**Quatre des six accords sont au doux.** L'hypothèse la plus simple : au doux le spectre est
pauvre et la couleur est dominée par la résonance propre de l'instrument ; au fort elle dépend
du souffle du joueur et du niveau d'enregistrement, qui varient d'une bibliothèque à l'autre.
À vérifier, mais le sens de l'effet est cohérent avec C5, où c'est au fort que les pics
s'élargissent le plus.

## 6. La trompette : correction d'une corroboration trop vite écrite

J'avais écrit que l'échec de la trompette au test de pente corroborait son Q ≈ 1 de la tâche
C6 — deux méthodes concluant qu'elle n'a pas de couleur localisée. **C'est à retirer.**

À nuance égale, la trompette au doux donne des pentes de −0,19 et −0,18 et un accord SOL/VN à
**0,12 demi-ton**. Elle est donc plate et répliquée, comme les autres. Son échec en nuances
réunies est un **effet de nuance**, pas une absence de couleur.

Les deux mesures ne portent du reste pas sur la même chose : C6 mesure la **largeur** du pic
(Q ≈ 1 : il ne sélectionne rien), celle-ci sa **position**. Un pic large peut avoir un centre
parfaitement stable. Les deux résultats coexistent sans se corroborer.

> Cela recoupe directement le quatrième cas de la tâche C11 : chez la trompette et le cor,
> l'absence de second maximum est un effet de nuance — 0 % au fort. Trois tâches trouvent
> maintenant la même chose par trois chemins : **chez les cuivres, rien ne doit être publié
> sans la nuance.**

## 7. Le tuba contrebasse : le piège que l'étude avait elle-même posé

Mesuré sur SOL et VN, le tuba contrebasse donne 280 et 279 Hz — **0,06 demi-ton**, sur 126 et
96 notes et 3,3 et 2,6 octaves de fondamentale. C'est le plus bel accord de toute la tâche.

**Il ne compte pas.** `_CORPUS-CROISES.md` l'écrit depuis le 25 octobre : le côté « SOL » du
tuba contrebasse vient de `Yan_Adds-Divers`, qui est du VSL, et VN est du VSL — 206 529
échantillons de part et d'autre, corrélation 0,99591. **Ce sont les mêmes enregistrements.**
La règle retenue est que la paire (SOL, VN) ne compte pas lorsque la source SOL est Yan_Adds,
et c'est le cas ici. Le tuba contrebasse n'a toujours aucun second corpus.

J'ai écrit la première version de ce document en annonçant ce chiffre comme une grandeur
neuve, avant de vérifier la provenance. Le garde-fou a fonctionné parce qu'il était écrit ;
c'est l'argument pour continuer à écrire les règles négatives au même rang que les résultats.

## 8. Une limite de la chaîne d'analyse, trouvée par accident

`pic_proche` cherche le maximum dans ±2 bins autour de k·f0, soit une fenêtre de 5 bins de
10,77 Hz. Quand l'écart entre deux harmoniques tombe sous cette largeur, les fenêtres de deux
rangs voisins **rendent le même pic** : la liste des harmoniques contient des fréquences
dupliquées.

| Écart f0/bin | 1 | 2 | 3 | 4 | 5 | 6 | **≥ 7** |
|---|---|---|---|---|---|---|---|
| Notes avec doublon | 100 % | 100 % | 100 % | 74 % | 51 % | 28 % | **0 %** |

Soit **f0 ≥ 75 Hz**. En dessous la mesure est contaminée, au-dessus jamais. Cela touche
**4,8 % des notes du corpus SOL**, mais jusqu'à 40 % des notes d'un seul instrument :
clarinette contrebasse 40 %, tuba basse avec sourdine et ensemble de contrebasses 34 %, tuba
contrebasse et contrebasson 29 %, tuba basse 22 %, contrebasse avec sourdine 19 %, trombone
basse 16 %, contrebasse et cor 8 %. Chez les cuivres, les registres touchés sont exactement
les pédales et l'extrême grave.

### Le contrôle doit partager le confondant de ce qu'il contrôle

Première tentative : retirer les notes sous 75 Hz et comparer les listes de formants. Huit
instruments sur neuf changent, et **le formant le plus grave disparaît chaque fois** — le tuba
contrebasse perd 370 Hz, le cor 466, le contrebasson 329. Conclusion apparente : les notes
graves fabriquent de faux formants bas.

**C'est faux.** Retirer ces notes retire aussi le domaine grave : la tâche C2 a montré que le
plancher de la méthode 8 vaut 3,2 à 5,7 fois f0_min. Le formant bas disparaît faute de
données, pas faute de validité — et l'inverse est vrai, **ce sont ces notes-là qui le portent**.

Le contrôle correct **garde les notes et n'écarte que les fréquences répétées** :

| Instrument | Tel que publié | Dédoublonné, mêmes notes | Notes retirées *(confondu)* |
|---|---|---|---|
| Cor | 466 739 2216 4183 | **466 739 2216 4183** | 739 987 2348 4183 |
| Trombone basse | 879 1759 3320 | **879 1759 3320** | 879 1759 3320 |
| Tuba basse | 466 659 987 1479 | **466 659 987 1479** | 698 987 1479 3134 |
| Clarinette contrebasse | 1108 1759 2635 4432 | **1108 1759 2635 4432** | 1660 2635 4695 |
| Tuba contrebasse | 370 879 1396 | 370 **932** 1396 1974 | 879 1567 2348 4695 |
| Contrebasson | 329 698 1174 4432 | 329 **659** 1174 **1863** | 698 1174 2635 4432 |
| Contrebasse | 466 1396 2635 3948 | 466 **1318** 2635 3948 | 1396 1974 2792 3948 |

Et le test de non-régression, à la tolérance du résultat qu'il protège — `evaluer_migration.py`,
celui-là même qui avait démenti la migration de fenêtre du 27 octobre :

| | Référence | Avec dédoublonnage |
|---|---|---|
| Formants route A | 23 sur 13 instruments | **23 sur 13 instruments** |
| Unissons formantiques | 3 — 1 108, 1 244, 2 216 Hz | **3, les mêmes** |
| Régions de convergence | 6 | **6** |
| Instruments perdus / gagnés | — | **aucun** |

**Le dédoublonnage ne touche rien de ce qui est publié comme établi.** Il déplace trois valeurs
d'un ou deux pas de grille, toutes de statut « indéterminé », et toutes chez les trois
instruments les plus graves. La raison est mécanique : `residu_local` retranche une médiane
locale, et un point compté deux fois la repondère sans la déplacer.

### Décision

`DEDOUBLONNER` reste **inactif par défaut**, activable par `FORMANTS_DEDOUBLONNER=1`.

Le motif est celui que la migration de fenêtre a appris : **on ne migre pas sans gain.** Le
gain est nul sur tout ce qui est établi, et le coût ne l'est pas — il faudrait remesurer les
cinq corpus et repropager vingt-deux fiches pour déplacer trois valeurs déjà marquées
indéterminées. Le drapeau existe pour que la vérification soit reproductible, et le seuil est
documenté dans `formants_lib.py` au point exact où il mord.

En revanche les trois valeurs concernées — contrebasse 1 396, contrebasson 698, tuba
contrebasse 879 Hz — **doivent porter la mention** qu'elles sont mesurées en partie sur des
notes dont les harmoniques ne sont pas séparables.

## Ce qui reste ouvert

| Question | Ce qui la trancherait |
|---|---|
| **Pourquoi quatre accords sur six sont-ils au doux ?** | hypothèse : au fort la couleur dépend du joueur et du niveau. Testable en comparant l'écart entre corpus au doux et au fort sur les autres familles, où C5 ne trouve pas d'effet de nuance — si l'écart y est le même aux deux nuances, l'explication tient |
| **Faut-il republier la carte vocalique des cuivres par nuance ?** | oui selon ce document, et C11 le dit pour une autre raison. C'est une décision de publication, pas de mesure : les cellules existent déjà, c'est l'agrégation qui les réunit à tort |
| **Le trou B1-D#2 des trombones** | une ligne `registres.md` par trombone, bornes prises sur le corpus. Décision à prendre : le fichier décrit-il l'instrument ou l'échantillon ? |
| **Le cor s'étale sur 6,8 demi-tons entre corpus** et ne donne aucun accord strict, à aucune nuance | c'est le cuivre le moins bien couvert de l'étude : la tâche C4 a montré qu'il manquait aussi à l'analyse LPC de mai |
| **Le tuba basse donne 147–150 Hz au doux (SOL, VO) et 231–249 Hz (VN, ZSamples)** | un écart de 7,5 demi-tons à nuance égale entre deux paires qui s'accordent chacune en interne : il y a vraisemblablement deux instruments différents sous le même nom |
| La résolution de 10,77 Hz par bin | fixée par NFFT ; une analyse à fenêtre plus longue sur les seules notes graves lèverait la limite, au prix d'une chaîne parallèle |

## Journal

- **2026-10-28** — Tâche C7. La demande — assigner les pédales — **n'ouvre aucune cellule
  lisible** hors le trombone basse, et laisse 81 échantillons de trombone hors registre, la
  ligne `registres.md` étant écrite pour le ténor sans barillet et appliquée aussi à la basse,
  qui en a un. Mais la relance donne quatre choses. **La carte vocalique d'un cuivre ne se lit
  que jusqu'au médium** (F1/f0 passe de 10,8 à 1,0, pente 1,03 et r² 0,85 au suraigu), et le
  registre pédale est au contraire le plus propre. **Mon premier verdict — quinze mesures fixes
  sur dix-sept — était vrai et trompeur** : la pente contre f0 contrôle la note et non la
  nuance, et F1 monte de 9,5 demi-tons en médiane du doux au fort. **À nuance égale la
  réplication apparaît** : six accords stricts contre un, et le trombone basse au doux donne
  412 · 408 · 403 Hz sur trois provenances distinctes, soit 0,39 demi-ton. **Deux corrections
  de ma propre main** : la corroboration trompette/C6 est retirée (son instabilité est un effet
  de nuance, et C6 mesure la largeur quand ceci mesure la position) ; et le superbe accord du
  tuba contrebasse à 0,06 demi-ton **ne compte pas**, SOL et VN étant la même bibliothèque —
  la règle de provenance du 25 octobre l'avait prévu. Enfin une **limite de résolution à 75 Hz**
  qui contamine jusqu'à 40 % des notes d'un instrument sans rien déplacer de ce qui est publié.
