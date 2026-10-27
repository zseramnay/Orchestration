---
id: fenetre-de-redressement
titre: La fenêtre de redressement — ce qu'élargir change, et ce que cela coûte
date: 2026-10-27
tache: C2
script: Scripts/v7-formants-2026/fenetre_redressement.py
donnees: Etude-Formants/Mesures/Fenetre-redressement/
statut: expérience faite · migration tentée, puis annulée le jour même
---

# La fenêtre de redressement

## Le diagnostic publié était faux

`residu_local()` porte dans sa docstring, et `_CARTE-VOCALIQUE.md` reprend, que le redressement
local à ±3/4 d'octave est « aveugle aux résonances plus larges que ~1,5 octave ». Nous en avions
déduit que la résonance de caisse de l'alto — 375–395 Hz, établie en C1 sur quatre corpus —
échappait à la méthode 8 **parce qu'elle est trop large**, et que l'élargissement la ferait
apparaître.

Ce n'est pas la raison. La vraie raison est arithmétique, et elle est plus simple.

Un harmonique n'entre dans l'enveloppe composite que si sa fenêtre de ±w octave contient **au
moins trois harmoniques de la même note** — c'est la condition `q >= 3` de `composite()`. Or le
nombre d'harmoniques présents dans l'intervalle [k/2^w, k·2^w] croît avec le rang k : aux rangs
les plus bas, la fenêtre est trop étroite pour en contenir trois.

| Fenêtre | Rang le plus bas admis | Plancher théorique du composite |
|---|---|---|
| ±0,50 octave | 4 | 4 × f0_min |
| **±0,75 octave — la fenêtre actuelle** | **3** | **3 × f0_min** |
| ±1,00 à ±1,50 octave | 2 | 2 × f0_min |
| ±2,00 octave | 1 | 1 × f0_min |

En pratique le plancher est encore plus haut, parce qu'une bande n'est publiée que si elle
réunit au moins cinq points et cinq notes. **Mesuré sur les 22 instruments et les cinq corpus,
le plancher réel à ±3/4 d'octave se situe entre 3,2 et 5,7 fois f0_min.**

Pour l'alto, f0_min = 131 Hz : le plancher vaut 523 à 659 Hz selon le corpus. Pour le
violoncelle, f0_min = 65 Hz : 261 à 329 Hz. **Les résonances établies en C1 — 384 Hz à l'alto,
219 Hz au violoncelle — sont en dessous de ce plancher dans les quatre corpus.** Elles ne sont
pas redressées : elles sont hors du domaine calculé. L'enveloppe composite n'y porte aucune
valeur, ce que l'on vérifie directement dans les JSON de mesure.

Élargir la fenêtre répond donc bien au problème, mais par un mécanisme différent de celui
annoncé : cela **abaisse le plancher du domaine**, au lieu de « voir plus large ».

## Ce que l'élargissement fait gagner

Passer de ±0,75 à ±1,00 octave abaisse le plancher du composite de **5 demi-tons en médiane**,
sur 66 mesures instrument × corpus couvrant les 22 instruments.

| Instrument · corpus | Plancher à ±0,75 | Plancher à ±1,00 | Gain |
|---|---|---|---|
| Trombone ténor · SOL | 233 Hz (×4,0) | 116 Hz (×2,0) | **12 dt** |
| Violoncelle · ZSamples | 277 Hz (×4,2) | 147 Hz (×2,2) | 11 dt |
| Violoncelle · SOL et VN | 261 Hz (×4,0) | 147 Hz (×2,2) | 10 dt |
| Violon · VO | 987 Hz (×5,0) | 587 Hz (×3,0) | 9 dt |
| Hautbois · SOL, Xsamples, VN | 932 Hz (×4,0) | 698 Hz (×3,0) | 5 dt |
| Tuba basse · ZSamples et VO | 87 Hz | 87 Hz | 0 dt |

Les cas à gain nul sont ceux dont le corpus descend déjà très bas : le plancher y est fixé par
la règle des cinq notes, pas par le rang.

Et dans la zone établie par la méthode 12, le résidu apparaît, important, et croît avec la
fenêtre :

| Alto, zone 384 Hz | ±0,75 | ±1,00 | ±1,25 | ±1,50 | ±2,00 |
|---|---|---|---|---|---|
| SOL | **aucune valeur** | 1,5 dB | 4,7 | 8,2 | 10,8 |
| ZSamples | **aucune valeur** | 8,3 dB | 10,5 | 11,2 | 13,1 |
| VO | **aucune valeur** | 11,7 dB | 15,8 | 12,9 | 14,2 |
| VN | **aucune valeur** | 0,7 dB | 4,3 | 4,8 | 8,4 |

Pour mémoire, le seuil de rétention est de 1 dB. Le violoncelle montre la même chose à 219 Hz,
mais plus tard — il faut ±1,25 à ±2,00 selon le corpus, son f0_min étant plus bas.

## Ce que l'élargissement coûte

C'est le test décisif, et il tranche. Pour chacun des formants déjà établis, on demande s'il se
retrouve encore à moins d'un tiers d'octave — la tolérance que `robustesse_fenetre()` emploie
déjà.

| Fenêtre | Formants route A retrouvés | Formants « 1 corpus » retrouvés |
|---|---|---|
| ±0,75 — l'actuelle | 72/76 · 95 % | 44/62 · 71 % |
| **±1,00** | **73/76 · 96 %** | **46/62 · 74 %** |
| ±1,25 | 69/76 · 91 % | 41/62 · 66 % |
| ±1,50 | 58/76 · 76 % | 39/62 · 63 % |
| ±2,00 | 52/76 · 68 % | 28/62 · 45 % |

**±1,00 octave ne coûte rien : il rapporte.** Au-delà, la dégradation est rapide et devient
sévère — à ±2,00, un formant route A sur trois ne se retrouve plus, et les listes de formants
changent entièrement : le violoncelle chez SOL passe de 2 216 · 1 108 · 3 948 · 554 à
2 091 · 1 318 · 349 · 494. Ce n'est pas un affinement, c'est une autre réponse.

Les ±1,5 et ±2,0 que le plan de reprise proposait sont donc **écartés par la mesure**.

### Le détail du passage à ±1,00

Sept formants établis se retrouvent qui manquaient, quatre se perdent.

| | Instrument · corpus | Formant |
|---|---|---|
| **retrouvé** | Violoncelle · VO | **1 108 Hz** — renforce une convergence publiée (basson · violoncelle) d'un corpus |
| retrouvé | Contrebasse · VO | 3 948 Hz |
| retrouvé | Cor anglais · ZSamples | 4 183 Hz |
| retrouvé | Flûte contrebasse · ZSamples | 1 759 Hz |
| retrouvé | Tuba basse · VN et ZSamples | 987 Hz |
| retrouvé | Violoncelle · VN | 554 Hz |
| perdu | Alto · VO | 1 318 Hz — c'est la valeur déjà signalée comme fausse de dix demi-tons sur ZSamples |
| perdu | Cor en Fa · Xsamples | 4 183 Hz |
| perdu | Tuba basse · ZSamples | 1 479 Hz |
| perdu | Violoncelle · VN | 2 216 Hz |

Le bilan qualitatif va dans le même sens que le bilan chiffré : ce qui est gagné consolide des
résultats existants, et la perte la plus notable est une valeur dont nous savions déjà qu'elle
était douteuse.

## Ce qui reste bloqué, et pourquoi ce n'est pas la fenêtre

Même à ±1,00, **aucun pic n'est rapporté** à 384 Hz pour l'alto ni à 219 Hz pour le
violoncelle, malgré des résidus de 8 à 16 dB. La cause n'est plus le redressement : c'est
`trouver_formants()`.

Un maximum n'y est retenu que si ses deux voisins sont eux aussi des bandes publiées
(`ok[i-1] and ok[i+1]`). **Une bande située au plancher du domaine n'a pas de voisin de gauche :
elle ne peut donc jamais être rapportée comme formant.** Or à ±1,00 le plancher de l'alto est
précisément 392 Hz, c'est-à-dire la bande de la résonance.

Cette règle est justifiée : au bord du domaine, rien ne distingue un maximum d'un flanc
montant. La supprimer inviterait exactement le genre d'artefact que cette reprise a servi à
éliminer. **Il ne faut donc pas la supprimer, et il ne faut pas non plus abaisser les seuils de
cinq notes et cinq points pour forcer le passage.**

La conclusion juste est plus modeste et plus sûre : dans cette zone, la méthode 8 peut
**confirmer une énergie** mais non **localiser un maximum**. C'est la méthode 12 qui localise, et
c'est exactement la hiérarchie déjà établie en C1 — méthode 12 pour la caisse des cordes,
méthode 8 pour les résonances internes au domaine.

Ce qu'on peut faire sans rien affaiblir : marquer ces bandes d'un statut explicite,
« limite basse du domaine — énergie confirmée, maximum non vérifiable », et le publier comme
tel. Un tel statut dirait quelque chose de vrai, là où aujourd'hui la fiche ne dit rien.

## La migration a été tentée — et annulée

**Le 27/10 à la mi-journée**, la fenêtre a été portée à ±1,00, les quatre corpus régénérés
depuis leurs fichiers de spectres, et SOL remesuré. Le résultat a fait annuler la migration
dans l'heure.

| | ±0,75 — l'actuelle | ±1,00 |
|---|---|---|
| Formants route A | **23** | 21 |
| Instruments porteurs | 13 | 13 |
| **Unissons formantiques** | **3** | **0** |
| Régions partagées | 6 | 6, mais d'autre composition |

**Les trois unissons disparaissent** — 1 108 Hz (basson · violoncelle), 1 244 Hz (clarinette
basse · tuba basse), 2 216 Hz (cor · violon). Ce sont les résultats les mieux établis de
l'étude et la matière du chapitre de synthèse. Les régions changent aussi de composition : la
région du basson passe de 1 108 à 1 874–1 974 Hz, appariée au trombone ténor et non plus au
violoncelle.

Tout a été restauré depuis `Archives/halfwin-0.75-2026-10-27/`, et le contrôle repasse : 23
formants, 13 instruments, 3 unissons, 6 régions, à l'identique.

### Pourquoi la recommandation était fausse

**Le test de dégâts mesurait la mauvaise chose.** Il demandait si chaque formant établi se
retrouve **à moins d'un tiers d'octave** — quatre demi-tons — parce que c'est la tolérance
qu'emploie `robustesse_fenetre()`. À cette tolérance, 96 % des formants route A survivaient à
±1,00, contre 95 % à ±0,75 : la migration paraissait gratuite.

Mais **la route A n'exige pas un tiers d'octave, elle exige un demi-ton** entre deux corpus de
provenances distinctes. Une dérive d'un demi-ton par corpus est invisible au premier critère et
fatale au second. Et c'est exactement ce qui se produit : les listes par corpus bougent à 60–77 %
d'une configuration à l'autre, le plus souvent d'un ou deux demi-tons — 494 → 466, 1 863 → 1 759,
2 216 → 1 974. Chaque valeur reste « au même endroit » à quatre demi-tons près, et plus aucune
paire ne tient à un demi-ton.

La leçon est une leçon de méthode, et elle vaut au-delà de cette tâche : **un test de
non-régression doit employer la tolérance du résultat qu'il protège, pas celle d'un autre test
du pipeline.**

## Ce qui est retenu malgré tout

1. **Le diagnostic corrigé.** Ce n'est pas la largeur mais le plancher du domaine qui fait
   manquer la caisse de l'alto et du violoncelle. C'est écrit dans la docstring de
   `residu_local()`, dans `00-METHODE.md`, `_CARTE-VOCALIQUE.md` et `_AMORCE-SESSION.md`.
2. **`HALFWIN` porte désormais un avertissement** qui dit de ne pas le porter à 1,00, et
   pourquoi, avec les chiffres.
3. **`robustesse_fenetre()` est corrigée.** Ses fenêtres de comparaison suivent maintenant la
   référence (`DELTA_ROB = 0,25`). À la référence de 0,75 cela redonne exactement 0,5 et 1,0,
   donc le comportement publié est inchangé — mais le test est désormais juste si la référence
   bouge un jour.
4. **Les deux limites de la méthode 8 sont mesurées** et non plus supposées : plancher à 3,2–5,7
   fois f0_min, et aveuglement aux résonances plus larges que ~1,5 octave.

## Ce qui reste ouvert

| Question | Ce qui la trancherait |
|---|---|
| Peut-on descendre le plancher **sans** déplacer les pics déjà établis ? | une fenêtre asymétrique, ou un redressement à fenêtre variable : large seulement sous le plancher actuel, inchangée au-dessus. C'est la piste à essayer, et elle n'a pas été testée |
| Le statut « limite basse du domaine » vaut-il d'être publié ? | il dirait quelque chose de vrai là où la fiche se tait aujourd'hui |
| Les flûtes relèvent-elles bien de la limite de **largeur** et non du plancher ? | leur f0_min est haut (587 Hz pour la petite flûte), donc le plancher ne les gêne pas dans l'aigu. La coupure de rayonnement reste l'explication, mais elle n'a pas été re-testée ici |

## Journal

- **2026-10-27 (après-midi)** — Migration en ±1,00 tentée puis **annulée**. Les trois unissons
  formantiques disparaissaient et la route A tombait de 23 à 21. La recommandation du matin
  était fausse : le test de dégâts employait la tolérance d'un tiers d'octave du test de
  robustesse, là où la route A exige un demi-ton. Tout est restauré et le contrôle repasse à
  l'identique. Conservés du passage : le diagnostic corrigé, l'avertissement sur `HALFWIN`, et
  `robustesse_fenetre()` dont les fenêtres suivent désormais la référence.
- **2026-10-27** — Tâche C2. Le diagnostic publié — « aveugle aux résonances larges » — est
  remplacé par le vrai : le plancher du domaine, fixé par la règle des trois harmoniques dans la
  fenêtre. ±1,5 et ±2,0, que le plan proposait, sont écartés par le test de dégâts. ±1,00 est
  recommandé : il est meilleur que la fenêtre actuelle sur les deux critères. L'obstacle restant
  n'est pas la fenêtre mais la règle de voisinage de `trouver_formants()`, qu'il ne faut pas
  lever.
