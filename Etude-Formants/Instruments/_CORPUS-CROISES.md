---
id: corpus-croises
titre: Cinq corpus croisés — Xsamples, VO, VN versés dans l'étude
date: 2026-10-25
corpus: SOL2020+Yan_Adds · ZSamples · Xsamples · VO · VN
---

# Cinq corpus croisés

> **À lire avec** `_REPLICATION.md` (le critère à deux routes) et
> `_VERIFICATION-GRAVES.md` (la résolution au demi-ton).

Trois banques nouvelles ont été versées le 25/10 : **Xsamples** (1 404 `.aif`), **VO**
(1 938 `.AIF`) et **VN** (1 034 `.wav`, dix instruments). Avec SOL2020+Yan_Adds et ZSamples,
cela fait **cinq corpus indépendants**. Mesures dans `Mesures/Xsamples/`,
`Mesures/VO/` et `Mesures/VN/`.

## Le résultat en une ligne

> **23 formants établis par la route A, sur 13 instruments** — contre 7 sur 5 instruments le
> 24 octobre. Et trois instruments restent sans aucun accord malgré quatre corpus, ce qui est
> aussi un résultat.

## ⚠ VN est du VSL, et Yan_Adds aussi — une règle de provenance

**À lire avant tout le reste.** Les abréviations de VN sont celles de VSL en allemand : KFA =
*Kontrafagott*, KTU = *Kontratuba*, BKL = *Bassklarinette*, TU = *Tuba*, TrC = *Trumpet C*,
EH, HO, VC, VA, VI. VN est donc un extrait de la **Vienna Symphonic Library**.

Or `Yan_Adds-Divers`, qui fournit le côté « SOL » de dix instruments sur vingt-deux, contient
du VSL. Vérification directe sur trois paires :

| Instrument | Yan_Adds | VN | Corrélation des spectres |
|---|---|---|---|
| **Contrebasson** | 362 564 échantillons | **362 564** | **1,00000** |
| **Tuba contrebasse** | 206 529 échantillons | **206 529** | 0,99591 |
| **Cor anglais** | 432 187 échantillons | **432 187** | 0,99863 |

Nombre d'échantillons identique **au sample près**, corrélation jusqu'à 1,00000 : **ce sont
les mêmes enregistrements**, à un gain de normalisation près. Un « accord entre deux corpus »
y était la même bibliothèque comparée à elle-même.

**Règle retenue : la paire (SOL, VN) ne compte pas lorsque la source SOL de l'instrument est
`Yan_Adds-Divers`.** Elle reste valide pour les douze instruments venus de `FullSOL2020`
(IRCAM) : flûte, hautbois, basson, clarinette en Sib, cor, trompette, trombone ténor, tuba
basse, violon, alto, violoncelle, contrebasse. Et VN reste comparable sans réserve à ZSamples,
Xsamples et VO pour **tous** les instruments.

Quatre formants annoncés plus tôt le 25/10 sont **annulés** par cette règle :

| Formant annulé | Pourquoi |
|---|---|
| Contrebasson 329 · 698 · 1 174 Hz | la seule paire stricte était Yan_Adds ↔ VN |
| Tuba contrebasse 1 396 Hz | idem — et c'était son unique second corpus |

Le contrebasson redevient non répliqué, et le **tuba contrebasse n'a toujours aucun second
corpus**. Ils sont exactement là où ils étaient le 24 octobre.

**Audit de doublons sur les cinq banques** : 750 fichiers indexés par durée exacte, dix paires
de corpus testées. Une seule durée commune hors VSL, entre un alto de VO et un hautbois de VN —
coïncidence. **SOL/FullSOL2020, ZSamples, Xsamples et VO sont mutuellement indépendants.**

À refaire pour toute banque future : avant de compter une réplication, **vérifier que les deux
corpus ne sont pas la même bibliothèque**. La durée en échantillons est une empreinte
suffisante et coûte quelques secondes.

## Ce que chaque banque a imposé d'adapter

Aucune des trois ne se lisait en l'état. Le pilote `Scripts/v7-formants-2026/corpus.py` absorbe
désormais les cinq nomenclatures sans qu'on touche à la chaîne de mesure : il écrit dans le
fichier de spectres un **nom normalisé au format SOL**, et tout l'aval fonctionne.

| Banque | Structure | Adaptation nécessaire |
|---|---|---|
| SOL2020 | `Instrument/technique/inst-tech-NOTE-dyn-N-N.wav` | aucune (référence) |
| ZSamples | technique dans le nom, `_string_N` | existante |
| **Xsamples** | `instrument/DYNAMIQUE/code_note.aif` | notes en **minuscules** · dynamique et technique dans le **dossier** · **H allemand** |
| **VO** | `instrument/sourdine/ordinario/String-N/code-NOTE-dyn.AIF` | technique sur **plusieurs niveaux** de dossier · **correction d'accord** |
| **VN** | `Instrument/TAG_dyn/TAG_dyn_NOTE.wav` | dynamique **avant** la note |

### Le H allemand — 15 fichiers de Xsamples

Le cor et le trombone de Xsamples écrivent `Corf_h2`, `posf_h3` : **H = si naturel**, notation
allemande. Tous les autres instruments de la même banque écrivent `b`. Vérifié mutuellement
exclusif par dossier — aucune ambiguïté, `b` reste l'anglais. Sans ce correctif, 15 fichiers
étaient perdus et le cor perdait son registre grave.

### La correction d'accord — VO entière

C'est le point le plus sérieux. `harmoniques()` relève chaque harmonique dans une fenêtre de
**±2 bins** (±21,5 Hz) autour de k·f0 **calculé depuis le nom de fichier**. Si la banque n'est
pas accordée à La = 440, l'écart dépasse la fenêtre dès que k·f0 > 21,5 / |r−1|, et tous les
harmoniques au-dessus sont relevés **dans un creux**.

| Instrument VO | Écart | Dernier harmonique correctement relevé |
|---|---|---|
| Flûte, petite flûte | **+30 / +35 cents** | **k = 1** |
| Flûte contrebasse | **+48 cents** | k = 3 |
| Clarinettes | +15 cents | k = 4 à 9 |
| Trombones, alto | +20 cents | k = 11 à 13 |
| Tuba basse | +30 cents | k = 12 |
| Violon, violoncelle, contrebasse | +10 / −5 / −7 cents | k = 5 à 33 |

Les vents de VO sont accordés autour de **La = 444**, diapason d'orchestre courant, les flûtes
plus haut encore ; les cordes sont à 440. **La première série de mesures VO était invalide**
pour tous les vents. Correction : recalage du spectre en fréquence **par échantillon**, puis
remise des formants dans le référentiel réel. Après correction, nomenclature = 1,000 sur les
quinze configurations.

À retenir pour toute banque future : **une nomenclature dans la tolérance 0,97–1,03 n'est pas
une garantie.** Quinze cents passent le contrôle et détruisent la mesure au-dessus de 1 kHz.
Le seuil utile est 0,2 %, soit 3 cents.

### Qualité comparée des cinq banques

| Banque | Accord | σ | Dynamiques | Jugement |
|---|---|---|---|---|
| **Xsamples** | **1,0000** | 0,06–0,2 % | 3 (p / mf / f) | **irréprochable** |
| **VN** | **1,0000** | 0,13 % | 3 à 4 | **irréprochable** |
| SOL2020 | 1,00 | — | variable | référence |
| ZSamples | 1,00 | — | 4 | bonne, échantillonnage espacé |
| VO | 1,003 à 1,028 | 0,6–2,6 % | 3 (pp / mf / ff) | exploitable **après correction** |

## Les 23 formants établis par la route A

Route A = au moins un **couple** de corpus sur le **même point de grille** (< 0,5 demi-ton),
de provenances distinctes.

> **Correction du 26/10/2026.** Ce tableau en annonçait **20 sur 12 instruments**. Le comptage
> regroupait les valeurs en zones *en chaînant depuis la première*, de sorte que deux valeurs
> qui s'apparient pouvaient tomber dans des zones voisines sans jamais être confrontées.
> L'appariement refait sans regroupement, donc indépendant de l'ordre, en donne **23 sur 13**.
> Les trois retrouvés sont l'**alto à 3 748 Hz**, la **clarinette basse à 2 225 Hz** et le
> **tuba basse à 498 Hz**. Le tableau ci-dessous est généré depuis
> `Resultats-v7/formants-v7.csv`, lui-même produit par `exporter_csv.py` : il ne peut plus
> diverger des données.

La colonne « portée par SOL » dit si une fiche d'instrument contient cette valeur. Sept zones
sont établies par des couples qui n'incluent pas SOL : elles existent sans qu'aucune fiche ne
les porte.

| Instrument | Formant | Corpus de la zone | Portée par SOL |
|---|---|---|---|
| Alto | **3 748 Hz** (La♯7) | VO · ZSamples | — |
| Basson | **494 Hz** (Si4) | Xsamples · ZSamples | — |
| Basson | **1 108 Hz** (Do♯6) | SOL · ZSamples · Xsamples | oui |
| Basson | **1 974 Hz** (Si6) | SOL · ZSamples · Xsamples | oui |
| Clarinette basse en Sib | **1 244 Hz** (Ré♯6) | SOL · Xsamples · VN | oui |
| Clarinette basse en Sib | **2 225 Hz** (Do♯7) | VN · VO | — |
| Clarinette contrebasse en Sib | **1 759 Hz** (La6) | SOL · VO | oui |
| Clarinette en Sib | **3 517 Hz** (La7) | SOL · ZSamples · Xsamples | oui |
| Contrebasse | **2 635 Hz** (Mi7) | SOL · ZSamples | oui |
| Cor en Fa | **2 216 Hz** (Do♯7) | SOL · Xsamples · VN | oui |
| Trombone basse | **879 Hz** (La5) | SOL · ZSamples · VO | oui |
| Trombone basse | **1 759 Hz** (La6) | SOL · ZSamples · VO | oui |
| Trombone ténor | **830 Hz** (Sol♯5) | SOL · ZSamples · Xsamples · VO | oui |
| Trompette en Ut | **2 792 Hz** (Fa7) · ⚠ non localisé | SOL · VN | oui |
| Trompette en Ut | **4 695 Hz** (Ré8) | SOL · Xsamples · VN | oui |
| Tuba basse | **498 Hz** (Si4) | VO · ZSamples | — |
| Tuba basse | **659 Hz** (Mi5) | SOL · ZSamples · VN | oui |
| Tuba basse | **1 244 Hz** (Ré♯6) | VN · ZSamples | — |
| Violon | **2 348 Hz** (Ré7) | SOL · ZSamples · VN | oui |
| Violon | **3 959 Hz** (Si7) | VO · ZSamples | — |
| Violoncelle | **1 108 Hz** (Do♯6) | SOL · VN | oui |
| Violoncelle | **2 631 Hz** (Mi7) | VN · VO | — |
| Violoncelle | **3 948 Hz** (Si7) | SOL · VN | oui |

### Les trois dossiers VSL ajoutés le 25/10

Tuba basse, trompette en ut et clarinette basse, tirés de
`/Volumes/Samps/VSL/Samples/…/Einzelnoten/`, versés dans VN. Ils ont apporté quatre formants :

- **Trompette 2 792 Hz réhabilitée.** Elle était tombée le matin même sur la comparaison
  retenu contre retenu ; VN la rétablit par une paire stricte avec SOL, dont la trompette vient
  de `FullSOL2020` — provenance indépendante, l'accord est donc valide.
- **Trompette 4 974 Hz**, nouvelle, paire VN + Xsamples.
- **Tuba basse 698 et 1 244 Hz.** C'était l'instrument où « la troisième source avait joué
  contre » : VO contredisait les deux zones que SOL et ZSamples donnaient à un demi-ton. VN
  tranche en donnant **exactement les valeurs de ZSamples** sur les deux. Le tuba basse passe
  de deux indéterminés à deux formants établis.
- La **clarinette basse** ne peut pas être appariée à SOL (Yan_Adds = VSL), mais son 2 216 Hz
  concorde à 0,15 demi-ton avec le 2 235 Hz de VO — à verser en route A après vérification.

## Les cinq instruments que les nouvelles banques ont tranchés

### Le violoncelle — de « aucun accord » à trois zones strictes

Trois corpus ne partageaient aucune zone. VN donne **1 108 · 2 635 · 3 948 Hz** :

- **1 108 Hz** — identique à SOL
- **3 948 Hz** — identique à SOL
- **2 635 Hz** — à 0,05 demi-ton du 2 627 Hz de VO

Trois zones strictes d'un coup, par trois paires de corpus différentes.

### Le violon — le *bridge hill* passe en route A

Il ne tenait qu'à un demi-ton entre SOL et ZSamples, et VO ne le voyait pas. **VN donne
2 216 Hz, exactement la valeur de ZSamples**, SOL étant à un demi-ton avec 2 348 Hz. Trois
corpus, une paire stricte : le rapprochement avec le *bridge hill* de la littérature
acoustique est désormais appuyé par la route A et non plus par la seule corroboration.

### Le cor — réhabilité

Classé **« non répliqué, aucun formant établi »**, ses valeurs SOL et ZSamples étant à
5,0 demi-tons de moyenne. SOL et VN donnent tous deux **2 216 Hz**, Xsamples 2 348 Hz à un
demi-ton. Le désaccord entre deux corpus ne suffisait pas à conclure.

### Le tuba basse — de deux indéterminés à deux établis

SOL et ZSamples s'accordaient à un demi-ton sur 466–494 et 659–698, et **VO contredisait les
deux** (282 · 502 · 845 · 1 341 Hz). VN donne **698 et 1 244 Hz**, soit exactement deux valeurs
de ZSamples. Deux paires strictes, provenances indépendantes : établis.

### Le tuba contrebasse — toujours sans second corpus

VN donnait **392 · 932 · 1 396 · 2 091 Hz** contre **370 · 879 · 1 396 Hz** pour SOL, dont
1 396 Hz exact. Mais la source SOL de cet instrument est `Yan_Adds-Divers`, donc du VSL : c'est
le même enregistrement (206 529 échantillons de part et d'autre, corrélation 0,996). **Accord
annulé.** Le tuba contrebasse reste le seul instrument de l'étude sans aucun second corpus
indépendant.

### Le contrebasson — annulé lui aussi

VN et SOL donnaient **329 · 698 · 1 174 Hz** tous les deux, les trois formants au Hz près. La
corrélation des spectres vaut **1,00000** : mêmes fichiers. Il redevient non répliqué —
6,0 demi-tons entre SOL et ZSamples. Xsamples place 739 Hz à un demi-ton de 698 : c'est la
seule piste réelle, et elle ne suffit pas.

## Des convergences entre instruments, cette fois réelles

Celles de l'ancienne étude étaient des bins de FFT voisins. Celles-ci portent sur des formants
établis **chacun indépendamment**, par route A, dans des corpus différents.

| Zone | Instruments | Statut de chacun |
|---|---|---|
| **2 216 Hz** (Do♯7) | **cor · violon** | route A tous les deux, par deux paires différentes (VN+SOL et VN+ZSamples) |
| **1 108 Hz** (Do♯6) | **basson · violoncelle** | route A tous les deux, 3 et 2 corpus |
| **879 Hz** (La5) | **trombone ténor · trombone basse** | route A tous les deux, 4 et 3 corpus |
| **1 244 Hz** (Ré♯6) | **clarinette basse · tuba basse** | route A tous les deux |
| **3 948–3 960 Hz** | **violon · violoncelle** | route A tous les deux |

La convergence **cor–violon à 2 216 Hz** est la plus intéressante : deux familles, quatre
corpus en tout, chacun validé séparément. Et elle n'a rien à voir avec la convergence à
4 200 Hz annoncée puis écartée — celle-là reposait sur le cor et le hautbois dans le seul
corpus SOL, et ni VN ni Xsamples ne la retrouvent (VN donne 4 695 Hz au cor, Xsamples 3 517).

## Ce qui ne se résout pas, et c'est un résultat

### Le hautbois — quatre corpus, quatre réponses incompatibles

| Corpus | Formants retenus | ρ |
|---|---|---|
| SOL | 4 183 Hz | — |
| ZSamples | **aucun** | — |
| Xsamples | 2 635 · 4 974 Hz | 0,72 |
| VN | 3 320 Hz | **0,78** |

**Aucune zone commune.** Quatre sources indépendantes, dont deux irréprochables d'accord, et
pas deux valeurs à moins de quatre demi-tons. Avec un ρ de 0,72 et 0,78, le spectre du
hautbois **suit la hauteur**.

C'est un résultat fort et contre-intuitif : le hautbois est l'instrument auquel la littérature
assigne un formant avec le plus d'assurance — Backus 1 300 Hz, Meyer et Giesler ~1 100 Hz. Nos
quatre corpus ne le confirment pas. **Il faut le reclasser en B** et dire dans sa fiche que la
notion de formant fixe ne s'y applique pas, comme pour les flûtes.

### L'alto — un formant, mais pas celui de SOL

| Corpus | Formants retenus |
|---|---|
| SOL | 1 318 · 2 091 · 3 517 · 4 974 Hz |
| ZSamples | 2 348 · 3 726 · 5 583 Hz |
| VO | 3 770 · 5 331 Hz |
| VN | 1 759 · 4 432 Hz |

> **Correction du 26/10/2026.** Cette section affirmait que l'alto était « le seul membre du
> quatuor sans formant établi ». **C'est faux.**

**VO donne 3 770 Hz et ZSamples 3 726 Hz : 0,20 demi-ton d'écart, deux provenances distinctes.
C'est un couple strict, donc une route A, à 3 748 Hz (La♯7).**

Ce qui reste vrai, et qui est la nuance à retenir : **SOL ne le trouve pas.** Sa valeur la plus
proche, 3 517 Hz, est à 1,1 demi-ton du couple — au-delà du seuil. Le formant existe, mais ce
n'est pas celui que mesure le corpus de l'IRCAM. La fiche de l'alto garde donc « indéterminé »
pour sa propre valeur, et la zone établie est portée à part dans `Resultats-v7/formants-v7.csv`.

L'autre zone, 5 331 (VO) / 5 583 (ZSamples), reste à un demi-ton sans couple strict.

Et l'alto demeure le membre du quatuor le moins bien établi : un formant, contre deux au violon
et trois au violoncelle, et aucun que SOL confirme.

### La famille des flûtes — résultat négatif, désormais massif

| Instrument | Corpus | Accord |
|---|---|---|
| **Petite flûte** | SOL · ZSamples · VO | **aucun formant dans les trois** |
| Flûte | SOL · ZSamples · VO | aucun accord (SOL 3 320 Hz non répliqué) |
| Flûte basse | SOL · ZSamples · Xsamples · VO | aucun accord |
| Flûte contrebasse | SOL · ZSamples · VO | aucun accord |

Un résultat négatif répliqué ne souffre pas du problème de comparaisons multiples : il n'y a
rien à apparier par hasard. **Toute la famille des flûtes est confirmée sans formant fixe**, et
le piccolo l'est par trois corpus qui ne trouvent rien du tout. C'est la réponse définitive à
l'anomalie initiale : le piccolo ne converge avec rien parce qu'il n'a pas de formant à faire
converger.

### Le tuba contrebasse et le contrebasson — sans second corpus indépendant

Le tuba contrebasse n'existe que dans SOL (via Yan_Adds). Le contrebasson existe dans trois
corpus mais n'y trouve aucune zone commune hors la paire VSL annulée. Pour ces deux
instruments, **l'étude ne peut rien affirmer**, et il faudrait une banque qui ne soit ni VSL ni
ZSamples.

## Une inconsistance du 24/10 qu'il faut acter

Les tableaux ci-dessus comparent, des deux côtés, les formants **retenus**. Le tableau du 24/10
dans `_REPLICATION.md` comparait les formants retenus de SOL aux **candidats** de ZSamples,
retenus ou non — plus permissif.

Conséquence vérifiable : la **trompette à 2 792 Hz**, annoncée le 24/10 comme « le seul formant
de cuivre aigu à passer la réplication », ne tient pas sur une comparaison retenu contre
retenu — ZSamples ne retient que 1 479 Hz à cet endroit. `_REPLICATION.md` doit être repris sur
ce point.

## À faire

> Les travaux restants sont **tous** dans `_PLAN-DE-REPRISE.md`, qui est la seule liste à tenir. Ne pas en ouvrir une ici : c'est ainsi que huit listes se sont contredites entre le 23 et le 26 octobre.

## Le vibrato comme voie de mesure — testé, et écarté pour l'instant

> **Ne pas réexplorer sans un corpus au vibrato plus profond.** Noté ici pour éviter de
> refaire le test.

L'idée est juste : le vibrato fait balayer chaque harmonique à travers le filtre, donc une
**seule note** peut révéler la pente locale de la fonction de transfert — ce qui serait
précieux pour les instruments de classe B, où aucune agrégation ne marche. Le principe est
d'appartenir à la corrélation entre fréquence instantanée et amplitude d'un harmonique : si
l'amplitude monte quand la fréquence monte, la résonance est au-dessus.

Testé sur `15 Oboe/Einzelnoten/OB-L_pV_pA_sus` (68 fichiers).

**1. La méthode 8 n'y gagne rien.** Hautbois vibrato : 3 517 et 5 583 Hz, **ρ = 0,80** — le
plus élevé des quatre mesures du hautbois, donc le vibrato ne réduit pas le caractère suiveur.
Verdicts tous en critère assoupli. Rien dans la zone 1 100–1 500 Hz de Backus, Meyer et
Giesler. Seul recoupement : 5 583 Hz, que ZSamples donne aussi mais en *fragile* des deux côtés.

**2. Premier test de corrélation faussé, et pourquoi.** Avec une fenêtre de 93 ms — une
demi-période de vibrato — l'excursion est lissée et les corrélations se contredisent : +0,79 à
1 761 Hz sur un La4, −0,77 à 1 759 Hz sur un La5. Deux causes : la fenêtre, et le **couplage
justesse–intensité** que les instrumentistes produisent naturellement, qui affecte tous les
harmoniques ensemble. Ce second effet se retire par un résidu local **par trame** (chaque
harmonique moins la moyenne des harmoniques de la même trame).

**3. La profondeur de vibrato est trop faible** — c'est ce qui bloque. Fenêtre de 23 ms, pic
interpolé par parabole, médiane sur les six premiers harmoniques :

| Jeu | Profondeur | Lecture |
|---|---|---|
| **Hautbois `pV`** | **±10,4 cents** | réel, 4× le plancher, mais faible |
| Hautbois `pA` sans vibrato | ±2,5 cents | plancher de bruit de la mesure |
| Cor anglais VN | ±1,4 cents | plancher |
| **Alto SOL2020** | **±14,5 cents** | le plus profond trouvé |
| Violon SOL2020 | ±11,5 cents | — |
| Alto VN | ±4,2 cents | quasi sans vibrato |
| Violoncelle VN | ±6,1 cents (q75 31, max 100) | bimodal — à vérifier, glissandos probables |

Un instrumentiste en concert fait ±30 à ±50 cents. **Aucun corpus en main n'en approche.** À
±10 cents, sur une résonance large d'un tiers d'octave, la modulation d'amplitude attendue est
de l'ordre de **0,3 dB** : détectable seulement par moyennage massif, et au prix du contrôle du
couplage ci-dessus. Le rapport effort/rendement est mauvais.

**Conclusion.** La voie reste valable en principe et redeviendrait intéressante avec des
enregistrements au vibrato franc — c'est une chose à demander si un enregistrement est
commandé. En l'état, pour l'alto et le hautbois, **la méthode 12 est la bonne réponse** : elle
ne suppose aucune résonance fixe, donc elle ne souffre pas de leur absence.

## Reproduire

```bash
cd Scripts/v7-formants-2026
python3 corpus.py --decrire "/chemin/Banque" > /tmp/d.json
python3 corpus.py --valider /tmp/d.json
python3 corpus.py --mesurer /tmp/d.json
```

VO exige `"corriger_accord": true`.

## Journal

- **2026-10-25 matin** — Versement de Xsamples, VO et VN. Pilote générique `corpus.py` : trois
  adaptations de nomenclature et une correction d'accord. Famille des flûtes confirmée sans
  formant ; hautbois et alto non résolus malgré quatre corpus.
- **2026-10-25 après-midi** — **VN identifié comme un extrait VSL, et `Yan_Adds-Divers` aussi.**
  Trois paires de fichiers vérifiées identiques au sample près (corrélation jusqu'à 1,00000).
  Règle de provenance instaurée ; quatre formants annulés (contrebasson ×3, tuba contrebasse).
  Audit de doublons : SOL/FullSOL2020, ZSamples, Xsamples et VO sont mutuellement indépendants.
- **2026-10-25** — Trois dossiers VSL ajoutés (tuba basse, trompette, clarinette basse).
  Trompette 2 792 Hz réhabilitée, trompette 4 974 Hz et tuba basse 698 et 1 244 Hz établis.
  Bilan net : **23 formants par route A sur 13 instruments**, contre 7 sur 5 la veille.
- **2026-10-26** — Correction du comptage : le regroupement séquentiel en masquait trois, dont
  celui de l'alto. Appariement refait sans regroupement. Le tableau est désormais généré depuis
  `Resultats-v7/formants-v7.csv`.
