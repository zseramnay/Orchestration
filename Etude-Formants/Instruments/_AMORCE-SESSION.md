# Amorce de session — reprendre l'étude formantique dans une conversation vierge

Ce fichier permet de relancer le travail **sans rien savoir** des sessions du 19–20 septembre
et du 23–25 octobre 2026, où la méthode a été refondue, corrigée, puis élargie à cinq corpus.

> **Mise à jour du 25/10/2026, soir.** Toute version antérieure de ce fichier est périmée.
> Celle du matin annonçait sept formants établis ; il y en a vingt, et quatre de ceux qu'elle
> donnait sont faux pour cause de doublons de bibliothèque. Ne jamais partir d'une copie
> ancienne.

## Ne pas repartir de la transcription

Une transcription complète de ces sessions existe, dans
`Etude-Formants/Conversations/`. **Elle ne doit pas servir d'amorce** : 673 Ko pour la version
lisible, et surtout elle est chronologique, donc elle contient toutes les conclusions
intermédiaires qui ont ensuite été réfutées. Elle sert à l'archéologie — retrouver *pourquoi*
une décision a été prise — jamais à transmettre l'état du travail.

C'est le présent fichier plus les documents de synthèse du dépôt qui portent l'état corrigé.

## Comment utiliser ce fichier

Coller le bloc ci-dessous au début de la nouvelle conversation, après avoir donné accès au
dépôt `Orchestration` et aux dossiers d'échantillons.

---

> Je reprends une étude de mesure formantique des instruments de l'orchestre. Toute la méthode
> et son historique sont dans le dépôt. **Lis d'abord, dans cet ordre :**
>
> 1. `Etude-Formants/Instruments/_CORPUS-CROISES.md` — **le plus important.** Les cinq corpus,
>    la règle de provenance, les vingt-trois formants établis, et ce qui ne se résout pas.
> 2. `Etude-Formants/Instruments/_REPLICATION.md` — le critère à deux routes et le test de
>    permutation qui montre qu'une concordance sur quatre est fortuite.
> 3. `Etude-Formants/Instruments/_VERIFICATION-GRAVES.md` — pourquoi l'ancien pipeline est
>    faux, pourquoi la résolution est d'un demi-ton, et ce que contient `Analyses_formants-mai26`.
> 4. `Etude-Formants/Instruments/00-METHODE.md` — la méthode, les critères de rejet, la
>    typologie, les méthodes écartées.
> 5. `Etude-Formants/Instruments/_CARTE-VOCALIQUE.md` — la couche voyelles, qui couvre les
>    22 instruments là où les formants n'en couvrent que douze.
> 6. `Etude-Formants/Instruments/_PLAN-DE-REPRISE.md` — où en est le projet et dans quel ordre
>    avancer.
> 7. `Scripts/v7-formants-2026/README.md` — la chaîne de mesure et les pièges à ne pas refaire.
> 8. Deux fiches pour le gabarit : `07-basson.md` (le cas le mieux établi) et
>    `01-petite-flute.md` (le cas où la notion de formant n'a pas de sens).
>
> N'utilise aucun script v1 à v6, ni `extract_formants_all_techniques_*.py` : ils reposent sur
> les `specenv` d'Orchidea et sur le Fp par centroïde, tous deux invalides, ce qui est démontré
> dans `00-METHODE.md` et `_VERIFICATION-GRAVES.md`. Tout passe par `Scripts/v7-formants-2026/`.
>
> Règles de travail : **français** pour tout le contenu du projet ; **aucun `git push` sans mon
> accord exprès** — je pousse moi-même depuis ma machine, et un hook réclame périodiquement un
> commit-push, qu'il faut ignorer ; prototype avant génération en série ; confirmation avant
> génération de fichiers ; par lots courts avec confirmation, pour ne pas heurter la limite de
> session.

---

## Les quatre choses à ne jamais refaire

Par ordre de gravité. Chacune a produit de faux résultats au moins une fois.

**1. Compter une réplication entre deux corpus qui sont la même bibliothèque.** `VN` est un
extrait de la Vienna Symphonic Library, et `Yan_Adds-Divers` en contient aussi. Trois paires de
fichiers vérifiées identiques **au sample près** — contrebasson 362 564 échantillons de part et
d'autre, corrélation **1,00000**. Quatre formants ont été annoncés puis annulés le 25/10 pour
cette raison. *Avant de compter une réplication, vérifier la durée en échantillons : c'est une
empreinte suffisante et cela coûte quelques secondes.*

**2. Croire qu'une nomenclature dans la tolérance 0,97–1,03 garantit la mesure.** La banque VO
est accordée à La ≈ 444, soit +15 à +48 cents. `harmoniques()` cherche chaque harmonique dans
une fenêtre de ±2 bins autour de k·f0 **calculé depuis le nom de fichier** : à +30 cents, pour
la flûte, **seul l'harmonique 1 est correctement relevé** et tout le reste est pris dans un
creux. Le seuil utile est 0,2 %, soit 3 cents. `corpus.py --valider` le signale,
`"corriger_accord": true` le corrige.

**3. Redresser linéairement en log-fréquence.** Cela FABRIQUE une bosse médiane sur tout
spectre concave, et avait produit de faux formants de hautbois à 1 318 et 1 863 Hz,
apparemment en accord avec la littérature. Le redressement est **local, à ±3/4 d'octave**
(`residu_local()`).

**4. Comparer les formants retenus d'un corpus aux candidats d'un autre.** C'est ce qui avait
gonflé la liste du 24/10. Toute comparaison se fait **retenu contre retenu**.

## L'essentiel en douze lignes

1. On mesure un formant par **enveloppe composite** (méthode 8) : les amplitudes harmoniques de
   toutes les notes de la tessiture, agrégées sur un axe fréquentiel commun.
2. Les amplitudes sont **redressées localement** (±3/4 d'octave). **Deux limites distinctes, à
   ne pas confondre.** Largeur : aveugle aux résonances plus larges que ~1,5 octave — c'est le
   cas des flûtes. **Plancher : rien n'est mesuré en dessous de 3,2 à 5,7 fois f0_min**, parce
   qu'un harmonique n'entre dans le composite que si sa fenêtre contient trois harmoniques de la
   même note. C'est le plancher, et non la largeur, qui lui a fait manquer la caisse de
   l'alto, voir plus bas.
3. Les spectres moyens sont **recalculés depuis les WAV, attaque exclue** (onset + 200 ms,
   fenêtre de 1 500 ms).
4. Un formant n'est **candidat** que s'il passe deux tests internes : indépendance au rang
   harmonique (découpage au rang médian, bootstrap par note, IC 90 % positif des deux côtés) et
   robustesse au changement de fenêtre de redressement.
5. **Ces deux tests ne suffisent pas**, étant internes à un corpus. Il faut en plus l'une des
   deux routes : **route A**, une paire de corpus **de provenances distinctes** au même point de
   grille ; **route B**, corroboration par deux familles de méthodes indépendantes.
6. **La résolution est d'un demi-ton**, l'axe composite étant une grille tempérée. Tous les
   formants tombent sur des fréquences de notes : 554 Hz = Do♯5, 2 216 = Do♯7. Écrire
   « Do♯7 (≈ 2 215 Hz) », jamais une valeur au Hz près.
7. Les écarts sous 10,77 Hz — un bin de FFT — ne signifient rien. **Les 2 099 valeurs de
   l'ancienne étude tombent toutes exactement sur des centres de bin** : ses « convergences à
   Δ = 11 Hz » étaient des bins adjacents.
8. Trois classes : **A** résonance fixe dominante · **B** spectre suiveur · **C** résonances
   multiples dont aucune ne domine. ρ seul ne distingue pas B de C.
9. Pour les flûtes et le hautbois, **la notion de formant fixe n'a pas de sens**, et il faut le
   dire. Leur description passe par la carte vocalique et la pente spectrale.
10. Une mesure stable n'est pas nécessairement juste : le Fp était stable et faux, le LPC était
    dispersé et visait juste.
11. Chaque instrument a une fiche `.md` à en-tête YAML, qui sert de notes de cours **et** de
    source de données pour le build HTML/docx.
12. Le suffixe `-cr` d'un verdict signale le critère assoupli pour corpus réduit. Le conserver.

## Les cinq corpus

| Banque | Provenance | Accord | Dynamiques | Jugement |
|---|---|---|---|---|
| `OrchideaSOL_and_friends` / **FullSOL2020** | IRCAM | 1,00 | variable | référence, 12 instruments |
| même dépôt / **Yan_Adds-Divers** | **VSL + ConTimbre + perso** | 1,00 | variable | 10 instruments · **ne pas apparier à VN** |
| **ZSamples** | indéterminée | 1,00 | 4 | bonne, échantillonnage espacé |
| **Xsamples** | indéterminée | **1,0000** σ 0,1 % | 3 | **irréprochable** · notes minuscules, H allemand |
| **VO** | indéterminée | 1,003 à 1,028 | 3 | exploitable **après correction d'accord** |
| **VN** | **VSL** | **1,0000** σ 0,13 % | 3 à 4 | irréprochable · **provenance commune avec Yan_Adds** |

Audit de doublons fait le 25/10 : hors VSL, une seule durée commune sur dix paires de corpus, et
entre deux instruments différents. **FullSOL2020, ZSamples, Xsamples et VO sont mutuellement
indépendants.**

## Les 23 formants établis par la route A

Un **couple de corpus de provenances distinctes au même point de grille** (< 0,5 demi-ton).
Liste générée depuis `Resultats-v7/formants-v7.csv` — ne pas la retranscrire à la main.

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

**Sept de ces zones ne sont portées par aucune fiche** : leur couple n'inclut pas SOL. C'est le
cas de l'alto, dont le formant à 3 748 Hz est établi par VO + ZSamples alors que SOL mesure
3 517 Hz, à 1,1 demi-ton de là.

**Convergences entre instruments au même point de grille** : **1 108 Hz** basson et violoncelle ·
**1 244 Hz** clarinette basse et tuba basse · **1 759 Hz** clarinette contrebasse et trombone
basse. Les zones cor–violon vers 2 216–2 348 Hz et violon–violoncelle vers 3 950 Hz sont réelles
mais à un demi-ton, non au même point.

## Ce qui est écarté, et ne doit pas être repris

- **La convergence inter-familles à 4 200 Hz**, présentée dans les fiches comme « la figure
  majeure ». Elle ne reposait que sur SOL ; ni VN ni Xsamples ne la retrouvent.
- **Le cor anglais à 2 216 et 4 183 Hz**, annoncé comme « le résultat le plus solide des anches
  doubles aiguës » : ZSamples n'y trouve rien, et son accord avec VN est une doublure VSL.
- **Le contrebasson à 329 · 698 · 1 174 Hz** et le **tuba contrebasse à 1 396 Hz** — doublons
  VSL. Ces deux instruments n'ont **aucun second corpus indépendant**.
- **L'unisson parfait à 226 Hz** entre contrebasson et les deux tubas.
- Les convergences de l'ancienne étude : flûte–hautbois, violon–basson, cor anglais–clarinette
  Sib, cor–alto, trompette–hautbois. Toutes des bins voisins.
- **L'hypothèse des 276 Hz** comme critère de validité de l'ancien pipeline. Le discriminant
  est la typologie, non la tessiture.
- **Le vibrato comme voie de mesure.** Testé le 25/10 sur `OB-L_pV_pA_sus` : l'idée est juste —
  le balayage traverse le filtre — mais la profondeur disponible est trop faible. Hautbois `pV`
  **±10,4 cents**, alto SOL ±14,5, violon SOL ±11,5, contre ±30 à ±50 chez un instrumentiste en
  concert. Modulation d'amplitude attendue ~0,3 dB. À redemander si un enregistrement est
  commandé ; inutile de réexplorer avec le matériel actuel.

## Les instruments sans formant, et c'est un résultat

| Instrument | Corpus | Conclusion |
|---|---|---|
| **Petite flûte** | 3 | **aucun formant dans les trois** — répond à l'anomalie initiale |
| Flûte · flûte basse · flûte contrebasse | 3 à 4 | aucun accord |
| **Hautbois** | **4** | **4 réponses incompatibles** · ρ 0,72–0,80 · **à reclasser en B** |
| **Alto** | 4 | aucun formant par méthode 8 — mais voir ci-dessous |

Le hautbois est le cas le plus contre-intuitif : c'est l'instrument auquel la littérature
assigne un formant avec le plus d'assurance (Backus 1 300, Meyer et Giesler ~1 100) et quatre
corpus indépendants ne le confirment pas.

## La découverte du 25/10 au soir — méthode 12 > méthode 8 sur l'alto

Sur quatre corpus, la **carte vocalique réplique là où les formants ne répliquent pas**.

- **Hautbois** : 991 Hz à l'aigu dans **trois corpus au Hz près**, hors espace vocalique dans
  les quatre ; médium /a/–/œ/ à 0,4 Bark d'écart maximum. Mais toutes ces cellules portent la
  marque `F1=f0` : la carte décrit **où est sa tessiture**, non une invariance de timbre.
- **Alto** : l'aigu est le meilleur accord de toute la carte — **quatre corpus entre 539 et
  605 Hz**, /ɛ/ ou /œ/, distances de 0,22 à 0,90 Bark. Et **au grave seulement**, sans aucune marque
  `F1=f0`, F1 se tient entre **374 et 395 Hz sur quatre corpus**, pente +0,01 : c'est la
  **résonance principale de caisse** de l'alto, que la littérature situe vers 350–450 Hz.
  Au médium F1 suit la note — corrigé le 27/10, voir `_RESONANCE-DE-CAISSE.md`.

La méthode 8 avait donné 1 318 Hz comme formant le plus grave de l'alto — faux de dix demi-tons
sur ZSamples. La raison est écrite dans `residu_local()` : le redressement à ±3/4 d'octave est
aveugle aux résonances larges, et la zone de caisse d'un alto est large et plate.

**C'est le premier cas où les deux méthodes ne sont pas complémentaires mais hiérarchisées : la
méthode 12 a raison et la méthode 8 se trompe.** À examiner d'urgence pour le violoncelle et la
contrebasse, dont les résonances de caisse sont tout aussi larges.

## Verser un nouveau corpus

`Scripts/v7-formants-2026/corpus.py` absorbe n'importe quelle arborescence et n'importe quelle
nomenclature sans qu'on touche à la chaîne de mesure : il écrit dans le fichier de spectres un
**nom normalisé au format SOL**, donc tout l'aval fonctionne. Cinq nomenclatures déjà gérées,
dont les notes en minuscules, le **H allemand** (cor et trombone de Xsamples écrivent `Corf_h2`
pour le si naturel), la dynamique **avant** la note, la technique sur plusieurs niveaux de
dossier, et la correction d'accord.

```bash
cd Scripts/v7-formants-2026
python3 corpus.py --decrire "/chemin/Banque" > /tmp/d.json   # explorer, obtenir un brouillon
# compléter : renommer les configurations avec les noms d'instruments du projet,
# fusionner les dossiers d'un même instrument, corriger la liste des techniques
python3 corpus.py --valider /tmp/d.json                      # OBLIGATOIRE
python3 corpus.py --mesurer /tmp/d.json
```

Vérifié : sur ZSamples, retrouve à l'identique le violon [2 216 · 2 792 · 3 948 · 5 583] et le
basson [494 · 1 108 · 1 974].

**À vérifier sans exception** : la **provenance** (point 1 des choses à ne pas refaire), la
**nomenclature** (point 2), l'existence d'une entrée dans `Data/registres.md` sans laquelle
`carte_vocalique.py` ne peut pas construire sa grille. **Ignorer** les instruments absents de
l'étude, les fichiers « NoPitch », et les sourdines qui n'y figurent pas déjà.

## À faire

> Les travaux restants sont **tous** dans `_PLAN-DE-REPRISE.md`, qui est la seule liste à tenir. Ne pas en ouvrir une ici : c'est ainsi que huit listes se sont contredites entre le 23 et le 26 octobre.

## Les sources de données

| Chemin | Contenu | Usage |
|---|---|---|
| `kDrive/Samples/Orchidea-dbs/OrchideaSOL_and_friends` | SOL2020 + Yan_Adds | corpus 1 et 2 |
| `Documents/Instrumental_formants/Zsamples` | 3 540 fichiers, 42 configurations | corpus 3 |
| `Documents/Instrumental_formants/Xsamples ` *(espace final)* | 1 404 `.aif` | corpus 4 |
| `Documents/Instrumental_formants/VO` | 1 938 `.AIF` | corpus 5 |
| `Documents/Instrumental_formants/VN` | 1 034 `.wav` + extraits `/Volumes/Samps/VSL/` | corpus 6, **VSL** |
| `Etude-Formants/Analyses_formants-mai26` | 45 CSV **par note** : LPC, cepstre 40/80/160 | route B ; proxy exact de l'ancien pipeline |
| `Etude-Formants/Resultats/formants_*_v3.csv` | l'ancien pipeline | archive ; toutes les valeurs sont des centres de bin |
| `Etude-Formants/Formants-Francais/` | voyelles Calliope (Tubach 1989) | carte vocalique. **F1 code l'aperture et presque rien d'autre** — 308 Hz pour /i/, 315 pour /u/, 684 pour /a/ chez l'homme. Il faut le couple (F1, F2) et le seuil de 3,5 Bark |
| `Etude-Formants/Conversations/` | transcriptions, export par `Scripts/exporter_conversation.py` | archéologie **seulement** — ne pas utiliser comme amorce |
