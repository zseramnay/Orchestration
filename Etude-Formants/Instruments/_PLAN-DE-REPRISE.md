---
id: plan-de-reprise
titre: Plan de reprise — liste unique des travaux
date: 2026-10-26
etat: liste consolidée ; remplace les listes dispersées dans les sept autres documents
---

# Plan de reprise — la seule liste à tenir

> **Les sept autres documents ne portent plus de liste de travaux.** Tout est ici. Si vous
> ajoutez un « à faire » ailleurs, recopiez-le dans ce fichier, sinon il sera perdu — c'est ce
> qui est arrivé entre le 23 et le 26 octobre, où huit documents portaient chacun la sienne.

## Où en est le projet

**La phase de mesure est terminée.** Cinq corpus, scripts documentés et reproductibles, données
rangées (`README.md` à la racine de `Etude-Formants/`). Ce qui reste est de l'écriture, plus un
petit nombre de mesures qui peuvent attendre.

| Acquis | État |
|---|---|
| 22 fiches, en-tête YAML, statut par formant | **fait** |
| **23 formants établis par route A, sur 13 instruments** | **fait** |
| Règle de provenance (Yan_Adds et VN sont du VSL) | **fait** |
| Carte vocalique registre × dynamique, 22 instruments | **fait** |
| Largeurs de bande, deux méthodes confrontées | **fait** |
| ρ par corpus, 21 instruments — étendue médiane 0,14 | **fait** |
| Pilote générique `corpus.py`, cinq nomenclatures absorbées | **fait** |
| Vérification de l'ancien pipeline | **fait** |

## Décisions prises, à ne pas rouvrir

1. **Le chapitre de synthèse s'organise par convergences entre instruments**, avec en plus une
   **section par propriétés mesurées** (typologie, dynamique, géométrie du tuyau, voyelles).
2. **La route A exige un couple de corpus de provenances distinctes** au même point de grille.
   La concordance à un demi-ton ne suffit pas : une sur quatre est fortuite.
3. **La résolution est d'un demi-ton.** Écrire « Ré7 (≈ 2 350 Hz) », jamais une valeur au Hz.
4. **`etendue_residu` n'est pas une largeur de bande** et ne doit jamais être présentée comme
   telle.
5. **Le hautbois est en classe B.** Quatre corpus, quatre réponses incompatibles.
6. **Les sourdines et ensembles sont suspendus** — réplication médiane à 4,0 demi-tons.

---

# BLOC B — la production

C'est la priorité. Tout le reste peut attendre.

| # | Tâche | Note |
|---|---|---|
| B1 | ✅ **fait le 26/10** — `exporter_csv.py` produit `Resultats-v7/formants-v7.csv` (65 lignes, une par formant) et `instruments-v7.csv`. Schéma neuf : l'ancien, en F1–F6 fixes, ne pouvait pas porter un nombre variable de formants ni leur route |
| B2 | ✅ **fait le 26/10** — `Scripts/v7-site/construire_site.py`, sortie dans `Etude-Formants/Site-v2/`. 30 pages, zéro lien cassé. La chaîne v6 n'a pas été adaptée : elle est soudée à `get_f()` et au modèle F1–F6. Son CSS est repris |
| B3 | ✅ **fait le 26/10** — `Scripts/v7-site/figures.py`, 45 figures dans `Site-v2/figures/` : 22 composites, 18 cartes vocaliques, 5 figures de synthèse. Branchées dans `construire_site.py` ; 46 images référencées, zéro cassée. Les 47 `specenv_*` ne sont reprises nulle part et portent désormais un avertissement dans leur dossier |
| B4 | ✅ **fait le 26/10** — réglé par B3 : les 18 cartes vocaliques portent les dix voyelles de Calliope (voix d'homme) en repère. Plus aucune figure sur voyelles allemandes. Reste ouvert : la référence féminine, en C12 |
| B5 | ✅ **fait le 26/10** — `_CHAPITRE-SYNTHESE.md`, rendu dans `Site-v2/syntheses/convergences.html`. Quatre parties : les convergences, les propriétés mesurées, ce qui n'est pas mesurable, ce que cela change pour l'enseignement. Les tableaux de convergences restent générés depuis le CSV et sont insérés au marqueur |

### Structure v2 du site

```
index.html                     sommaire général, les deux études
  formants/
    index.html                 sommaire par famille, avec le statut de chaque instrument
    methode.html               méthode, limites, quand le formant n'a pas de sens
    bois/01-petite-flute.html … 12-clarinette-contrebasse.html
    cuivres/13-cor.html … 18-tuba-contrebasse.html
    cordes/19-violon.html … 22-contrebasse.html
    syntheses/
      convergences.html        LE CHAPITRE — les convergences établies
      formants-etablis.html    les 20, avec route, corpus et écart
      typologie.html           les trois classes, et ρ par corpus
      carte-vocalique.html     registre × dynamique, les quatre comportements
      dynamique.html           l'effet de la force de jeu (cuivres)
      tuyau-cylindrique.html   balance impair−pair, la famille clarinette
      largeurs.html            étroit sélectionne, large colore
      cordes-et-bridge-hill.html   classe C, méthodes 8 et 12 ensemble
      heritage.html            ce que disait l'ancienne étude, et pourquoi
      non-mesurable.html       sourdines, ensembles, hautbois, flûtes
```

Une page par instrument se génère du front matter de sa fiche. Le sommaire affiche le statut :
nombre de formants par route, classe, cellules vocaliques assignables. Le lecteur voit d'un
coup d'œil où l'étude dit quelque chose et où elle se tait.

### Les convergences, matière du chapitre

Chaque membre validé **indépendamment**, par route A. À ne pas confondre avec celles de
l'ancienne étude, qui étaient des bins de FFT voisins.

**Au même point de grille** — les seules à publier comme unissons formantiques :

| Zone | Note | Instruments |
|---|---|---|
| **1 108 Hz** | Do♯6 | basson · violoncelle |
| **1 244 Hz** | Ré♯6 | clarinette basse · tuba basse |
| **1 759 Hz** | La6 | clarinette contrebasse · trombone basse |

**Régions partagées à un demi-ton près** — réelles, mais à présenter comme des régions :

| Région | Étendue | Instruments |
|---|---|---|
| 494–498 Hz | 0,1 dt | basson · tuba basse |
| 830–879 Hz | 1,0 dt | les deux trombones |
| **2 216–2 348 Hz** | 1,0 dt | **clarinette basse · cor · violon** — trois instruments, trois familles |
| 2 631–2 792 Hz | 1,0 dt | contrebasse · trompette · violoncelle |
| **3 748–3 959 Hz** | 0,9 dt | **alto · violon · violoncelle** — toute la famille des cordes aiguës |

> Ces deux tableaux sont **générés** par `construire_site.py` depuis
> `Resultats-v7/formants-v7.csv` et se retrouvent dans `Site-v2/syntheses/convergences.html`.
> Ne pas les retranscrire à la main : c'est la retranscription qui a produit l'erreur du 25/10.

---

# BLOC C — mesures ouvertes, par priorité

Aucune ne bloque le bloc B.

| # | Tâche | Pourquoi |
|---|---|---|
| **C1** | ✅ **fait le 27/10** — `resonance_caisse.py`, 30 lignes de mesure sur quatre cordes et quatre corpus, document `_RESONANCE-DE-CAISSE.md`, figure `synthese-resonance-caisse.png`. **Acquis : le violoncelle a une résonance vers 215–220 Hz au médium, quatre corpus, 1,1 dt d'écart, pente −0,04** — invisible à la méthode 8, comme chez l'alto. **Corrigé : l'affirmation sur l'alto valait pour le grave seulement** ; au médium F1 suit la note (pentes +0,79 à +1,11). Le grave sort renforcé : 374–395 Hz sur quatre corpus et non 347–498 sur trois. Contrebasse et violon : non établis, causes identifiées |
| **C2** | ✅ **fait le 27/10** — `fenetre_redressement.py`, document `_FENETRE-DE-REDRESSEMENT.md`. **Le diagnostic était faux** : ce n'est pas la largeur mais le **plancher du domaine** (3,2 à 5,7 × f0_min) qui fait manquer la caisse de l'alto et du violoncelle. ±1,5 et ±2,0 sont **écartés par la mesure** : à ±2,0 un formant route A sur trois disparaît. **±1,00 est recommandé** — 96 % de formants route A retrouvés contre 95 % à la fenêtre actuelle, et 5 demi-tons de domaine gagnés en médiane. **Migration en ±1,00 tentée l'après-midi même, puis ANNULÉE** : les trois unissons formantiques disparaissaient et la route A tombait de 23 à 21. Le test de dégâts employait la tolérance d'un tiers d'octave du test de robustesse, là où la route A exige un demi-ton. `HALFWIN` reste à 0,75 et porte désormais un avertissement chiffré |
| **C3** | ✅ **fait le 27/10** — `suraigu.py` et `impact_regle_f0.py`, document `_SURAIGU.md`. Corrigé par une **borne physique** : aucun partiel n'existe sous f0, donc aucun maximum n'y est un formant. IQR du suraigu du violon 7,4 → 2,3 (SOL) et 9,0 → 3,4 (ZSamples), les trois dynamiques s'accordent. Vérifié avant adoption : 4 à 5 % des cellules bougent, 0 à 2 % changent de voyelle, **rien au grave ni au médium** — les acquis de C1 sont intacts. Effets de bord : cor anglais réintégré, **17 cartes par corpus gagnées** (avance C13), contrebasson et trombone basse sortent des « hors espace par le bas » |
| C4 | ✅ **fait le 27/10** — `appariement_lpc.py`, document `_APPARIEMENT-LPC.md`. 39 formants testés sur 16 instruments, appariement exact par nom de fichier. **La corroboration tient pour quatre familles sur cinq** — cordes 79ᵉ centile, clarinettes 78ᵉ, cuivres 71ᵉ, flûtes 83ᵉ — et **échoue entièrement pour les anches doubles** (15ᵉ centile, zéro au-dessus de la médiane), ce que le modèle tout-pôle du LPC prédit chez une anche couplée à la perce. Le premier test, sans témoin apparié, donnait 13 résultats écrasants et faux. Manque : le cor, absent de `LPC-v7/` en version ordinaire |
| C5 | ✅ **fait le 27/10** — `largeur_par_registre.py`, document `_LARGEUR-PAR-REGISTRE.md`. **Hypothèse confirmée, et c'est bien un effet de cuivres** : +31 points d'ensemble, jusqu'à **+63 points à registre constant** au trombone ténor dans le grave. Cordes +3, anches doubles +1 — rien. Mais deux cuivres sur cinq s'y opposent, et **le partage suit la perce** : cylindrique (trombones, trompette) oui, conique (cor, tuba) non, sauf dans l'aigu. Chiffres directement utilisables : le formant du trombone basse n'existe pratiquement pas au piano (76 % de pics non localisables contre 22 % au fort), le médium du tuba basse ne l'est à aucune nuance (78–79 %) |
| C6 | **Répliquer la détection de localisation** sur les quatre autres corpus | le Q ≈ 1 de la trompette est-il l'instrument ou ce corpus ? |
| C7 | **Relancer les cuivres** avec les registres pédale, ajoutés à `registres.md` le 24/10 | les pédales formaient une catégorie « hors registre » |
| C8 | **Signaler les coïncidences F2 / formant établi** dans `carte_vocalique.py` | indice de confiance pour les cordes |
| C9 | ✅ **fait le 27/10** — `melanges.py`, document `_MELANGES.md`, figure `synthese-melanges.png`. **215 paires, 12 926 cas.** Le mélange n'est pas simulé : le motif d'excitation étant linéaire en puissance, motif(A+B) = motif(A) + motif(B) exactement. **Résultat : il n'y a pas deux familles d'alliages mais un continuum, gouverné par l'écart de couleur entre les deux instruments** — fusion de 8 % à 55 % quand l'écart passe de 0,5 à 3,5 Bark, homogénéité de 66 % à 14 %, renforcement culminant au milieu. Les anches doubles fusionnent avec tout (30–34 %), les flûtes avec rien (14–18 %). **Suite immédiate la plus utile : le doublage à l'octave**, non couvert |
| C10 | **Coupure de Benade pour les flûtes** | le descripteur qui remplacerait le formant absent |
| C11 | **Reprendre les « hors espace par le bas »** autrement que par le comptage de maxima | le critère dépend du corpus, pas de l'instrument |
| C12 | **Tester la référence féminine** de Calliope | un instrument classé /ɔ/ en voix d'homme peut se classer autrement |
| C13 | **Étendre la vérification vocalique** aux instruments non comparés — **largement avancé le 27/10 par effet de bord de C3** : Xsamples passe de 1 à 5 instruments, VO de 1 à 8, VN de 2 à 8 | restent sans second corpus : clarinette basse, flûte contrebasse, clarinette contrebasse |

---

# BLOC D — corrections de documents restant à faire

| # | Tâche |
|---|---|
| **D1** | ✅ **fait le 27/10** — recompté retenu contre retenu : **5 concordances strictes et non 7**. Deux tombent : la trompette à 2 792 Hz, prévue, et **la contrebasse à 1 396 Hz, qui ne l'était pas**. La trompette tient par ailleurs (SOL + VN) ; la contrebasse à 1 396 passe en « indéterminé ». Le tableau du 24/10 est marqué comme archive d'un état à deux corpus ; la liste qui fait foi est celle des 23. La route B est en outre chiffrée par C4 et refusée aux anches doubles. ~~Ancien libellé :~~ Le tableau du 24/10 compare les retenus de SOL aux *candidats* de ZSamples — plus permissif. Conséquence vérifiée : la trompette à 2 792 Hz n'y tient pas sur une comparaison stricte, alors qu'elle tient par ailleurs via SOL + VN. Recompter et publier la liste des 23 |
| D2 | ✅ **fait le 26/10** — l'alto a bien un formant établi, à 3 748 Hz (VO + ZSamples) ; c'est SOL qui ne le trouve pas. Affirmation corrigée dans `_CORPUS-CROISES.md` et dans la fiche |
| D3 | ✅ **fait le 27/10** — **la typologie tient** : A mesure 0,48–0,50, B 0,77–0,88, C 0,84, et les groupes ne se chevauchent pas. Deux manques comblés : **le cor et la trompette n'étaient dans aucune ligne** (ρ 0,57, d'où la classe A/B qu'ils portaient sans justification), et l'exception du hautbois en B n'était pas écrite |

---

# SUSPENDU, par décision ou faute de matière

| Sujet | Pourquoi | Ce qui débloquerait |
|---|---|---|
| **Sourdines et ensembles** — 28 configurations | réplication médiane à 4,0 demi-tons sur 18 paires. Une sourdine est un objet physique particulier | un troisième corpus, ou publication explicite comme exploratoire |
| **Tuba contrebasse et contrebasson** | aucun second corpus indépendant : Yan_Adds et VN sont tous deux du VSL | une banque **ni VSL ni ZSamples** |
| **Le vibrato comme voie de mesure** | l'idée est juste mais la profondeur disponible est trop faible — hautbois ±10 cents, alto SOL ±14,5, contre ±30 à ±50 en concert | des enregistrements au vibrato franc, à demander si une séance se présente |
| **Flûte basse dans VN** | absente de la bibliothèque | — |

---

## Ce qui a été fait, pour mémoire

Coché le 26/10, pour qu'on ne le redemande pas.

- ✅ Propager les formants dans les 22 fiches, avec route, corpus et écart
- ✅ Reclasser le hautbois en B, avec les ρ des quatre corpus à l'appui
- ✅ Noter la provenance dans chaque fiche, et la faire porter par le script
- ✅ Renommer `bande_db3` en `etendue_residu` partout
- ✅ Inscrire « aucun pic localisé » pour la trompette et le trombone basse
- ✅ Carte vocalique sur les nouvelles banques, pour l'alto et le hautbois
- ✅ Corriger le filtre de technique de `corpus.py` et écarter les données contaminées
- ✅ Ranger le dépôt : rédaction dans `Instruments/`, données dans `Mesures/`

## Journal

- **2026-10-24** — Première version.
- **2026-10-26 (soir)** — B5 : chapitre de synthèse écrit. Les convergences y sont présentées
  comme des **prédictions** et non comme des doublures établies : nos mesures portent sur des
  notes isolées, pas sur des mélanges. C9 reste donc la tâche qui referme la boucle.
- **2026-10-26** — Correction du comptage des formants : 23 sur 13 instruments et non 20 sur 12.
- **2026-10-27 (nuit)** — **Consolidation.** Le chapitre de synthèse datait du 26 et ne
  connaissait aucun des six travaux du 27 : il affirmait deux fois que les mélanges ne sont pas
  mesurés. Réécrit — 317 → 539 lignes, cinq parties, les mélanges en partie III. Les 22 fiches
  reçoivent une section générée depuis les CSV de C4, C5 et C9 (`propager_fiches.py`, idempotent).
  D1 et D3 faits. **Le bloc D est terminé.**
- **2026-10-27 (soir)** — C9 : **la question d'origine est traitée.** Deux corrections de fond
  en route : la catégorie « homogène », absente du premier classement, qui rangeait le fondu par
  identité avec les mesures manquantes — flûte + clarinette y tombait, alors que c'est l'exemple
  d'école ; et le garde-fou de la « chimère », qui vérifie qu'un cas de fusion est bien une
  fusion et non un F1 emprunté à l'un et un F2 à l'autre. Il ne domine aucune paire.
- **2026-10-27 (soir)** — C5 : largeur par registre et dynamique, et **le cor comblé dans les
  deux jeux de largeurs**. Cause du manque : l'analyse de mai 2026 avait omis le cor ordinaire,
  et tous les scripts aval tiraient leur liste d'instruments de ce dossier — l'absence ne se
  voyait nulle part. `largeurs_lpc.py` et `largeurs_enveloppe.py` ont désormais une option
  `--depuis-dossier`. Cela comble aussi le manque signalé par C4 : le cor place son formant
  route A au 31ᵉ centile, et **l'unisson de 1 244 Hz est le seul des trois dont les deux membres
  sont corroborés**. Troisième agrégat de la journée à masquer une structure.
- **2026-10-27 (soir)** — C4 : l'appariement note par note. Deuxième fois dans la journée qu'un
  témoin mal choisi produit un faux positif spectaculaire — après la tolérance d'un tiers
  d'octave de C2, c'est ici un témoin au triton qui s'associait deux fois plus fort que le
  formant lui-même. **Règle à retenir : le témoin doit partager le facteur de confusion de ce
  qu'il contrôle.** Corrigé en comparant fc au rang qu'il occupe dans son propre spectre.
- **2026-10-27 (soir)** — C3 : le plancher de la fondamentale. Le seul défaut de méthode
  reproductible de l'étude est corrigé, par une borne physique et non par un réglage. Trois
  effets de bord favorables : le cor anglais, que le script faisait tomber de la carte depuis
  qu'un commentaire périmé l'en excluait, est réintégré ; dix-sept cartes par corpus sont
  gagnées en relançant simplement le script complet ; et le contrebasson comme le trombone basse
  sortent des « hors espace par le bas », ce qui confirme que ce classement dépend du corpus.
- **2026-10-27 (après-midi)** — Migration en ±1,00 **tentée puis annulée dans l'heure**. Leçon
  de méthode à retenir : un test de non-régression doit employer la tolérance du résultat qu'il
  protège, pas celle d'un autre test du pipeline. `tests_de_rejet.py` archivé dans
  `Archives/scripts-morts/`.
- **2026-10-27 (suite)** — C2 : la fenêtre de redressement. Deux limites de `residu_local()`
  avaient été confondues, et c'est la mauvaise qui était invoquée depuis le 24/10. Corrigé dans
  la docstring, `00-METHODE.md`, `_CARTE-VOCALIQUE.md`, `_AMORCE-SESSION.md` et
  `_RESONANCE-DE-CAISSE.md`. `HALFWIN` est laissé à 0,75 : la migration en 1,00 est recommandée
  mais non engagée, elle modifierait des chiffres publiés. Note au passage :
  `tests_de_rejet.py` pointe vers un scratchpad de mai et ne tourne plus — ses deux tests
  vivent dans `formants_lib.verdict()`, le fichier devrait être archivé.
- **2026-10-27** — C1 : résonances de caisse. Le test publié sur l'alto ne reposait que sur
  la marque `F1=f0`, qui ne teste que l'harmonique 1 ; au violoncelle, la zone de caisse
  attendue tombe sur les harmoniques 2 à 4 du grave. Nouveau test, note par note : la pente de
  log2(F1) contre log2(f0) — plate pour une résonance, à +1 pour un harmonique. Deux garde-fous
  posés d'avance : le nombre de maxima fusionnés (vaut 1 partout, donc F1 n'est pas un centroïde
  déguisé) et l'écart au centroïde. Résultat : un acquis (violoncelle), une affirmation publiée
  corrigée de moitié (alto, médium), deux non-résultats dont les causes sont nommées. La
  régularité qui se dégage : la méthode 12 voit la caisse **quand le registre grave est sous le
  mode principal**, et seulement là — ce qui prédit où elle marchera.
- **2026-10-26 (soir)** — B3 et B4 : figures régénérées. Trois constats au passage. Quatre
  instruments n'ont **pas** de carte vocalique — contrebasson, trombone basse, tuba basse,
  tuba contrebasse — parce qu'un seul maximum survit à la fusion de Chistovich dans toutes
  leurs cellules exploitables : ce sont les « hors espace par le bas », et la page l'écrit
  maintenant au lieu de laisser un trou. Les registres ne sont pas quatre mais douze
  (chalumeau, gorge, clarine, sons pédale…). La figure des convergences reprend le
  regroupement de `page_convergences()` mot pour mot, pour qu'image et tableau ne puissent
  pas diverger.
  Le regroupement séquentiel en masquait trois. Les listes sont désormais générées depuis
  `Resultats-v7/formants-v7.csv`.
- **2026-10-26** — Consolidation : les listes des sept autres documents sont fusionnées ici.
  Le bloc A (mise à jour des fiches) est terminé. La forme du chapitre est tranchée.
