---
id: clarinette-contrebasse
nom: Clarinette contrebasse en Sib
nom_en: Contrabass clarinet
famille: bois
sous_famille: clarinettes
ordre: 12
source_echantillons: Yan_Adds-Divers/Winds/Contrabass-Clarinet-Bb/ordinario
provenance: VSL
technique: ordinario
n_echantillons: 45
tessiture_mesuree: A#0–E3
f0_min_hz: 29.1
f0_max_hz: 329.6

classe: B/A
pertinence_formant: partielle
rho_hauteur: 0.51

formants:
  - { hz: 1759, statut: "route A", residu_db: 2.71, etendue_residu: [879, 5583], verdict: retenu-cr, notes: 11,
      ic_bas: [3.0, 0.6], commentaire: "confirme par critere assoupli 23/10 (45 echantillons)" }
  - { hz: 1108, residu_db: 2.95, etendue_residu: [879, 5583], verdict: fragile, notes: 10 }
  - { hz: 4432, residu_db: 2.28, etendue_residu: [783, 5915], verdict: fragile, notes: 12 }

candidats_rejetes:
  - { hz: 2635, residu_db: 1.49, verdict: rejete, raison: "deux groupes de rang seulement, faiblement positifs, échec de robustesse" }

pente_db_octave: -11.4
pente_sd: 4.8
harmoniques_par_note: 90
harmoniques_min_max: [33, 90]
balance_impair_pair_db: 5.1
enveloppe_vraie_hz: 172
enveloppe_vraie_sd: 328

corpus_suffisant: false
version_mesure: 2026-10-23-wav-v1
---

# Clarinette contrebasse en Sib

> ## Statut des formants — révision du 26/10/2026
>
> Confrontation des **cinq corpus**. Provenance de ce corpus : **VSL / ConTimbre / perso (Yan_Adds)**.
>
> | Formant | Statut |
> |---|---|
> | **1759 Hz** | **route A** — SOL + VO au même point de grille, 1761 Hz · zone vue par SOL, VO |
>
> `etendue_residu` dans l'en-tête **n'est pas une largeur de bande** : c'est l'étendue du résidu de la méthode 8, souvent une butée de plage. Les largeurs mesurées sont dans `_LARGEURS-DE-BANDE.md`.

## Méthode retenue, et pourquoi

> **Méthode 8 — un formant retenu faiblement, deux fragiles. Corpus insuffisant.**
> Voir `00-METHODE.md`.

Cet instrument occupe une position intermédiaire intéressante : **ρ = 0,51**, la valeur la
plus basse des clarinettes et la seule qui s'approche du domaine des anches doubles. Sa
tessiture très grave (29 à 330 Hz) la place en grande partie sous la limite de validité des
méthodes à ordre fixe, et elle offre **90 harmoniques par note** — le plafond de mesure,
comme le contrebasson.

Toutes les conditions physiques sont donc favorables. C'est le corpus qui manque :
**45 échantillons** seulement, soit trop peu de notes distinctes pour peupler les groupes de
rang harmonique. Les résidus mesurés sont faibles (+2,3 à +3,0 dB) et deux des trois
candidats restent fragiles.

Conclusion à retenir : c'est, avec la clarinette en Mib, **l'instrument dont un corpus
élargi changerait le plus probablement les résultats**, et ici dans un sens favorable, les
conditions de mesure étant par ailleurs excellentes.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes distinctes | Verdict |
|---|---|---|---|---|
| **1 759 Hz** | +2,7 dB | 879–5 583 Hz | 11 | **retenu** (faible) |
| 1 108 Hz | +3,0 dB | 879–5 583 Hz | 10 | fragile |
| 4 432 Hz | +2,3 dB | 783–5 915 Hz | 12 | fragile |

## Tests de rejet

| Zone | k = 4–7 | k = 8–15 | k ≥ 16 | Robustesse |
|---|---|---|---|---|
| 1 759 Hz | — (2 notes) | +5,0 | +1,5 | oui |
| 1 108 Hz | +0,1 | **−0,5** | +1,3 | oui |
| 4 432 Hz | — | +8,0 (1 note) | +1,6 | oui |
| 2 635 Hz | — | +1,3 | +2,1 | non |

Les trois candidats passent le test de robustesse de fenêtre mais butent sur la couverture.
Le cas de 1 108 Hz est ambigu : son résidu est négatif aux rangs 8–15, ce qui plaide contre,
mais les effectifs sont trop faibles pour conclure. Les bandes à −3 dB très larges
(879–5 583 Hz pour deux d'entre eux) indiquent que la méthode ne parvient pas à circonscrire
les pics, faute de données.

## Descripteurs globaux

| Descripteur | Valeur | Lecture |
|---|---|---|
| **Pente spectrale médiane** | **−11,4 dB/octave** (σ 4,8) | la plus plate de tous les bois |
| **Harmoniques par note** | **90 [33–90]** | plafond de mesure |
| **Balance impair − pair** | **+5,1 dB** | **la plus marquée de tout l'orchestre mesuré** |
| Enveloppe vraie (méthode 6) | 172 Hz, σ 328 | ordres adaptatifs de 46 à 529 |

Deux records, qui décrivent bien l'instrument.

La **pente de −11,4 dB/octave** est la plus plate des douze bois. Avec 90 harmoniques
mesurables par note, cela signifie un spectre dense et étalé : c'est la description
quantitative du son « massif et enveloppant ».

La **balance impair − pair de +5,1 dB** est la plus forte valeur de tout le corpus analysé.
La signature du tuyau cylindrique fermé croît régulièrement vers le grave de la famille :
+0,7 dB pour la Sib, +4,2 pour la Mib, +4,5 pour la basse, +5,1 pour la contrebasse. Cette
progression est cohérente : plus l'instrument est grave, plus la part de la tessiture jouée
dans le registre fondamental (chalumeau) est importante, et plus la suppression des
harmoniques pairs est nette.

## Comparaison avec la littérature

| Source | Valeur annoncée | Notre mesure | Lecture |
|---|---|---|---|
| Yan_Adds (ancien) | F1 = 323 ± 95 Hz, /u/ | — | zone de la fondamentale |
| Ancien Fp | 1 090 Hz | 1 108 Hz (fragile) | **convergent** |
| Cepstre mai 2026 (ordre 80) | 1 658 Hz, 31 % sur un harmonique | 1 759 Hz | proche |
| Backus, Giesler, Meyer | aucune donnée | — | instrument non étudié |

Deux convergences à signaler malgré la fragilité des mesures : l'ancien Fp (1 090 Hz) tombe
sur le candidat fragile à 1 108 Hz, et le cepstre de mai 2026 (1 658 Hz) sur le formant
retenu à 1 759 Hz. Trois méthodes pointent donc ces deux zones, ce qui renforce l'hypothèse
qu'elles sont réelles et que seul le corpus manque pour les établir.

L'ancienne étude notait « F1 = 323 Hz identique à la clarinette basse, mais F2 = 947 Hz
montre une résonance plus développée dans le medium ». Les deux F1 identiques étaient deux
zones de fondamentale, mais l'intuition sur le médium n'est pas démentie par nos mesures.

## Registres

Non défini dans `Data/registres.md`. À établir, éventuellement par transposition du
découpage de la clarinette basse (grave A#1–D3 · médium D#3–G#3 · aigu A3–A#4 · suraigu à
partir de B4).

<!-- DEBUT-TRAVAUX-27-10 -->
## Ce que les travaux du 27/10/2026 disent de cet instrument

### Corroboration par le LPC, note par note

Rang du formant dans son propre spectre : 50 % signifie que les deux méthodes ne s'accordent pas mieux là qu'ailleurs. Voir `../_APPARIEMENT-LPC.md`.

| Formant | Statut | Rang |
|---|---|---|
| 1759 Hz | route A | 81,0 % |

### Effet de la nuance sur la localisation

Part de pics **non localisables** — l'enveloppe n'y redescend pas de 3 dB des deux côtés, donc le formant n'a pas de largeur mesurable. Voir `../_LARGEUR-PAR-REGISTRE.md`.

| | doux | moyen | fort |
|---|---|---|---|
| **tous registres** | 2 % | 12 % | 0 % |
| chalumeau | 0 % | 10 % | 0 % |

**Effet doux → fort : +2 points.** Effet négligeable.

### Dans la table des mélanges

18 paires calculées. **Fusion moyenne : 28 %.** Voir `../_MELANGES.md`.

| Couleur nouvelle avec | | Se fond avec | |
|---|---|---|---|
| Trompette en Ut | 67 % | Tuba contrebasse | 71 % |
| Basson | 58 % | Tuba basse | 50 % |
| Cor anglais | 53 % | Flûte contrebasse | 46 % |

<!-- FIN-TRAVAUX-27-10 -->

## Observations pour l'orchestration

À compléter. Éléments acquis :

- Avec la dominance impaire la plus forte de l'orchestre (+5,1 dB), c'est le meilleur
  candidat pour une doublure de **comblement** par un instrument à spectre complet. Un
  basson ou un contrebasson doublant à l'octave supérieure viendrait remplir les partiels
  pairs manquants. À tester par la méthode 12.
- Les doublures annoncées antérieurement avec le contrebasson (Δ = 97 Hz) et le tuba basse
  (Δ = 97 Hz) reposaient sur des zones de fondamentale, et sont à reprendre.
- Instrument rare : son intérêt principal en orchestration contemporaine est textural, ce
  que la pente très plate et le spectre dense confirment.

## Graphiques

Aucun à ce stade. Figures prévues dans `../Versions-html-and-docx/media/` :

| Fichier prévu | Contenu |
|---|---|
| `clarinette_cb_composite.png` | enveloppe composite avec couverture, montrant l'insuffisance du corpus |
| `clarinettes_balance.png` | **figure de famille** : progression +0,7 → +4,2 → +4,5 → +5,1 dB vers le grave |
| `clarinettes_pentes.png` | figure de famille : les quatre pentes |

La figure de famille sur la balance impair/pair est l'une des plus démonstratives de toute
l'étude : elle montre une signature physique du tuyau cylindrique fermé qui se renforce
régulièrement vers le grave, mesurée directement et sans aucun postulat.

## Révision du 23/10/2026 — critère assoupli

Le formant à **1 759 Hz est confirmé** : +4,2 dB [IC +3,0] en dessous du rang 36 et +1,5 dB
[IC +0,6] au-dessus, sur 36 et 12 points. Avec 45 échantillons seulement, c'est un résultat
solide.

Cet instrument a d'ailleurs servi à corriger un défaut de l'assouplissement lui-même. Un
premier découpage fusionné à seuil fixe (k = 2–5 / k ≥ 6) le classait **« non mesurable »**,
parce que tous ses harmoniques de bande sont à rang élevé — conséquence de sa tessiture très
grave — et que le groupe bas se retrouvait vide. Le découpage au rang médian a été adopté
pour cette raison.

Les zones à 1 108 et 4 432 Hz restent **fragiles**. Pour 4 432 Hz, le motif est net :
+5,7 dB [IC +4,1] en dessous du rang 51 mais **−1,6 dB [IC −3,6]** au-dessus.
La zone à 2 635 Hz reste **rejetée**.

## Journal

- **2026-10-23** — Un formant retenu faiblement (1 759 Hz), deux fragiles. Corpus de 45
  échantillons insuffisant, alors que les conditions physiques sont excellentes
  (90 harmoniques par note, tessiture entièrement grave).
- **2026-10-23** — ρ = 0,51, position intermédiaire entre les clarinettes et les anches
  doubles : à vérifier sur un corpus élargi.
- **2026-10-23** — Balance impair/pair de +5,1 dB, record du corpus. Progression régulière
  de la famille vers le grave établie.
- **2026-10-23 (soir)** — Critère assoupli appliqué (voir `00-METHODE.md`).
