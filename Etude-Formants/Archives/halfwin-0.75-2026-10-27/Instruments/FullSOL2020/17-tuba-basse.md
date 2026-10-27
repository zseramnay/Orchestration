---
id: tuba-basse
nom: Tuba basse
nom_en: Bass tuba
famille: cuivres
sous_famille: cuivres principaux
ordre: 17
source_echantillons: FullSOL2020/Brass/Bass_Tuba/ordinario
provenance: IRCAM
technique: ordinario
n_echantillons: 108
tessiture_mesuree: F#1–F4
f0_min_hz: 46.2
f0_max_hz: 349.2

classe: A
pertinence_formant: pertinente
rho_hauteur: 0.44

formants:
  - { hz: 466, statut: "indéterminé",  residu_db: 4.03, etendue_residu: [329, 3517], verdict: retenu, notes: 8 }
  - { hz: 659, statut: "route A",  residu_db: 5.79, etendue_residu: [415, 1863], verdict: retenu, notes: 11 }
  - { hz: 987, statut: "1 corpus",  residu_db: 3.69, etendue_residu: [329, 3726], verdict: retenu, notes: 18 }
  - { hz: 1479, statut: "1 corpus", residu_db: 4.30, etendue_residu: [329, 3320], verdict: retenu, notes: 25 }

pente_db_octave: -17.1
pente_sd: 2.6
harmoniques_par_note: 85
harmoniques_min_max: [31, 90]
balance_impair_pair_db: -4.7
enveloppe_vraie_hz: 280
enveloppe_vraie_sd: 110

version_mesure: 2026-10-23-wav-v1
---

# Tuba basse

> ## Statut des formants — révision du 26/10/2026
>
> Confrontation des **cinq corpus**. Provenance de ce corpus : **IRCAM (FullSOL2020)**.
>
> | Formant | Statut |
> |---|---|
> | **466 Hz** | **indéterminé** — concordance à un demi-ton avec ZSamples mais aucun couple au même point de grille. Environ une concordance sur quatre à ce seuil est fortuite. |
> | **659 Hz** | **route A** — VN + ZSamples au même point de grille, 698 Hz · zone vue par SOL, VN, ZSamples. ⚠ La valeur de cette fiche, 659 Hz, est à -1.0 demi-ton du couple : c'est **698 Hz** qui est établi |
> | **987 Hz** | présent dans le seul corpus SOL |
> | **1479 Hz** | présent dans le seul corpus SOL |
>
> `etendue_residu` dans l'en-tête **n'est pas une largeur de bande** : c'est l'étendue du résidu de la méthode 8, souvent une butée de plage. Les largeurs mesurées sont dans `_LARGEURS-DE-BANDE.md`.

## Méthode retenue, et pourquoi

> **Méthode 8 — quatre formants retenus, aucun rejeté.** Voir `00-METHODE.md`.

Classe A, ρ = 0,44. Conditions de mesure excellentes : 108 échantillons, une médiane de
**85 harmoniques par note** avec un minimum de 31, une tessiture (46 à 349 Hz) presque
entièrement sous la limite de validité des méthodes à ordre fixe, une pente très régulière
(σ = 2,6 dB/octave) et une enveloppe vraie stable (280 Hz, σ = 110 Hz).

C'est, avec le trombone ténor, le second cuivre dont les quatre candidats passent tous les
deux tests. Et c'est l'instrument des cuivres dont les résidus sont les plus forts : jusqu'à
+5,8 dB à 659 Hz.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes distinctes | Verdict |
|---|---|---|---|---|
| **466 Hz** | +4,0 dB | 329–3 517 Hz | 8 | **retenu** |
| **659 Hz** | +5,8 dB | 415–1 863 Hz | 11 | **retenu** |
| **987 Hz** | +3,7 dB | 329–3 726 Hz | 18 | **retenu** |
| **1 479 Hz** | +4,3 dB | 329–3 320 Hz | 25 | **retenu** |

## Tests de rejet

| Zone | k = 2–3 | k = 4–7 | k = 8–15 | k ≥ 16 | Robustesse |
|---|---|---|---|---|---|
| 466 Hz | +4,2 | +4,4 | +5,8 | — | oui |
| 659 Hz | +6,5 | +5,8 | +5,8 | — | oui |
| 987 Hz | +5,4 | +3,7 | +4,7 | +3,6 | oui |
| 1 479 Hz | — | +5,6 | +3,2 | +4,5 | oui |

Les quatre formants sont stables sur tous les groupes de rang disponibles, sans changement de
signe. Le cas de 659 Hz est le plus net du corpus des cuivres : résidu entre +5,8 et +6,5 dB
sur trois groupes consécutifs, avec une bande à −3 dB relativement circonscrite
(415–1 863 Hz).

Les positions (466, 659, 987, 1 479 Hz) sont approximativement dans les rapports
**1 : 1,41 : 2,12 : 3,17**, soit des intervalles proches de la quarte augmentée et de
l'octave. À examiner comme possible série de modes.

## Descripteurs globaux

| Descripteur | Valeur | Lecture |
|---|---|---|
| Pente spectrale médiane | −17,1 dB/octave (σ 2,6) | régulière |
| **Harmoniques par note** | **85 [31–90]** | quasi au plafond de mesure |
| Balance impair − pair | −4,7 dB | la plus négative des cuivres |
| Enveloppe vraie (méthode 6) | 280 Hz, σ 110 | stable |

La balance de −4,7 dB est la plus négative de la famille : le tuba basse privilégie nettement
les harmoniques pairs. Associée à une pente de −17,1 dB/octave, c'est la description d'un son
plein et rond, qui concentre son énergie dans les premiers partiels pairs.

## Comparaison avec la littérature

| Source | Valeur annoncée | Notre mesure | Accord |
|---|---|---|---|
| Backus (1969) | 200–400 Hz | env. vraie 280 Hz | **bon** |
| Meyer (2009) | 210–250 Hz, /u/ | env. vraie 280 Hz | **bon** |
| Giesler (1985) | 200–350 Hz, *u-ähnlich* | env. vraie 280 Hz | **bon** |
| Ancien pipeline `specenv` | F1 = 226 Hz, /u/ | env. vraie 280 Hz | proche |
| Ancien Fp | 1 239 Hz | formant retenu 1 479 Hz | proche |

Les trois sources classiques s'accordent et sont confirmées par l'enveloppe vraie à 280 Hz.
L'ancien F1 de 226 Hz en était proche, et son rapport au milieu géométrique de la tessiture
valait **2,00** — signe qu'il mesurait bien une résonance et non la hauteur médiane. Comme
pour le cor et les trombones, les anciennes valeurs des cuivres graves étaient donc
globalement justes, la méthode étant valide dans ce registre.

Les formants à 466, 659 et 987 Hz sont inédits : aucune source ne les mentionne, les mesures
classiques s'étant concentrées sur la zone la plus grave.

## Registres

`Data/registres.md`, entrée commune « tubas » : extrême grave en dessous de F1 · grave F1–F2 ·
médium F2–F3 · aigu F3–D4 · suraigu au-dessus de D#4.

Mesure par registre : non effectuée.

Note sur l'ancienne analyse par registre : elle donnait F1 = 215 Hz au grave et 205 Hz au
médium, puis 280 Hz à l'aigu et 334 Hz au suraigu. La « convergence exacte Δ = 0 Hz » avec le
violoncelle suraigu et avec la flûte grave reposait donc sur des rencontres de zones de
fondamentale entre instruments de tessitures très différentes.

## Observations pour l'orchestration

À compléter. Éléments acquis :

- **La zone 659 Hz est commune au tuba basse et au contrebasson** (698 Hz, à un demi-ton
  près), tous deux validés indépendamment et avec les résidus les plus forts de leurs
  familles respectives (+5,8 et +5,9 dB). C'est une convergence précise entre un cuivre grave
  et un bois grave, et elle pourrait fonder acoustiquement la doublure tuba–contrebasson.
- La zone 1 479 Hz rapproche le tuba basse du tuba contrebasse (1 396 Hz) : homogénéité du
  pupitre, mesurée.
- **L'« unisson formantique parfait Δ = 0 Hz » annoncé entre tuba basse, tuba contrebasse et
  contrebasson, tous à 226 Hz, est écarté.** Les trois valeurs tombaient dans le même bin de
  FFT (bin 21), c'est-à-dire dans la zone de leurs fondamentales respectives. Les formants
  réels sont 466/659/987/1 479 Hz pour le tuba basse, 370/879/1 396 Hz pour le tuba
  contrebasse et 329/698/1 174 Hz pour le contrebasson. Il existe bien des correspondances,
  mais pas celle qui était annoncée, et elles ne sont pas des unissons.

## Graphiques

Aucun à ce stade. Figures prévues dans `../Versions-html-and-docx/media/` :

| Fichier prévu | Contenu |
|---|---|
| `tuba_basse_composite.png` | enveloppe composite — quatre formants retenus |
| `tubas_pupitre.png` | superposition tuba basse / tuba contrebasse |
| `graves_convergence.png` | figure inter-familles : tuba basse 659 Hz, contrebasson 698 Hz |
| `unisson_226_ecarte.png` | figure méthodologique : les trois anciens F1 à 226 Hz contre les formants réels |

La dernière figure a une valeur pédagogique : elle montre comment trois valeurs dans un même
bin de FFT peuvent produire une « convergence parfaite » entièrement artificielle.

## Journal

- **2026-10-23** — Quatre formants retenus (466, 659, 987, 1 479 Hz), aucun rejet. Résidus
  les plus forts des cuivres.
- **2026-10-23** — Les trois sources classiques confirmées par l'enveloppe vraie (280 Hz).
- **2026-10-23** — L'« unisson parfait à 226 Hz » avec le tuba contrebasse et le
  contrebasson est écarté : même bin de FFT, zones de fondamentale.
- **2026-10-23** — Convergence tuba basse 659 Hz / contrebasson 698 Hz relevée comme piste
  solide.
