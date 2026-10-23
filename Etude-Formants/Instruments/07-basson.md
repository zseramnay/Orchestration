---
id: basson
nom: Basson
nom_en: Bassoon
famille: bois
sous_famille: anches doubles
ordre: 7
source_echantillons: FullSOL2020/Winds/Bassoon/ordinario
technique: ordinario
n_echantillons: 126
tessiture_mesuree: A#1–D#5
f0_min_hz: 58.3
f0_max_hz: 622.3

# --- classement méthodologique ---
classe: A                      # A = résonance fixe · B = spectre suiveur
pertinence_formant: pertinente # pertinente · partielle · non pertinente
rho_hauteur: 0.40              # Spearman(pic enveloppe vraie, f0) — proche de 0 = formant fixe

# --- formants validés (méthode 8, attaque exclue) ---
formants:
  - { hz: 554,  residu_db: 7.49, bande_db3: [523, 659],   verdict: retenu, notes: 7 }
  - { hz: 1108, residu_db: 4.65, bande_db3: [987, 1396],  verdict: retenu, notes: 18 }
  - { hz: 1974, residu_db: 7.60, bande_db3: [1759, 2216], verdict: retenu, notes: 26 }
  - { hz: 3320, residu_db: 2.48, bande_db3: [415, 5915],  verdict: fragile, notes: 34 }

# --- descripteurs globaux ---
pente_db_octave: -14.7
pente_sd: 1.7
harmoniques_par_note: 57       # médiane
harmoniques_min_max: [17, 90]
balance_impair_pair_db: -3.3
enveloppe_vraie_hz: 538
enveloppe_vraie_sd: 87

version_mesure: 2026-10-23-wav-v1
---

# Basson

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite par agrégation des amplitudes harmoniques.**
> Mesure sur les 126 échantillons *ordinario*, spectres moyens recalculés depuis
> les WAV en excluant le transitoire d'attaque (onset + 200 ms, fenêtre de 1,5 s).

Le basson est le cas idéal pour cette méthode, pour une raison quantifiable : avec
une médiane de **57 harmoniques par note** (de 17 à 90 selon le registre), chaque note
échantillonne densément la réponse du tuyau sur toute la plage analysée. L'agrégation
des 126 notes éclaire donc le filtre par autant de sources différentes, ce qui est la
définition physique même du formant.

La notion de formant est ici **pleinement pertinente**. Trois arguments convergent :

1. **ρ = 0,40** entre le pic de l'enveloppe vraie et la hauteur jouée — la valeur la plus
   basse des bois après le hautbois et le cor anglais. Le spectre ne suit pas la note.
2. Les quatre candidats détectés passent le test d'indépendance au rang harmonique :
   le résidu reste positif quel que soit le rang qui vient éclairer la bande.
3. **Pente σ = 1,7 dB/octave**, la plus régulière de toute la famille des bois : le
   comportement spectral de l'instrument varie très peu d'une note à l'autre.

La méthode 6 (enveloppe vraie à ordre cepstral adaptatif) est utilisable note par note,
avec des ordres allant de 24 à 264 selon la hauteur. Elle donne **538 Hz, σ = 87 Hz**,
en accord direct avec le premier formant composite.

## Formants validés

| Zone | Résidu | Largeur −3 dB | Notes distinctes | Verdict |
|---|---|---|---|---|
| **554 Hz** | +7,5 dB | 523–659 Hz | 7 | **retenu** |
| **1 108 Hz** | +4,7 dB | 987–1 396 Hz | 18 | **retenu** |
| **1 974 Hz** | +7,6 dB | 1 759–2 216 Hz | 26 | **retenu** |
| 3 320 Hz | +2,5 dB | très large | 34 | fragile |

Le résidu est l'écart, en dB, entre l'amplitude mesurée d'un harmonique et la médiane
locale des harmoniques de la même note dans une fenêtre de ±3/4 d'octave. Un résidu
positif signale une résonance plus étroite que ~1,5 octave.

## Tests de rejet

Chaque bande doit rester élevée **quel que soit le rang harmonique** qui y tombe : c'est
le test direct de l'indépendance à la hauteur. Un pic qui ne serait qu'un effet de
hauteur s'effondre dès que le rang change.

| Zone | k = 2–3 | k = 4–7 | k = 8–15 | k ≥ 16 | Robustesse fenêtre |
|---|---|---|---|---|---|
| 554 Hz | +10,4 | +8,1 | +6,3 | — | oui |
| 1 108 Hz | +5,3 | +2,1 | +5,9 | −2,2 | oui |
| 1 974 Hz | — | +7,7 | +8,3 | +9,1 | oui |
| 3 320 Hz | — | +0,4 | +1,2 | +2,1 | **non** |

Le pic à 3 320 Hz est classé fragile : ses résidus sont faibles et il ne survit pas au
changement de fenêtre de redressement (±1/2 et ±1 octave). Il traduit probablement une
tendance large plutôt qu'une résonance.

Le cas de 1 108 Hz mérite une note : le résidu devient négatif aux rangs ≥ 16, c'est-à-dire
pour les notes les plus graves de la tessiture. La résonance est donc bien réelle mais
moins marquée dans le registre grave — à vérifier par une mesure par registre.

## Descripteurs globaux

| Descripteur | Valeur | Lecture |
|---|---|---|
| Pente spectrale médiane | **−14,7 dB/octave** (σ 1,7) | décroissance modérée et très régulière |
| Harmoniques par note | 57 [17–90] | spectre très riche, mesure densément échantillonnée |
| Balance impair − pair | −3,3 dB | tuyau conique : harmoniques pairs présents |
| Enveloppe vraie (méthode 6) | 538 Hz, σ 87 | confirme le premier formant |

## Comparaison avec la littérature

Valeurs de la littérature telles que relevées dans les tableaux de l'étude précédente
(à revérifier aux sources primaires).

| Source | Valeur annoncée | Notre mesure | Accord |
|---|---|---|---|
| Backus (1969) | 440–500 Hz | 554 Hz | proche |
| Giesler (1985) | 500 Hz, *o-ähnlich* | 554 Hz | **bon** |
| Meyer (2009) | ~500 Hz et ~1 000 Hz | 554 et 1 108 Hz | **bon sur les deux** |
| Cepstre mai 2026 | 527 Hz | 554 Hz | **bon** |
| Ancien pipeline `specenv` | F1 = 495 Hz | 554 Hz | proche |

C'est le seul instrument dont la valeur de l'ancien pipeline ressort intacte, et pour une
raison précise : avec un f0 maximal de 622 Hz, une grande partie de la tessiture du basson
reste sous la limite de validité du lissage Orchidea à l'ordre 160 (276 Hz), et le reste
n'en est pas très loin. Le formant à 1 974 Hz, en revanche, est nouveau et n'apparaît dans
aucune source consultée.

## Registres

Découpage de référence (`Data/registres.md`) :
grave A#1–A2 · bas médium A#2–A3 · haut médium A#3–A4 · aigu à partir de A#4.

Mesure par registre : **non encore effectuée** avec la nouvelle méthode. À faire, en
particulier pour trancher le comportement de la bande 1 108 Hz dans le grave.

## Observations pour l'orchestration

À compléter. Pistes ouvertes, non encore mesurées :

- La doublure **flûte deux octaves au-dessus du basson**, préférée de Mozart : pour un
  basson entre D3 et D4, la flûte joue entre 587 et 1 175 Hz, donc sa fondamentale tombe
  exactement entre les formants mesurés à 554 et 1 108 Hz, et sur l'harmonique 4 du basson.
  Convergence de type « fondamentale suiveuse × formant fixe », à valider par la méthode 12.
- Le formant à 1 974 Hz, inédit, est le plus fort de l'instrument (+7,6 dB) et le mieux
  attesté (26 notes, positif sur trois groupes de rang). Son rôle dans la projection du
  basson reste à examiner.
- Zone 1 108 Hz commune au basson et au contrebasson (1 174 Hz) : parenté de famille
  mesurée, et non postulée.

## Graphiques

Aucun graphique produit à ce stade (décision du 23/10/2026 : tableaux d'abord).
Figures prévues, à placer dans `../Versions-html-and-docx/media/` :

| Fichier prévu | Contenu |
|---|---|
| `basson_composite.png` | enveloppe composite : résidu médian par bande de 1/12 d'octave, avec couverture |
| `basson_rangs.png` | résidu par groupe de rang harmonique, pour chaque formant retenu |
| `basson_te_registres.png` | enveloppes vraies superposées, une par registre |

## Journal

- **2026-10-23** — Reprise complète. Abandon des `specenv` d'Orchidea (lissage cepstral
  à ordre fixe 160, non valide au-dessus de 276 Hz) et du Fp par centroïde (grandeur bornée
  par sa propre fenêtre d'analyse, ne mesure pas une résonance). Passage à la méthode 8 sur
  les fichiers `spectrum` non lissés, puis recalcul depuis les WAV avec exclusion de l'attaque.
- **2026-10-23** — Les deux mesures (attaque incluse / exclue) concordent à un ou deux
  douzièmes d'octave près : 554 / 1 174 / 1 974 contre 554 / 1 108 / 1 974.
- **2026-10-23** — Premier essai avec redressement linéaire en log-fréquence abandonné :
  sur un spectre concave, il fabrique une bosse médiane artificielle. Remplacé par un
  redressement local à ±3/4 d'octave.
