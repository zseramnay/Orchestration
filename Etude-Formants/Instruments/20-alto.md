---
id: alto
nom: Alto
nom_en: Viola
famille: cordes
sous_famille: cordes solistes
ordre: 20
source_echantillons: FullSOL2020/Strings/Viola/ordinario
technique: ordinario
n_echantillons: 309
tessiture_mesuree: C3–C7
f0_min_hz: 130.8
f0_max_hz: 2093.0

classe: C
pertinence_formant: pertinente
rho_hauteur: 0.87

formants:
  - { hz: 1318, residu_db: 1.80, bande_db3: [1108, 5915], verdict: retenu, notes: 8 }
  - { hz: 2091, residu_db: 3.62, bande_db3: [1863, 5915], verdict: retenu, notes: 15 }
  - { hz: 3517, residu_db: 5.81, bande_db3: [2792, 5915], verdict: retenu, notes: 24 }
  - { hz: 4974, residu_db: 4.14, bande_db3: [1863, 5915], verdict: retenu, notes: 30 }

pente_db_octave: -15.6
pente_sd: 1.8
harmoniques_par_note: 22
harmoniques_min_max: [5, 83]
balance_impair_pair_db: -3.6
enveloppe_vraie_hz: 538
enveloppe_vraie_sd: 970

version_mesure: 2026-10-23-wav-v1
---

# Alto

## Classe C — quatre formants retenus, aucun rejeté

> **Méthode 8 — quatre formants retenus.** Voir `00-METHODE.md`, section
> « Typologie — trois classes ».

L'alto est l'illustration la plus nette de la classe C. Son ρ atteint **0,87**, aussi élevé
que celui de la clarinette basse, et son enveloppe vraie est très dispersée (538 Hz,
σ = 970 Hz). Pourtant **les quatre candidats détectés passent tous les deux tests de rejet**,
sans aucun rejet ni cas fragile — performance que seuls le trombone ténor et le tuba basse
égalent dans tout le corpus.

Conditions de mesure favorables : 309 échantillons, le deuxième corpus le plus fourni, et une
pente d'une régularité exceptionnelle, σ = 1,8 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes distinctes | Verdict |
|---|---|---|---|---|
| **1 318 Hz** | +1,8 dB | large | 8 | **retenu** (faible) |
| **2 091 Hz** | +3,6 dB | large | 15 | **retenu** |
| **3 517 Hz** | +5,8 dB | 2 792–5 915 Hz | 24 | **retenu** |
| **4 974 Hz** | +4,1 dB | large | 30 | **retenu** |

## Tests de rejet

| Zone | k = 2–3 | k = 4–7 | k = 8–15 | k ≥ 16 | Robustesse |
|---|---|---|---|---|---|
| 1 318 Hz | +1,5 | +2,1 | +4,7 | — | oui |
| 2 091 Hz | +3,5 | +4,4 | +6,1 | — | oui |
| 3 517 Hz | +2,2 | +4,6 | +5,1 | +5,1 | oui |
| 4 974 Hz | +3,6 | +2,1 | +1,9 | +1,8 | oui |

Les quatre formants sont stables, sans changement de signe. Le plus fort, à 3 517 Hz, voit
son résidu croître avec le rang harmonique puis se stabiliser à +5,1 dB, ce qui est le
comportement attendu d'une résonance de caisse éclairée par des harmoniques de plus en plus
serrés.

## Descripteurs globaux

| Descripteur | Valeur | Lecture |
|---|---|---|
| Pente spectrale médiane | −15,6 dB/octave (**σ 1,8**) | la deuxième plus régulière du corpus |
| Harmoniques par note | 22 [5–83] | correct |
| Balance impair − pair | −3,6 dB | corde frottée |
| Enveloppe vraie (méthode 6) | 538 Hz, σ 970 | inexploitable note par note |

## Comparaison avec la littérature

| Source | Valeur annoncée | Notre mesure | Lecture |
|---|---|---|---|
| Backus, Giesler, Meyer | pas de valeur propre à l'alto dans les tableaux de l'étude | — | à chercher aux sources primaires |
| Ancien pipeline `specenv` | F1 = 377 Hz, zone /o/ | — | hauteur médiane (rapport 0,91) |
| Ancien Fp | — (non publié séparément) | — | — |
| Cepstre mai 2026 (ordre 80) | 409 Hz, σ **1 191** | — | σ supérieur à la valeur : inexploitable |

Aucune source classique ne donne de valeur spécifique à l'alto dans les tableaux de l'étude
antérieure, qui le plaçait dans la zone /o/ à 377 Hz sur la seule base du pipeline `specenv`.
Les quatre formants mesurés ici sont donc **entièrement inédits**.

Point d'attention relevé dans les données de mai 2026 : le cepstre d'ordre 80 donnait 409 Hz
pour l'alto soliste mais **393 Hz pour l'ensemble d'altos avec un σ de 25 Hz seulement**,
contre σ = 1 191 Hz pour le soliste. Cet écart de dispersion d'un facteur 48 entre soliste et
pupitre est une anomalie qui mérite examen : il suggère que les légers désaccords d'un
pupitre « remplissent » l'espace entre harmoniques et révèlent mieux le corps de résonance.
**Les ensembles seraient alors de meilleurs témoins formantiques que les solistes.** À
tester, et c'est une piste méthodologique potentiellement importante.

## Registres

`Data/registres.md` : grave C3–G3 · médium G3–G4 · aigu G4–G5 · suraigu G5 et au-dessus.

Mesure par registre : non effectuée.

## Observations pour l'orchestration

À compléter. Éléments acquis :

- La zone **3 517 Hz est identique à celle de la clarinette en Sib** (3 517 Hz, exactement)
  et proche de celle du trombone basse (3 320 Hz) et du basson (3 320 Hz). Quatre
  instruments de trois familles, tous validés.
- La zone **2 091 Hz est identique à celle du trombone ténor** (2 091 Hz) et proche de
  celles du cor (2 216 Hz) et du cor anglais (2 216 Hz).
- L'ancienne étude annonçait une convergence « quasi-parfaite Δ = 11 Hz » cor–alto
  (388 contre 377 Hz), fondement supposé de « la fusion cor-cordes ». Le F1 du cor était une
  mesure valide (rapport 1,87), celui de l'alto non (rapport 0,91) : la convergence reposait
  donc sur une valeur juste et une valeur fausse, et elle est **écartée**. Mais une
  correspondance réelle subsiste ailleurs, à 2 091–2 216 Hz, et elle pourrait rendre compte
  du même phénomène perceptif. À vérifier.
- La convergence alto–clarinette Sib annoncée à Δ = 86 Hz (377 contre 463 Hz) est écartée,
  mais remplacée par une coïncidence exacte à 3 517 Hz.

## Graphiques

Aucun à ce stade. Figures prévues dans `../Versions-html-and-docx/media/` :

| Fichier prévu | Contenu |
|---|---|
| `alto_composite.png` | enveloppe composite — quatre formants retenus, bonne figure de classe C |
| `alto_soliste_vs_ensemble.png` | **piste méthodologique** : dispersion soliste (σ 1 191) contre pupitre (σ 25) |
| `cordes_composite.png` | figure de famille : les quatre cordes solistes superposées |

La deuxième figure concerne une hypothèse qui pourrait changer la méthode : si les pupitres
sont de meilleurs témoins que les solistes, il faudra privilégier les ensembles.

## Journal

- **2026-10-23** — Quatre formants retenus, aucun rejet. Avec le trombone ténor et le tuba
  basse, l'un des trois seuls cas du corpus.
- **2026-10-23** — Les quatre formants sont inédits : aucune source classique ne traite
  l'alto spécifiquement.
- **2026-10-23** — Anomalie relevée dans les données de mai 2026 : σ de 1 191 Hz pour le
  soliste contre 25 Hz pour le pupitre. Hypothèse à tester — les ensembles révéleraient mieux
  le corps de résonance.
