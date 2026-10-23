---
id: contrebasse
nom: Contrebasse
nom_en: Contrabass
famille: cordes
sous_famille: cordes solistes
ordre: 22
source_echantillons: FullSOL2020/Strings/Contrabass/ordinario
technique: ordinario
n_echantillons: 309
tessiture_mesuree: E1–C5
f0_min_hz: 41.2
f0_max_hz: 523.3

classe: C
pertinence_formant: pertinente
rho_hauteur: 0.90

formants:
  - { hz: 466,  residu_db: 3.49, bande_db3: [392, 587],   verdict: retenu, notes: 9 }
  - { hz: 1396, residu_db: 6.16, bande_db3: [1174, 1660], verdict: retenu, notes: 29 }
  - { hz: 2635, residu_db: 3.44, bande_db3: [987, 4695],  verdict: retenu, notes: 37 }
  - { hz: 3948, residu_db: 2.51, bande_db3: [879, 5915],  verdict: retenu, notes: 39 }

pente_db_octave: -12.8
pente_sd: 1.7
harmoniques_par_note: 83
harmoniques_min_max: [20, 90]
balance_impair_pair_db: -2.8
enveloppe_vraie_hz: 151
enveloppe_vraie_sd: 236

version_mesure: 2026-10-23-wav-v1
---

# Contrebasse

## Classe C — quatre formants retenus, dont deux très bien circonscrits

> **Méthode 8 — quatre formants retenus, aucun rejeté.** Voir `00-METHODE.md`, section
> « Typologie — trois classes ».

La contrebasse présente le **ρ le plus élevé des cordes, 0,90**, à égalité avec les flûtes —
et pourtant les quatre candidats détectés passent tous les deux tests, dont deux avec des
bandes à −3 dB étroites (392–587 Hz et 1 174–1 660 Hz). C'est, avec le violon, la
démonstration la plus claire que ρ seul ne peut pas servir d'indice d'absence de formant.

Conditions de mesure optimales : 309 échantillons, une médiane de **83 harmoniques par
note**, une tessiture descendant à 41 Hz entièrement sous la limite de validité des méthodes
à ordre fixe, et une pente très régulière (σ = 1,7 dB/octave).

## Formants

| Zone | Résidu | Largeur −3 dB | Notes distinctes | Verdict |
|---|---|---|---|---|
| **466 Hz** | +3,5 dB | 392–587 Hz | 9 | **retenu** |
| **1 396 Hz** | +6,2 dB | 1 174–1 660 Hz | 29 | **retenu** |
| **2 635 Hz** | +3,4 dB | 987–4 695 Hz | 37 | **retenu** |
| **3 948 Hz** | +2,5 dB | large | 39 | **retenu** (faible) |

## Tests de rejet

| Zone | k = 2–3 | k = 4–7 | k = 8–15 | k ≥ 16 | Robustesse |
|---|---|---|---|---|---|
| 466 Hz | +4,3 | +3,9 | +2,5 | — | oui |
| 1 396 Hz | — | +6,7 | +7,4 | +7,5 | oui |
| 2 635 Hz | — | +4,1 | +4,1 | +2,9 | oui |
| 3 948 Hz | — | — | +1,6 | +1,8 | oui |

Le formant à **1 396 Hz** est le mieux établi : résidu croissant régulièrement de +6,7 à
+7,5 dB sur trois groupes de rang, avec une bande étroite (1 174–1 660 Hz, moins d'une
demi-octave). C'est l'un des formants les plus nets de tout le corpus des cordes.

## Descripteurs globaux

| Descripteur | Valeur | Lecture |
|---|---|---|
| Pente spectrale médiane | −12,8 dB/octave (σ 1,7) | la plus plate des cordes |
| **Harmoniques par note** | **83 [20–90]** | quasi au plafond de mesure |
| Balance impair − pair | −2,8 dB | corde frottée |
| Enveloppe vraie (méthode 6) | 151 Hz, σ 236 | σ supérieur à la valeur |
| **ρ(pic, f0)** | **0,90** | le plus élevé des cordes |

## Comparaison avec la littérature

| Source | Valeur annoncée | Notre mesure | Lecture |
|---|---|---|---|
| Backus, Giesler, Meyer | pas de valeur propre à la contrebasse dans les tableaux de l'étude | — | à chercher aux sources primaires |
| Ancien pipeline `specenv` | F1 = 172 Hz, zone /u/ | — | zone de la fondamentale (rapport 1,36) |
| Cepstre mai 2026 (ordre 80) | 1 388 Hz, σ 840 | **1 396 Hz** | **excellent accord** |
| Cepstre mai 2026, **ensemble** | 344 Hz, σ 1 050 | — | incohérent avec le soliste |

Le cepstre de mai 2026 donnait **1 388 Hz** pour la contrebasse soliste, à 8 Hz de notre
formant retenu à 1 396 Hz. C'est le meilleur accord entre l'ancienne et la nouvelle méthode
de tout le corpus. L'explication est la même que pour le trombone basse : la tessiture de la
contrebasse étant presque entièrement sous 276 Hz, le cepstre d'ordre fixe 80 **était
valide** pour cet instrument, et seulement 4 % de ses pics tombaient sur un harmonique.

En revanche, l'ancienne étude avait publié le F1 de 172 Hz et placé la contrebasse dans la
ligne /u/ « 200–400 Hz » du tableau vocalique — alors que 172 Hz est **en dehors** de la
plage de sa propre ligne. Cette incohérence était visible sans aucune mesure.

L'ancienne étude relevait aussi que « le Fp du trombone converge avec la contrebasse
(Fp ≈ 1 283–1 297 Hz) dans tous les registres ». Nos mesures donnent 1 396 Hz pour la
contrebasse et 1 567 Hz pour le trombone ténor : la convergence existe mais est moins serrée
qu'annoncé, et elle portait sur une grandeur qui n'était pas un formant.

## Registres

`Data/registres.md` : grave C1–C2 · médium C2–C3 · aigu C3–C4 · suraigu C4 et au-dessus.

Mesure par registre : non effectuée.

## Observations pour l'orchestration

À compléter. Éléments acquis :

- La zone **466 Hz est commune à la contrebasse et au tuba basse** (466 Hz, exactement),
  tous deux validés indépendamment. Convergence précise entre les deux fondations graves de
  l'orchestre, cordes et cuivres.
- La zone **1 396 Hz est commune à la contrebasse et au tuba contrebasse** (1 396 Hz,
  exactement) et proche du tuba basse (1 479 Hz). Trois instruments, deux familles.
- La zone **3 948 Hz est commune à la contrebasse et au violoncelle** (3 948 Hz,
  exactement) : parenté de famille mesurée.
- L'ancienne étude annonçait la contrebasse comme « fondation » convergeant avec le
  contrebasson (Δ = 54 Hz), le trombone basse (Δ = 86 Hz) et la clarinette basse (Δ = 0 Hz
  dans le registre grave). Tous ces écarts portaient sur des zones de fondamentale. Ils sont
  à reprendre, mais les correspondances réelles à 466 et 1 396 Hz pourraient les remplacer —
  et elles sont, elles, mesurées.

## Graphiques

Aucun à ce stade. Figures prévues dans `../Versions-html-and-docx/media/` :

| Fichier prévu | Contenu |
|---|---|
| `contrebasse_composite.png` | enveloppe composite — quatre formants retenus |
| `contrebasse_1396.png` | résidu par rang à 1 396 Hz : +6,7 → +7,5 dB, formant le plus net des cordes |
| `graves_466.png` | convergence contrebasse / tuba basse à 466 Hz |
| `cordes_composite.png` | figure de famille partagée : les quatre cordes solistes |

## Journal

- **2026-10-23** — Quatre formants retenus (466, 1 396, 2 635, 3 948 Hz), aucun rejet.
- **2026-10-23** — ρ = 0,90, le plus élevé des cordes, et pourtant quatre formants retenus :
  confirmation décisive de la nécessité de la classe C.
- **2026-10-23** — Meilleur accord ancienne/nouvelle méthode du corpus : cepstre mai 2026 à
  1 388 Hz contre 1 396 Hz mesuré, l'ordre fixe 80 étant valide pour cette tessiture.
- **2026-10-23** — Convergences exactes relevées : 466 Hz avec le tuba basse, 1 396 Hz avec
  le tuba contrebasse, 3 948 Hz avec le violoncelle.
