---
id: tuba-contrebasse
nom: Tuba contrebasse
nom_en: Contrabass tuba
famille: cuivres
sous_famille: cuivres principaux
ordre: 18
source_echantillons: Yan_Adds-Divers/Brass/Contrabass-Tuba/ordinario
technique: ordinario
n_echantillons: 158
tessiture_mesuree: C#0–C#5
f0_min_hz: 17.3
f0_max_hz: 554.4

classe: A
pertinence_formant: pertinente
rho_hauteur: 0.50

formants:
  - { hz: 370,  residu_db: 4.95, bande_db3: [329, 2487], verdict: retenu, notes: 15 }
  - { hz: 879,  residu_db: 4.13, bande_db3: [293, 2792], verdict: retenu, notes: 27 }
  - { hz: 1396, residu_db: 4.67, bande_db3: [311, 2635], verdict: retenu, notes: 36 }

pente_db_octave: -17.7
pente_sd: 2.3
harmoniques_par_note: 90
harmoniques_min_max: [19, 90]
balance_impair_pair_db: -2.9
enveloppe_vraie_hz: 269
enveloppe_vraie_sd: 110

version_mesure: 2026-10-23-wav-v1
---

# Tuba contrebasse

## Méthode retenue, et pourquoi

> **Méthode 8 — trois formants retenus, aucun rejeté.** Voir `00-METHODE.md`.

Classe A, ρ = 0,50. C'est l'instrument au corpus le plus fourni des cuivres,
**158 échantillons**, et à la tessiture la plus étendue de tout le corpus mesuré :
de 17,3 Hz (C#0) à 554 Hz, soit cinq octaves. Avec **90 harmoniques par note** au plafond de
mesure et une pente des plus régulières (σ = 2,3 dB/octave), les conditions sont optimales.

Les trois candidats détectés passent tous les deux tests.

Une remarque sur la fondamentale la plus grave : à 17,3 Hz, elle est sous le seuil
d'audition tonale. Pour ces notes, c'est l'ensemble des harmoniques qui porte la hauteur
perçue, et l'enveloppe composite y est particulièrement bien échantillonnée — chaque note
fournit ses 90 harmoniques dans la plage d'analyse.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes distinctes | Verdict |
|---|---|---|---|---|
| **370 Hz** | +5,0 dB | 329–2 487 Hz | 15 | **retenu** |
| **879 Hz** | +4,1 dB | 293–2 792 Hz | 27 | **retenu** |
| **1 396 Hz** | +4,7 dB | 311–2 635 Hz | 36 | **retenu** |

## Tests de rejet

| Zone | k = 2–3 | k = 4–7 | k = 8–15 | k ≥ 16 | Robustesse |
|---|---|---|---|---|---|
| 370 Hz | +4,2 | +6,8 | +5,6 | +5,8 | oui |
| 879 Hz | — | +5,9 | +5,9 | +2,1 | oui |
| 1 396 Hz | — | +4,4 | +3,6 | +4,8 | oui |

Le formant à **370 Hz est le seul du corpus des cuivres à être positif et supérieur à +4 dB
sur les quatre groupes de rang**. C'est le cas le plus complet : la bande reste élevée qu'elle
soit éclairée par l'harmonique 2 ou par l'harmonique 40. Aucune ambiguïté possible avec un
effet de hauteur.

## Descripteurs globaux

| Descripteur | Valeur | Lecture |
|---|---|---|
| Pente spectrale médiane | −17,7 dB/octave (σ 2,3) | la plus raide des cuivres graves |
| **Harmoniques par note** | **90 [19–90]** | plafond de mesure |
| Balance impair − pair | −2,9 dB | harmoniques pairs favorisés |
| Enveloppe vraie (méthode 6) | 269 Hz, σ 110 | stable ; ordres adaptatifs jusqu'à 529 |
| Tessiture mesurée | **5 octaves** | la plus étendue du corpus |

## Comparaison avec la littérature

| Source | Valeur annoncée | Notre mesure | Accord |
|---|---|---|---|
| Backus (1969), pour le tuba | 200–400 Hz | **formant retenu à 370 Hz** | **très bon** |
| Giesler (1985), pour le tuba | 200–350 Hz, *u-ähnlich* | **formant retenu à 370 Hz** | **bon** |
| Meyer (2009), pour le tuba | 210–250 Hz, /u/ | env. vraie 269 Hz | **bon** |
| Ancien pipeline `specenv` | F1 = 226 Hz | env. vraie 269 Hz | proche |
| Ancien Fp | 1 214 Hz *(mesuré sur Bass_Tuba)* | formant retenu 1 396 Hz | proche |

C'est l'un des rares cas où un formant **retenu** par les deux tests tombe directement dans
la plage annoncée par deux sources classiques. Le rapport entre l'ancien F1 (226 Hz) et le
milieu géométrique de la tessiture valait **2,59**, le plus élevé de tout le corpus : c'était
le signe le plus net possible que cette valeur mesurait une résonance réelle.

Rappel de l'avertissement consigné dans `methodology-learnings` : les fichiers `cor` et
`horn` des analyses antérieures étaient des doublons pointant sur les mêmes fichiers Hn. Il
faudra vérifier qu'une confusion analogue n'affecte pas les deux tubas dans les anciens CSV,
leurs valeurs publiées étant identiques au bin près (226 Hz pour les deux).

## Registres

`Data/registres.md`, entrée commune « tubas » : extrême grave en dessous de F1 · grave F1–F2 ·
médium F2–F3 · aigu F3–D4 · suraigu au-dessus de D#4.

Cette entrée est **inadaptée** au tuba contrebasse, dont la tessiture mesurée descend à C#0,
soit bien en dessous de F1, et monte à C#5, soit au-dessus de D#4. À redéfinir
spécifiquement.

Mesure par registre : non effectuée.

## Observations pour l'orchestration

À compléter. Éléments acquis :

- **Homogénéité du pupitre des tubas, mesurée.** Tuba basse 1 479 Hz contre tuba contrebasse
  1 396 Hz, soit un écart d'un demi-ton ; et tuba basse 987 Hz contre tuba contrebasse
  879 Hz. Les quatre valeurs sont validées indépendamment. L'ancienne étude affirmait cette
  homogénéité par un « Δ = 0 Hz » entre deux F1 à 226 Hz, dans le même bin : la conclusion
  était juste, la preuve ne l'était pas.
- La zone **879 Hz est commune au tuba contrebasse et au trombone basse** (879 Hz,
  exactement), tous deux validés. Convergence précise entre les deux instruments les plus
  graves des cuivres.
- L'« unisson formantique parfait » avec le contrebasson et le tuba basse à 226 Hz est
  écarté (voir `17-tuba-basse.md`). Il subsiste toutefois une correspondance réelle, non
  annoncée celle-là : tuba contrebasse 370 Hz contre contrebasson 329 Hz.
- L'ancienne étude proposait aussi une doublure avec la contrebasse (F1 = 172 Hz) et la
  flûte contrebasse (F1 = 334 Hz) sur des « zones graves /u/ partagées ». À reprendre quand
  les cordes auront été mesurées.

## Graphiques

Aucun à ce stade. Figures prévues dans `../Versions-html-and-docx/media/` :

| Fichier prévu | Contenu |
|---|---|
| `tuba_cb_composite.png` | enveloppe composite — le formant à 370 Hz, positif sur quatre groupes de rang |
| `tubas_pupitre.png` | figure partagée : superposition des deux tubas |
| `tuba_cb_rangs.png` | résidu par groupe de rang à 370 Hz — meilleure illustration du test d'indépendance |

## Journal

- **2026-10-23** — Trois formants retenus (370, 879, 1 396 Hz), aucun rejet.
- **2026-10-23** — Le formant à 370 Hz est le seul des cuivres positif et supérieur à +4 dB
  sur les quatre groupes de rang : cas le plus complet du test d'indépendance.
- **2026-10-23** — Zone 879 Hz identique à celle du trombone basse.
- **2026-10-23** — À vérifier : les anciens CSV donnaient 226 Hz pour les deux tubas, valeur
  identique au bin près. Risque de doublon analogue à celui déjà constaté entre `cor` et
  `horn`.
- **2026-10-23** — L'entrée « tubas » de `registres.md` est inadaptée à cet instrument
  (tessiture mesurée C#0–C#5).
