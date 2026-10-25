---
id: trombone-ténor-wah-ouvert
nom: Trombone ténor+wah ouvert
source_echantillons: ZSamples/trombones/tenor_trombone/wahwah_stem_in_open
technique: wahwah_stem_in_open
n_echantillons: 54
n_notes_distinctes: 22
f0_min_hz: 32.7
f0_max_hz: 523.3
controle_nomenclature: 1.006   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.61
formants:
  - { hz: 1107.9, residu_db: 6.27, bande_db3: [987.0, 1395.8], verdict: retenu-cr, notes: 16, rang_median: 20, ic_bas: [6.43, 5.32] }
  - { hz: 2091.4, residu_db: 3.57, bande_db3: [879.3, 2215.8], verdict: retenu-cr, notes: 18, rang_median: 38, ic_bas: [4.13, 2.45] }
  - { hz: 3726.5, residu_db: 2.17, bande_db3: [2957.7, 4695.1], verdict: retenu-cr, notes: 17, rang_median: 54, ic_bas: [1.63, 0.53] }
pente_db_octave: -12.6
pente_sd: 5.0
harmoniques_par_note: 86
harmoniques_min_max: [20, 90]
balance_impair_pair_db: -1.6
enveloppe_vraie_hz: 1001.0
enveloppe_vraie_sd: 448.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Trombone ténor+wah ouvert

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 3 formant(s) retenu(s), 0 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.61 · 86 harmoniques par note en médiane · 54 échantillons sur 22 notes distinctes · pente σ = 5.0 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 1108 Hz | +6.3 dB | 987–1396 Hz | 16 | retenu-cr |
| 2091 Hz | +3.6 dB | 879–2216 Hz | 18 | retenu-cr |
| 3726 Hz | +2.2 dB | 2958–4695 Hz | 17 | retenu-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 1108 Hz (k méd. 20) | +6.7 dB [IC +6.4], 36 notes | +5.7 dB [IC +5.3], 15 notes | oui |
| 2091 Hz (k méd. 38) | +5.0 dB [IC +4.1], 45 notes | +3.3 dB [IC +2.5], 15 notes | oui |
| 3726 Hz (k méd. 54) | +2.2 dB [IC +1.6], 39 notes | +1.3 dB [IC +0.5], 15 notes | oui |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -12.6 dB/octave (σ 5.0) |
| Harmoniques par note | 86 [20–90] |
| Balance impair − pair | -1.6 dB |
| Enveloppe vraie (méthode 6) | 1001 Hz, σ 448 ; ordres 29–471 |
| Pics tombant sur un harmonique | 1.9% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
