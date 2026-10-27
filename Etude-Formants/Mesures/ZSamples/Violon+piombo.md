---
id: violon-piombo
nom: Violon+piombo
source_echantillons: ZSamples/violin/practice_mute
technique: practice_mute
n_echantillons: 93
n_notes_distinctes: 29
f0_min_hz: 196.0
f0_max_hz: 2793.8
controle_nomenclature: 1.003   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.64
formants:
  - { hz: 3726.5, residu_db: 3.31, bande_db3: [2487.1, 5915.4], verdict: retenu-cr, notes: 12, rang_median: 10, ic_bas: [2.07, 2.67] }
  - { hz: 5583.4, residu_db: 2.13, bande_db3: [2091.4, 5915.4], verdict: retenu-cr, notes: 17, rang_median: 14, ic_bas: [1.32, 2.36] }
  - { hz: 2791.7, residu_db: 2.03, bande_db3: [2091.4, 5915.4], verdict: fragile-cr, notes: 9, rang_median: 7, ic_bas: [0.0, 0.87] }
pente_db_octave: -8.7
pente_sd: 2.5
harmoniques_par_note: 17
harmoniques_min_max: [3, 56]
balance_impair_pair_db: -3.9
enveloppe_vraie_hz: 947.0
enveloppe_vraie_sd: 1020.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Violon+piombo

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 2 formant(s) retenu(s), 1 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.64 · 17 harmoniques par note en médiane · 93 échantillons sur 29 notes distinctes · pente σ = 2.5 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 3726 Hz | +3.3 dB | 2487–5915 Hz | 12 | retenu-cr |
| 5583 Hz | +2.1 dB | 2091–5915 Hz | 17 | retenu-cr |
| 2792 Hz | +2.0 dB | 2091–5915 Hz | 9 | fragile-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 3726 Hz (k méd. 10) | +3.3 dB [IC +2.1], 61 notes | +3.9 dB [IC +2.7], 24 notes | oui |
| 5583 Hz (k méd. 14) | +2.2 dB [IC +1.3], 66 notes | +3.0 dB [IC +2.4], 27 notes | oui |
| 2792 Hz (k méd. 7) | +0.9 dB [IC +0.0], 48 notes | +2.4 dB [IC +0.9], 24 notes | oui |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -8.7 dB/octave (σ 2.5) |
| Harmoniques par note | 17 [3–56] |
| Balance impair − pair | -3.9 dB |
| Enveloppe vraie (méthode 6) | 947 Hz, σ 1020 ; ordres 5–78 |
| Pics tombant sur un harmonique | 22.6% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
