---
id: basson
nom: Basson
source_echantillons: ZSamples/Bassons/Basson/ordinario
technique: ordinario
n_echantillons: 263
n_notes_distinctes: 44
f0_min_hz: 58.3
f0_max_hz: 698.5
controle_nomenclature: 1.006   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.58
formants:
  - { hz: 493.5, residu_db: 10.12, bande_db3: [439.7, 553.9], verdict: retenu-cr, notes: 11, rang_median: 5, ic_bas: [9.08, 6.89] }
  - { hz: 1107.9, residu_db: 6.36, bande_db3: [931.6, 1317.5], verdict: retenu-cr, notes: 19, rang_median: 10, ic_bas: [5.84, 2.71] }
  - { hz: 1974.0, residu_db: 6.26, bande_db3: [1660.0, 2347.5], verdict: retenu-cr, notes: 29, rang_median: 16, ic_bas: [5.16, 6.83] }
  - { hz: 3948.1, residu_db: 1.15, bande_db3: [329.4, 5915.4], verdict: fragile-cr, notes: 37, rang_median: 32, ic_bas: [0.93, 0.9] }
pente_db_octave: -14.2
pente_sd: 3.9
harmoniques_par_note: 49
harmoniques_min_max: [15, 90]
balance_impair_pair_db: -5.5
enveloppe_vraie_hz: 506.0
enveloppe_vraie_sd: 170.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Basson

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 3 formant(s) retenu(s), 1 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.58 · 49 harmoniques par note en médiane · 263 échantillons sur 44 notes distinctes · pente σ = 3.9 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 494 Hz | +10.1 dB | 440–554 Hz | 11 | retenu-cr |
| 1108 Hz | +6.4 dB | 932–1318 Hz | 19 | retenu-cr |
| 1974 Hz | +6.3 dB | 1660–2348 Hz | 29 | retenu-cr |
| 3948 Hz | +1.1 dB | 329–5915 Hz | 37 | fragile-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 494 Hz (k méd. 5) | +10.2 dB [IC +9.1], 73 notes | +8.7 dB [IC +6.9], 37 notes | oui |
| 1108 Hz (k méd. 10) | +6.3 dB [IC +5.8], 158 notes | +3.8 dB [IC +2.7], 57 notes | oui |
| 1974 Hz (k méd. 16) | +5.7 dB [IC +5.2], 204 notes | +7.7 dB [IC +6.8], 72 notes | oui |
| 3948 Hz (k méd. 32) | +1.6 dB [IC +0.9], 210 notes | +1.4 dB [IC +0.9], 77 notes | non |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -14.2 dB/octave (σ 3.9) |
| Harmoniques par note | 49 [15–90] |
| Balance impair − pair | -5.5 dB |
| Enveloppe vraie (méthode 6) | 506 Hz, σ 170 ; ordres 22–264 |
| Pics tombant sur un harmonique | 9.1% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
