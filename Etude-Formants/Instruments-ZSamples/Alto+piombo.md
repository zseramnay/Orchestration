---
id: alto-piombo
nom: Alto+piombo
source_echantillons: ZSamples/viola/practice_mute
technique: practice_mute
n_echantillons: 93
n_notes_distinctes: 28
f0_min_hz: 130.8
f0_max_hz: 1864.7
controle_nomenclature: 1.003   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: partielle   # À VALIDER À LA MAIN
rho_hauteur: 0.72
formants:
  - { hz: 3726.5, residu_db: 5.25, bande_db3: [3133.6, 4182.8], verdict: retenu-cr, notes: 18, rang_median: 14, ic_bas: [3.79, 6.22] }
  - { hz: 1660.0, residu_db: 3.63, bande_db3: [1566.8, 2091.4], verdict: fragile-cr, notes: 8, rang_median: 7, ic_bas: [0.0, 4.29] }
  - { hz: 2635.0, residu_db: 1.05, bande_db3: [1395.8, 5915.4], verdict: rejeté, notes: 11, rang_median: 10, ic_bas: [0.0, 0.0] }
pente_db_octave: -11.0
pente_sd: 2.7
harmoniques_par_note: 26
harmoniques_min_max: [5, 83]
balance_impair_pair_db: -2.9
enveloppe_vraie_hz: 662.0
enveloppe_vraie_sd: 669.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Alto+piombo

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 1 formant(s) retenu(s), 1 fragile(s), 1 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.72 · 26 harmoniques par note en médiane · 93 échantillons sur 28 notes distinctes · pente σ = 2.7 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 3726 Hz | +5.2 dB | 3134–4183 Hz | 18 | retenu-cr |
| 1660 Hz | +3.6 dB | 1567–2091 Hz | 8 | fragile-cr |
| 2635 Hz | +1.1 dB | 1396–5915 Hz | 11 | rejeté |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 3726 Hz (k méd. 14) | +5.2 dB [IC +3.8], 66 notes | +7.2 dB [IC +6.2], 27 notes | oui |
| 1660 Hz (k méd. 7) | +1.4 dB [IC +0.0], 45 notes | +5.8 dB [IC +4.3], 18 notes | oui |
| 2635 Hz (k méd. 10) | +0.0 dB [IC +0.0], 57 notes | +0.3 dB [IC +0.0], 24 notes | non |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -11.0 dB/octave (σ 2.7) |
| Harmoniques par note | 26 [5–83] |
| Balance impair − pair | -2.9 dB |
| Enveloppe vraie (méthode 6) | 662 Hz, σ 669 ; ordres 8–117 |
| Pics tombant sur un harmonique | 8.7% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
