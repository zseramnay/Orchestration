---
id: violoncelle-piombo
nom: Violoncelle+piombo
source_echantillons: ZSamples/violoncello/practice_mute_metal
technique: practice_mute_metal
n_echantillons: 96
n_notes_distinctes: 29
f0_min_hz: 65.4
f0_max_hz: 1108.7
controle_nomenclature: 1.006   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: partielle   # À VALIDER À LA MAIN
rho_hauteur: 0.79
formants:
  - { hz: 3319.9, residu_db: 6.62, bande_db3: [2957.7, 3948.1], verdict: retenu-cr, notes: 25, rang_median: 25, ic_bas: [5.59, 5.83] }
  - { hz: 739.4, residu_db: 3.54, bande_db3: [697.9, 1863.2], verdict: fragile-cr, notes: 7, rang_median: 6, ic_bas: [0.14, 0.0] }
  - { hz: 931.6, residu_db: 3.14, bande_db3: [697.9, 1863.2], verdict: fragile-cr, notes: 8, rang_median: 7, ic_bas: [0.36, -1.16] }
  - { hz: 1566.8, residu_db: 2.33, bande_db3: [697.9, 1974.0], verdict: fragile-cr, notes: 13, rang_median: 12, ic_bas: [0.31, 0.0] }
pente_db_octave: -10.5
pente_sd: 1.9
harmoniques_par_note: 50
harmoniques_min_max: [9, 90]
balance_impair_pair_db: -3.1
enveloppe_vraie_hz: 291.0
enveloppe_vraie_sd: 617.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Violoncelle+piombo

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 1 formant(s) retenu(s), 3 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.79 · 50 harmoniques par note en médiane · 96 échantillons sur 29 notes distinctes · pente σ = 1.9 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 3320 Hz | +6.6 dB | 2958–3948 Hz | 25 | retenu-cr |
| 739 Hz | +3.5 dB | 698–1863 Hz | 7 | fragile-cr |
| 932 Hz | +3.1 dB | 698–1863 Hz | 8 | fragile-cr |
| 1567 Hz | +2.3 dB | 698–1974 Hz | 13 | fragile-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 3320 Hz (k méd. 25) | +6.6 dB [IC +5.6], 72 notes | +6.3 dB [IC +5.8], 27 notes | oui |
| 739 Hz (k méd. 6) | +1.4 dB [IC +0.1], 44 notes | +0.8 dB [IC +0.0], 21 notes | oui |
| 932 Hz (k méd. 7) | +1.9 dB [IC +0.4], 47 notes | +0.1 dB [IC -1.2], 24 notes | oui |
| 1567 Hz (k méd. 12) | +1.3 dB [IC +0.3], 66 notes | +2.2 dB [IC +0.0], 24 notes | oui |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -10.5 dB/octave (σ 1.9) |
| Harmoniques par note | 50 [9–90] |
| Balance impair − pair | -3.1 dB |
| Enveloppe vraie (méthode 6) | 291 Hz, σ 617 ; ordres 13–235 |
| Pics tombant sur un harmonique | 19.8% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
