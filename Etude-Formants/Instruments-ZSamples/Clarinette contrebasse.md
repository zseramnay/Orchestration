---
id: clarinette-contrebasse
nom: Clarinette contrebasse
source_echantillons: ZSamples/Clarinets/contrabass_clarinet/ordinario
technique: ordinario
n_echantillons: 52
n_notes_distinctes: 15
f0_min_hz: 27.5
f0_max_hz: 233.1
controle_nomenclature: 1.006   # doit valoir ~1,00
classe: A        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: partielle   # À VALIDER À LA MAIN
rho_hauteur: 0.13
formants:
  - { hz: 1395.8, residu_db: 3.59, bande_db3: [1107.9, 2347.5], verdict: retenu-cr, notes: 14, rang_median: 30, ic_bas: [3.2, 2.1] }
  - { hz: 1974.0, residu_db: 2.29, bande_db3: [553.9, 5915.4], verdict: fragile-cr, notes: 14, rang_median: 43, ic_bas: [1.84, -0.46] }
  - { hz: 739.4, residu_db: 2.03, bande_db3: [553.9, 5915.4], verdict: fragile-cr, notes: 9, rang_median: 14, ic_bas: [1.15, -1.33] }
  - { hz: 931.6, residu_db: 1.72, bande_db3: [553.9, 5915.4], verdict: rejeté, notes: 9, rang_median: 20, ic_bas: [0.0, 0.06] }
pente_db_octave: -10.5
pente_sd: 3.2
harmoniques_par_note: 90
harmoniques_min_max: [47, 90]
balance_impair_pair_db: 10.6
enveloppe_vraie_hz: 221.0
enveloppe_vraie_sd: 110.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Clarinette contrebasse

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 1 formant(s) retenu(s), 2 fragile(s), 1 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.13 · 90 harmoniques par note en médiane · 52 échantillons sur 15 notes distinctes · pente σ = 3.2 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 1396 Hz | +3.6 dB | 1108–2348 Hz | 14 | retenu-cr |
| 1974 Hz | +2.3 dB | 554–5915 Hz | 14 | fragile-cr |
| 739 Hz | +2.0 dB | 554–5915 Hz | 9 | fragile-cr |
| 932 Hz | +1.7 dB | 554–5915 Hz | 9 | rejeté |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 1396 Hz (k méd. 30) | +4.3 dB [IC +3.2], 41 notes | +3.9 dB [IC +2.1], 16 notes | oui |
| 1974 Hz (k méd. 43) | +2.4 dB [IC +1.8], 43 notes | +0.9 dB [IC -0.5], 16 notes | oui |
| 739 Hz (k méd. 14) | +3.0 dB [IC +1.1], 40 notes | +0.5 dB [IC -1.3], 16 notes | oui |
| 932 Hz (k méd. 20) | +2.2 dB [IC +0.0], 43 notes | +1.9 dB [IC +0.1], 16 notes | non |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -10.5 dB/octave (σ 3.2) |
| Harmoniques par note | 90 [47–90] |
| Balance impair − pair | 10.6 dB |
| Enveloppe vraie (méthode 6) | 221 Hz, σ 110 ; ordres 66–560 |
| Pics tombant sur un harmonique | 44.2% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
