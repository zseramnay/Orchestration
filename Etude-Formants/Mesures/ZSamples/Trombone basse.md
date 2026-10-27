---
id: trombone-basse
nom: Trombone basse
source_echantillons: ZSamples/trombones/bass_trombone/ordinario
technique: ordinario
n_echantillons: 68
n_notes_distinctes: 17
f0_min_hz: 32.7
f0_max_hz: 523.3
controle_nomenclature: 1.006   # doit valoir ~1,00
classe: A        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.5
formants:
  - { hz: 1660.0, residu_db: 3.57, bande_db3: [739.4, 4431.5], verdict: retenu-cr, notes: 10, rang_median: 28, ic_bas: [2.53, 3.36] }
  - { hz: 879.3, residu_db: 3.01, bande_db3: [739.4, 4695.1], verdict: retenu-cr, notes: 7, rang_median: 17, ic_bas: [2.71, 2.78] }
  - { hz: 2215.8, residu_db: 2.81, bande_db3: [739.4, 4695.1], verdict: fragile-cr, notes: 15, rang_median: 40, ic_bas: [2.8, 0.23] }
  - { hz: 3517.3, residu_db: 1.83, bande_db3: [739.4, 5915.4], verdict: rejeté, notes: 14, rang_median: 49, ic_bas: [2.58, -3.16] }
pente_db_octave: -12.9
pente_sd: 4.8
harmoniques_par_note: 83
harmoniques_min_max: [20, 90]
balance_impair_pair_db: 0.0
enveloppe_vraie_hz: 484.0
enveloppe_vraie_sd: 196.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Trombone basse

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 2 formant(s) retenu(s), 1 fragile(s), 1 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.5 · 83 harmoniques par note en médiane · 68 échantillons sur 17 notes distinctes · pente σ = 4.8 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 1660 Hz | +3.6 dB | 739–4432 Hz | 10 | retenu-cr |
| 879 Hz | +3.0 dB | 739–4695 Hz | 7 | retenu-cr |
| 2216 Hz | +2.8 dB | 739–4695 Hz | 15 | fragile-cr |
| 3517 Hz | +1.8 dB | 739–5915 Hz | 14 | rejeté |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 1660 Hz (k méd. 28) | +3.1 dB [IC +2.5], 56 notes | +3.9 dB [IC +3.4], 16 notes | oui |
| 879 Hz (k méd. 17) | +2.9 dB [IC +2.7], 48 notes | +3.4 dB [IC +2.8], 12 notes | oui |
| 2216 Hz (k méd. 40) | +3.1 dB [IC +2.8], 56 notes | +1.4 dB [IC +0.2], 16 notes | non |
| 3517 Hz (k méd. 49) | +3.2 dB [IC +2.6], 52 notes | -1.6 dB [IC -3.2], 20 notes | non |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -12.9 dB/octave (σ 4.8) |
| Harmoniques par note | 83 [20–90] |
| Balance impair − pair | 0.0 dB |
| Enveloppe vraie (méthode 6) | 484 Hz, σ 196 ; ordres 29–471 |
| Pics tombant sur un harmonique | 5.9% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
