---
id: trompette
nom: Trompette
source_echantillons: ZSamples/trumpets/trumpet_in_c/ordinario
technique: ordinario
n_echantillons: 74
n_notes_distinctes: 30
f0_min_hz: 36.7
f0_max_hz: 1108.7
controle_nomenclature: 1.003   # doit valoir ~1,00
classe: A        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: partielle   # À VALIDER À LA MAIN
rho_hauteur: 0.54
formants:
  - { hz: 1974.0, residu_db: 3.21, bande_db3: [879.3, 5915.4], verdict: fragile-cr, notes: 18, rang_median: 26, ic_bas: [3.33, 2.33] }
  - { hz: 1478.9, residu_db: 3.19, bande_db3: [879.3, 5915.4], verdict: retenu-cr, notes: 17, rang_median: 21, ic_bas: [2.99, 2.23] }
  - { hz: 4182.8, residu_db: 3.04, bande_db3: [879.3, 5915.4], verdict: fragile-cr, notes: 22, rang_median: 42, ic_bas: [4.5, -2.03] }
  - { hz: 2791.7, residu_db: 2.37, bande_db3: [739.4, 5915.4], verdict: fragile-cr, notes: 21, rang_median: 35, ic_bas: [2.97, -0.24] }
pente_db_octave: -13.1
pente_sd: 5.5
harmoniques_par_note: 49
harmoniques_min_max: [9, 90]
balance_impair_pair_db: -1.9
enveloppe_vraie_hz: 996.0
enveloppe_vraie_sd: 465.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Trompette

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 1 formant(s) retenu(s), 3 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.54 · 49 harmoniques par note en médiane · 74 échantillons sur 30 notes distinctes · pente σ = 5.5 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 1974 Hz | +3.2 dB | 879–5915 Hz | 18 | fragile-cr |
| 1479 Hz | +3.2 dB | 879–5915 Hz | 17 | retenu-cr |
| 4183 Hz | +3.0 dB | 879–5915 Hz | 22 | fragile-cr |
| 2792 Hz | +2.4 dB | 739–5915 Hz | 21 | fragile-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 1974 Hz (k méd. 26) | +4.0 dB [IC +3.3], 50 notes | +3.1 dB [IC +2.3], 14 notes | non |
| 1479 Hz (k méd. 21) | +3.4 dB [IC +3.0], 47 notes | +2.7 dB [IC +2.2], 12 notes | oui |
| 4183 Hz (k méd. 42) | +4.8 dB [IC +4.5], 59 notes | -0.4 dB [IC -2.0], 16 notes | oui |
| 2792 Hz (k méd. 35) | +3.2 dB [IC +3.0], 58 notes | +0.5 dB [IC -0.2], 15 notes | oui |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -13.1 dB/octave (σ 5.5) |
| Harmoniques par note | 49 [9–90] |
| Balance impair − pair | -1.9 dB |
| Enveloppe vraie (méthode 6) | 996 Hz, σ 465 ; ordres 13–420 |
| Pics tombant sur un harmonique | 2.7% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
