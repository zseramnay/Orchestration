---
id: cor
nom: Cor
source_echantillons: ZSamples/horns/horn_in_f/ordinario
technique: ordinario
n_echantillons: 71
n_notes_distinctes: 19
f0_min_hz: 43.7
f0_max_hz: 784.0
controle_nomenclature: 1.003   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: partielle   # À VALIDER À LA MAIN
rho_hauteur: 0.66
formants:
  - { hz: 1758.7, residu_db: 5.48, bande_db3: [1395.8, 2215.8], verdict: retenu-cr, notes: 14, rang_median: 23, ic_bas: [4.3, 4.43] }
  - { hz: 1107.9, residu_db: 4.07, bande_db3: [830.0, 4182.8], verdict: fragile-cr, notes: 7, rang_median: 15, ic_bas: [2.47, 3.75] }
  - { hz: 3319.9, residu_db: 2.51, bande_db3: [830.0, 5915.4], verdict: fragile-cr, notes: 14, rang_median: 42, ic_bas: [2.65, -0.34] }
pente_db_octave: -18.4
pente_sd: 5.0
harmoniques_par_note: 62
harmoniques_min_max: [14, 90]
balance_impair_pair_db: -3.0
enveloppe_vraie_hz: 420.0
enveloppe_vraie_sd: 147.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Cor

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 1 formant(s) retenu(s), 2 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.66 · 62 harmoniques par note en médiane · 71 échantillons sur 19 notes distinctes · pente σ = 5.0 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 1759 Hz | +5.5 dB | 1396–2216 Hz | 14 | retenu-cr |
| 1108 Hz | +4.1 dB | 830–4183 Hz | 7 | fragile-cr |
| 3320 Hz | +2.5 dB | 830–5915 Hz | 14 | fragile-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 1759 Hz (k méd. 23) | +5.5 dB [IC +4.3], 48 notes | +5.3 dB [IC +4.4], 16 notes | oui |
| 1108 Hz (k méd. 15) | +3.4 dB [IC +2.5], 40 notes | +4.5 dB [IC +3.8], 16 notes | non |
| 3320 Hz (k méd. 42) | +3.7 dB [IC +2.6], 59 notes | +0.3 dB [IC -0.3], 16 notes | oui |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -18.4 dB/octave (σ 5.0) |
| Harmoniques par note | 62 [14–90] |
| Balance impair − pair | -3.0 dB |
| Enveloppe vraie (méthode 6) | 420 Hz, σ 147 ; ordres 19–353 |
| Pics tombant sur un harmonique | 2.8% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
