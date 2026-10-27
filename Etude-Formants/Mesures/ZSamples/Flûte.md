---
id: flûte
nom: Flûte
source_echantillons: ZSamples/flutes/flute/ordinario
technique: ordinario
n_echantillons: 56
n_notes_distinctes: 14
f0_min_hz: 246.9
f0_max_hz: 2093.0
controle_nomenclature: 1.006   # doit valoir ~1,00
classe: B        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: partielle   # À VALIDER À LA MAIN
rho_hauteur: 0.88
formants:
  - { hz: 5270.0, residu_db: 2.83, bande_db3: [4974.2, 5915.4], verdict: fragile-cr, notes: 10, rang_median: 14, ic_bas: [1.73, 1.36] }
pente_db_octave: -18.4
pente_sd: 3.7
harmoniques_par_note: 15
harmoniques_min_max: [5, 44]
balance_impair_pair_db: -6.7
enveloppe_vraie_hz: 926.0
enveloppe_vraie_sd: 535.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Flûte

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 0 formant(s) retenu(s), 1 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.88 · 15 harmoniques par note en médiane · 56 échantillons sur 14 notes distinctes · pente σ = 3.7 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 5270 Hz | +2.8 dB | 4974–5915 Hz | 10 | fragile-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 5270 Hz (k méd. 14) | +2.4 dB [IC +1.7], 36 notes | +2.4 dB [IC +1.4], 16 notes | non |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -18.4 dB/octave (σ 3.7) |
| Harmoniques par note | 15 [5–44] |
| Balance impair − pair | -6.7 dB |
| Enveloppe vraie (méthode 6) | 926 Hz, σ 535 ; ordres 7–62 |
| Pics tombant sur un harmonique | 12.5% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
