---
id: hautbois
nom: Hautbois
source_echantillons: ZSamples/oboes/oboe/ordinario
technique: ordinario
n_echantillons: 52
n_notes_distinctes: 13
f0_min_hz: 233.1
f0_max_hz: 1864.7
controle_nomenclature: 1.004   # doit valoir ~1,00
classe: B        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: partielle   # À VALIDER À LA MAIN
rho_hauteur: 0.8
formants:
  - { hz: 5583.4, residu_db: 8.84, bande_db3: [5270.0, 5915.4], verdict: fragile-cr, notes: 10, rang_median: 14, ic_bas: [8.18, 7.28] }
pente_db_octave: -18.9
pente_sd: 5.3
harmoniques_par_note: 16
harmoniques_min_max: [5, 47]
balance_impair_pair_db: -5.4
enveloppe_vraie_hz: 1314.0
enveloppe_vraie_sd: 561.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Hautbois

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 0 formant(s) retenu(s), 1 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.8 · 16 harmoniques par note en médiane · 52 échantillons sur 13 notes distinctes · pente σ = 5.3 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 5583 Hz | +8.8 dB | 5270–5915 Hz | 10 | fragile-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 5583 Hz (k méd. 14) | +10.0 dB [IC +8.2], 36 notes | +8.7 dB [IC +7.3], 16 notes | non |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -18.9 dB/octave (σ 5.3) |
| Harmoniques par note | 16 [5–47] |
| Balance impair − pair | -5.4 dB |
| Enveloppe vraie (méthode 6) | 1314 Hz, σ 561 ; ordres 8–66 |
| Pics tombant sur un harmonique | 0.0% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
