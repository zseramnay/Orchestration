---
id: alto-sourdine
nom: Alto+sourdine
source_echantillons: ZSamples/viola/con_sordino
technique: con_sordino
n_echantillons: 93
n_notes_distinctes: 28
f0_min_hz: 130.8
f0_max_hz: 1864.7
controle_nomenclature: 1.006   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: partielle   # À VALIDER À LA MAIN
rho_hauteur: 0.84
formants:
  - { hz: 3726.5, residu_db: 5.83, bande_db3: [3319.9, 4431.5], verdict: retenu-cr, notes: 18, rang_median: 14, ic_bas: [5.01, 4.65] }
  - { hz: 2091.4, residu_db: 1.65, bande_db3: [1395.8, 5915.4], verdict: rejeté, notes: 8, rang_median: 8, ic_bas: [-0.84, 1.61] }
pente_db_octave: -13.9
pente_sd: 2.2
harmoniques_par_note: 26
harmoniques_min_max: [5, 83]
balance_impair_pair_db: -4.4
enveloppe_vraie_hz: 463.0
enveloppe_vraie_sd: 761.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Alto+sourdine

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 1 formant(s) retenu(s), 0 fragile(s), 1 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.84 · 26 harmoniques par note en médiane · 93 échantillons sur 28 notes distinctes · pente σ = 2.2 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 3726 Hz | +5.8 dB | 3320–4432 Hz | 18 | retenu-cr |
| 2091 Hz | +1.6 dB | 1396–5915 Hz | 8 | rejeté |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 3726 Hz (k méd. 14) | +5.6 dB [IC +5.0], 66 notes | +5.2 dB [IC +4.7], 27 notes | oui |
| 2091 Hz (k méd. 8) | +0.0 dB [IC -0.8], 57 notes | +2.6 dB [IC +1.6], 24 notes | non |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -13.9 dB/octave (σ 2.2) |
| Harmoniques par note | 26 [5–83] |
| Balance impair − pair | -4.4 dB |
| Enveloppe vraie (méthode 6) | 463 Hz, σ 761 ; ordres 8–117 |
| Pics tombant sur un harmonique | 17.2% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
