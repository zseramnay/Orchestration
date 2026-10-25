---
id: alto
nom: Alto
source_echantillons: ZSamples/viola/ordinario
technique: ordinario
n_echantillons: 160
n_notes_distinctes: 35
f0_min_hz: 130.8
f0_max_hz: 1864.7
controle_nomenclature: 1.003   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.86
formants:
  - { hz: 3726.5, residu_db: 7.9, bande_db3: [3133.6, 4182.8], verdict: retenu-cr, notes: 19, rang_median: 13, ic_bas: [6.74, 7.07] }
  - { hz: 2347.5, residu_db: 3.53, bande_db3: [1758.7, 5915.4], verdict: retenu-cr, notes: 12, rang_median: 8, ic_bas: [1.62, 3.2] }
  - { hz: 5583.4, residu_db: 3.13, bande_db3: [1758.7, 5915.4], verdict: retenu-cr, notes: 24, rang_median: 18, ic_bas: [1.89, 3.37] }
pente_db_octave: -15.5
pente_sd: 2.3
harmoniques_par_note: 23
harmoniques_min_max: [5, 83]
balance_impair_pair_db: -4.0
enveloppe_vraie_hz: 554.0
enveloppe_vraie_sd: 627.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Alto

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 3 formant(s) retenu(s), 0 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.86 · 23 harmoniques par note en médiane · 160 échantillons sur 35 notes distinctes · pente σ = 2.3 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 3726 Hz | +7.9 dB | 3134–4183 Hz | 19 | retenu-cr |
| 2348 Hz | +3.5 dB | 1759–5915 Hz | 12 | retenu-cr |
| 5583 Hz | +3.1 dB | 1759–5915 Hz | 24 | retenu-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 3726 Hz (k méd. 13) | +7.6 dB [IC +6.7], 116 notes | +7.7 dB [IC +7.1], 44 notes | oui |
| 2348 Hz (k méd. 8) | +3.1 dB [IC +1.6], 92 notes | +4.1 dB [IC +3.2], 36 notes | oui |
| 5583 Hz (k méd. 18) | +2.6 dB [IC +1.9], 125 notes | +3.9 dB [IC +3.4], 48 notes | oui |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -15.5 dB/octave (σ 2.3) |
| Harmoniques par note | 23 [5–83] |
| Balance impair − pair | -4.0 dB |
| Enveloppe vraie (méthode 6) | 554 Hz, σ 627 ; ordres 8–117 |
| Pics tombant sur un harmonique | 13.8% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
