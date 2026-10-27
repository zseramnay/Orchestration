---
id: contrebasse-sourdine
nom: Contrebasse+sourdine
source_echantillons: ZSamples/double_bass/con_sordino
technique: con_sordino
n_echantillons: 138
n_notes_distinctes: 40
f0_min_hz: 30.9
f0_max_hz: 830.6
controle_nomenclature: 1.006   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.79
formants:
  - { hz: 2091.4, residu_db: 5.42, bande_db3: [1660.0, 2635.0], verdict: retenu-cr, notes: 33, rang_median: 31, ic_bas: [4.74, 3.29] }
  - { hz: 391.7, residu_db: 4.73, bande_db3: [349.0, 493.5], verdict: retenu-cr, notes: 11, rang_median: 6, ic_bas: [3.3, 2.18] }
  - { hz: 4182.8, residu_db: 2.39, bande_db3: [3726.5, 4974.2], verdict: fragile-cr, notes: 33, rang_median: 51, ic_bas: [4.34, 1.2] }
  - { hz: 830.0, residu_db: 1.82, bande_db3: [329.4, 3133.6], verdict: fragile-cr, notes: 18, rang_median: 12, ic_bas: [1.22, 0.65] }
pente_db_octave: -12.2
pente_sd: 1.6
harmoniques_par_note: 90
harmoniques_min_max: [13, 90]
balance_impair_pair_db: -4.1
enveloppe_vraie_hz: 183.0
enveloppe_vraie_sd: 180.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Contrebasse+sourdine

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 2 formant(s) retenu(s), 2 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.79 · 90 harmoniques par note en médiane · 138 échantillons sur 40 notes distinctes · pente σ = 1.6 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 2091 Hz | +5.4 dB | 1660–2635 Hz | 33 | retenu-cr |
| 392 Hz | +4.7 dB | 349–494 Hz | 11 | retenu-cr |
| 4183 Hz | +2.4 dB | 3726–4974 Hz | 33 | fragile-cr |
| 830 Hz | +1.8 dB | 329–3134 Hz | 18 | fragile-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 2091 Hz (k méd. 31) | +5.3 dB [IC +4.7], 103 notes | +4.1 dB [IC +3.3], 39 notes | oui |
| 392 Hz (k méd. 6) | +4.7 dB [IC +3.3], 63 notes | +3.2 dB [IC +2.2], 33 notes | oui |
| 4183 Hz (k méd. 51) | +4.9 dB [IC +4.3], 99 notes | +1.7 dB [IC +1.2], 46 notes | non |
| 830 Hz (k méd. 12) | +2.1 dB [IC +1.2], 90 notes | +1.6 dB [IC +0.7], 39 notes | non |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -12.2 dB/octave (σ 1.6) |
| Harmoniques par note | 90 [13–90] |
| Balance impair − pair | -4.1 dB |
| Enveloppe vraie (méthode 6) | 183 Hz, σ 180 ; ordres 18–499 |
| Pics tombant sur un harmonique | 8.0% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
