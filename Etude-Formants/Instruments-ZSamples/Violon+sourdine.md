---
id: violon-sourdine
nom: Violon+sourdine
source_echantillons: ZSamples/violin/con_sordino
technique: con_sordino
n_echantillons: 93
n_notes_distinctes: 29
f0_min_hz: 196.0
f0_max_hz: 2793.8
controle_nomenclature: 1.003   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: partielle   # À VALIDER À LA MAIN
rho_hauteur: 0.88
formants:
  - { hz: 3517.3, residu_db: 3.52, bande_db3: [2091.4, 5915.4], verdict: retenu-cr, notes: 10, rang_median: 9, ic_bas: [1.98, 0.67] }
  - { hz: 2791.7, residu_db: 2.77, bande_db3: [2091.4, 5915.4], verdict: fragile-cr, notes: 9, rang_median: 7, ic_bas: [0.68, 0.77] }
  - { hz: 2215.8, residu_db: 2.03, bande_db3: [2091.4, 5915.4], verdict: rejeté, notes: 7, rang_median: 6, ic_bas: [-0.32, 1.37] }
pente_db_octave: -13.0
pente_sd: 2.4
harmoniques_par_note: 17
harmoniques_min_max: [3, 56]
balance_impair_pair_db: -4.7
enveloppe_vraie_hz: 689.0
enveloppe_vraie_sd: 531.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Violon+sourdine

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 1 formant(s) retenu(s), 1 fragile(s), 1 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.88 · 17 harmoniques par note en médiane · 93 échantillons sur 29 notes distinctes · pente σ = 2.4 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 3517 Hz | +3.5 dB | 2091–5915 Hz | 10 | retenu-cr |
| 2792 Hz | +2.8 dB | 2091–5915 Hz | 9 | fragile-cr |
| 2216 Hz | +2.0 dB | 2091–5915 Hz | 7 | rejeté |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 3517 Hz (k méd. 9) | +2.9 dB [IC +2.0], 57 notes | +1.7 dB [IC +0.7], 24 notes | oui |
| 2792 Hz (k méd. 7) | +1.0 dB [IC +0.7], 48 notes | +3.0 dB [IC +0.8], 24 notes | oui |
| 2216 Hz (k méd. 6) | +1.8 dB [IC -0.3], 42 notes | +3.5 dB [IC +1.4], 21 notes | non |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -13.0 dB/octave (σ 2.4) |
| Harmoniques par note | 17 [3–56] |
| Balance impair − pair | -4.7 dB |
| Enveloppe vraie (méthode 6) | 689 Hz, σ 531 ; ordres 5–78 |
| Pics tombant sur un harmonique | 13.2% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
