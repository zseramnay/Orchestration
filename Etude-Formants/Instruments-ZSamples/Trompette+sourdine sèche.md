---
id: trompette-sourdine-sèche
nom: Trompette+sourdine sèche
source_echantillons: ZSamples/trumpets/trumpet_in_c/straight+straight_metal+straight_wood
technique: straight+straight_metal+straight_wood
n_echantillons: 63
n_notes_distinctes: 27
f0_min_hz: 38.9
f0_max_hz: 1046.5
controle_nomenclature: 1.003   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: partielle   # À VALIDER À LA MAIN
rho_hauteur: 0.7
formants:
  - { hz: 2091.4, residu_db: 3.33, bande_db3: [1243.6, 5915.4], verdict: retenu-cr, notes: 19, rang_median: 32, ic_bas: [3.03, 2.63] }
  - { hz: 3517.3, residu_db: 2.86, bande_db3: [1243.6, 5915.4], verdict: fragile-cr, notes: 22, rang_median: 47, ic_bas: [4.08, -1.99] }
  - { hz: 1478.9, residu_db: 1.22, bande_db3: [658.8, 5915.4], verdict: fragile-cr, notes: 16, rang_median: 24, ic_bas: [0.93, 1.41] }
  - { hz: 5583.4, residu_db: 1.18, bande_db3: [658.8, 5915.4], verdict: rejeté, notes: 17, rang_median: 41, ic_bas: [2.21, -1.28] }
pente_db_octave: -10.9
pente_sd: 5.9
harmoniques_par_note: 59
harmoniques_min_max: [10, 90]
balance_impair_pair_db: -1.1
enveloppe_vraie_hz: 1281.0
enveloppe_vraie_sd: 819.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Trompette+sourdine sèche

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 1 formant(s) retenu(s), 2 fragile(s), 1 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.7 · 59 harmoniques par note en médiane · 63 échantillons sur 27 notes distinctes · pente σ = 5.9 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 2091 Hz | +3.3 dB | 1244–5915 Hz | 19 | retenu-cr |
| 3517 Hz | +2.9 dB | 1244–5915 Hz | 22 | fragile-cr |
| 1479 Hz | +1.2 dB | 659–5915 Hz | 16 | fragile-cr |
| 5583 Hz | +1.2 dB | 659–5915 Hz | 17 | rejeté |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 2091 Hz (k méd. 32) | +3.9 dB [IC +3.0], 48 notes | +3.1 dB [IC +2.6], 12 notes | oui |
| 3517 Hz (k méd. 47) | +5.2 dB [IC +4.1], 51 notes | -0.9 dB [IC -2.0], 15 notes | oui |
| 1479 Hz (k méd. 24) | +1.2 dB [IC +0.9], 43 notes | +2.1 dB [IC +1.4], 12 notes | non |
| 5583 Hz (k méd. 41) | +2.8 dB [IC +2.2], 42 notes | -0.1 dB [IC -1.3], 15 notes | non |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -10.9 dB/octave (σ 5.9) |
| Harmoniques par note | 59 [10–90] |
| Balance impair − pair | -1.1 dB |
| Enveloppe vraie (méthode 6) | 1281 Hz, σ 819 ; ordres 14–396 |
| Pics tombant sur un harmonique | 4.8% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
