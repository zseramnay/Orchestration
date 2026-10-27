---
id: tuba-basse-sourdine-sèche
nom: Tuba basse+sourdine sèche
source_echantillons: ZSamples/tubas/bass_tuba/straight+straight_metal+straight_wood
technique: straight+straight_metal+straight_wood
n_echantillons: 51
n_notes_distinctes: 20
f0_min_hz: 27.5
f0_max_hz: 466.2
controle_nomenclature: 1.003   # doit valoir ~1,00
classe: A        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: partielle   # À VALIDER À LA MAIN
rho_hauteur: 0.34
formants:
  - { hz: 1317.5, residu_db: 5.58, bande_db3: [658.8, 1660.0], verdict: retenu-cr, notes: 14, rang_median: 26, ic_bas: [5.17, 4.27] }
  - { hz: 830.0, residu_db: 3.93, bande_db3: [658.8, 1863.2], verdict: fragile-cr, notes: 11, rang_median: 18, ic_bas: [1.83, 3.14] }
  - { hz: 2957.7, residu_db: 1.22, bande_db3: [658.8, 5915.4], verdict: rejeté, notes: 14, rang_median: 46, ic_bas: [2.57, -0.81] }
pente_db_octave: -17.1
pente_sd: 2.1
harmoniques_par_note: 90
harmoniques_min_max: [23, 90]
balance_impair_pair_db: -3.4
enveloppe_vraie_hz: 431.0
enveloppe_vraie_sd: 112.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Tuba basse+sourdine sèche

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 1 formant(s) retenu(s), 1 fragile(s), 1 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.34 · 90 harmoniques par note en médiane · 51 échantillons sur 20 notes distinctes · pente σ = 2.1 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 1318 Hz | +5.6 dB | 659–1660 Hz | 14 | retenu-cr |
| 830 Hz | +3.9 dB | 659–1863 Hz | 11 | fragile-cr |
| 2958 Hz | +1.2 dB | 659–5915 Hz | 14 | rejeté |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 1318 Hz (k méd. 26) | +6.0 dB [IC +5.2], 42 notes | +5.9 dB [IC +4.3], 12 notes | oui |
| 830 Hz (k méd. 18) | +2.4 dB [IC +1.8], 33 notes | +4.7 dB [IC +3.1], 9 notes | non |
| 2958 Hz (k méd. 46) | +3.0 dB [IC +2.6], 39 notes | +0.0 dB [IC -0.8], 15 notes | non |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -17.1 dB/octave (σ 2.1) |
| Harmoniques par note | 90 [23–90] |
| Balance impair − pair | -3.4 dB |
| Enveloppe vraie (méthode 6) | 431 Hz, σ 112 ; ordres 33–560 |
| Pics tombant sur un harmonique | 3.9% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
