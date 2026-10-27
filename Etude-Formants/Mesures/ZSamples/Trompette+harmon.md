---
id: trompette-harmon
nom: Trompette+harmon
source_echantillons: ZSamples/trumpets/trumpet_in_c/harmon
technique: harmon
n_echantillons: 66
n_notes_distinctes: 28
f0_min_hz: 34.6
f0_max_hz: 1046.5
controle_nomenclature: 1.003   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.85
formants:
  - { hz: 391.7, residu_db: 4.07, bande_db3: [369.7, 465.8], verdict: fragile-cr, notes: 7, rang_median: 8, ic_bas: [0.54, 1.66] }
  - { hz: 1107.9, residu_db: 3.82, bande_db3: [658.8, 1478.9], verdict: retenu-cr, notes: 14, rang_median: 19, ic_bas: [2.33, 2.49] }
  - { hz: 2791.7, residu_db: 3.55, bande_db3: [2347.5, 3517.3], verdict: retenu-cr, notes: 20, rang_median: 46, ic_bas: [1.39, 1.44] }
  - { hz: 553.9, residu_db: 2.01, bande_db3: [522.9, 586.9], verdict: rejeté, notes: 13, rang_median: 10, ic_bas: [-0.2, 1.11] }
pente_db_octave: -8.5
pente_sd: 4.2
harmoniques_par_note: 60
harmoniques_min_max: [10, 90]
balance_impair_pair_db: -3.6
enveloppe_vraie_hz: 221.0
enveloppe_vraie_sd: 1614.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Trompette+harmon

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 2 formant(s) retenu(s), 1 fragile(s), 1 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.85 · 60 harmoniques par note en médiane · 66 échantillons sur 28 notes distinctes · pente σ = 4.2 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 392 Hz | +4.1 dB | 370–466 Hz | 7 | fragile-cr |
| 1108 Hz | +3.8 dB | 659–1479 Hz | 14 | retenu-cr |
| 2792 Hz | +3.5 dB | 2348–3517 Hz | 20 | retenu-cr |
| 554 Hz | +2.0 dB | 523–587 Hz | 13 | rejeté |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 392 Hz (k méd. 8) | +2.6 dB [IC +0.5], 20 notes | +4.3 dB [IC +1.7], 9 notes | non |
| 1108 Hz (k méd. 19) | +3.2 dB [IC +2.3], 36 notes | +4.1 dB [IC +2.5], 12 notes | oui |
| 2792 Hz (k méd. 46) | +2.2 dB [IC +1.4], 54 notes | +2.1 dB [IC +1.4], 15 notes | oui |
| 554 Hz (k méd. 10) | +1.4 dB [IC -0.2], 26 notes | +1.6 dB [IC +1.1], 12 notes | non |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -8.5 dB/octave (σ 4.2) |
| Harmoniques par note | 60 [10–90] |
| Balance impair − pair | -3.6 dB |
| Enveloppe vraie (méthode 6) | 221 Hz, σ 1614 ; ordres 14–445 |
| Pics tombant sur un harmonique | 6.1% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
