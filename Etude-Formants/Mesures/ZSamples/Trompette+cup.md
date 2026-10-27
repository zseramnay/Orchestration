---
id: trompette-cup
nom: Trompette+cup
source_echantillons: ZSamples/trumpets/trumpet_in_c/cup+cup_metal+cup_open
technique: cup+cup_metal+cup_open
n_echantillons: 66
n_notes_distinctes: 30
f0_min_hz: 32.7
f0_max_hz: 1046.5
controle_nomenclature: 1.003   # doit valoir ~1,00
classe: A        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.46
formants:
  - { hz: 1243.6, residu_db: 7.86, bande_db3: [1107.9, 1395.8], verdict: retenu-cr, notes: 16, rang_median: 21, ic_bas: [8.38, 7.57] }
  - { hz: 2791.7, residu_db: 4.74, bande_db3: [2487.1, 3948.1], verdict: retenu-cr, notes: 21, rang_median: 49, ic_bas: [6.67, 5.44] }
  - { hz: 3517.3, residu_db: 2.91, bande_db3: [2347.5, 4182.8], verdict: fragile-cr, notes: 22, rang_median: 47, ic_bas: [5.7, -4.13] }
pente_db_octave: -13.6
pente_sd: 5.3
harmoniques_par_note: 59
harmoniques_min_max: [10, 90]
balance_impair_pair_db: -2.8
enveloppe_vraie_hz: 980.0
enveloppe_vraie_sd: 450.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Trompette+cup

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 2 formant(s) retenu(s), 1 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.46 · 59 harmoniques par note en médiane · 66 échantillons sur 30 notes distinctes · pente σ = 5.3 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 1244 Hz | +7.9 dB | 1108–1396 Hz | 16 | retenu-cr |
| 2792 Hz | +4.7 dB | 2487–3948 Hz | 21 | retenu-cr |
| 3517 Hz | +2.9 dB | 2348–4183 Hz | 22 | fragile-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 1244 Hz (k méd. 21) | +8.9 dB [IC +8.4], 39 notes | +8.1 dB [IC +7.6], 14 notes | oui |
| 2792 Hz (k méd. 49) | +7.4 dB [IC +6.7], 51 notes | +5.9 dB [IC +5.4], 14 notes | oui |
| 3517 Hz (k méd. 47) | +6.5 dB [IC +5.7], 51 notes | -2.9 dB [IC -4.1], 14 notes | oui |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -13.6 dB/octave (σ 5.3) |
| Harmoniques par note | 59 [10–90] |
| Balance impair − pair | -2.8 dB |
| Enveloppe vraie (méthode 6) | 980 Hz, σ 450 ; ordres 14–471 |
| Pics tombant sur un harmonique | 4.5% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
