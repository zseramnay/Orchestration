---
id: trompette-wah-fermé
nom: Trompette+wah fermé
source_echantillons: ZSamples/trumpets/trumpet_in_c/wahwah_stem_in_closed
technique: wahwah_stem_in_closed
n_echantillons: 60
n_notes_distinctes: 28
f0_min_hz: 46.2
f0_max_hz: 1108.7
controle_nomenclature: 1.003   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.73
formants:
  - { hz: 1660.0, residu_db: 7.82, bande_db3: [1478.9, 1974.0], verdict: retenu-cr, notes: 16, rang_median: 19, ic_bas: [6.17, 8.5] }
  - { hz: 739.4, residu_db: 3.86, bande_db3: [658.8, 987.0], verdict: retenu-cr, notes: 7, rang_median: 10, ic_bas: [1.71, 0.46] }
  - { hz: 2957.7, residu_db: 3.46, bande_db3: [658.8, 3517.3], verdict: retenu-cr, notes: 21, rang_median: 34, ic_bas: [4.13, 2.16] }
pente_db_octave: -12.9
pente_sd: 3.4
harmoniques_par_note: 54
harmoniques_min_max: [9, 90]
balance_impair_pair_db: -2.3
enveloppe_vraie_hz: 716.0
enveloppe_vraie_sd: 421.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Trompette+wah fermé

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 3 formant(s) retenu(s), 0 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.73 · 54 harmoniques par note en médiane · 60 échantillons sur 28 notes distinctes · pente σ = 3.4 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 1660 Hz | +7.8 dB | 1479–1974 Hz | 16 | retenu-cr |
| 739 Hz | +3.9 dB | 659–987 Hz | 7 | retenu-cr |
| 2958 Hz | +3.5 dB | 659–3517 Hz | 21 | retenu-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 1660 Hz (k méd. 19) | +7.8 dB [IC +6.2], 38 notes | +8.8 dB [IC +8.5], 14 notes | oui |
| 739 Hz (k méd. 10) | +3.4 dB [IC +1.7], 29 notes | +3.7 dB [IC +0.5], 10 notes | oui |
| 2958 Hz (k méd. 34) | +5.4 dB [IC +4.1], 49 notes | +2.8 dB [IC +2.2], 15 notes | oui |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -12.9 dB/octave (σ 3.4) |
| Harmoniques par note | 54 [9–90] |
| Balance impair − pair | -2.3 dB |
| Enveloppe vraie (méthode 6) | 716 Hz, σ 421 ; ordres 13–333 |
| Pics tombant sur un harmonique | 5.0% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
