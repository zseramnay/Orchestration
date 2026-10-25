---
id: trompette-wah-ouvert
nom: Trompette+wah ouvert
source_echantillons: ZSamples/trumpets/trumpet_in_c/wahwah_stem_in_open
technique: wahwah_stem_in_open
n_echantillons: 60
n_notes_distinctes: 28
f0_min_hz: 46.2
f0_max_hz: 1108.7
controle_nomenclature: 1.006   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.57
formants:
  - { hz: 1478.9, residu_db: 7.13, bande_db3: [1317.5, 1758.7], verdict: retenu-cr, notes: 15, rang_median: 18, ic_bas: [6.9, 7.55] }
  - { hz: 2957.7, residu_db: 6.15, bande_db3: [2791.7, 3133.6], verdict: retenu-cr, notes: 20, rang_median: 36, ic_bas: [6.4, 4.38] }
pente_db_octave: -9.2
pente_sd: 5.1
harmoniques_par_note: 54
harmoniques_min_max: [9, 90]
balance_impair_pair_db: -0.9
enveloppe_vraie_hz: 1453.0
enveloppe_vraie_sd: 570.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Trompette+wah ouvert

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 2 formant(s) retenu(s), 0 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.57 · 54 harmoniques par note en médiane · 60 échantillons sur 28 notes distinctes · pente σ = 5.1 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 1479 Hz | +7.1 dB | 1318–1759 Hz | 15 | retenu-cr |
| 2958 Hz | +6.2 dB | 2792–3134 Hz | 20 | retenu-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 1479 Hz (k méd. 18) | +8.2 dB [IC +6.9], 39 notes | +7.9 dB [IC +7.5], 13 notes | oui |
| 2958 Hz (k méd. 36) | +7.8 dB [IC +6.4], 50 notes | +6.2 dB [IC +4.4], 13 notes | oui |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -9.2 dB/octave (σ 5.1) |
| Harmoniques par note | 54 [9–90] |
| Balance impair − pair | -0.9 dB |
| Enveloppe vraie (méthode 6) | 1453 Hz, σ 570 ; ordres 13–333 |
| Pics tombant sur un harmonique | 6.7% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
