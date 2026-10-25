---
id: trombone-ténor-cup
nom: Trombone ténor+cup
source_echantillons: ZSamples/trombones/tenor_trombone/cup+cup_metal+cup_open
technique: cup+cup_metal+cup_open
n_echantillons: 54
n_notes_distinctes: 20
f0_min_hz: 30.9
f0_max_hz: 523.3
controle_nomenclature: 1.004   # doit valoir ~1,00
classe: A        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.27
formants:
  - { hz: 697.9, residu_db: 6.84, bande_db3: [621.8, 931.6], verdict: retenu-cr, notes: 8, rang_median: 13, ic_bas: [6.14, 6.04] }
  - { hz: 2635.0, residu_db: 5.12, bande_db3: [2347.5, 3133.6], verdict: retenu-cr, notes: 19, rang_median: 46, ic_bas: [4.71, 3.1] }
  - { hz: 1107.9, residu_db: 4.0, bande_db3: [553.9, 1317.5], verdict: fragile-cr, notes: 15, rang_median: 20, ic_bas: [3.34, 2.11] }
pente_db_octave: -13.1
pente_sd: 4.5
harmoniques_par_note: 86
harmoniques_min_max: [20, 90]
balance_impair_pair_db: -1.2
enveloppe_vraie_hz: 630.0
enveloppe_vraie_sd: 204.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Trombone ténor+cup

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 2 formant(s) retenu(s), 1 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.27 · 86 harmoniques par note en médiane · 54 échantillons sur 20 notes distinctes · pente σ = 4.5 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 698 Hz | +6.8 dB | 622–932 Hz | 8 | retenu-cr |
| 2635 Hz | +5.1 dB | 2348–3134 Hz | 19 | retenu-cr |
| 1108 Hz | +4.0 dB | 554–1318 Hz | 15 | fragile-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 698 Hz (k méd. 13) | +6.9 dB [IC +6.1], 30 notes | +7.0 dB [IC +6.0], 12 notes | oui |
| 2635 Hz (k méd. 46) | +6.2 dB [IC +4.7], 45 notes | +4.4 dB [IC +3.1], 15 notes | oui |
| 1108 Hz (k méd. 20) | +4.2 dB [IC +3.3], 37 notes | +3.3 dB [IC +2.1], 15 notes | non |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -13.1 dB/octave (σ 4.5) |
| Harmoniques par note | 86 [20–90] |
| Balance impair − pair | -1.2 dB |
| Enveloppe vraie (méthode 6) | 630 Hz, σ 204 ; ordres 29–499 |
| Pics tombant sur un harmonique | 9.3% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
