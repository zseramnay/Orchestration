---
id: trombone-ténor-harmon
nom: Trombone ténor+harmon
source_echantillons: ZSamples/trombones/tenor_trombone/harmon
technique: harmon
n_echantillons: 54
n_notes_distinctes: 19
f0_min_hz: 30.9
f0_max_hz: 523.3
controle_nomenclature: 1.004   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.84
formants:
  - { hz: 1395.8, residu_db: 9.61, bande_db3: [1317.5, 1566.8], verdict: retenu-cr, notes: 11, rang_median: 25, ic_bas: [8.58, 15.23] }
  - { hz: 3517.3, residu_db: 8.62, bande_db3: [3133.6, 3517.3], verdict: retenu-cr, notes: 14, rang_median: 54, ic_bas: [10.97, 5.5] }
  - { hz: 697.9, residu_db: 4.15, bande_db3: [522.9, 783.4], verdict: retenu-cr, notes: 8, rang_median: 13, ic_bas: [3.29, 4.6] }
pente_db_octave: -12.9
pente_sd: 4.0
harmoniques_par_note: 86
harmoniques_min_max: [20, 90]
balance_impair_pair_db: -3.5
enveloppe_vraie_hz: 172.0
enveloppe_vraie_sd: 458.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Trombone ténor+harmon

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 3 formant(s) retenu(s), 0 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.84 · 86 harmoniques par note en médiane · 54 échantillons sur 19 notes distinctes · pente σ = 4.0 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 1396 Hz | +9.6 dB | 1318–1567 Hz | 11 | retenu-cr |
| 3517 Hz | +8.6 dB | 3134–3517 Hz | 14 | retenu-cr |
| 698 Hz | +4.2 dB | 523–783 Hz | 8 | retenu-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 1396 Hz (k méd. 25) | +11.1 dB [IC +8.6], 42 notes | +16.7 dB [IC +15.2], 15 notes | oui |
| 3517 Hz (k méd. 54) | +11.9 dB [IC +11.0], 42 notes | +7.2 dB [IC +5.5], 15 notes | oui |
| 698 Hz (k méd. 13) | +4.0 dB [IC +3.3], 30 notes | +4.9 dB [IC +4.6], 12 notes | oui |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -12.9 dB/octave (σ 4.0) |
| Harmoniques par note | 86 [20–90] |
| Balance impair − pair | -3.5 dB |
| Enveloppe vraie (méthode 6) | 172 Hz, σ 458 ; ordres 29–499 |
| Pics tombant sur un harmonique | 18.5% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
