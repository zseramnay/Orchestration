---
id: trombone-ténor
nom: Trombone ténor
source_echantillons: ZSamples/trombones/tenor_trombone/ordinario
technique: ordinario
n_echantillons: 84
n_notes_distinctes: 25
f0_min_hz: 30.9
f0_max_hz: 880.0
controle_nomenclature: 1.006   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.6
formants:
  - { hz: 1660.0, residu_db: 3.51, bande_db3: [621.8, 4182.8], verdict: retenu-cr, notes: 12, rang_median: 29, ic_bas: [2.59, 3.66] }
  - { hz: 879.3, residu_db: 3.13, bande_db3: [522.9, 4431.5], verdict: retenu-cr, notes: 10, rang_median: 16, ic_bas: [2.58, 2.51] }
  - { hz: 2791.7, residu_db: 2.53, bande_db3: [522.9, 4431.5], verdict: fragile-cr, notes: 15, rang_median: 47, ic_bas: [3.35, -1.09] }
  - { hz: 3517.3, residu_db: 2.02, bande_db3: [522.9, 4695.1], verdict: rejeté, notes: 16, rang_median: 52, ic_bas: [3.52, -3.13] }
pente_db_octave: -14.8
pente_sd: 4.9
harmoniques_par_note: 70
harmoniques_min_max: [12, 90]
balance_impair_pair_db: -1.6
enveloppe_vraie_hz: 646.0
enveloppe_vraie_sd: 239.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Trombone ténor

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 2 formant(s) retenu(s), 1 fragile(s), 1 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.6 · 70 harmoniques par note en médiane · 84 échantillons sur 25 notes distinctes · pente σ = 4.9 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 1660 Hz | +3.5 dB | 622–4183 Hz | 12 | retenu-cr |
| 879 Hz | +3.1 dB | 523–4432 Hz | 10 | retenu-cr |
| 2792 Hz | +2.5 dB | 523–4432 Hz | 15 | fragile-cr |
| 3517 Hz | +2.0 dB | 523–4695 Hz | 16 | rejeté |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 1660 Hz (k méd. 29) | +3.5 dB [IC +2.6], 59 notes | +4.1 dB [IC +3.7], 20 notes | oui |
| 879 Hz (k méd. 16) | +3.0 dB [IC +2.6], 48 notes | +3.0 dB [IC +2.5], 16 notes | oui |
| 2792 Hz (k méd. 47) | +3.7 dB [IC +3.4], 67 notes | +0.0 dB [IC -1.1], 22 notes | oui |
| 3517 Hz (k méd. 52) | +4.2 dB [IC +3.5], 64 notes | -1.4 dB [IC -3.1], 20 notes | non |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -14.8 dB/octave (σ 4.9) |
| Harmoniques par note | 70 [12–90] |
| Balance impair − pair | -1.6 dB |
| Enveloppe vraie (méthode 6) | 646 Hz, σ 239 ; ordres 17–499 |
| Pics tombant sur un harmonique | 2.4% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
