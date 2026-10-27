---
id: trombone-ténor-sourdine-sèche
nom: Trombone ténor+sourdine sèche
source_echantillons: ZSamples/trombones/tenor_trombone/straight+straight_metal+straight_wood
technique: straight+straight_metal+straight_wood
n_echantillons: 54
n_notes_distinctes: 19
f0_min_hz: 32.7
f0_max_hz: 523.3
controle_nomenclature: 1.006   # doit valoir ~1,00
classe: A        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.49
formants:
  - { hz: 1243.6, residu_db: 4.42, bande_db3: [1045.7, 4431.5], verdict: retenu-cr, notes: 11, rang_median: 22, ic_bas: [4.78, 3.04] }
  - { hz: 1863.2, residu_db: 3.57, bande_db3: [987.0, 4695.1], verdict: retenu-cr, notes: 16, rang_median: 35, ic_bas: [3.32, 2.96] }
  - { hz: 3948.1, residu_db: 2.52, bande_db3: [931.6, 4695.1], verdict: fragile-cr, notes: 14, rang_median: 54, ic_bas: [5.53, -4.36] }
  - { hz: 2791.7, residu_db: 2.35, bande_db3: [931.6, 4695.1], verdict: fragile-cr, notes: 15, rang_median: 49, ic_bas: [3.53, -1.32] }
pente_db_octave: -12.8
pente_sd: 5.9
harmoniques_par_note: 86
harmoniques_min_max: [20, 90]
balance_impair_pair_db: -0.4
enveloppe_vraie_hz: 1023.0
enveloppe_vraie_sd: 418.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Trombone ténor+sourdine sèche

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 2 formant(s) retenu(s), 2 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.49 · 86 harmoniques par note en médiane · 54 échantillons sur 19 notes distinctes · pente σ = 5.9 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 1244 Hz | +4.4 dB | 1046–4432 Hz | 11 | retenu-cr |
| 1863 Hz | +3.6 dB | 987–4695 Hz | 16 | retenu-cr |
| 3948 Hz | +2.5 dB | 932–4695 Hz | 14 | fragile-cr |
| 2792 Hz | +2.4 dB | 932–4695 Hz | 15 | fragile-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 1244 Hz (k méd. 22) | +5.7 dB [IC +4.8], 42 notes | +3.7 dB [IC +3.0], 15 notes | oui |
| 1863 Hz (k méd. 35) | +3.9 dB [IC +3.3], 42 notes | +3.7 dB [IC +3.0], 14 notes | oui |
| 3948 Hz (k méd. 54) | +6.5 dB [IC +5.5], 38 notes | -2.1 dB [IC -4.4], 16 notes | oui |
| 2792 Hz (k méd. 49) | +4.3 dB [IC +3.5], 45 notes | -0.1 dB [IC -1.3], 15 notes | oui |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -12.8 dB/octave (σ 5.9) |
| Harmoniques par note | 86 [20–90] |
| Balance impair − pair | -0.4 dB |
| Enveloppe vraie (méthode 6) | 1023 Hz, σ 418 ; ordres 29–471 |
| Pics tombant sur un harmonique | 3.7% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
