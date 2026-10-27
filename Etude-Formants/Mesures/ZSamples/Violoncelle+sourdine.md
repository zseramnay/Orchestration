---
id: violoncelle-sourdine
nom: Violoncelle+sourdine
source_echantillons: ZSamples/violoncello/con_sordino_rubber
technique: con_sordino_rubber
n_echantillons: 96
n_notes_distinctes: 29
f0_min_hz: 65.4
f0_max_hz: 1108.7
controle_nomenclature: 1.006   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.85
formants:
  - { hz: 3726.5, residu_db: 4.8, bande_db3: [1758.7, 4431.5], verdict: retenu-cr, notes: 23, rang_median: 28, ic_bas: [4.76, 4.74] }
  - { hz: 2791.7, residu_db: 4.71, bande_db3: [1758.7, 4431.5], verdict: retenu-cr, notes: 23, rang_median: 21, ic_bas: [3.49, 5.78] }
  - { hz: 1863.2, residu_db: 2.66, bande_db3: [1317.5, 5915.4], verdict: retenu-cr, notes: 17, rang_median: 14, ic_bas: [2.86, 2.3] }
  - { hz: 931.6, residu_db: 1.37, bande_db3: [697.9, 1045.7], verdict: rejeté, notes: 8, rang_median: 7, ic_bas: [2.13, -1.08] }
pente_db_octave: -13.1
pente_sd: 1.8
harmoniques_par_note: 50
harmoniques_min_max: [9, 90]
balance_impair_pair_db: -4.2
enveloppe_vraie_hz: 248.0
enveloppe_vraie_sd: 711.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Violoncelle+sourdine

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 3 formant(s) retenu(s), 0 fragile(s), 1 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.85 · 50 harmoniques par note en médiane · 96 échantillons sur 29 notes distinctes · pente σ = 1.8 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 3726 Hz | +4.8 dB | 1759–4432 Hz | 23 | retenu-cr |
| 2792 Hz | +4.7 dB | 1759–4432 Hz | 23 | retenu-cr |
| 1863 Hz | +2.7 dB | 1318–5915 Hz | 17 | retenu-cr |
| 932 Hz | +1.4 dB | 698–1046 Hz | 8 | rejeté |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 3726 Hz (k méd. 28) | +5.3 dB [IC +4.8], 75 notes | +5.6 dB [IC +4.7], 27 notes | oui |
| 2792 Hz (k méd. 21) | +4.1 dB [IC +3.5], 69 notes | +6.7 dB [IC +5.8], 27 notes | oui |
| 1863 Hz (k méd. 14) | +3.9 dB [IC +2.9], 66 notes | +3.1 dB [IC +2.3], 27 notes | oui |
| 932 Hz (k méd. 7) | +2.9 dB [IC +2.1], 48 notes | +0.4 dB [IC -1.1], 24 notes | non |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -13.1 dB/octave (σ 1.8) |
| Harmoniques par note | 50 [9–90] |
| Balance impair − pair | -4.2 dB |
| Enveloppe vraie (méthode 6) | 248 Hz, σ 711 ; ordres 13–235 |
| Pics tombant sur un harmonique | 6.2% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
