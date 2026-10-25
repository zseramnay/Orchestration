---
id: contrebasson
nom: Contrebasson
source_echantillons: ZSamples/Bassons/Contrebasson/ordinario
technique: ordinario
n_echantillons: 60
n_notes_distinctes: 15
f0_min_hz: 29.1
f0_max_hz: 392.0
controle_nomenclature: 1.003   # doit valoir ~1,00
classe: A        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.31
formants:
  - { hz: 783.4, residu_db: 5.21, bande_db3: [621.8, 931.6], verdict: retenu-cr, notes: 8, rang_median: 17, ic_bas: [3.48, 6.01] }
  - { hz: 1863.2, residu_db: 2.62, bande_db3: [553.9, 3133.6], verdict: retenu-cr, notes: 12, rang_median: 37, ic_bas: [3.37, 0.4] }
  - { hz: 1243.6, residu_db: 1.85, bande_db3: [553.9, 5915.4], verdict: rejeté, notes: 9, rang_median: 25, ic_bas: [0.0, 1.91] }
  - { hz: 4182.8, residu_db: 1.38, bande_db3: [553.9, 5915.4], verdict: fragile-cr, notes: 11, rang_median: 50, ic_bas: [3.21, -0.21] }
pente_db_octave: -12.7
pente_sd: 1.9
harmoniques_par_note: 90
harmoniques_min_max: [28, 90]
balance_impair_pair_db: -1.9
enveloppe_vraie_hz: 296.0
enveloppe_vraie_sd: 76.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Contrebasson

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 2 formant(s) retenu(s), 1 fragile(s), 1 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.31 · 90 harmoniques par note en médiane · 60 échantillons sur 15 notes distinctes · pente σ = 1.9 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 783 Hz | +5.2 dB | 622–932 Hz | 8 | retenu-cr |
| 1863 Hz | +2.6 dB | 554–3134 Hz | 12 | retenu-cr |
| 1244 Hz | +1.9 dB | 554–5915 Hz | 9 | rejeté |
| 4183 Hz | +1.4 dB | 554–5915 Hz | 11 | fragile-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 783 Hz (k méd. 17) | +4.7 dB [IC +3.5], 48 notes | +7.4 dB [IC +6.0], 12 notes | oui |
| 1863 Hz (k méd. 37) | +3.9 dB [IC +3.4], 48 notes | +1.6 dB [IC +0.4], 16 notes | oui |
| 1244 Hz (k méd. 25) | +0.7 dB [IC +0.0], 48 notes | +2.3 dB [IC +1.9], 16 notes | non |
| 4183 Hz (k méd. 50) | +4.0 dB [IC +3.2], 36 notes | +0.1 dB [IC -0.2], 16 notes | oui |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -12.7 dB/octave (σ 1.9) |
| Harmoniques par note | 90 [28–90] |
| Balance impair − pair | -1.9 dB |
| Enveloppe vraie (méthode 6) | 296 Hz, σ 76 ; ordres 39–529 |
| Pics tombant sur un harmonique | 0.0% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
