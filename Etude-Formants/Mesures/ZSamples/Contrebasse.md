---
id: contrebasse
nom: Contrebasse
source_echantillons: ZSamples/double_bass/ordinario
technique: ordinario
n_echantillons: 188
n_notes_distinctes: 41
f0_min_hz: 30.9
f0_max_hz: 830.6
controle_nomenclature: 1.006   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.76
formants:
  - { hz: 987.0, residu_db: 4.89, bande_db3: [783.4, 1758.7], verdict: retenu-cr, notes: 23, rang_median: 14, ic_bas: [2.53, 4.53] }
  - { hz: 439.7, residu_db: 4.0, bande_db3: [391.7, 522.9], verdict: retenu-cr, notes: 11, rang_median: 7, ic_bas: [2.79, 2.92] }
  - { hz: 1395.8, residu_db: 3.53, bande_db3: [739.4, 2957.7], verdict: fragile-cr, notes: 27, rang_median: 20, ic_bas: [2.75, 2.63] }
  - { hz: 2635.0, residu_db: 3.23, bande_db3: [739.4, 3133.6], verdict: retenu-cr, notes: 33, rang_median: 38, ic_bas: [4.48, 1.0] }
pente_db_octave: -13.1
pente_sd: 1.7
harmoniques_par_note: 90
harmoniques_min_max: [13, 90]
balance_impair_pair_db: -3.1
enveloppe_vraie_hz: 183.0
enveloppe_vraie_sd: 215.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Contrebasse

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 3 formant(s) retenu(s), 1 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.76 · 90 harmoniques par note en médiane · 188 échantillons sur 41 notes distinctes · pente σ = 1.7 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 987 Hz | +4.9 dB | 783–1759 Hz | 23 | retenu-cr |
| 440 Hz | +4.0 dB | 392–523 Hz | 11 | retenu-cr |
| 1396 Hz | +3.5 dB | 739–2958 Hz | 27 | fragile-cr |
| 2635 Hz | +3.2 dB | 739–3134 Hz | 33 | retenu-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 987 Hz (k méd. 14) | +3.4 dB [IC +2.5], 124 notes | +5.3 dB [IC +4.5], 53 notes | oui |
| 440 Hz (k méd. 7) | +3.8 dB [IC +2.8], 92 notes | +4.8 dB [IC +2.9], 40 notes | oui |
| 1396 Hz (k méd. 20) | +3.3 dB [IC +2.8], 132 notes | +2.9 dB [IC +2.6], 60 notes | non |
| 2635 Hz (k méd. 38) | +4.7 dB [IC +4.5], 144 notes | +1.9 dB [IC +1.0], 60 notes | oui |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -13.1 dB/octave (σ 1.7) |
| Harmoniques par note | 90 [13–90] |
| Balance impair − pair | -3.1 dB |
| Enveloppe vraie (méthode 6) | 183 Hz, σ 215 ; ordres 18–499 |
| Pics tombant sur un harmonique | 8.0% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
