---
id: contrebasse-piombo
nom: Contrebasse+piombo
source_echantillons: ZSamples/double_bass/practice_mute
technique: practice_mute
n_echantillons: 138
n_notes_distinctes: 39
f0_min_hz: 30.9
f0_max_hz: 830.6
controle_nomenclature: 1.006   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.79
formants:
  - { hz: 1758.7, residu_db: 5.04, bande_db3: [1478.9, 2091.4], verdict: retenu-cr, notes: 32, rang_median: 27, ic_bas: [4.88, 4.77] }
  - { hz: 783.4, residu_db: 4.1, bande_db3: [697.9, 987.0], verdict: retenu-cr, notes: 16, rang_median: 12, ic_bas: [2.9, 3.59] }
  - { hz: 2487.1, residu_db: 2.31, bande_db3: [621.8, 3133.6], verdict: retenu-cr, notes: 36, rang_median: 37, ic_bas: [2.93, 0.64] }
  - { hz: 4182.8, residu_db: 2.15, bande_db3: [3726.5, 4695.1], verdict: fragile-cr, notes: 33, rang_median: 51, ic_bas: [3.99, 1.11] }
pente_db_octave: -11.6
pente_sd: 1.6
harmoniques_par_note: 90
harmoniques_min_max: [13, 90]
balance_impair_pair_db: -3.8
enveloppe_vraie_hz: 161.0
enveloppe_vraie_sd: 263.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Contrebasse+piombo

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 3 formant(s) retenu(s), 1 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.79 · 90 harmoniques par note en médiane · 138 échantillons sur 39 notes distinctes · pente σ = 1.6 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 1759 Hz | +5.0 dB | 1479–2091 Hz | 32 | retenu-cr |
| 783 Hz | +4.1 dB | 698–987 Hz | 16 | retenu-cr |
| 2487 Hz | +2.3 dB | 622–3134 Hz | 36 | retenu-cr |
| 4183 Hz | +2.1 dB | 3726–4695 Hz | 33 | fragile-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 1759 Hz (k méd. 27) | +5.2 dB [IC +4.9], 101 notes | +5.3 dB [IC +4.8], 39 notes | oui |
| 783 Hz (k méd. 12) | +4.1 dB [IC +2.9], 87 notes | +4.5 dB [IC +3.6], 36 notes | oui |
| 2487 Hz (k méd. 37) | +3.2 dB [IC +2.9], 105 notes | +1.3 dB [IC +0.6], 39 notes | oui |
| 4183 Hz (k méd. 51) | +4.3 dB [IC +4.0], 99 notes | +1.9 dB [IC +1.1], 45 notes | non |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -11.6 dB/octave (σ 1.6) |
| Harmoniques par note | 90 [13–90] |
| Balance impair − pair | -3.8 dB |
| Enveloppe vraie (méthode 6) | 161 Hz, σ 263 ; ordres 18–499 |
| Pics tombant sur un harmonique | 15.9% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
