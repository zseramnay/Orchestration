---
id: tuba-basse
nom: Tuba basse
source_echantillons: ZSamples/tubas/bass_tuba/ordinario
technique: ordinario
n_echantillons: 76
n_notes_distinctes: 26
f0_min_hz: 27.5
f0_max_hz: 659.3
controle_nomenclature: 1.006   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.59
formants:
  - { hz: 1243.6, residu_db: 5.84, bande_db3: [931.6, 1974.0], verdict: retenu-cr, notes: 14, rang_median: 25, ic_bas: [4.79, 5.1] }
  - { hz: 493.5, residu_db: 4.65, bande_db3: [439.7, 2091.4], verdict: retenu-cr, notes: 11, rang_median: 12, ic_bas: [3.67, 4.29] }
  - { hz: 1566.8, residu_db: 4.5, bande_db3: [439.7, 2215.8], verdict: fragile-cr, notes: 17, rang_median: 32, ic_bas: [3.67, 3.84] }
  - { hz: 697.9, residu_db: 4.09, bande_db3: [415.0, 2791.7], verdict: retenu-cr, notes: 15, rang_median: 15, ic_bas: [3.55, 3.28] }
pente_db_octave: -19.1
pente_sd: 2.3
harmoniques_par_note: 79
harmoniques_min_max: [16, 90]
balance_impair_pair_db: -6.0
enveloppe_vraie_hz: 323.0
enveloppe_vraie_sd: 133.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Tuba basse

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 3 formant(s) retenu(s), 1 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.59 · 79 harmoniques par note en médiane · 76 échantillons sur 26 notes distinctes · pente σ = 2.3 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 1244 Hz | +5.8 dB | 932–1974 Hz | 14 | retenu-cr |
| 494 Hz | +4.7 dB | 440–2091 Hz | 11 | retenu-cr |
| 1567 Hz | +4.5 dB | 440–2216 Hz | 17 | fragile-cr |
| 698 Hz | +4.1 dB | 415–2792 Hz | 15 | retenu-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 1244 Hz (k méd. 25) | +5.6 dB [IC +4.8], 52 notes | +7.4 dB [IC +5.1], 16 notes | oui |
| 494 Hz (k méd. 12) | +4.2 dB [IC +3.7], 33 notes | +5.3 dB [IC +4.3], 12 notes | oui |
| 1567 Hz (k méd. 32) | +4.4 dB [IC +3.7], 60 notes | +4.6 dB [IC +3.8], 16 notes | non |
| 698 Hz (k méd. 15) | +4.3 dB [IC +3.5], 39 notes | +4.0 dB [IC +3.3], 14 notes | oui |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -19.1 dB/octave (σ 2.3) |
| Harmoniques par note | 79 [16–90] |
| Balance impair − pair | -6.0 dB |
| Enveloppe vraie (méthode 6) | 323 Hz, σ 133 ; ordres 23–560 |
| Pics tombant sur un harmonique | 9.2% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
