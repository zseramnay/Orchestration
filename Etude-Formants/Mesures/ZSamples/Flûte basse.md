---
id: flûte-basse
nom: Flûte basse
source_echantillons: ZSamples/flutes/bass_flute/ordinario
technique: ordinario
n_echantillons: 52
n_notes_distinctes: 13
f0_min_hz: 130.8
f0_max_hz: 1046.5
controle_nomenclature: 1.006   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.93
formants:
  - { hz: 2791.7, residu_db: 4.41, bande_db3: [2091.4, 4431.5], verdict: retenu-cr, notes: 5, rang_median: 13, ic_bas: [3.66, 3.43] }
  - { hz: 3726.5, residu_db: 3.94, bande_db3: [2091.4, 4431.5], verdict: retenu-cr, notes: 10, rang_median: 17, ic_bas: [4.03, 1.95] }
pente_db_octave: -15.4
pente_sd: 1.8
harmoniques_par_note: 29
harmoniques_min_max: [10, 83]
balance_impair_pair_db: -6.2
enveloppe_vraie_hz: 382.0
enveloppe_vraie_sd: 281.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Flûte basse

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 2 formant(s) retenu(s), 0 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.93 · 29 harmoniques par note en médiane · 52 échantillons sur 13 notes distinctes · pente σ = 1.8 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 2792 Hz | +4.4 dB | 2091–4432 Hz | 5 | retenu-cr |
| 3726 Hz | +3.9 dB | 2091–4432 Hz | 10 | retenu-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 2792 Hz (k méd. 13) | +4.4 dB [IC +3.7], 36 notes | +4.1 dB [IC +3.4], 16 notes | oui |
| 3726 Hz (k méd. 17) | +5.0 dB [IC +4.0], 36 notes | +3.0 dB [IC +1.9], 16 notes | oui |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -15.4 dB/octave (σ 1.8) |
| Harmoniques par note | 29 [10–83] |
| Balance impair − pair | -6.2 dB |
| Enveloppe vraie (méthode 6) | 382 Hz, σ 281 ; ordres 14–117 |
| Pics tombant sur un harmonique | 28.8% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
