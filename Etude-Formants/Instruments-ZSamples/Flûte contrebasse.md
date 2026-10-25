---
id: flûte-contrebasse
nom: Flûte contrebasse
source_echantillons: ZSamples/flutes/contrabass_flute/ordinario
technique: ordinario
n_echantillons: 52
n_notes_distinctes: 13
f0_min_hz: 98.0
f0_max_hz: 784.0
controle_nomenclature: 1.006   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: partielle   # À VALIDER À LA MAIN
rho_hauteur: 0.82
formants:
  - { hz: 3517.3, residu_db: 4.11, bande_db3: [1566.8, 4431.5], verdict: fragile-cr, notes: 8, rang_median: 21, ic_bas: [4.61, 0.91] }
  - { hz: 2487.1, residu_db: 3.98, bande_db3: [1566.8, 4431.5], verdict: retenu-cr, notes: 6, rang_median: 14, ic_bas: [3.85, 2.03] }
pente_db_octave: -16.8
pente_sd: 2.8
harmoniques_par_note: 39
harmoniques_min_max: [14, 90]
balance_impair_pair_db: -5.5
enveloppe_vraie_hz: 296.0
enveloppe_vraie_sd: 271.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Flûte contrebasse

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 1 formant(s) retenu(s), 1 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.82 · 39 harmoniques par note en médiane · 52 échantillons sur 13 notes distinctes · pente σ = 2.8 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 3517 Hz | +4.1 dB | 1567–4432 Hz | 8 | fragile-cr |
| 2487 Hz | +4.0 dB | 1567–4432 Hz | 6 | retenu-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 3517 Hz (k méd. 21) | +5.1 dB [IC +4.6], 40 notes | +2.6 dB [IC +0.9], 16 notes | non |
| 2487 Hz (k méd. 14) | +4.6 dB [IC +3.9], 40 notes | +3.5 dB [IC +2.0], 16 notes | oui |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -16.8 dB/octave (σ 2.8) |
| Harmoniques par note | 39 [14–90] |
| Balance impair − pair | -5.5 dB |
| Enveloppe vraie (méthode 6) | 296 Hz, σ 271 ; ordres 19–157 |
| Pics tombant sur un harmonique | 11.5% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
