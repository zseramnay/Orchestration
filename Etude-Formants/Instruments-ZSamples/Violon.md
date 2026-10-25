---
id: violon
nom: Violon
source_echantillons: ZSamples/violin/ordinario
technique: ordinario
n_echantillons: 164
n_notes_distinctes: 36
f0_min_hz: 196.0
f0_max_hz: 3322.4
controle_nomenclature: 1.003   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.9
formants:
  - { hz: 3948.1, residu_db: 6.81, bande_db3: [3517.3, 4431.5], verdict: retenu-cr, notes: 13, rang_median: 9, ic_bas: [4.93, 6.58] }
  - { hz: 5583.4, residu_db: 4.14, bande_db3: [3133.6, 5915.4], verdict: retenu-cr, notes: 18, rang_median: 13, ic_bas: [2.03, 4.32] }
  - { hz: 2791.7, residu_db: 2.98, bande_db3: [2091.4, 5915.4], verdict: retenu-cr, notes: 9, rang_median: 7, ic_bas: [0.41, 2.27] }
  - { hz: 2215.8, residu_db: 2.32, bande_db3: [2091.4, 5915.4], verdict: retenu-cr, notes: 7, rang_median: 6, ic_bas: [0.49, 0.43] }
pente_db_octave: -14.5
pente_sd: 2.8
harmoniques_par_note: 15
harmoniques_min_max: [3, 56]
balance_impair_pair_db: -4.2
enveloppe_vraie_hz: 1066.0
enveloppe_vraie_sd: 1031.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Violon

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 4 formant(s) retenu(s), 0 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.9 · 15 harmoniques par note en médiane · 164 échantillons sur 36 notes distinctes · pente σ = 2.8 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 3948 Hz | +6.8 dB | 3517–4432 Hz | 13 | retenu-cr |
| 5583 Hz | +4.1 dB | 3134–5915 Hz | 18 | retenu-cr |
| 2792 Hz | +3.0 dB | 2091–5915 Hz | 9 | retenu-cr |
| 2216 Hz | +2.3 dB | 2091–5915 Hz | 7 | retenu-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 3948 Hz (k méd. 9) | +5.9 dB [IC +4.9], 100 notes | +7.1 dB [IC +6.6], 44 notes | oui |
| 5583 Hz (k méd. 13) | +2.4 dB [IC +2.0], 117 notes | +5.4 dB [IC +4.3], 44 notes | oui |
| 2792 Hz (k méd. 7) | +1.2 dB [IC +0.4], 79 notes | +3.0 dB [IC +2.3], 32 notes | oui |
| 2216 Hz (k méd. 6) | +1.2 dB [IC +0.5], 66 notes | +1.7 dB [IC +0.4], 28 notes | oui |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -14.5 dB/octave (σ 2.8) |
| Harmoniques par note | 15 [3–56] |
| Balance impair − pair | -4.2 dB |
| Enveloppe vraie (méthode 6) | 1066 Hz, σ 1031 ; ordres 4–78 |
| Pics tombant sur un harmonique | 15.2% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
