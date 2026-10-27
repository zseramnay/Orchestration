---
id: clarinette-en-sib
nom: Clarinette en Sib
source_echantillons: ZSamples/Clarinets/bb_clarinet_boehm_system/ordinario
technique: ordinario
n_echantillons: 64
n_notes_distinctes: 16
f0_min_hz: 146.8
f0_max_hz: 1864.7
controle_nomenclature: 1.006   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 1.0
formants:
  - { hz: 5583.4, residu_db: 5.18, bande_db3: [3133.6, 5915.4], verdict: retenu-cr, notes: 10, rang_median: 23, ic_bas: [4.8, 4.88] }
  - { hz: 3517.3, residu_db: 5.0, bande_db3: [2957.7, 5915.4], verdict: retenu-cr, notes: 10, rang_median: 14, ic_bas: [4.85, 2.8] }
pente_db_octave: -21.1
pente_sd: 7.1
harmoniques_par_note: 20
harmoniques_min_max: [5, 74]
balance_impair_pair_db: 0.6
enveloppe_vraie_hz: 549.0
enveloppe_vraie_sd: 647.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Clarinette en Sib

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 2 formant(s) retenu(s), 0 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 1.0 · 20 harmoniques par note en médiane · 64 échantillons sur 16 notes distinctes · pente σ = 7.1 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 5583 Hz | +5.2 dB | 3134–5915 Hz | 10 | retenu-cr |
| 3517 Hz | +5.0 dB | 2958–5915 Hz | 10 | retenu-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 5583 Hz (k méd. 23) | +5.6 dB [IC +4.8], 52 notes | +6.3 dB [IC +4.9], 16 notes | oui |
| 3517 Hz (k méd. 14) | +6.6 dB [IC +4.8], 36 notes | +3.8 dB [IC +2.8], 16 notes | oui |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -21.1 dB/octave (σ 7.1) |
| Harmoniques par note | 20 [5–74] |
| Balance impair − pair | 0.6 dB |
| Enveloppe vraie (méthode 6) | 549 Hz, σ 647 ; ordres 8–105 |
| Pics tombant sur un harmonique | 26.6% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
