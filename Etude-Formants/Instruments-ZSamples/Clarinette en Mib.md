---
id: clarinette-en-mib
nom: Clarinette en Mib
source_echantillons: ZSamples/Clarinets/eb_clarinet_boehm_system/ordinario
technique: ordinario
n_echantillons: 60
n_notes_distinctes: 16
f0_min_hz: 196.0
f0_max_hz: 2093.0
controle_nomenclature: 1.009   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: partielle   # À VALIDER À LA MAIN
rho_hauteur: 0.98
formants:
  - { hz: 5270.0, residu_db: 4.27, bande_db3: [3726.5, 5915.4], verdict: retenu-cr, notes: 7, rang_median: 16, ic_bas: [4.0, 0.84] }
pente_db_octave: -23.9
pente_sd: 7.0
harmoniques_par_note: 16
harmoniques_min_max: [5, 56]
balance_impair_pair_db: 0.3
enveloppe_vraie_hz: 732.0
enveloppe_vraie_sd: 795.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Clarinette en Mib

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 1 formant(s) retenu(s), 0 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.98 · 16 harmoniques par note en médiane · 60 échantillons sur 16 notes distinctes · pente σ = 7.0 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 5270 Hz | +4.3 dB | 3726–5915 Hz | 7 | retenu-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 5270 Hz (k méd. 16) | +6.0 dB [IC +4.0], 44 notes | +2.5 dB [IC +0.8], 16 notes | oui |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -23.9 dB/octave (σ 7.0) |
| Harmoniques par note | 16 [5–56] |
| Balance impair − pair | 0.3 dB |
| Enveloppe vraie (méthode 6) | 732 Hz, σ 795 ; ordres 7–78 |
| Pics tombant sur un harmonique | 18.3% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
