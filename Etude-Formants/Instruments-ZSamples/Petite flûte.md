---
id: petite-flûte
nom: Petite flûte
source_echantillons: ZSamples/flutes/piccolo_flute/ordinario
technique: ordinario
n_echantillons: 52
n_notes_distinctes: 14
f0_min_hz: 587.3
f0_max_hz: 4186.0
controle_nomenclature: 1.006   # doit valoir ~1,00
classe: B        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: non pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.96
formants: []
pente_db_octave: -20.8
pente_sd: 4.5
harmoniques_par_note: 7
harmoniques_min_max: [2, 18]
balance_impair_pair_db: -8.4
enveloppe_vraie_hz: 1647.0
enveloppe_vraie_sd: 1144.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Petite flûte

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 0 formant(s) retenu(s), 0 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.96 · 7 harmoniques par note en médiane · 52 échantillons sur 14 notes distinctes · pente σ = 4.5 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -20.8 dB/octave (σ 4.5) |
| Harmoniques par note | 7 [2–18] |
| Balance impair − pair | -8.4 dB |
| Enveloppe vraie (méthode 6) | 1647 Hz, σ 1144 ; ordres 3–26 |
| Pics tombant sur un harmonique | 5.8% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
