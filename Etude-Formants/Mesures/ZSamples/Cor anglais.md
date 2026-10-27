---
id: cor-anglais
nom: Cor anglais
source_echantillons: ZSamples/oboes/english_horn/ordinario
technique: ordinario
n_echantillons: 48
n_notes_distinctes: 12
f0_min_hz: 164.8
f0_max_hz: 1046.5
controle_nomenclature: 1.006   # doit valoir ~1,00
classe: A        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: non pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.45
formants: []
pente_db_octave: -19.3
pente_sd: 4.0
harmoniques_par_note: 25
harmoniques_min_max: [10, 66]
balance_impair_pair_db: -5.6
enveloppe_vraie_hz: 894.0
enveloppe_vraie_sd: 259.0
corpus_suffisant: false
version_mesure: 2026-10-24-wav-v1
---

# Cor anglais

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 0 formant(s) retenu(s), 0 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.45 · 25 harmoniques par note en médiane · 48 échantillons sur 12 notes distinctes · pente σ = 4.0 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -19.3 dB/octave (σ 4.0) |
| Harmoniques par note | 25 [10–66] |
| Balance impair − pair | -5.6 dB |
| Enveloppe vraie (méthode 6) | 894 Hz, σ 259 ; ordres 14–93 |
| Pics tombant sur un harmonique | 2.1% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
