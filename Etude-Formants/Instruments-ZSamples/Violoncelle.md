---
id: violoncelle
nom: Violoncelle
source_echantillons: ZSamples/violoncello/ordinario
technique: ordinario
n_echantillons: 256
n_notes_distinctes: 38
f0_min_hz: 65.4
f0_max_hz: 1318.5
controle_nomenclature: 1.006   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.84
formants:
  - { hz: 3517.3, residu_db: 5.42, bande_db3: [1974.0, 4431.5], verdict: retenu-cr, notes: 26, rang_median: 24, ic_bas: [2.81, 5.36] }
  - { hz: 2791.7, residu_db: 4.92, bande_db3: [1974.0, 4431.5], verdict: retenu-cr, notes: 24, rang_median: 19, ic_bas: [3.45, 6.05] }
  - { hz: 621.8, residu_db: 3.94, bande_db3: [553.9, 658.8], verdict: fragile-cr, notes: 7, rang_median: 5, ic_bas: [3.65, 3.26] }
  - { hz: 931.6, residu_db: 2.5, bande_db3: [553.9, 5915.4], verdict: retenu-cr, notes: 10, rang_median: 7, ic_bas: [3.0, 0.87] }
pente_db_octave: -13.4
pente_sd: 2.0
harmoniques_par_note: 45
harmoniques_min_max: [8, 90]
balance_impair_pair_db: -3.7
enveloppe_vraie_hz: 291.0
enveloppe_vraie_sd: 446.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Violoncelle

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 3 formant(s) retenu(s), 1 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.84 · 45 harmoniques par note en médiane · 256 échantillons sur 38 notes distinctes · pente σ = 2.0 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 3517 Hz | +5.4 dB | 1974–4432 Hz | 26 | retenu-cr |
| 2792 Hz | +4.9 dB | 1974–4432 Hz | 24 | retenu-cr |
| 622 Hz | +3.9 dB | 554–659 Hz | 7 | fragile-cr |
| 932 Hz | +2.5 dB | 554–5915 Hz | 10 | retenu-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 3517 Hz (k méd. 24) | +3.4 dB [IC +2.8], 205 notes | +5.8 dB [IC +5.4], 84 notes | oui |
| 2792 Hz (k méd. 19) | +3.8 dB [IC +3.5], 196 notes | +6.6 dB [IC +6.0], 84 notes | oui |
| 622 Hz (k méd. 5) | +4.3 dB [IC +3.6], 80 notes | +4.2 dB [IC +3.3], 40 notes | non |
| 932 Hz (k méd. 7) | +4.1 dB [IC +3.0], 123 notes | +2.4 dB [IC +0.9], 48 notes | oui |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -13.4 dB/octave (σ 2.0) |
| Harmoniques par note | 45 [8–90] |
| Balance impair − pair | -3.7 dB |
| Enveloppe vraie (méthode 6) | 291 Hz, σ 446 ; ordres 11–235 |
| Pics tombant sur un harmonique | 11.7% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
