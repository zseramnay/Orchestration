---
id: trombone-ténor-wah-fermé
nom: Trombone ténor+wah fermé
source_echantillons: ZSamples/trombones/tenor_trombone/wahwah_stem_in_closed
technique: wahwah_stem_in_closed
n_echantillons: 54
n_notes_distinctes: 22
f0_min_hz: 30.9
f0_max_hz: 523.3
controle_nomenclature: 1.003   # doit valoir ~1,00
classe: C        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.7
formants:
  - { hz: 1173.8, residu_db: 7.1, bande_db3: [1107.9, 1317.5], verdict: retenu-cr, notes: 13, rang_median: 21, ic_bas: [6.37, 8.94] }
  - { hz: 2091.4, residu_db: 2.51, bande_db3: [522.9, 2487.1], verdict: retenu-cr, notes: 18, rang_median: 37, ic_bas: [4.06, 0.5] }
  - { hz: 3517.3, residu_db: 2.42, bande_db3: [2957.7, 4431.5], verdict: fragile-cr, notes: 17, rang_median: 54, ic_bas: [3.38, -0.01] }
  - { hz: 783.4, residu_db: 1.78, bande_db3: [522.9, 4431.5], verdict: rejeté, notes: 14, rang_median: 14, ic_bas: [0.0, 1.37] }
pente_db_octave: -14.2
pente_sd: 3.8
harmoniques_par_note: 86
harmoniques_min_max: [20, 90]
balance_impair_pair_db: -2.2
enveloppe_vraie_hz: 420.0
enveloppe_vraie_sd: 346.0
corpus_suffisant: true
version_mesure: 2026-10-24-wav-v1
---

# Trombone ténor+wah fermé

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 2 formant(s) retenu(s), 1 fragile(s), 1 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.7 · 86 harmoniques par note en médiane · 54 échantillons sur 22 notes distinctes · pente σ = 3.8 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 1174 Hz | +7.1 dB | 1108–1318 Hz | 13 | retenu-cr |
| 2091 Hz | +2.5 dB | 523–2487 Hz | 18 | retenu-cr |
| 3517 Hz | +2.4 dB | 2958–4432 Hz | 17 | fragile-cr |
| 783 Hz | +1.8 dB | 523–4432 Hz | 14 | rejeté |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 1174 Hz (k méd. 21) | +8.7 dB [IC +6.4], 39 notes | +10.4 dB [IC +8.9], 15 notes | oui |
| 2091 Hz (k méd. 37) | +4.7 dB [IC +4.1], 45 notes | +1.2 dB [IC +0.5], 15 notes | oui |
| 3517 Hz (k méd. 54) | +4.0 dB [IC +3.4], 42 notes | +0.4 dB [IC -0.0], 15 notes | oui |
| 783 Hz (k méd. 14) | +1.1 dB [IC +0.0], 30 notes | +2.0 dB [IC +1.4], 14 notes | non |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -14.2 dB/octave (σ 3.8) |
| Harmoniques par note | 86 [20–90] |
| Balance impair − pair | -2.2 dB |
| Enveloppe vraie (méthode 6) | 420 Hz, σ 346 ; ordres 29–499 |
| Pics tombant sur un harmonique | 3.7% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
