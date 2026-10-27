---
id: cor-sourdine-sèche
nom: Cor+sourdine sèche
source_echantillons: ZSamples/horns/horn_in_f/straight_mute
technique: straight_mute
n_echantillons: 48
n_notes_distinctes: 17
f0_min_hz: 43.7
f0_max_hz: 587.3
controle_nomenclature: 1.003   # doit valoir ~1,00
classe: A        # A/B/C — À VALIDER À LA MAIN
pertinence_formant: pertinente   # À VALIDER À LA MAIN
rho_hauteur: 0.42
formants:
  - { hz: 1758.7, residu_db: 4.58, bande_db3: [931.6, 2215.8], verdict: retenu-cr, notes: 14, rang_median: 24, ic_bas: [3.72, 3.12] }
  - { hz: 3948.1, residu_db: 3.58, bande_db3: [931.6, 5270.0], verdict: fragile-cr, notes: 15, rang_median: 48, ic_bas: [6.64, -0.81] }
  - { hz: 1045.7, residu_db: 3.48, bande_db3: [931.6, 5270.0], verdict: retenu-cr, notes: 11, rang_median: 14, ic_bas: [2.95, 1.42] }
pente_db_octave: -15.9
pente_sd: 4.8
harmoniques_par_note: 68
harmoniques_min_max: [18, 90]
balance_impair_pair_db: -2.6
enveloppe_vraie_hz: 388.0
enveloppe_vraie_sd: 211.0
corpus_suffisant: false
version_mesure: 2026-10-24-wav-v1
---

# Cor+sourdine sèche

## Méthode retenue, et pourquoi

> **Méthode 8 — enveloppe composite.** 2 formant(s) retenu(s), 1 fragile(s), 0 rejeté(s). Voir `00-METHODE.md`.

À RÉDIGER. Éléments de fait : ρ = 0.42 · 68 harmoniques par note en médiane · 48 échantillons sur 17 notes distinctes · pente σ = 4.8 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes | Verdict |
|---|---|---|---|---|
| 1759 Hz | +4.6 dB | 932–2216 Hz | 14 | retenu-cr |
| 3948 Hz | +3.6 dB | 932–5270 Hz | 15 | fragile-cr |
| 1046 Hz | +3.5 dB | 932–5270 Hz | 11 | retenu-cr |

## Tests de rejet

| Zone | rang ≤ médian | rang > médian | Robustesse |
|---|---|---|---|
| 1759 Hz (k méd. 24) | +4.6 dB [IC +3.7], 36 notes | +5.2 dB [IC +3.1], 12 notes | oui |
| 3948 Hz (k méd. 48) | +7.9 dB [IC +6.6], 38 notes | -0.1 dB [IC -0.8], 14 notes | oui |
| 1046 Hz (k méd. 14) | +3.4 dB [IC +3.0], 27 notes | +2.2 dB [IC +1.4], 12 notes | oui |

## Descripteurs globaux

| Descripteur | Valeur |
|---|---|
| Pente spectrale médiane | -15.9 dB/octave (σ 4.8) |
| Harmoniques par note | 68 [18–90] |
| Balance impair − pair | -2.6 dB |
| Enveloppe vraie (méthode 6) | 388 Hz, σ 211 ; ordres 26–353 |
| Pics tombant sur un harmonique | 6.2% |

## Comparaison avec la littérature

À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), Meyer (2009), et comparer aux valeurs de l'ancien pipeline.

## Observations pour l'orchestration

À RÉDIGER.

## Journal

- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).
