---
id: tuba-basse
nom: Tuba basse
corpus: ZSamples
n_echantillons: 127
rho_hauteur: 0.47
pente_db_octave: -18.3
harmoniques_par_note: 79
balance_impair_pair_db: -4.7
enveloppe_vraie_hz: 345.0
enveloppe_vraie_sd: 128.0
nomenclature: 1.003
accord_corrige: false
accord_rapport: 1.0
classe: A_VALIDER_A_LA_MAIN
formants:
  - { hz: 466, residu_db: 6.06, verdict: retenu-cr, notes: 7 }
  - { hz: 1318, residu_db: 5.7, verdict: retenu-cr, notes: 18 }
  - { hz: 659, residu_db: 3.97, verdict: retenu-cr, notes: 11 }
  - { hz: 2792, residu_db: 1.19, verdict: rejeté, notes: 21 }
version_mesure: ZSamples-2026-10-25
---

# Tuba basse — corpus ZSamples

> **Méthode 8** (enveloppe composite). Méthode et critères dans `../00-METHODE.md`. Statut des formants dans `../_REPLICATION.md`.

## Formants

| Zone | Résidu | Notes | Verdict |
|---|---|---|---|
| 466 Hz | 6.06 dB | 7 | retenu-cr |
| 1318 Hz | 5.7 dB | 18 | retenu-cr |
| 659 Hz | 3.97 dB | 11 | retenu-cr |
| 2792 Hz | 1.19 dB | 21 | rejeté |

## Descripteurs

| Descripteur | Valeur |
|---|---|
| ρ (pic / hauteur) | 0.47 |
| Pente spectrale | -18.3 dB/octave (σ 2.4) |
| Harmoniques par note | 79 |
| Balance impair − pair | -4.7 dB |
| Enveloppe vraie | 345.0 Hz (σ 128.0) |

> La **pente spectrale n'est pas comparable entre corpus** : elle dépend du micro et du local. À lire à l'intérieur de ce corpus seulement.

## Confrontation aux corpus existants — À RÉDIGER

Comparer chaque formant ci-dessus à la fiche de l'instrument nu (`../NN-*.md`) et à ZSamples (`../../Instruments-ZSamples/`). **Ce corpus est une troisième source : il peut faire basculer en établis les dix formants classés « indéterminés » dans `../_REPLICATION.md`.** Route A = même point de grille ; route B = deux familles de méthodes indépendantes. La grille est au demi-ton : écrire « Ré7 (≈ 2 350 Hz) », pas une valeur au Hz près.

## Observations — À RÉDIGER

## Journal

- **2026-10-25** — Mesure initiale, corpus ZSamples, 127 échantillons. 3 formant(s) retenu(s) sur 4 candidat(s). Nomenclature 1.003.
