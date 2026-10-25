---
id: contrebasson
nom: Contrebasson
corpus: VN
n_echantillons: 111
rho_hauteur: 0.42
pente_db_octave: -14.0
harmoniques_par_note: 90
balance_impair_pair_db: -2.0
enveloppe_vraie_hz: 301.0
enveloppe_vraie_sd: 101.0
nomenclature: 1.0
accord_corrige: false
accord_rapport: 1.0
classe: A_VALIDER_A_LA_MAIN
formants:
  - { hz: 698, residu_db: 5.91, verdict: retenu-cr, notes: 22 }
  - { hz: 329, residu_db: 4.58, verdict: retenu-cr, notes: 8 }
  - { hz: 1174, residu_db: 4.14, verdict: retenu-cr, notes: 30 }
  - { hz: 4432, residu_db: 2.58, verdict: rejeté, notes: 28 }
version_mesure: VN-2026-10-25
---

# Contrebasson — corpus VN

> **Méthode 8** (enveloppe composite). Méthode et critères dans `../00-METHODE.md`. Statut des formants dans `../_REPLICATION.md`.

## Formants

| Zone | Résidu | Notes | Verdict |
|---|---|---|---|
| 698 Hz | 5.91 dB | 22 | retenu-cr |
| 329 Hz | 4.58 dB | 8 | retenu-cr |
| 1174 Hz | 4.14 dB | 30 | retenu-cr |
| 4432 Hz | 2.58 dB | 28 | rejeté |

## Descripteurs

| Descripteur | Valeur |
|---|---|
| ρ (pic / hauteur) | 0.42 |
| Pente spectrale | -14.0 dB/octave (σ 2.2) |
| Harmoniques par note | 90 |
| Balance impair − pair | -2.0 dB |
| Enveloppe vraie | 301.0 Hz (σ 101.0) |

> La **pente spectrale n'est pas comparable entre corpus** : elle dépend du micro et du local. À lire à l'intérieur de ce corpus seulement.

## Confrontation aux corpus existants — À RÉDIGER

Comparer chaque formant ci-dessus à la fiche de l'instrument nu (`../NN-*.md`) et à ZSamples (`../../Instruments-ZSamples/`). **Ce corpus est une troisième source : il peut faire basculer en établis les dix formants classés « indéterminés » dans `../_REPLICATION.md`.** Route A = même point de grille ; route B = deux familles de méthodes indépendantes. La grille est au demi-ton : écrire « Ré7 (≈ 2 350 Hz) », pas une valeur au Hz près.

## Observations — À RÉDIGER

## Journal

- **2026-10-25** — Mesure initiale, corpus VN, 111 échantillons. 3 formant(s) retenu(s) sur 4 candidat(s). Nomenclature 1.0.
