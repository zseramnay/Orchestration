---
id: violoncelle
nom: Violoncelle
corpus: VN
n_echantillons: 159
rho_hauteur: 0.89
pente_db_octave: -13.2
harmoniques_par_note: 37
balance_impair_pair_db: -3.2
enveloppe_vraie_hz: 366.0
enveloppe_vraie_sd: 647.0
nomenclature: 1.0
accord_corrige: false
accord_rapport: 1.0
classe: A_VALIDER_A_LA_MAIN
formants:
  - { hz: 1108, residu_db: 5.95, verdict: retenu-cr, notes: 16 }
  - { hz: 3948, residu_db: 5.05, verdict: retenu-cr, notes: 37 }
  - { hz: 2635, residu_db: 4.34, verdict: retenu-cr, notes: 30 }
  - { hz: 5583, residu_db: 1.88, verdict: rejeté, notes: 41 }
version_mesure: VN-2026-10-25
---

# Violoncelle — corpus VN

> **Méthode 8** (enveloppe composite). Méthode et critères dans `../00-METHODE.md`. Statut des formants dans `../_REPLICATION.md`.

## Formants

| Zone | Résidu | Notes | Verdict |
|---|---|---|---|
| 1108 Hz | 5.95 dB | 16 | retenu-cr |
| 3948 Hz | 5.05 dB | 37 | retenu-cr |
| 2635 Hz | 4.34 dB | 30 | retenu-cr |
| 5583 Hz | 1.88 dB | 41 | rejeté |

## Descripteurs

| Descripteur | Valeur |
|---|---|
| ρ (pic / hauteur) | 0.89 |
| Pente spectrale | -13.2 dB/octave (σ 1.6) |
| Harmoniques par note | 37 |
| Balance impair − pair | -3.2 dB |
| Enveloppe vraie | 366.0 Hz (σ 647.0) |

> La **pente spectrale n'est pas comparable entre corpus** : elle dépend du micro et du local. À lire à l'intérieur de ce corpus seulement.

## Confrontation aux corpus existants — À RÉDIGER

Comparer chaque formant ci-dessus à la fiche de l'instrument nu (`../NN-*.md`) et à ZSamples (`../../Instruments-ZSamples/`). **Ce corpus est une troisième source : il peut faire basculer en établis les dix formants classés « indéterminés » dans `../_REPLICATION.md`.** Route A = même point de grille ; route B = deux familles de méthodes indépendantes. La grille est au demi-ton : écrire « Ré7 (≈ 2 350 Hz) », pas une valeur au Hz près.

## Observations — À RÉDIGER

## Journal

- **2026-10-25** — Mesure initiale, corpus VN, 159 échantillons. 3 formant(s) retenu(s) sur 4 candidat(s). Nomenclature 1.0.
