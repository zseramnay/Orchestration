---
id: violon
nom: Violon
corpus: ZSamples
n_echantillons: 350
rho_hauteur: 0.81
pente_db_octave: -12.7
harmoniques_par_note: 16
balance_impair_pair_db: -4.3
enveloppe_vraie_hz: 899.0
enveloppe_vraie_sd: 959.0
nomenclature: 1.003
accord_corrige: false
accord_rapport: 1.0
classe: A_VALIDER_A_LA_MAIN
formants:
  - { hz: 3948, residu_db: 4.42, verdict: retenu-cr, notes: 13 }
  - { hz: 5583, residu_db: 3.16, verdict: fragile-cr, notes: 18 }
  - { hz: 2792, residu_db: 2.63, verdict: retenu-cr, notes: 9 }
  - { hz: 2216, residu_db: 1.76, verdict: rejeté, notes: 7 }
version_mesure: ZSamples-2026-10-25
---

# Violon — corpus ZSamples

> **Méthode 8** (enveloppe composite). Méthode et critères dans `../00-METHODE.md`. Statut des formants dans `../_REPLICATION.md`.

## Formants

| Zone | Résidu | Notes | Verdict |
|---|---|---|---|
| 3948 Hz | 4.42 dB | 13 | retenu-cr |
| 5583 Hz | 3.16 dB | 18 | fragile-cr |
| 2792 Hz | 2.63 dB | 9 | retenu-cr |
| 2216 Hz | 1.76 dB | 7 | rejeté |

## Descripteurs

| Descripteur | Valeur |
|---|---|
| ρ (pic / hauteur) | 0.81 |
| Pente spectrale | -12.7 dB/octave (σ 3.6) |
| Harmoniques par note | 16 |
| Balance impair − pair | -4.3 dB |
| Enveloppe vraie | 899.0 Hz (σ 959.0) |

> La **pente spectrale n'est pas comparable entre corpus** : elle dépend du micro et du local. À lire à l'intérieur de ce corpus seulement.

## Confrontation aux corpus existants — À RÉDIGER

Comparer chaque formant ci-dessus à la fiche de l'instrument nu (`../NN-*.md`) et à ZSamples (`../../Instruments-ZSamples/`). **Ce corpus est une troisième source : il peut faire basculer en établis les dix formants classés « indéterminés » dans `../_REPLICATION.md`.** Route A = même point de grille ; route B = deux familles de méthodes indépendantes. La grille est au demi-ton : écrire « Ré7 (≈ 2 350 Hz) », pas une valeur au Hz près.

## Observations — À RÉDIGER

## Journal

- **2026-10-25** — Mesure initiale, corpus ZSamples, 350 échantillons. 2 formant(s) retenu(s) sur 4 candidat(s). Nomenclature 1.003.
