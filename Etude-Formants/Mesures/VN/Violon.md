---
id: violon
nom: Violon
corpus: VN
n_echantillons: 153
rho_hauteur: 0.9
pente_db_octave: -12.6
harmoniques_par_note: 13
balance_impair_pair_db: -2.4
enveloppe_vraie_hz: 1044.0
enveloppe_vraie_sd: 1050.0
nomenclature: 1.0
accord_corrige: false
accord_rapport: 1.0
classe: A_VALIDER_A_LA_MAIN
formants:
  - { hz: 4183, residu_db: 5.21, verdict: retenu-cr, notes: 19 }
  - { hz: 2216, residu_db: 3.15, verdict: retenu-cr, notes: 9 }
  - { hz: 1479, residu_db: 1.24, verdict: rejeté, notes: 5 }
version_mesure: VN-2026-10-25
---

# Violon — corpus VN

> **Méthode 8** (enveloppe composite). Méthode et critères dans `../../Instruments/00-METHODE.md`. Statut des formants dans `../../Instruments/_REPLICATION.md`.

## Formants

| Zone | Résidu | Notes | Verdict |
|---|---|---|---|
| 4183 Hz | 5.21 dB | 19 | retenu-cr |
| 2216 Hz | 3.15 dB | 9 | retenu-cr |
| 1479 Hz | 1.24 dB | 5 | rejeté |

## Descripteurs

| Descripteur | Valeur |
|---|---|
| ρ (pic / hauteur) | 0.9 |
| Pente spectrale | -12.6 dB/octave (σ 3.2) |
| Harmoniques par note | 13 |
| Balance impair − pair | -2.4 dB |
| Enveloppe vraie | 1044.0 Hz (σ 1050.0) |

> La **pente spectrale n'est pas comparable entre corpus** : elle dépend du micro et du local. À lire à l'intérieur de ce corpus seulement.

## Confrontation aux corpus existants — À RÉDIGER

Comparer chaque formant ci-dessus à la fiche de l'instrument nu (`../../Instruments/FullSOL2020/` ou `Yan_Adds/`) et à ZSamples (`../ZSamples/`). **Ce corpus est une troisième source : il peut faire basculer en établis les dix formants classés « indéterminés » dans `../../Instruments/_REPLICATION.md`.** Route A = même point de grille ; route B = deux familles de méthodes indépendantes. La grille est au demi-ton : écrire « Ré7 (≈ 2 350 Hz) », pas une valeur au Hz près.

## Observations — À RÉDIGER

## Journal

- **2026-10-25** — Mesure initiale, corpus VN, 153 échantillons. 2 formant(s) retenu(s) sur 3 candidat(s). Nomenclature 1.0.
