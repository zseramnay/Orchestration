---
id: alto
nom: Alto
corpus: VN
n_echantillons: 132
rho_hauteur: 0.85
pente_db_octave: -13.7
harmoniques_par_note: 23
balance_impair_pair_db: -3.9
enveloppe_vraie_hz: 528.0
enveloppe_vraie_sd: 619.0
nomenclature: 1.0
accord_corrige: false
accord_rapport: 1.0
classe: A_VALIDER_A_LA_MAIN
formants:
  - { hz: 4432, residu_db: 5.27, verdict: retenu-cr, notes: 27 }
  - { hz: 1759, residu_db: 4.08, verdict: retenu-cr, notes: 12 }
  - { hz: 2635, residu_db: 2.2, verdict: fragile-cr, notes: 18 }
version_mesure: VN-2026-10-25
---

# Alto — corpus VN

> **Méthode 8** (enveloppe composite). Méthode et critères dans `../../Instruments/00-METHODE.md`. Statut des formants dans `../../Instruments/_REPLICATION.md`.

## Formants

| Zone | Résidu | Notes | Verdict |
|---|---|---|---|
| 4432 Hz | 5.27 dB | 27 | retenu-cr |
| 1759 Hz | 4.08 dB | 12 | retenu-cr |
| 2635 Hz | 2.2 dB | 18 | fragile-cr |

## Descripteurs

| Descripteur | Valeur |
|---|---|
| ρ (pic / hauteur) | 0.85 |
| Pente spectrale | -13.7 dB/octave (σ 3.5) |
| Harmoniques par note | 23 |
| Balance impair − pair | -3.9 dB |
| Enveloppe vraie | 528.0 Hz (σ 619.0) |

> La **pente spectrale n'est pas comparable entre corpus** : elle dépend du micro et du local. À lire à l'intérieur de ce corpus seulement.

## Confrontation aux corpus existants — À RÉDIGER

Comparer chaque formant ci-dessus à la fiche de l'instrument nu (`../../Instruments/FullSOL2020/` ou `Yan_Adds/`) et à ZSamples (`../ZSamples/`). **Ce corpus est une troisième source : il peut faire basculer en établis les dix formants classés « indéterminés » dans `../../Instruments/_REPLICATION.md`.** Route A = même point de grille ; route B = deux familles de méthodes indépendantes. La grille est au demi-ton : écrire « Ré7 (≈ 2 350 Hz) », pas une valeur au Hz près.

## Observations — À RÉDIGER

## Journal

- **2026-10-25** — Mesure initiale, corpus VN, 132 échantillons. 2 formant(s) retenu(s) sur 3 candidat(s). Nomenclature 1.0.
