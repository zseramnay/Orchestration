---
id: trombone-tenor
nom: Trombone tenor
corpus: VO
n_echantillons: 63
rho_hauteur: 0.52
pente_db_octave: -16.2
harmoniques_par_note: 66
balance_impair_pair_db: -1.6
enveloppe_vraie_hz: 506.0
enveloppe_vraie_sd: 218.0
nomenclature: 1.0
accord_corrige: true
accord_rapport: 1.0116
classe: A_VALIDER_A_LA_MAIN
formants:
  - { hz: 1885, residu_db: 4.28, verdict: retenu-cr, notes: 12 }
  - { hz: 840, residu_db: 2.57, verdict: retenu-cr, notes: 11 }
  - { hz: 3770, residu_db: 2.17, verdict: fragile-cr, notes: 14 }
version_mesure: VO-2026-10-25
---

# Trombone tenor — corpus VO

> **Méthode 8** (enveloppe composite). Méthode et critères dans `../../Instruments/00-METHODE.md`. Statut des formants dans `../../Instruments/_REPLICATION.md`.

## Formants

| Zone | Résidu | Notes | Verdict |
|---|---|---|---|
| 1885 Hz | 4.28 dB | 12 | retenu-cr |
| 840 Hz | 2.57 dB | 11 | retenu-cr |
| 3770 Hz | 2.17 dB | 14 | fragile-cr |

## Descripteurs

| Descripteur | Valeur |
|---|---|
| ρ (pic / hauteur) | 0.52 |
| Pente spectrale | -16.2 dB/octave (σ 5.4) |
| Harmoniques par note | 66 |
| Balance impair − pair | -1.6 dB |
| Enveloppe vraie | 506.0 Hz (σ 218.0) |

> La **pente spectrale n'est pas comparable entre corpus** : elle dépend du micro et du local. À lire à l'intérieur de ce corpus seulement.

> **Correction d'accord appliquée.** Cette banque est accordée +24100 cents par rapport aux hauteurs de ses noms de fichiers (rapport 1.0116). Sans correction, harmoniques() relève les pics dans un creux au-delà de ~1 300 Hz et l'enveloppe composite est fausse. Le recalage est fait par échantillon et les fréquences ci-dessus sont remises dans le référentiel réel.

## Confrontation aux corpus existants — À RÉDIGER

Comparer chaque formant ci-dessus à la fiche de l'instrument nu (`../../Instruments/FullSOL2020/` ou `Yan_Adds/`) et à ZSamples (`../ZSamples/`). **Ce corpus est une troisième source : il peut faire basculer en établis les dix formants classés « indéterminés » dans `../../Instruments/_REPLICATION.md`.** Route A = même point de grille ; route B = deux familles de méthodes indépendantes. La grille est au demi-ton : écrire « Ré7 (≈ 2 350 Hz) », pas une valeur au Hz près.

## Observations — À RÉDIGER

## Journal

- **2026-10-25** — Mesure initiale, corpus VO, 63 échantillons. 2 formant(s) retenu(s) sur 3 candidat(s). Nomenclature 1.0.
