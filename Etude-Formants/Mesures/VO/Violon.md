---
id: violon
nom: Violon
corpus: VO
n_echantillons: 108
rho_hauteur: 0.74
pente_db_octave: -14.8
harmoniques_par_note: 14
balance_impair_pair_db: -3.9
enveloppe_vraie_hz: 1222.0
enveloppe_vraie_sd: 1226.0
nomenclature: 1.0
accord_corrige: true
accord_rapport: 1.0058
classe: A_VALIDER_A_LA_MAIN
formants:
  - { hz: 3971, residu_db: 8.17, verdict: retenu-cr, notes: 13 }
  - { hz: 5616, residu_db: 4.41, verdict: fragile-cr, notes: 17 }
  - { hz: 2104, residu_db: 2.58, verdict: fragile-cr, notes: 6 }
version_mesure: VO-2026-10-25
---

# Violon — corpus VO

> **Méthode 8** (enveloppe composite). Méthode et critères dans `../../Instruments/00-METHODE.md`. Statut des formants dans `../../Instruments/_REPLICATION.md`.

## Formants

| Zone | Résidu | Notes | Verdict |
|---|---|---|---|
| 3971 Hz | 8.17 dB | 13 | retenu-cr |
| 5616 Hz | 4.41 dB | 17 | fragile-cr |
| 2104 Hz | 2.58 dB | 6 | fragile-cr |

## Descripteurs

| Descripteur | Valeur |
|---|---|
| ρ (pic / hauteur) | 0.74 |
| Pente spectrale | -14.8 dB/octave (σ 3.0) |
| Harmoniques par note | 14 |
| Balance impair − pair | -3.9 dB |
| Enveloppe vraie | 1222.0 Hz (σ 1226.0) |

> La **pente spectrale n'est pas comparable entre corpus** : elle dépend du micro et du local. À lire à l'intérieur de ce corpus seulement.

> **Correction d'accord appliquée.** Cette banque est accordée +12050 cents par rapport aux hauteurs de ses noms de fichiers (rapport 1.0058). Sans correction, harmoniques() relève les pics dans un creux au-delà de ~1 300 Hz et l'enveloppe composite est fausse. Le recalage est fait par échantillon et les fréquences ci-dessus sont remises dans le référentiel réel.

## Confrontation aux corpus existants — À RÉDIGER

Comparer chaque formant ci-dessus à la fiche de l'instrument nu (`../../Instruments/FullSOL2020/` ou `Yan_Adds/`) et à ZSamples (`../ZSamples/`). **Ce corpus est une troisième source : il peut faire basculer en établis les dix formants classés « indéterminés » dans `../../Instruments/_REPLICATION.md`.** Route A = même point de grille ; route B = deux familles de méthodes indépendantes. La grille est au demi-ton : écrire « Ré7 (≈ 2 350 Hz) », pas une valeur au Hz près.

## Observations — À RÉDIGER

## Journal

- **2026-10-25** — Mesure initiale, corpus VO, 108 échantillons. 1 formant(s) retenu(s) sur 3 candidat(s). Nomenclature 1.0.
