---
id: clarinette-en-sib
nom: Clarinette en Sib
corpus: VO
n_echantillons: 54
rho_hauteur: 0.82
pente_db_octave: -17.8
harmoniques_par_note: 16
balance_impair_pair_db: 1.0
enveloppe_vraie_hz: 910.0
enveloppe_vraie_sd: 977.0
nomenclature: 1.0
accord_corrige: true
accord_rapport: 1.0087
classe: A_VALIDER_A_LA_MAIN
formants:
  - { hz: 5316, residu_db: 5.21, verdict: retenu-cr, notes: 9 }
  - { hz: 2983, residu_db: 4.93, verdict: retenu-cr, notes: 9 }
version_mesure: VO-2026-10-25
---

# Clarinette en Sib — corpus VO

> **Méthode 8** (enveloppe composite). Méthode et critères dans `../../Instruments/00-METHODE.md`. Statut des formants dans `../../Instruments/_REPLICATION.md`.

## Formants

| Zone | Résidu | Notes | Verdict |
|---|---|---|---|
| 5316 Hz | 5.21 dB | 9 | retenu-cr |
| 2983 Hz | 4.93 dB | 9 | retenu-cr |

## Descripteurs

| Descripteur | Valeur |
|---|---|
| ρ (pic / hauteur) | 0.82 |
| Pente spectrale | -17.8 dB/octave (σ 4.2) |
| Harmoniques par note | 16 |
| Balance impair − pair | 1.0 dB |
| Enveloppe vraie | 910.0 Hz (σ 977.0) |

> La **pente spectrale n'est pas comparable entre corpus** : elle dépend du micro et du local. À lire à l'intérieur de ce corpus seulement.

> **Correction d'accord appliquée.** Cette banque est accordée +18075 cents par rapport aux hauteurs de ses noms de fichiers (rapport 1.0087). Sans correction, harmoniques() relève les pics dans un creux au-delà de ~1 300 Hz et l'enveloppe composite est fausse. Le recalage est fait par échantillon et les fréquences ci-dessus sont remises dans le référentiel réel.

## Confrontation aux corpus existants — À RÉDIGER

Comparer chaque formant ci-dessus à la fiche de l'instrument nu (`../../Instruments/FullSOL2020/` ou `Yan_Adds/`) et à ZSamples (`../ZSamples/`). **Ce corpus est une troisième source : il peut faire basculer en établis les dix formants classés « indéterminés » dans `../../Instruments/_REPLICATION.md`.** Route A = même point de grille ; route B = deux familles de méthodes indépendantes. La grille est au demi-ton : écrire « Ré7 (≈ 2 350 Hz) », pas une valeur au Hz près.

## Observations — À RÉDIGER

## Journal

- **2026-10-25** — Mesure initiale, corpus VO, 54 échantillons. 2 formant(s) retenu(s) sur 2 candidat(s). Nomenclature 1.0.
