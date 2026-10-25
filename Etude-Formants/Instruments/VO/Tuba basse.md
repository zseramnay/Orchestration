---
id: tuba-basse
nom: Tuba basse
corpus: VO
n_echantillons: 61
rho_hauteur: 0.68
pente_db_octave: -15.9
harmoniques_par_note: 90
balance_impair_pair_db: -4.9
enveloppe_vraie_hz: 237.0
enveloppe_vraie_sd: 131.0
nomenclature: 1.0
accord_corrige: true
accord_rapport: 1.0175
classe: A_VALIDER_A_LA_MAIN
formants:
  - { hz: 502, residu_db: 5.62, verdict: retenu-cr, notes: 9 }
  - { hz: 845, residu_db: 4.82, verdict: retenu-cr, notes: 11 }
  - { hz: 1341, residu_db: 3.84, verdict: retenu-cr, notes: 13 }
  - { hz: 282, residu_db: 3.29, verdict: retenu-cr, notes: 5 }
version_mesure: VO-2026-10-25
---

# Tuba basse — corpus VO

> **Méthode 8** (enveloppe composite). Méthode et critères dans `../00-METHODE.md`. Statut des formants dans `../_REPLICATION.md`.

## Formants

| Zone | Résidu | Notes | Verdict |
|---|---|---|---|
| 502 Hz | 5.62 dB | 9 | retenu-cr |
| 845 Hz | 4.82 dB | 11 | retenu-cr |
| 1341 Hz | 3.84 dB | 13 | retenu-cr |
| 282 Hz | 3.29 dB | 5 | retenu-cr |

## Descripteurs

| Descripteur | Valeur |
|---|---|
| ρ (pic / hauteur) | 0.68 |
| Pente spectrale | -15.9 dB/octave (σ 2.7) |
| Harmoniques par note | 90 |
| Balance impair − pair | -4.9 dB |
| Enveloppe vraie | 237.0 Hz (σ 131.0) |

> La **pente spectrale n'est pas comparable entre corpus** : elle dépend du micro et du local. À lire à l'intérieur de ce corpus seulement.

> **Correction d'accord appliquée.** Cette banque est accordée +36357 cents par rapport aux hauteurs de ses noms de fichiers (rapport 1.0175). Sans correction, harmoniques() relève les pics dans un creux au-delà de ~1 300 Hz et l'enveloppe composite est fausse. Le recalage est fait par échantillon et les fréquences ci-dessus sont remises dans le référentiel réel.

## Confrontation aux corpus existants — À RÉDIGER

Comparer chaque formant ci-dessus à la fiche de l'instrument nu (`../NN-*.md`) et à ZSamples (`../../Instruments-ZSamples/`). **Ce corpus est une troisième source : il peut faire basculer en établis les dix formants classés « indéterminés » dans `../_REPLICATION.md`.** Route A = même point de grille ; route B = deux familles de méthodes indépendantes. La grille est au demi-ton : écrire « Ré7 (≈ 2 350 Hz) », pas une valeur au Hz près.

## Observations — À RÉDIGER

## Journal

- **2026-10-25** — Mesure initiale, corpus VO, 61 échantillons. 4 formant(s) retenu(s) sur 4 candidat(s). Nomenclature 1.0.
