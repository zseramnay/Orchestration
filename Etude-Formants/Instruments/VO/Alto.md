---
id: alto
nom: Alto
corpus: VO
n_echantillons: 120
rho_hauteur: 0.74
pente_db_octave: -14.0
harmoniques_par_note: 20
balance_impair_pair_db: -3.3
enveloppe_vraie_hz: 517.0
enveloppe_vraie_sd: 640.0
nomenclature: 1.0
accord_corrige: true
accord_rapport: 1.0116
classe: A_VALIDER_A_LA_MAIN
formants:
  - { hz: 3770, residu_db: 4.87, verdict: retenu-cr, notes: 19 }
  - { hz: 1779, residu_db: 3.68, verdict: fragile-cr, notes: 7 }
  - { hz: 1412, residu_db: 3.29, verdict: fragile-cr, notes: 7 }
  - { hz: 5331, residu_db: 3.11, verdict: retenu-cr, notes: 23 }
version_mesure: VO-2026-10-25
---

# Alto — corpus VO

> **Méthode 8** (enveloppe composite). Méthode et critères dans `../00-METHODE.md`. Statut des formants dans `../_REPLICATION.md`.

## Formants

| Zone | Résidu | Notes | Verdict |
|---|---|---|---|
| 3770 Hz | 4.87 dB | 19 | retenu-cr |
| 1779 Hz | 3.68 dB | 7 | fragile-cr |
| 1412 Hz | 3.29 dB | 7 | fragile-cr |
| 5331 Hz | 3.11 dB | 23 | retenu-cr |

## Descripteurs

| Descripteur | Valeur |
|---|---|
| ρ (pic / hauteur) | 0.74 |
| Pente spectrale | -14.0 dB/octave (σ 2.8) |
| Harmoniques par note | 20 |
| Balance impair − pair | -3.3 dB |
| Enveloppe vraie | 517.0 Hz (σ 640.0) |

> La **pente spectrale n'est pas comparable entre corpus** : elle dépend du micro et du local. À lire à l'intérieur de ce corpus seulement.

> **Correction d'accord appliquée.** Cette banque est accordée +24100 cents par rapport aux hauteurs de ses noms de fichiers (rapport 1.0116). Sans correction, harmoniques() relève les pics dans un creux au-delà de ~1 300 Hz et l'enveloppe composite est fausse. Le recalage est fait par échantillon et les fréquences ci-dessus sont remises dans le référentiel réel.

## Confrontation aux corpus existants — À RÉDIGER

Comparer chaque formant ci-dessus à la fiche de l'instrument nu (`../NN-*.md`) et à ZSamples (`../../Instruments-ZSamples/`). **Ce corpus est une troisième source : il peut faire basculer en établis les dix formants classés « indéterminés » dans `../_REPLICATION.md`.** Route A = même point de grille ; route B = deux familles de méthodes indépendantes. La grille est au demi-ton : écrire « Ré7 (≈ 2 350 Hz) », pas une valeur au Hz près.

## Observations — À RÉDIGER

## Journal

- **2026-10-25** — Mesure initiale, corpus VO, 120 échantillons. 2 formant(s) retenu(s) sur 4 candidat(s). Nomenclature 1.0.
