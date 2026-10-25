---
id: violoncelle
nom: Violoncelle
corpus: VO
n_echantillons: 207
rho_hauteur: 0.83
pente_db_octave: -13.4
harmoniques_par_note: 22
balance_impair_pair_db: -2.9
enveloppe_vraie_hz: 619.0
enveloppe_vraie_sd: 923.0
nomenclature: 1.0
accord_corrige: true
accord_rapport: 0.9971
classe: A_VALIDER_A_LA_MAIN
formants:
  - { hz: 2627, residu_db: 5.62, verdict: retenu-cr, notes: 24 }
  - { hz: 828, residu_db: 3.64, verdict: retenu-cr, notes: 8 }
version_mesure: VO-2026-10-25
---

# Violoncelle — corpus VO

> **Méthode 8** (enveloppe composite). Méthode et critères dans `../00-METHODE.md`. Statut des formants dans `../_REPLICATION.md`.

## Formants

| Zone | Résidu | Notes | Verdict |
|---|---|---|---|
| 2627 Hz | 5.62 dB | 24 | retenu-cr |
| 828 Hz | 3.64 dB | 8 | retenu-cr |

## Descripteurs

| Descripteur | Valeur |
|---|---|
| ρ (pic / hauteur) | 0.83 |
| Pente spectrale | -13.4 dB/octave (σ 3.2) |
| Harmoniques par note | 22 |
| Balance impair − pair | -2.9 dB |
| Enveloppe vraie | 619.0 Hz (σ 923.0) |

> La **pente spectrale n'est pas comparable entre corpus** : elle dépend du micro et du local. À lire à l'intérieur de ce corpus seulement.

> **Correction d'accord appliquée.** Cette banque est accordée -6025 cents par rapport aux hauteurs de ses noms de fichiers (rapport 0.9971). Sans correction, harmoniques() relève les pics dans un creux au-delà de ~1 300 Hz et l'enveloppe composite est fausse. Le recalage est fait par échantillon et les fréquences ci-dessus sont remises dans le référentiel réel.

## Confrontation aux corpus existants — À RÉDIGER

Comparer chaque formant ci-dessus à la fiche de l'instrument nu (`../NN-*.md`) et à ZSamples (`../../Instruments-ZSamples/`). **Ce corpus est une troisième source : il peut faire basculer en établis les dix formants classés « indéterminés » dans `../_REPLICATION.md`.** Route A = même point de grille ; route B = deux familles de méthodes indépendantes. La grille est au demi-ton : écrire « Ré7 (≈ 2 350 Hz) », pas une valeur au Hz près.

## Observations — À RÉDIGER

## Journal

- **2026-10-25** — Mesure initiale, corpus VO, 207 échantillons. 2 formant(s) retenu(s) sur 2 candidat(s). Nomenclature 1.0.
