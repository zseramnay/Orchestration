---
id: trombone-basse
nom: Trombone basse
corpus: VO
n_echantillons: 57
rho_hauteur: 0.47
pente_db_octave: -16.2
harmoniques_par_note: 79
balance_impair_pair_db: -1.6
enveloppe_vraie_hz: 506.0
enveloppe_vraie_sd: 171.0
nomenclature: 1.0
accord_corrige: true
accord_rapport: 1.0116
classe: A_VALIDER_A_LA_MAIN
formants:
  - { hz: 1679, residu_db: 4.44, verdict: retenu-cr, notes: 15 }
  - { hz: 890, residu_db: 2.78, verdict: retenu-cr, notes: 8 }
  - { hz: 3358, residu_db: 1.97, verdict: fragile-cr, notes: 16 }
  - { hz: 2516, residu_db: 1.72, verdict: rejeté, notes: 13 }
version_mesure: VO-2026-10-25
---

# Trombone basse — corpus VO

> **Méthode 8** (enveloppe composite). Méthode et critères dans `../00-METHODE.md`. Statut des formants dans `../_REPLICATION.md`.

## Formants

| Zone | Résidu | Notes | Verdict |
|---|---|---|---|
| 1679 Hz | 4.44 dB | 15 | retenu-cr |
| 890 Hz | 2.78 dB | 8 | retenu-cr |
| 3358 Hz | 1.97 dB | 16 | fragile-cr |
| 2516 Hz | 1.72 dB | 13 | rejeté |

## Descripteurs

| Descripteur | Valeur |
|---|---|
| ρ (pic / hauteur) | 0.47 |
| Pente spectrale | -16.2 dB/octave (σ 5.4) |
| Harmoniques par note | 79 |
| Balance impair − pair | -1.6 dB |
| Enveloppe vraie | 506.0 Hz (σ 171.0) |

> La **pente spectrale n'est pas comparable entre corpus** : elle dépend du micro et du local. À lire à l'intérieur de ce corpus seulement.

> **Correction d'accord appliquée.** Cette banque est accordée +24100 cents par rapport aux hauteurs de ses noms de fichiers (rapport 1.0116). Sans correction, harmoniques() relève les pics dans un creux au-delà de ~1 300 Hz et l'enveloppe composite est fausse. Le recalage est fait par échantillon et les fréquences ci-dessus sont remises dans le référentiel réel.

## Confrontation aux corpus existants — À RÉDIGER

Comparer chaque formant ci-dessus à la fiche de l'instrument nu (`../NN-*.md`) et à ZSamples (`../../Instruments-ZSamples/`). **Ce corpus est une troisième source : il peut faire basculer en établis les dix formants classés « indéterminés » dans `../_REPLICATION.md`.** Route A = même point de grille ; route B = deux familles de méthodes indépendantes. La grille est au demi-ton : écrire « Ré7 (≈ 2 350 Hz) », pas une valeur au Hz près.

## Observations — À RÉDIGER

## Journal

- **2026-10-25** — Mesure initiale, corpus VO, 57 échantillons. 2 formant(s) retenu(s) sur 4 candidat(s). Nomenclature 1.0.
