---
id: contrebasse
nom: Contrebasse
corpus: VO
n_echantillons: 132
rho_hauteur: 0.78
pente_db_octave: -12.7
harmoniques_par_note: 72
balance_impair_pair_db: -3.1
enveloppe_vraie_hz: 215.0
enveloppe_vraie_sd: 405.0
nomenclature: 1.0
accord_corrige: true
accord_rapport: 0.9957
classe: A_VALIDER_A_LA_MAIN
formants:
  - { hz: 1473, residu_db: 4.94, verdict: retenu-cr, notes: 25 }
  - { hz: 2206, residu_db: 3.93, verdict: retenu-cr, notes: 28 }
  - { hz: 413, residu_db: 2.7, verdict: fragile-cr, notes: 6 }
  - { hz: 5559, residu_db: 1.93, verdict: rejeté, notes: 33 }
version_mesure: VO-2026-10-25
---

# Contrebasse — corpus VO

> **Méthode 8** (enveloppe composite). Méthode et critères dans `../00-METHODE.md`. Statut des formants dans `../_REPLICATION.md`.

## Formants

| Zone | Résidu | Notes | Verdict |
|---|---|---|---|
| 1473 Hz | 4.94 dB | 25 | retenu-cr |
| 2206 Hz | 3.93 dB | 28 | retenu-cr |
| 413 Hz | 2.7 dB | 6 | fragile-cr |
| 5559 Hz | 1.93 dB | 33 | rejeté |

## Descripteurs

| Descripteur | Valeur |
|---|---|
| ρ (pic / hauteur) | 0.78 |
| Pente spectrale | -12.7 dB/octave (σ 1.7) |
| Harmoniques par note | 72 |
| Balance impair − pair | -3.1 dB |
| Enveloppe vraie | 215.0 Hz (σ 405.0) |

> La **pente spectrale n'est pas comparable entre corpus** : elle dépend du micro et du local. À lire à l'intérieur de ce corpus seulement.

> **Correction d'accord appliquée.** Cette banque est accordée -8934 cents par rapport aux hauteurs de ses noms de fichiers (rapport 0.9957). Sans correction, harmoniques() relève les pics dans un creux au-delà de ~1 300 Hz et l'enveloppe composite est fausse. Le recalage est fait par échantillon et les fréquences ci-dessus sont remises dans le référentiel réel.

## Confrontation aux corpus existants — À RÉDIGER

Comparer chaque formant ci-dessus à la fiche de l'instrument nu (`../NN-*.md`) et à ZSamples (`../../Instruments-ZSamples/`). **Ce corpus est une troisième source : il peut faire basculer en établis les dix formants classés « indéterminés » dans `../_REPLICATION.md`.** Route A = même point de grille ; route B = deux familles de méthodes indépendantes. La grille est au demi-ton : écrire « Ré7 (≈ 2 350 Hz) », pas une valeur au Hz près.

## Observations — À RÉDIGER

## Journal

- **2026-10-25** — Mesure initiale, corpus VO, 132 échantillons. 2 formant(s) retenu(s) sur 4 candidat(s). Nomenclature 1.0.
