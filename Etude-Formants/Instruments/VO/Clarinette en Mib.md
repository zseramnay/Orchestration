---
id: clarinette-en-mib
nom: Clarinette en Mib
corpus: VO
n_echantillons: 45
rho_hauteur: 0.67
pente_db_octave: -17.4
harmoniques_par_note: 16
balance_impair_pair_db: 3.9
enveloppe_vraie_hz: 1120.0
enveloppe_vraie_sd: 1065.0
nomenclature: 1.0
accord_corrige: true
accord_rapport: 1.0087
classe: A_VALIDER_A_LA_MAIN
formants:
  - { hz: 3759, residu_db: 3.81, verdict: retenu-cr, notes: 5 }
version_mesure: VO-2026-10-25
---

# Clarinette en Mib — corpus VO

> **Méthode 8** (enveloppe composite). Méthode et critères dans `../00-METHODE.md`. Statut des formants dans `../_REPLICATION.md`.

## Formants

| Zone | Résidu | Notes | Verdict |
|---|---|---|---|
| 3759 Hz | 3.81 dB | 5 | retenu-cr |

## Descripteurs

| Descripteur | Valeur |
|---|---|
| ρ (pic / hauteur) | 0.67 |
| Pente spectrale | -17.4 dB/octave (σ 4.3) |
| Harmoniques par note | 16 |
| Balance impair − pair | 3.9 dB |
| Enveloppe vraie | 1120.0 Hz (σ 1065.0) |

> La **pente spectrale n'est pas comparable entre corpus** : elle dépend du micro et du local. À lire à l'intérieur de ce corpus seulement.

> **Correction d'accord appliquée.** Cette banque est accordée +18075 cents par rapport aux hauteurs de ses noms de fichiers (rapport 1.0087). Sans correction, harmoniques() relève les pics dans un creux au-delà de ~1 300 Hz et l'enveloppe composite est fausse. Le recalage est fait par échantillon et les fréquences ci-dessus sont remises dans le référentiel réel.

## Confrontation aux corpus existants — À RÉDIGER

Comparer chaque formant ci-dessus à la fiche de l'instrument nu (`../NN-*.md`) et à ZSamples (`../../Instruments-ZSamples/`). **Ce corpus est une troisième source : il peut faire basculer en établis les dix formants classés « indéterminés » dans `../_REPLICATION.md`.** Route A = même point de grille ; route B = deux familles de méthodes indépendantes. La grille est au demi-ton : écrire « Ré7 (≈ 2 350 Hz) », pas une valeur au Hz près.

## Observations — À RÉDIGER

## Journal

- **2026-10-25** — Mesure initiale, corpus VO, 45 échantillons. 1 formant(s) retenu(s) sur 1 candidat(s). Nomenclature 1.0.
