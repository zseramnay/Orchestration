---
id: flute-basse
nom: Flute basse
corpus: VO
n_echantillons: 39
rho_hauteur: 0.85
pente_db_octave: -13.8
harmoniques_par_note: 29
balance_impair_pair_db: -2.0
enveloppe_vraie_hz: 420.0
enveloppe_vraie_sd: 790.0
nomenclature: 1.0
accord_corrige: true
accord_rapport: 1.0175
classe: A_VALIDER_A_LA_MAIN
formants:
  - { hz: 2840, residu_db: 2.2, verdict: rejeté, notes: 5 }
  - { hz: 3579, residu_db: 1.67, verdict: fragile-cr, notes: 7 }
version_mesure: VO-2026-10-25
---

# Flute basse — corpus VO

> **Méthode 8** (enveloppe composite). Méthode et critères dans `../../Instruments/00-METHODE.md`. Statut des formants dans `../../Instruments/_REPLICATION.md`.

## Formants

| Zone | Résidu | Notes | Verdict |
|---|---|---|---|
| 2840 Hz | 2.2 dB | 5 | rejeté |
| 3579 Hz | 1.67 dB | 7 | fragile-cr |

## Descripteurs

| Descripteur | Valeur |
|---|---|
| ρ (pic / hauteur) | 0.85 |
| Pente spectrale | -13.8 dB/octave (σ 4.1) |
| Harmoniques par note | 29 |
| Balance impair − pair | -2.0 dB |
| Enveloppe vraie | 420.0 Hz (σ 790.0) |

> La **pente spectrale n'est pas comparable entre corpus** : elle dépend du micro et du local. À lire à l'intérieur de ce corpus seulement.

> **Correction d'accord appliquée.** Cette banque est accordée +36357 cents par rapport aux hauteurs de ses noms de fichiers (rapport 1.0175). Sans correction, harmoniques() relève les pics dans un creux au-delà de ~1 300 Hz et l'enveloppe composite est fausse. Le recalage est fait par échantillon et les fréquences ci-dessus sont remises dans le référentiel réel.

## Confrontation aux corpus existants — À RÉDIGER

Comparer chaque formant ci-dessus à la fiche de l'instrument nu (`../../Instruments/FullSOL2020/` ou `Yan_Adds/`) et à ZSamples (`../ZSamples/`). **Ce corpus est une troisième source : il peut faire basculer en établis les dix formants classés « indéterminés » dans `../../Instruments/_REPLICATION.md`.** Route A = même point de grille ; route B = deux familles de méthodes indépendantes. La grille est au demi-ton : écrire « Ré7 (≈ 2 350 Hz) », pas une valeur au Hz près.

## Observations — À RÉDIGER

## Journal

- **2026-10-25** — Mesure initiale, corpus VO, 39 échantillons. 0 formant(s) retenu(s) sur 2 candidat(s). Nomenclature 1.0.
