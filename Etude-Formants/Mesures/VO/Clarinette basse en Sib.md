---
id: clarinette-basse-en-sib
nom: Clarinette basse en Sib
corpus: VO
n_echantillons: 57
rho_hauteur: 0.86
pente_db_octave: -13.3
harmoniques_par_note: 39
balance_impair_pair_db: 7.3
enveloppe_vraie_hz: 301.0
enveloppe_vraie_sd: 657.0
nomenclature: 1.0
accord_corrige: true
accord_rapport: 1.0087
classe: A_VALIDER_A_LA_MAIN
formants:
  - { hz: 2235, residu_db: 5.07, verdict: retenu-cr, notes: 9 }
  - { hz: 3349, residu_db: 3.36, verdict: fragile-cr, notes: 15 }
  - { hz: 5316, residu_db: 2.53, verdict: fragile-cr, notes: 14 }
version_mesure: VO-2026-10-25
---

# Clarinette basse en Sib — corpus VO

> **Méthode 8** (enveloppe composite). Méthode et critères dans `../../Instruments/00-METHODE.md`. Statut des formants dans `../../Instruments/_REPLICATION.md`.

## Formants

| Zone | Résidu | Notes | Verdict |
|---|---|---|---|
| 2235 Hz | 5.07 dB | 9 | retenu-cr |
| 3349 Hz | 3.36 dB | 15 | fragile-cr |
| 5316 Hz | 2.53 dB | 14 | fragile-cr |

## Descripteurs

| Descripteur | Valeur |
|---|---|
| ρ (pic / hauteur) | 0.86 |
| Pente spectrale | -13.3 dB/octave (σ 5.4) |
| Harmoniques par note | 39 |
| Balance impair − pair | 7.3 dB |
| Enveloppe vraie | 301.0 Hz (σ 657.0) |

> La **pente spectrale n'est pas comparable entre corpus** : elle dépend du micro et du local. À lire à l'intérieur de ce corpus seulement.

> **Correction d'accord appliquée.** Cette banque est accordée +18075 cents par rapport aux hauteurs de ses noms de fichiers (rapport 1.0087). Sans correction, harmoniques() relève les pics dans un creux au-delà de ~1 300 Hz et l'enveloppe composite est fausse. Le recalage est fait par échantillon et les fréquences ci-dessus sont remises dans le référentiel réel.

## Confrontation aux corpus existants — À RÉDIGER

Comparer chaque formant ci-dessus à la fiche de l'instrument nu (`../../Instruments/FullSOL2020/` ou `Yan_Adds/`) et à ZSamples (`../ZSamples/`). **Ce corpus est une troisième source : il peut faire basculer en établis les dix formants classés « indéterminés » dans `../../Instruments/_REPLICATION.md`.** Route A = même point de grille ; route B = deux familles de méthodes indépendantes. La grille est au demi-ton : écrire « Ré7 (≈ 2 350 Hz) », pas une valeur au Hz près.

## Observations — À RÉDIGER

## Journal

- **2026-10-25** — Mesure initiale, corpus VO, 57 échantillons. 1 formant(s) retenu(s) sur 3 candidat(s). Nomenclature 1.0.
