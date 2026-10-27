---
id: flute-contrebasse
nom: Flute contrebasse
corpus: VO
n_echantillons: 38
rho_hauteur: 0.77
pente_db_octave: -13.3
harmoniques_par_note: 39
balance_impair_pair_db: 0.0
enveloppe_vraie_hz: 366.0
enveloppe_vraie_sd: 285.0
nomenclature: 1.0
accord_corrige: true
accord_rapport: 1.0278
classe: A_VALIDER_A_LA_MAIN
formants:
  - { hz: 1915, residu_db: 2.97, verdict: retenu-cr, notes: 5 }
  - { hz: 2869, residu_db: 2.27, verdict: fragile-cr, notes: 10 }
version_mesure: VO-2026-10-25
---

# Flute contrebasse — corpus VO

> **Méthode 8** (enveloppe composite). Méthode et critères dans `../../Instruments/00-METHODE.md`. Statut des formants dans `../../Instruments/_REPLICATION.md`.

## Formants

| Zone | Résidu | Notes | Verdict |
|---|---|---|---|
| 1915 Hz | 2.97 dB | 5 | retenu-cr |
| 2869 Hz | 2.27 dB | 10 | fragile-cr |

## Descripteurs

| Descripteur | Valeur |
|---|---|
| ρ (pic / hauteur) | 0.77 |
| Pente spectrale | -13.3 dB/octave (σ 2.3) |
| Harmoniques par note | 39 |
| Balance impair − pair | 0.0 dB |
| Enveloppe vraie | 366.0 Hz (σ 285.0) |

> La **pente spectrale n'est pas comparable entre corpus** : elle dépend du micro et du local. À lire à l'intérieur de ce corpus seulement.

> **Correction d'accord appliquée.** Cette banque est accordée +57756 cents par rapport aux hauteurs de ses noms de fichiers (rapport 1.0278). Sans correction, harmoniques() relève les pics dans un creux au-delà de ~1 300 Hz et l'enveloppe composite est fausse. Le recalage est fait par échantillon et les fréquences ci-dessus sont remises dans le référentiel réel.

## Confrontation aux corpus existants — À RÉDIGER

Comparer chaque formant ci-dessus à la fiche de l'instrument nu (`../../Instruments/FullSOL2020/` ou `Yan_Adds/`) et à ZSamples (`../ZSamples/`). **Ce corpus est une troisième source : il peut faire basculer en établis les dix formants classés « indéterminés » dans `../../Instruments/_REPLICATION.md`.** Route A = même point de grille ; route B = deux familles de méthodes indépendantes. La grille est au demi-ton : écrire « Ré7 (≈ 2 350 Hz) », pas une valeur au Hz près.

## Observations — À RÉDIGER

## Journal

- **2026-10-25** — Mesure initiale, corpus VO, 38 échantillons. 1 formant(s) retenu(s) sur 2 candidat(s). Nomenclature 1.0.
