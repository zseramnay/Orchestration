---
id: clarinette-contrebasse-en-sib
nom: Clarinette contrebasse en Sib
corpus: VO
n_echantillons: 45
rho_hauteur: 0.56
pente_db_octave: -11.3
harmoniques_par_note: 90
balance_impair_pair_db: 6.1
enveloppe_vraie_hz: 172.0
enveloppe_vraie_sd: 323.0
nomenclature: 1.0
accord_corrige: true
accord_rapport: 1.0029
classe: A_VALIDER_A_LA_MAIN
formants:
  - { hz: 4444, residu_db: 2.54, verdict: fragile-cr, notes: 11 }
  - { hz: 1111, residu_db: 2.36, verdict: rejeté, notes: 10 }
  - { hz: 1764, residu_db: 2.16, verdict: retenu-cr, notes: 12 }
  - { hz: 2643, residu_db: 1.34, verdict: rejeté, notes: 13 }
version_mesure: VO-2026-10-25
---

# Clarinette contrebasse en Sib — corpus VO

> **Méthode 8** (enveloppe composite). Méthode et critères dans `../00-METHODE.md`. Statut des formants dans `../_REPLICATION.md`.

## Formants

| Zone | Résidu | Notes | Verdict |
|---|---|---|---|
| 4444 Hz | 2.54 dB | 11 | fragile-cr |
| 1111 Hz | 2.36 dB | 10 | rejeté |
| 1764 Hz | 2.16 dB | 12 | retenu-cr |
| 2643 Hz | 1.34 dB | 13 | rejeté |

## Descripteurs

| Descripteur | Valeur |
|---|---|
| ρ (pic / hauteur) | 0.56 |
| Pente spectrale | -11.3 dB/octave (σ 4.3) |
| Harmoniques par note | 90 |
| Balance impair − pair | 6.1 dB |
| Enveloppe vraie | 172.0 Hz (σ 323.0) |

> La **pente spectrale n'est pas comparable entre corpus** : elle dépend du micro et du local. À lire à l'intérieur de ce corpus seulement.

> **Correction d'accord appliquée.** Cette banque est accordée +6025 cents par rapport aux hauteurs de ses noms de fichiers (rapport 1.0029). Sans correction, harmoniques() relève les pics dans un creux au-delà de ~1 300 Hz et l'enveloppe composite est fausse. Le recalage est fait par échantillon et les fréquences ci-dessus sont remises dans le référentiel réel.

## Confrontation aux corpus existants — À RÉDIGER

Comparer chaque formant ci-dessus à la fiche de l'instrument nu (`../NN-*.md`) et à ZSamples (`../../Instruments-ZSamples/`). **Ce corpus est une troisième source : il peut faire basculer en établis les dix formants classés « indéterminés » dans `../_REPLICATION.md`.** Route A = même point de grille ; route B = deux familles de méthodes indépendantes. La grille est au demi-ton : écrire « Ré7 (≈ 2 350 Hz) », pas une valeur au Hz près.

## Observations — À RÉDIGER

## Journal

- **2026-10-25** — Mesure initiale, corpus VO, 45 échantillons. 1 formant(s) retenu(s) sur 4 candidat(s). Nomenclature 1.0.
