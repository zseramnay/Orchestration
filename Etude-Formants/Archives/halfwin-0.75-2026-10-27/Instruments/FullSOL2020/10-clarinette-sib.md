---
id: clarinette-sib
nom: Clarinette en Sib
nom_en: Bb clarinet
famille: bois
sous_famille: clarinettes
ordre: 10
source_echantillons: FullSOL2020/Winds/Clarinet_Bb/ordinario
provenance: IRCAM
technique: ordinario
n_echantillons: 126
tessiture_mesuree: D3–G6
f0_min_hz: 146.8
f0_max_hz: 1568.0

classe: B
pertinence_formant: partielle
rho_hauteur: 0.77

formants:
  - { hz: 1567, statut: "1 corpus", residu_db: 6.67, etendue_residu: [1108, 1759], verdict: retenu,  notes: 8 }
  - { hz: 3517, statut: "route A", residu_db: 4.92, etendue_residu: [2792, 5915], verdict: retenu,  notes: 22 }
  - { hz: 4974, residu_db: 4.03, etendue_residu: [2635, 5915], verdict: fragile, notes: 26 }

candidats_rejetes:
  - { hz: 1108, residu_db: 3.84, verdict: rejete, raison: "+15,4 dB aux rangs 2-3 mais +4,3 dB seulement aux rangs 4-7 ; échec de robustesse" }

pente_db_octave: -18.5
pente_sd: 5.4
harmoniques_par_note: 22
harmoniques_min_max: [7, 74]
balance_impair_pair_db: 0.7
enveloppe_vraie_hz: 678
enveloppe_vraie_sd: 531

version_mesure: 2026-10-23-wav-v1
---

# Clarinette en Sib

> ## Statut des formants — révision du 26/10/2026
>
> Confrontation des **cinq corpus**. Provenance de ce corpus : **IRCAM (FullSOL2020)**.
>
> | Formant | Statut |
> |---|---|
> | **1567 Hz** | présent dans le seul corpus SOL |
> | **3517 Hz** | **route A** — SOL + ZSamples au même point de grille, 3517 Hz · zone vue par SOL, Xsamples, ZSamples |
>
> `etendue_residu` dans l'en-tête **n'est pas une largeur de bande** : c'est l'étendue du résidu de la méthode 8, souvent une butée de plage. Les largeurs mesurées sont dans `_LARGEURS-DE-BANDE.md`.

## Méthode retenue, et pourquoi

> **Méthode 8 — deux formants retenus. Réconcilie deux sources que l'ancienne étude
> croyait contradictoires.** Voir `00-METHODE.md`.

La clarinette en Sib est de classe B — ρ = 0,77, son spectre suit largement la hauteur, et
l'ancienne étude avait raison sur ce point : « pas de formant fixe » au sens d'une résonance
de corps unique. Mais avec 126 échantillons et une médiane de 22 harmoniques par note,
l'agrégation détecte des zones de renforcement réelles, d'où la mention « partielle ».

Le résultat le plus important de cet instrument est une réconciliation. L'ancienne étude
présentait Backus et Meyer (1 500 Hz) comme contredisant Giesler (3 000–4 000 Hz), et
expliquait l'écart par le fait que chaque source « capturait un registre différent ». La
mesure montre autre chose : **les deux sources mesuraient deux formants réels et distincts
du même instrument.**

| Source | Valeur | Notre mesure | Accord |
|---|---|---|---|
| Backus (1969) et Meyer (2009) | 1 500 Hz | **1 567 Hz** | **excellent** |
| Giesler (1985) | 3 000–4 000 Hz | **3 517 Hz** | **excellent** |

Les deux passent les deux tests de rejet indépendamment. Il n'y avait pas de désaccord à
expliquer, seulement deux formants à distinguer.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes distinctes | Verdict |
|---|---|---|---|---|
| **1 567 Hz** | +6,7 dB | 1 108–1 759 Hz | 8 | **retenu** |
| **3 517 Hz** | +4,9 dB | 2 792–5 915 Hz | 22 | **retenu** |
| 4 974 Hz | +4,0 dB | 2 635–5 915 Hz | 26 | fragile |

## Tests de rejet

| Zone | k = 2–3 | k = 4–7 | k = 8–15 | k ≥ 16 | Robustesse |
|---|---|---|---|---|---|
| 1 567 Hz | +11,3 | +6,6 | +5,1 | — | oui |
| 3 517 Hz | +7,5 | +5,0 | +3,8 | +6,8 | oui |
| 4 974 Hz | — | +5,9 | +5,5 | +2,8 | **non** |
| 1 108 Hz | **+15,4** | +4,3 | — | — | **non** |

Le rejet de 1 108 Hz mérite attention. Son résidu de +15,4 dB aux rangs 2–3 est le plus
élevé de tout le corpus, mais il s'effondre à +4,3 dB aux rangs 4–7 et n'est robuste à aucun
changement de fenêtre. C'est la signature d'un effet de hauteur : les notes dont
l'harmonique 2 ou 3 tombe vers 1 100 Hz sont les notes du registre de gorge et du clairon,
et c'est leur propre comportement qui est mesuré, non une résonance de l'instrument. Noter
que 1 108 Hz se situe par ailleurs dans la bande à −3 dB du formant retenu à 1 567 Hz : il
peut aussi s'agir du flanc de celui-ci.

## Descripteurs globaux

| Descripteur | Valeur | Lecture |
|---|---|---|
| Pente spectrale médiane | −18,5 dB/octave (σ 5,4) | forte dispersion : registres contrastés |
| Harmoniques par note | 22 [7–74] | correct |
| **Balance impair − pair** | **+0,7 dB** | positive, mais la plus faible des clarinettes |
| Enveloppe vraie (méthode 6) | 678 Hz, σ 531 | σ élevé : classe B confirmée |

La balance de +0,7 dB seulement est un résultat à souligner : la clarinette en Sib est, des
quatre clarinettes, celle dont la dominance des harmoniques impairs est la **moins** marquée
(+0,7 contre +4,2, +4,5 et +5,1 dB). Mesurée sur toute la tessiture, la dominance impaire
s'atténue parce que le registre du clairon et le suraigu, où l'instrument joue sur le
troisième mode, y contribuent autant que le chalumeau. C'est une raison supplémentaire de
mener l'analyse par registre.

## Comparaison avec la littérature

| Source | Valeur annoncée | Notre mesure | Accord |
|---|---|---|---|
| Backus (1969) | 1 500 Hz | 1 567 Hz | **excellent** |
| Meyer (2009) | 1 500 Hz | 1 567 Hz | **excellent** |
| Giesler (1985) | 3 000–4 000 Hz, *i-ähnlich* | 3 517 Hz | **excellent** |
| Ancien pipeline `specenv` | F1 = 463 Hz, zone /o/ | — | hauteur médiane (rapport 0,88) |
| Ancien Fp | 1 412 Hz | 1 567 Hz | proche |
| Cepstre mai 2026 (ordre 80) | 1 442 Hz, 44 % sur un harmonique | 1 567 Hz | proche mais non valide |

L'ancien F1 de 463 Hz plaçait la clarinette dans la zone /o/, et c'est ce qui fondait sa
« convergence quasi-parfaite Δ = 11 Hz » avec le cor anglais (452 Hz). Les deux valeurs
étant des médianes de tessiture dans des bins voisins, cette convergence est écartée.

## Registres

`Data/registres.md` : chalumeau D3–D4 · gorge D#4–G#4 · clairon A4–A#5 · suraigu à partir
de B5.

Mesure par registre : **non effectuée, et prioritaire pour cet instrument** — c'est la
famille où la dispersion de pente et la variation de la balance impair/pair l'exigent le
plus.

L'ancienne analyse par registre donnait F1 = 215 Hz au chalumeau, 355 Hz à la gorge, 581 Hz
au clairon et 1 152 Hz au suraigu : une progression qui suit exactement les fondamentales
des registres. Ce n'étaient pas des formants.

## Observations pour l'orchestration

À compléter. Éléments acquis et pistes :

- **Flûte–clarinette à l'octave, flûte au-dessus**, l'une des doublures les plus employées
  du répertoire classique : la fondamentale de la flûte tombe sur l'harmonique 2 de la
  clarinette, précisément le partiel pair affaibli par le tuyau cylindrique fermé. La flûte
  comble le spectre lacunaire de la clarinette. La balance mesurée, −4,2 dB pour la flûte
  contre +0,7 dB pour la clarinette, donne la base quantitative du raisonnement. Ce mécanisme
  est invisible à toute comparaison F1 contre F1.
- **Flûte–clarinette à l'unisson** : deux instruments de classe B jouant la même note ont par
  construction le même pic suiveur. La convergence est réelle et triviale, et l'ancienne
  matrice ne pouvait pas la voir puisqu'elle comparait des médianes globales.
- Le formant à 3 517 Hz est proche du candidat de la clarinette en Mib (3 320 Hz) : parenté
  de famille probable, à confirmer quand le corpus de la Mib sera élargi.
- La convergence cor anglais–clarinette Sib de l'ancienne étude est écartée.

## Graphiques

Aucun à ce stade. Figures prévues dans `../Versions-html-and-docx/media/` :

| Fichier prévu | Contenu |
|---|---|
| `clarinette_sib_composite.png` | enveloppe composite avec les deux formants retenus |
| `clarinette_sib_registres.png` | enveloppes par registre — **prioritaire** |
| `clarinette_sib_litterature.png` | positions de Backus/Meyer et de Giesler reportées sur l'enveloppe composite — figure de réconciliation |
| `flute_clarinette_peigne.png` | partagée avec la flûte : peignes harmoniques à l'octave |

La figure de réconciliation a une valeur pédagogique forte : elle montre que deux sources
jugées contradictoires pendant toute l'étude antérieure décrivaient en fait deux formants
distincts du même instrument.

## Journal

- **2026-10-23** — Deux formants retenus : 1 567 Hz (Backus, Meyer) et 3 517 Hz (Giesler).
  Le désaccord supposé entre ces sources est levé : ce sont deux formants réels.
- **2026-10-23** — Candidat à 1 108 Hz rejeté malgré un résidu de +15,4 dB aux rangs 2–3 :
  effet de hauteur caractéristique, ou flanc du formant à 1 567 Hz.
- **2026-10-23** — Balance impair/pair de +0,7 dB seulement, la plus faible des clarinettes,
  parce que la mesure porte sur toute la tessiture et non sur le seul chalumeau.
