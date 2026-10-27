---
id: petite-flute
nom: Petite flûte
nom_en: Piccolo
famille: bois
sous_famille: flûtes
ordre: 1
source_echantillons: Yan_Adds-Divers/Winds/Piccolo/ordinario
provenance: VSL
technique: ordinario
n_echantillons: 111
tessiture_mesuree: D5–C8
f0_min_hz: 587.3
f0_max_hz: 4186.0

classe: B
pertinence_formant: partielle  # revise 23/10 : zone 4432 Hz fragile-cr
rho_hauteur: 0.96

formants:
  - { hz: 4432, residu_db: 6.61, verdict: fragile-cr, notes_k_bas: 38, notes_k_haut: 21,
      ic_bas: [4.6, 6.4], commentaire: "region large, non une resonance etroite" }

candidats_rejetes:
  - { hz: 4432, residu_db: 6.61, verdict: rejete, raison: "échoue au test de robustesse de fenêtre" }

pente_db_octave: -21.2
pente_sd: 2.5
harmoniques_par_note: 7
harmoniques_min_max: [2, 18]
balance_impair_pair_db: -7.0
enveloppe_vraie_hz: 1685
enveloppe_vraie_sd: 1157

version_mesure: 2026-10-23-wav-v1
---

# Petite flûte

> ## Statut des formants — révision du 26/10/2026
>
> Confrontation des **cinq corpus**. Provenance de ce corpus : **VSL / ConTimbre / perso (Yan_Adds)**.
>
> Aucun formant retenu. Si le résultat négatif est répliqué, c'est un acquis — voir `_CORPUS-CROISES.md`.

## La notion de formant n'a pas de sens pour cet instrument

> **Méthode 8 appliquée — aucun formant retenu. Seuls les descripteurs de classe B
> sont valides.** Voir `00-METHODE.md`.

Ce n'est pas un échec de mesure, c'est un résultat. Quatre raisons concordantes.

**1. Toute la tessiture est hors du domaine de validité des méthodes à ordre fixe.**
Le f0 va de 587 à 4 186 Hz. Le lissage Orchidea à l'ordre 160 n'est valide qu'en dessous
de 276 Hz. La petite flûte est donc le seul instrument de l'orchestre dont **aucune** note
ne tombe dans le domaine de validité des `specenv`. C'est la cause directe de l'anomalie
constatée dans l'ancienne étude, où le piccolo ne convergeait avec rien : son « F1 » de
1 109 Hz était simplement la hauteur médiane de sa tessiture.

**2. Le spectre suit intégralement la hauteur.** ρ = **0,96**, la valeur la plus élevée de
tout le corpus des bois. Le pic de l'enveloppe vraie est une fonction de la note jouée,
pas une propriété de l'instrument. Son écart-type, σ = 1 157 Hz, le dit assez.

**3. Une note isolée ne porte presque aucune information d'enveloppe.** Médiane de
**7 harmoniques par note**, et **2 seulement** dans l'extrême aigu. Au-dessus de 3 500 Hz,
l'ordre cepstral adaptatif descend à 3 : l'enveloppe se réduit à une pente. Aucune méthode,
LPC, cepstre ou centroïde, ne peut extraire un formant de deux points de mesure.

**4. La zone d'énergie est une coupure, non une résonance.** L'agrégation fait bien
apparaître un maximum large entre 2 800 et 4 700 Hz, cohérent avec Backus (~3 400 Hz) et
en franche contradiction avec l'ancien pipeline (1 109 Hz). Mais ce maximum est trop large
pour survivre au redressement local, et le seul candidat détecté est rejeté. Il s'agit de
la coupure de rayonnement du tuyau, pas d'un formant.

## Candidat rejeté

| Zone | Résidu | k = 2–3 | k = 4–7 | k = 8–15 | Robustesse | Verdict |
|---|---|---|---|---|---|---|
| 4 432 Hz | +6,6 dB | +7,2 | +6,9 | — | **non** | **rejeté** |

Le test d'indépendance au rang est pourtant passé sur deux groupes. C'est la robustesse de
fenêtre qui échoue : à ±1/2 octave, plus aucun pic n'est détecté. Le signal est donc
réellement présent mais large, confirmant une coupure plutôt qu'une résonance.

## Descripteurs valides

| Descripteur | Valeur | Lecture |
|---|---|---|
| **Pente spectrale médiane** | **−21,2 dB/octave** (σ 2,5) | la plus raide de tous les bois |
| Harmoniques par note | 7 [2–18] | spectre très pauvre |
| Balance impair − pair | −7,0 dB | tuyau ouvert, harmoniques pairs dominants |
| Zone d'énergie dominante | 2 800–4 700 Hz | indicative, non validée comme formant |

La pente de −21,2 dB/octave est le descripteur juste pour cet instrument : elle quantifie
le fait que l'énergie est concentrée sur la fondamentale et les deux ou trois premiers
harmoniques, ce qui explique à la fois la pureté du son et sa capacité à percer l'orchestre
sans masquer les timbres voisins.

## Comparaison avec la littérature

| Source | Valeur annoncée | Notre mesure | Lecture |
|---|---|---|---|
| Backus (1969) | ~3 400 Hz | zone 2 800–4 700 Hz | **cohérent** |
| Ancien pipeline `specenv` | F1 = 1 109 Hz | — | artefact de tessiture |
| Ancien Fp | 1 320 Hz | — | artefact de fenêtre |

À noter : le document antérieur annonçait une plage analysée de 100–3 000 Hz, incompatible
avec un instrument dont les fondamentales montent à 4 186 Hz. Une partie du registre aigu
était hors champ. La plage actuelle va à 11 015 Hz.

Autre incohérence relevée dans l'ancien document : le texte affirmait « Fp non calculé dans
le corpus actuel » tout en affichant « Fp = 1 320 Hz ».

## Registres

Non défini dans `Data/registres.md`. À établir.

## Observations pour l'orchestration

Le répertoire regorge de doublures de piccolo à l'octave, qui ajoutent de la brillance.
L'ancienne matrice F1 ne pouvait pas les voir, et il faut comprendre pourquoi : elle ne
modélisait que la **fusion par proximité de résonances**, alors que le piccolo opère par un
mécanisme différent.

Hypothèse à valider par la méthode 12 : la fondamentale du piccolo coïncide avec
l'harmonique 2 de l'instrument doublé. La fusion est assurée par la coïncidence harmonique,
et l'effet est un **déplacement du centroïde du mélange vers l'aigu** — c'est exactement ce
qu'on nomme brillance. C'est un mécanisme d'**extension spectrale**, non de convergence.
La pente très raide mesurée ici le rend d'ailleurs plausible : le piccolo apporte peu
d'énergie en dehors de ses premiers partiels, donc il colore sans encombrer.

Cette distinction entre fusion et renforcement/brillance correspond à celle de Koechlin
entre homogénéité et éclat.

## Graphiques

Aucun à ce stade. Figures prévues dans `../Versions-html-and-docx/media/` :

| Fichier prévu | Contenu |
|---|---|
| `piccolo_composite.png` | enveloppe composite, montrant le maximum large 2,8–4,7 kHz |
| `piccolo_te_ordres.png` | ordre cepstral adaptatif en fonction de f0 — démonstration de la limite |
| `piccolo_nharm.png` | nombre d'harmoniques disponibles par note sur la tessiture |

La deuxième figure est pédagogiquement la plus utile : elle montre d'un coup d'œil pourquoi
la question formantique est mal posée pour cet instrument.

## Révision du 23/10/2026 — critère assoupli

Avec le critère assoupli (voir `00-METHODE.md`), la zone à **4 432 Hz** passe de *rejetée* à
**fragile-cr**. Le test d'indépendance au rang est franchement positif des deux côtés du rang
médian : +6,1 dB [IC +4,6] en dessous du rang 5, +7,1 dB [IC +6,4] au-dessus, sur 38 et
21 points. C'est le test de robustesse de fenêtre qui échoue toujours.

La lecture devient donc plus précise qu'un simple rejet : **il existe bien, autour de
4,4 kHz, une région de renforcement indépendante de la hauteur**, mais elle est trop large
pour être une résonance étroite. Cela confirme l'interprétation en coupure de rayonnement et
rejoint Backus (~3 400 Hz). La pertinence formantique passe de « non pertinente » à
« partielle » : l'instrument n'a pas de formant au sens strict, mais la région existe et est
mesurable.

## Journal

- **2026-10-23** — Point de départ de la reprise : c'est l'absence de convergence du piccolo
  dans les matrices F1 de l'étude antérieure, acoustiquement absurde au vu du répertoire,
  qui a déclenché la remise en cause de toute la chaîne de mesure.
- **2026-10-23** — Conclusion assumée : pas de formant. Le descripteur à développer est la
  fréquence de coupure de Benade.
- **2026-10-23 (soir)** — Critère assoupli appliqué (voir `00-METHODE.md`).
