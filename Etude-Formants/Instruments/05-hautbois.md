---
id: hautbois
nom: Hautbois
nom_en: Oboe
famille: bois
sous_famille: anches doubles
ordre: 5
source_echantillons: FullSOL2020/Winds/Oboe/ordinario
technique: ordinario
n_echantillons: 107
tessiture_mesuree: A#3–A6
f0_min_hz: 233.1
f0_max_hz: 1760.0

classe: A
pertinence_formant: pertinente
rho_hauteur: 0.26

formants:
  - { hz: 4183, residu_db: 6.21, bande_db3: [3517, 5915], verdict: retenu,  notes: 17 }
  - { hz: 5583, residu_db: 6.71, bande_db3: [3517, 5915], verdict: fragile, notes: 20 }

candidats_rejetes:
  - { hz: 2792, residu_db: 1.40, verdict: rejete, raison: "résidu négatif aux rangs 8-15" }

question_ouverte: "zone médiane 1100-1700 Hz annoncée par Meyer et Giesler : non retrouvée dans l'agrégation globale"

pente_db_octave: -18.9
pente_sd: 3.7
harmoniques_par_note: 17
harmoniques_min_max: [6, 47]
balance_impair_pair_db: -4.9
enveloppe_vraie_hz: 1453
enveloppe_vraie_sd: 701

version_mesure: 2026-10-23-wav-v1
---

# Hautbois

> ## ⚠ NON RÉPLIQUÉ — révision du 24/10/2026
>
> Le formant de 4 183 Hz n'est pas établi : ZSamples place son unique candidat à 5 583 Hz,
> soit 5 demi-tons plus haut, et en statut fragile. Statut : **1 corpus**.
>
> En revanche la **zone médiane de Meyer et Giesler reste introuvable dans les deux corpus**,
> ce qui renforce ce point de la fiche : le désaccord avec la littérature n'est pas un
> artefact d'un seul jeu d'échantillons. Voir `_REPLICATION.md`.

## Méthode retenue, et pourquoi

> **Méthode 8 — un formant retenu, un fragile. Résultat partiellement contre-intuitif.**
> Voir `00-METHODE.md`.

Le hautbois est l'instrument des bois dont le caractère de classe A est le mieux établi :
**ρ = 0,26**, la valeur la plus basse du corpus. Son spectre ne suit pas la hauteur, ce qui
est attendu d'une anche double sur tuyau conique. La mesure formantique est donc pleinement
pertinente, et l'enveloppe vraie note par note est également exploitable, avec des ordres
adaptatifs de 8 à 66.

Mais le résultat de l'agrégation est inattendu, et il faut le signaler plutôt que le lisser :
**les formants retenus se situent dans l'aigu, à 4 183 Hz, et non dans la zone 1 100–1 500 Hz
annoncée par Meyer et Giesler.** C'est le point méthodologique le plus sérieux encore ouvert
pour cette famille, et il est détaillé plus bas.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes distinctes | Verdict |
|---|---|---|---|---|
| **4 183 Hz** | +6,2 dB | 3 517–5 915 Hz | 17 | **retenu** |
| 5 583 Hz | +6,7 dB | 3 517–5 915 Hz | 20 | fragile |

## Tests de rejet

| Zone | k = 2–3 | k = 4–7 | k = 8–15 | k ≥ 16 | Robustesse |
|---|---|---|---|---|---|
| 4 183 Hz | +2,6 | +3,9 | +8,4 | +9,6 | oui |
| 5 583 Hz | — | +10,4 | +6,8 | +6,9 | **non** |
| 2 792 Hz | +5,7 | 0,0 | **−3,0** | — | non |

Le formant à 4 183 Hz est solide : résidu croissant avec le rang harmonique, positif sur
quatre groupes, robuste au changement de fenêtre. La zone à 5 583 Hz présente des résidus
plus forts encore mais échoue à la robustesse ; les deux se recouvrent largement
(bande −3 dB identique), et il s'agit probablement d'une seule région de renforcement entre
3 500 et 6 000 Hz dont la méthode ne sépare pas encore les composantes.

Le rejet de 2 792 Hz est instructif : +5,7 dB aux rangs 2–3 mais **−3,0 dB** aux rangs 8–15.
Un vrai formant ne change pas de signe selon le rang qui l'éclaire. C'était un effet de
hauteur.

## La question ouverte : la zone médiane

Meyer et Giesler situent tous deux le formant principal du hautbois autour de 1 100 Hz, et
Giesler ajoute un second vers 2 800 Hz. Backus donne ~1 300 Hz. Notre agrégation globale ne
retrouve **aucun** de ces trois.

Trois éléments à verser au dossier avant de trancher.

**Premier élément — un artefact écarté.** Un premier essai avec redressement linéaire en
log-fréquence produisait précisément des pics à 1 318 et 1 863 Hz, apparemment en bon accord
avec la littérature. Ils se sont évanouis dès le passage au redressement local. Sur un
spectre concave, un redressement linéaire fabrique une bosse médiane artificielle. **L'accord
avec la littérature était donc, dans cet essai, un artefact.** C'est une raison de prudence
dans les deux sens.

**Deuxième élément — la zone apparaît en sous-ensemble.** L'analyse de la moitié grave de la
tessiture fait ressortir 1 660 Hz à **+7,4 dB**, et les dynamiques douces donnent 1 759 Hz à
+8,4 dB. Le signal existe donc, mais il ne s'impose pas dans l'agrégation globale.

**Troisième élément — une différence de protocole.** Meyer et Giesler mesurent sur des plages
restreintes, alors que l'agrégation porte ici sur toute la tessiture. Un formant marqué dans
le grave et absent dans l'aigu se dilue dans une moyenne globale.

**À faire pour trancher :** mesure par registre avec la nouvelle méthode. C'est la tâche
prioritaire de cette famille.

## Descripteurs globaux

| Descripteur | Valeur | Lecture |
|---|---|---|
| Pente spectrale médiane | −18,9 dB/octave (σ 3,7) | raide pour une anche double |
| Harmoniques par note | 17 [6–47] | correct, mais le plus faible des anches doubles |
| Balance impair − pair | −4,9 dB | la plus négative des anches doubles |
| Enveloppe vraie (méthode 6) | 1 453 Hz, σ 701 | compatible avec Backus (~1 300) |

L'enveloppe vraie à 1 453 Hz est à noter : elle tombe, elle, dans la zone de la littérature.
Avec un σ de 701 Hz, elle reste trop dispersée pour conclure, mais elle va dans le sens de
l'existence d'une résonance médiane.

## Comparaison avec la littérature

| Source | Valeur annoncée | Notre mesure | Accord |
|---|---|---|---|
| Meyer (2009) | ~1 100 Hz, /a/ | non retrouvé globalement ; 1 660 Hz dans le grave | **à trancher** |
| Giesler (1985) | ~1 100 et ~2 800 Hz | non retrouvés ; 2 792 Hz rejeté | **désaccord** |
| Backus (1969) | ~1 300 Hz | enveloppe vraie 1 453 Hz | plausible |
| Ancien pipeline `specenv` | F1 = 743 Hz | — | hauteur médiane de tessiture (rapport 1,16) |
| Ancien Fp | 1 485 Hz | — | artefact de fenêtre, mais proche de notre enveloppe vraie |

Rappel : le F1 de 742,9 Hz publié était identique à celui de la flûte, au bin près. C'est
l'origine de la fausse « convergence quasi-parfaite Δ = 0 Hz ★ » entre les deux instruments.
Elle est écartée. L'expérience d'orchestration la contredisait d'ailleurs : à l'unisson, le
hautbois domine la flûte plus qu'il ne fusionne avec elle.

## Registres

`Data/registres.md` : grave A#3–G4 · médium G4–G5 · aigu G5–D6 · suraigu au-dessus.

Mesure par registre : **non effectuée, et prioritaire** — c'est elle qui tranchera la
question de la zone médiane.

## Observations pour l'orchestration

À compléter. Éléments acquis :

- La convergence flûte–hautbois de l'ancienne étude est écartée.
- La région de renforcement 3 500–6 000 Hz, bien attestée, est une piste pour expliquer la
  capacité de projection et le caractère « nasal et pénétrant » du hautbois : cette zone
  correspond au maximum de sensibilité de l'oreille, ce qui lui donne un poids perceptif
  supérieur à son poids énergétique. À quantifier par la méthode 12.

## Graphiques

Aucun à ce stade. Figures prévues dans `../Versions-html-and-docx/media/` :

| Fichier prévu | Contenu |
|---|---|
| `hautbois_composite.png` | enveloppe composite globale |
| `hautbois_registres.png` | enveloppes composites par registre — **figure clé** pour la question ouverte |
| `hautbois_artefact.png` | comparaison redressement linéaire / local, démonstration de l'artefact |

La troisième figure a une valeur pédagogique propre : elle montre comment un choix de
normalisation peut fabriquer un accord avec la littérature.

## Journal

- **2026-10-23** — ρ = 0,26, le plus faible du corpus : classe A confirmée.
- **2026-10-23** — Les pics médians à 1 318 et 1 863 Hz du premier essai identifiés comme
  artefact du redressement linéaire sur spectre concave.
- **2026-10-23** — Question ouverte consignée : la zone 1 100–1 700 Hz de Meyer et Giesler
  n'apparaît qu'en sous-ensembles (moitié grave, dynamiques douces). Mesure par registre
  requise avant toute publication sur ce point.
- **2026-10-24** — Confrontation au corpus ZSamples : voir la bannière en tête et `_REPLICATION.md`.
