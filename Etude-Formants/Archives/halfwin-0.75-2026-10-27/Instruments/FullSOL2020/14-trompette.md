---
id: trompette
nom: Trompette en Ut
nom_en: Trumpet in C
famille: cuivres
sous_famille: cuivres principaux
ordre: 14
source_echantillons: FullSOL2020/Brass/Trumpet_C/ordinario
provenance: IRCAM
technique: ordinario
n_echantillons: 96
tessiture_mesuree: F#3–D6
f0_min_hz: 185.0
f0_max_hz: 1174.7

classe: A/B
pertinence_formant: partielle
rho_hauteur: 0.62

formants:
  - { hz: 2792, statut: "route A", residu_db: 2.96, etendue_residu: [1318, 5915], verdict: retenu,  notes: 14 }
  - { hz: 4695, statut: "route A", residu_db: 4.17, etendue_residu: [1318, 5915], verdict: retenu,  notes: 22 }
  - { hz: 2091, residu_db: 2.77, etendue_residu: [1318, 5915], verdict: fragile, notes: 9 }

candidats_rejetes:
  - { hz: 1479, residu_db: 2.51, verdict: fragile-cr, ic_bas: [1.4, 1.3],
      raison: "echec de robustesse ; reclasse fragile le 23/10",
      commentaire: "zone de Giesler 1200-1500 Hz" }

pente_db_octave: -16.8
pente_sd: 5.0
harmoniques_par_note: 23
harmoniques_min_max: [9, 59]
balance_impair_pair_db: -3.1
enveloppe_vraie_hz: 1034
enveloppe_vraie_sd: 498

version_mesure: 2026-10-23-wav-v1
---

# Trompette en Ut

> ## Statut des formants — révision du 26/10/2026
>
> Confrontation des **cinq corpus**. Provenance de ce corpus : **IRCAM (FullSOL2020)**.
>
> | Formant | Statut |
> |---|---|
> | **2792 Hz** | **route A** — SOL + VN au même point de grille, 2792 Hz · zone vue par SOL, VN · ⚠ **aucun pic localisé** : Q = 1.0 sur l'enveloppe vraie, c'est une région d'énergie et non une résonance |
> | **4695 Hz** | **route A** — VN + Xsamples au même point de grille, 4974 Hz · zone vue par SOL, VN, Xsamples. ⚠ La valeur de cette fiche, 4695 Hz, est à -1.0 demi-ton du couple : c'est **4974 Hz** qui est établi |
>
> `etendue_residu` dans l'en-tête **n'est pas une largeur de bande** : c'est l'étendue du résidu de la méthode 8, souvent une butée de plage. Les largeurs mesurées sont dans `_LARGEURS-DE-BANDE.md`.

## Méthode retenue, et pourquoi

> **Méthode 8 — deux formants retenus, un fragile. Pertinence partielle.**
> Voir `00-METHODE.md`.

La trompette est le cuivre le plus difficile à mesurer, et l'ancienne étude l'avait déjà
noté sans en tirer les conséquences. Trois éléments l'expliquent.

**La dépendance à la dynamique est maximale.** C'est le résultat le plus solide concernant
cet instrument, déjà établi dans les travaux antérieurs : le F1 variait de 523 Hz en *pp* à
1 134 Hz en *ff*, soit plus d'une octave. Notre **dispersion de pente, σ = 5,0 dB/octave**,
la plus forte des cuivres, confirme quantitativement ce comportement. Le spectre de la
trompette n'est pas une propriété fixe de l'instrument mais une fonction de la force de jeu.

**Le corpus est le plus petit des cuivres** : 96 échantillons, avec une médiane de seulement
23 harmoniques par note — le plus faible de la famille, la tessiture étant la plus aiguë.

**Les bandes à −3 dB sont très larges** (1 318–5 915 Hz pour les trois candidats), signe que
la méthode ne parvient pas à circonscrire les pics. Les résidus sont modestes, de +2,8 à
+4,2 dB, contre +5,4 pour le cor et +5,8 pour le tuba basse.

D'où la mention « partielle » : les deux formants retenus passent les tests, mais ils
décrivent des régions de renforcement larges plutôt que des résonances nettes.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes distinctes | Verdict |
|---|---|---|---|---|
| **4 695 Hz** | +4,2 dB | large | 22 | **retenu** |
| **2 792 Hz** | +3,0 dB | large | 14 | **retenu** |
| 2 091 Hz | +2,8 dB | large | 9 | fragile |

## Tests de rejet

| Zone | k = 2–3 | k = 4–7 | k = 8–15 | k ≥ 16 | Robustesse |
|---|---|---|---|---|---|
| 4 695 Hz | — | +4,9 | +4,6 | +4,2 | oui |
| 2 792 Hz | +2,4 | +3,0 | +2,6 | — | oui |
| 2 091 Hz | — | +2,2 | +2,3 | — | **non** |
| 1 479 Hz | +1,9 | +1,4 | — | — | non |

Le formant à 4 695 Hz est le plus régulier : résidu stable autour de +4,5 dB sur trois
groupes de rang. C'est aussi la zone où l'on attendrait le « brillant » caractéristique de
l'instrument, située au maximum de sensibilité de l'oreille.

## Descripteurs globaux

| Descripteur | Valeur | Lecture |
|---|---|---|
| Pente spectrale médiane | −16,8 dB/octave (**σ 5,0**) | la plus dispersée des cuivres |
| Harmoniques par note | 23 [9–59] | la plus faible des cuivres |
| Balance impair − pair | −3,1 dB | tuyau cylindrique mais à pavillon |
| Enveloppe vraie (méthode 6) | 1 034 Hz, σ 498 | dispersion élevée |
| Fraction de pics sur un harmonique | **0 %** | l'ordre adaptatif fonctionne parfaitement |

Le dernier chiffre mérite un mot : avec l'ordre cepstral adapté à f0, **aucun** des 96 pics
détectés ne tombe sur un harmonique. À comparer aux 33 % de la clarinette en Sib et aux 44 %
qu'obtenait le cepstre d'ordre fixe 80 de mai 2026 sur ce même instrument. C'est une
validation directe de la méthode 6.

## Comparaison avec la littérature

| Source | Valeur annoncée | Notre mesure | Accord |
|---|---|---|---|
| Backus (1969) | 1 000–1 500 Hz et ~2 000 Hz | 2 091 Hz (fragile) | **partiel** |
| Giesler (1985) | 1 200–1 500 Hz, *e-ähnlich* | 1 479 Hz **rejeté** | désaccord |
| Meyer (2009) | ~1 000 Hz, /a/–/e/ | env. vraie 1 034 Hz | **bon** |
| Ancien pipeline `specenv` | F1 = 786 Hz | — | hauteur médiane (rapport 1,79) |
| Ancien Fp | 1 046 Hz | env. vraie 1 034 Hz | **convergent** |

Le second formant de Backus (~2 000 Hz) correspond à notre candidat fragile à 2 091 Hz, et
Meyer (~1 000 Hz) à notre enveloppe vraie (1 034 Hz). En revanche la zone 1 200–1 500 Hz de
Giesler est rejetée par les tests.

L'ancienne étude affirmait que « les sources divergent fortement car chaque source capture un
registre différent » et que « Fp = 1 046 Hz est la mesure la plus représentative ». La
première partie est probablement juste, mais la cause réelle est la **dynamique** plus que le
registre. La seconde partie se trouve partiellement corroborée, puisque notre enveloppe vraie
tombe à 1 034 Hz — mais, ici encore, cela ne réhabilite pas le Fp comme mesure de résonance.

## Registres

`Data/registres.md` : grave F#3–C3 *(coquille probable dans la source : lire F#3–C4)* ·
médium C3–G4 · aigu G4–C5 · suraigu au-dessus.

Mesure par registre : non effectuée. **L'analyse par dynamique est ici la priorité absolue**,
plus encore que pour le cor : c'est le seul instrument de l'orchestre dont la littérature et
nos mesures s'accordent sur une dépendance dynamique d'une octave.

## Observations pour l'orchestration

À compléter. Éléments acquis :

- L'ancienne étude annonçait une convergence « excellente » trompette–hautbois à Δ = 43 Hz
  (786 contre 743 Hz). Les deux valeurs étaient des hauteurs médianes de tessiture dans des
  bins voisins. Cette convergence est **écartée** : les formants mesurés sont 2 792 et
  4 695 Hz pour la trompette, 4 183 Hz pour le hautbois.
- Le « super-cluster à 323 Hz » entre trompette, cor et alto dans le registre médium
  reposait sur des F1 identiques au bin près, qui suivaient les fondamentales. À reprendre.
- Une piste réelle en revanche : la zone 4 183–4 695 Hz rapproche la trompette du cor
  (4 183 Hz), du hautbois (4 183 Hz) et du cor anglais (4 183 Hz). À examiner par la
  méthode 12 : s'agit-il d'une région de renforcement partagée par les instruments à forte
  projection ?

## Graphiques

Aucun à ce stade. Figures prévues dans `../Versions-html-and-docx/media/` :

| Fichier prévu | Contenu |
|---|---|
| `trompette_composite.png` | enveloppe composite — montrer la largeur des bandes |
| `trompette_dynamiques.png` | **prioritaire** : enveloppes composites par dynamique, pp / mf / ff |
| `trompette_te_vs_cepstre80.png` | comparaison ordre adaptatif / ordre fixe 80 : 0 % contre 44 % de pics sur un harmonique |

La troisième figure est une bonne démonstration méthodologique de l'apport de l'ordre
adaptatif.

## Révision du 23/10/2026 — critère assoupli

La zone à **1 479 Hz** passe de *rejetée* à **fragile-cr** : +2,1 dB [IC +1,4] en dessous du
rang 5 et +1,7 dB [IC +1,3] au-dessus. Le test d'indépendance au rang est passé, la
robustesse de fenêtre échoue.

C'est la zone que **Giesler situe à 1 200–1 500 Hz**, et ce fichier la donnait en désaccord
avec nos mesures. Ce désaccord est à nuancer : le signal existe, il est indépendant du rang,
mais il n'est pas localisable par le test de fenêtre. Pour un instrument dont le spectre se
déplace d'une octave entre *pp* et *ff*, c'est exactement ce que l'on attend : la résonance
existe mais sa position dépend de la dynamique, donc la bande agrégée est étalée.

**L'analyse par dynamique devient donc le test décisif pour cet instrument**, et non une
simple amélioration. La zone à 2 091 Hz (Backus ~2 000) reste fragile pour le même motif.

## Journal

- **2026-10-23** — Deux formants retenus (2 792 et 4 695 Hz), bandes larges, résidus modestes.
- **2026-10-23** — Dispersion de pente la plus forte des cuivres (σ 5,0), confirmant la
  dépendance dynamique déjà relevée dans les travaux antérieurs.
- **2026-10-23** — Zéro pic sur un harmonique avec l'ordre adaptatif, contre 44 % avec
  l'ordre fixe 80 : validation de la méthode 6.
- **2026-10-23** — Convergence trompette–hautbois de l'ancienne étude écartée.
- **2026-10-23 (soir)** — Critère assoupli appliqué (voir `00-METHODE.md`).
- **2026-10-24** — Confrontation au corpus ZSamples : voir la bannière en tête et `_REPLICATION.md`.
